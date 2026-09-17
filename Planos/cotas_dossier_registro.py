# -*- coding: utf-8 -*-
"""
Registro de evidencias JPG en ``public.cotas_dossier`` (NestingPro :5433).

Flujo acordado (Cotas antes del nest):
  1) Cliente/producto se toman del **VSM** (``foldertree.jobs``: client/product).
  2) Se resuelve ``…\\JOB\\DOSSIER FILES\\JPGS`` (VSM / env / picker).
  3) Se **publican** los JPG al share (anti-crash, reintentos) y se INSERTA/UPDATE
     en ``cotas_dossier`` con la ruta corporativa exacta.
  4) Fallos de red/DB/picker **no** abortan el acotado (fail-soft).

Activación:
  - Por defecto **activo** (COTAS_DOSSIER unset o 1/true/on).
  - Desactivar: ``COTAS_DOSSIER=0``.
  - Sin publicar al share: ``COTAS_DOSSIER_PUBLISH=0``.
  - Sin explorador si falta job: ``COTAS_DOSSIER_PICKER=0``.

Overrides opcionales:
  COTAS_DOSSIER_CLIENTE / COTAS_DOSSIER_PRODUCTO
  COTAS_VSM_JOB_ROOT  (carpeta JOB o DOSSIER FILES)
  COTAS_DOSSIER_PICKER_INITIAL  (initialdir del explorador)
  NESTING_DB_* / VSM_DB_*  (creds; defaults = ANS/VSM productivos)
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from functools import lru_cache
from typing import Any

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

_TRUE = ("1", "true", "yes", "on", "si", "sí")
_NOMBRE_DOSSIER_FILES = "DOSSIER FILES"
_NOMBRE_JPGS = "JPGS"
# UNC típico del share corporativo (picker / initialdir).
_UNC_CORPORATE_DEFAULT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM"
)


def dossier_habilitado() -> bool:
    raw = os.environ.get("COTAS_DOSSIER", "1").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    return raw in _TRUE or raw == ""


def dossier_publish_habilitado() -> bool:
    """Publicar JPG al share ``DOSSIER FILES\\JPGS`` (anti-crash, fail-soft)."""
    raw = os.environ.get("COTAS_DOSSIER_PUBLISH", "1").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    return raw in _TRUE or raw == ""


def dossier_picker_habilitado() -> bool:
    """Abrir explorador si no se resuelve la carpeta del JOB."""
    raw = os.environ.get("COTAS_DOSSIER_PICKER", "1").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    return raw in _TRUE or raw == ""


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name, default) or "").strip()


def _db_nesting() -> dict:
    return {
        "host": _env("NESTING_DB_HOST", "192.168.2.80"),
        "port": _env("NESTING_DB_PORT", "5433"),
        "database": _env("NESTING_DB_NAME", "nestingpro_db"),
        "user": _env("NESTING_DB_USER", "postgres"),
        "password": _env("NESTING_DB_PASSWORD", "nesting123"),
        "connect_timeout": int(_env("NESTING_DB_CONNECT_TIMEOUT", "5") or "5"),
    }


def _db_vsm() -> dict:
    return {
        "host": _env("VSM_DB_HOST", "192.168.2.80"),
        "port": _env("VSM_DB_PORT", "5437"),
        "database": _env("VSM_DB_NAME", "foldertree"),
        "user": _env("VSM_DB_USER", "user"),
        "password": _env("VSM_DB_PASSWORD", "password"),
        "connect_timeout": int(_env("VSM_DB_CONNECT_TIMEOUT", "5") or "5"),
    }


def _runtime_dir() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    d = os.path.join(here, ".runtime")
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def _contexto_path() -> str:
    return os.path.join(_runtime_dir(), "dossier_contexto.json")


# ---------------------------------------------------------------------------
# Contexto de sesión (job / cliente / producto)
# ---------------------------------------------------------------------------


def guardar_contexto_dossier(
    *,
    job: str,
    cliente: str = "",
    producto: str = "",
    job_root: str = "",
    dossier_files: str = "",
    dossier_jpgs: str = "",
) -> None:
    data = {
        "job": str(job or "").strip(),
        "cliente": str(cliente or "").strip(),
        "producto": str(producto or "").strip(),
        "job_root": str(job_root or "").strip(),
        "dossier_files": str(dossier_files or "").strip(),
        "dossier_jpgs": str(dossier_jpgs or "").strip(),
    }
    try:
        with open(_contexto_path(), "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
    except OSError as exc:
        print(f"AVISO dossier: no se pudo guardar contexto ({exc})")


def cargar_contexto_dossier() -> dict:
    try:
        with open(_contexto_path(), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def normalizar_ruta_dossier(ruta: str) -> dict:
    """
    Acepta JOB root, ``DOSSIER FILES`` o ``JPGS`` y devuelve rutas canónicas.

    Returns:
      ``{job_root, dossier_files, dossier_jpgs}`` (strings; pueden estar vacías).
    """
    out = {"job_root": "", "dossier_files": "", "dossier_jpgs": ""}
    try:
        raw = str(ruta or "").strip()
        if not raw:
            return out
        unc = raw.startswith("\\\\")
        parts = [x for x in raw.replace("/", "\\").split("\\") if x]
        low = [x.lower() for x in parts]
        if "jpgs" in low:
            i = low.index("jpgs")
            parts = parts[:i]
            low = low[:i]
        if "dossier files" in low:
            i = low.index("dossier files")
            parts = parts[:i]
        if not parts:
            return out
        if unc:
            job_root = "\\\\" + "\\".join(parts)
        else:
            job_root = "\\".join(parts)
        dossier_files = os.path.join(job_root, _NOMBRE_DOSSIER_FILES)
        dossier_jpgs = os.path.join(dossier_files, _NOMBRE_JPGS)
        out["job_root"] = job_root
        out["dossier_files"] = dossier_files
        out["dossier_jpgs"] = dossier_jpgs
    except Exception as exc:
        print(f"AVISO dossier: normalizar ruta ({exc})")
    return out


def _ruta_accesible(ruta: str) -> bool:
    if not ruta:
        return False
    try:
        return os.path.isdir(ruta)
    except Exception:
        return False


def asegurar_estructura_dossier_jpgs(job_root: str) -> str:
    """
    Crea ``<job>\\DOSSIER FILES\\JPGS`` si hace falta. Fail-soft → ``\"\"``.
    """
    try:
        info = normalizar_ruta_dossier(job_root)
        root = info.get("job_root") or ""
        jpgs = info.get("dossier_jpgs") or ""
        if not root or not jpgs:
            return ""
        os.makedirs(jpgs, exist_ok=True)
        return jpgs if _ruta_accesible(jpgs) else ""
    except Exception as exc:
        print(f"AVISO dossier: no se pudo crear JPGS ({exc})")
        return ""


def _pedir_ruta_dossier_interactivo(job: str) -> str:
    """
    Abre el explorador de carpetas para elegir dónde inyectar la estructura
    (JOB, DOSSIER FILES o JPGS). Fail-soft → ``\"\"``.
    """
    if not dossier_picker_habilitado():
        return ""
    initial = (
        _env("COTAS_DOSSIER_PICKER_INITIAL")
        or _env("COTAS_VSM_JOB_ROOT")
        or _UNC_CORPORATE_DEFAULT
    )
    try:
        if not _ruta_accesible(initial):
            # Subir un nivel si el UNC corporativo no monta completo.
            initial = os.path.dirname(initial.rstrip("\\/")) or initial
    except Exception:
        pass
    titulo = (
        f"Selecciona la carpeta DOSSIER FILES (o el JOB) para inyectar cotas"
        f" — job {job or '?'}"
    )
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass
        elegido = filedialog.askdirectory(
            title=titulo,
            initialdir=initial if _ruta_accesible(initial) else None,
            mustexist=True,
        )
        try:
            root.destroy()
        except Exception:
            pass
        return str(elegido or "").strip()
    except Exception as exc:
        print(f"AVISO dossier: no se pudo abrir el explorador ({exc})")
        return ""


def resolver_destino_dossier(
    job: str,
    job_root: str = "",
    *,
    pedir_si_falta: bool = True,
) -> dict:
    """
    Resuelve job_root + DOSSIER FILES\\JPGS.

    Orden: arg → env → VSM local_path → contexto previo → (opcional) picker.
    Nunca lanza: si falla, dict con rutas vacías.
    """
    vacio = {
        "job": str(job or "").strip(),
        "job_root": "",
        "dossier_files": "",
        "dossier_jpgs": "",
        "fuente_ruta": "ninguna",
    }
    try:
        job = str(job or "").strip()
        candidatos: list[tuple[str, str]] = []
        if job_root:
            candidatos.append((job_root, "arg"))
        env_root = _env("COTAS_VSM_JOB_ROOT")
        if env_root:
            candidatos.append((env_root, "env"))
        vsm = _lookup_vsm_job(job) if job else None
        if vsm and vsm.get("local_path"):
            candidatos.append(
                (
                    _local_path_vsm_a_windows(str(vsm["local_path"])),
                    "vsm",
                )
            )
        ctx = cargar_contexto_dossier()
        if ctx.get("job_root"):
            candidatos.append((str(ctx.get("job_root")), "contexto"))
        if ctx.get("dossier_files"):
            candidatos.append((str(ctx.get("dossier_files")), "contexto_dossier"))
        if ctx.get("dossier_jpgs"):
            candidatos.append((str(ctx.get("dossier_jpgs")), "contexto_jpgs"))

        elegido = ""
        fuente = "ninguna"
        for cand, fu in candidatos:
            info = normalizar_ruta_dossier(cand)
            root = info.get("job_root") or ""
            if not root:
                continue
            # Válido si existe JOB, DOSSIER FILES o se puede crear JPGS bajo un root accesible.
            if (
                _ruta_accesible(root)
                or _ruta_accesible(info.get("dossier_files") or "")
                or _ruta_accesible(info.get("dossier_jpgs") or "")
            ):
                elegido = root
                fuente = fu
                break
            # Root traducido (X:\) inexistente: seguir buscando.
        if not elegido and pedir_si_falta:
            print(
                f"[dossier] No se encontró carpeta del job {job!r}. "
                "Abriendo explorador de archivos…"
            )
            pick = _pedir_ruta_dossier_interactivo(job)
            if pick:
                info_p = normalizar_ruta_dossier(pick)
                elegido = info_p.get("job_root") or pick
                fuente = "picker"

        if not elegido:
            print(
                f"AVISO dossier: sin ruta de inyección para job={job!r}. "
                "Las cotas quedan solo locales hasta mapear DOSSIER FILES."
            )
            return vacio

        info = normalizar_ruta_dossier(elegido)
        jpgs = asegurar_estructura_dossier_jpgs(info.get("job_root") or elegido)
        if not jpgs:
            # Reintentar picker una vez si la ruta no es escribible.
            if pedir_si_falta and fuente != "picker":
                print(
                    "[dossier] Ruta resuelta no accesible para escribir. "
                    "Selecciona DOSSIER FILES manualmente…"
                )
                pick = _pedir_ruta_dossier_interactivo(job)
                if pick:
                    info = normalizar_ruta_dossier(pick)
                    jpgs = asegurar_estructura_dossier_jpgs(
                        info.get("job_root") or pick
                    )
                    fuente = "picker"
        if not jpgs:
            return {
                **vacio,
                "job_root": info.get("job_root") or "",
                "dossier_files": info.get("dossier_files") or "",
                "fuente_ruta": fuente,
            }
        return {
            "job": job,
            "job_root": info.get("job_root") or "",
            "dossier_files": info.get("dossier_files") or "",
            "dossier_jpgs": jpgs,
            "fuente_ruta": fuente,
        }
    except Exception as exc:
        print(f"AVISO dossier: resolver_destino ({exc})")
        return vacio


def producto_cliente_desde_job_root(ruta_job: str) -> tuple[str, str]:
    """
    De ``…/PRODUCTO/CLIENTE/JOB`` → (producto, cliente).

    Misma convención que ANS ``_producto_cliente_desde_job_root`` y que
    ``jobs.local_path`` del VSM (p. ej. ENCLOSURES NEMA 1 / GIGA / 9919-BOARD…).
    """
    try:
        p = os.path.normpath(str(ruta_job or ""))
        # Si apuntan a DOSSIER FILES o JPGS, subir al JOB root.
        parts = [x for x in p.replace("/", "\\").split("\\") if x]
        low = [x.lower() for x in parts]
        if "jpgs" in low:
            i = low.index("jpgs")
            parts = parts[:i]
            low = low[:i]
        if "dossier files" in low:
            i = low.index("dossier files")
            parts = parts[:i]
        if len(parts) < 3:
            return "", ""
        return str(parts[-3]), str(parts[-2])
    except Exception:
        return "", ""


def _local_path_vsm_a_windows(local_path: str) -> str:
    """
    ``/mnt/server_data/ARGA METALS…`` → ``X:\\ARGA METALS…`` si existe,
    o UNC configurable.
    """
    raw = str(local_path or "").strip().replace("/", "\\")
    if not raw:
        return ""
    # Quitar prefijo linux del mount VSM.
    markers = (
        "\\mnt\\server_data\\",
        "/mnt/server_data/",
    )
    raw_cmp = raw.replace("/", "\\")
    for pref in markers:
        pref_n = pref.replace("/", "\\")
        if raw_cmp.lower().startswith(pref_n.lower()):
            raw = raw_cmp[len(pref_n) :]
            break
    else:
        raw = raw_cmp
    raw = raw.lstrip("\\/")
    drive = _env("COTAS_VSM_DRIVE", "X:")
    if not drive.endswith(":"):
        drive = drive.rstrip("\\") + ":"
    win = drive + "\\" + raw.replace("/", "\\")
    return os.path.normpath(win)


@lru_cache(maxsize=32)
def _lookup_vsm_job(job: str) -> dict | None:
    job = str(job or "").strip()
    if not job:
        return None
    try:
        import psycopg2
        from psycopg2.extras import RealDictCursor
    except ImportError:
        print("AVISO dossier: falta psycopg2; no se consulta VSM.")
        return None
    try:
        with psycopg2.connect(**_db_vsm(), cursor_factory=RealDictCursor) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT job_number, client, product, local_path, status
                    FROM public.jobs
                    WHERE UPPER(TRIM(job_number)) = UPPER(TRIM(%s))
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (job,),
                )
                row = cur.fetchone()
                if row:
                    return dict(row)
                # Alias parcial: 9919-BOARD2_2 vs 9919-BOARD 2_2
                token = re.sub(r"[\s_]+", "", job.upper())
                cur.execute(
                    """
                    SELECT job_number, client, product, local_path, status
                    FROM public.jobs
                    WHERE REPLACE(REPLACE(UPPER(job_number), ' ', ''), '_', '')
                          = %s
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (token,),
                )
                row = cur.fetchone()
                return dict(row) if row else None
    except Exception as exc:
        print(f"AVISO dossier: VSM jobs no disponible ({exc})")
        return None


def _lookup_ans_erp(job: str) -> tuple[str, str]:
    """Respaldo: erp_jobs / diccionario_swo (mismos valores que sembró VSM)."""
    job = str(job or "").strip()
    if not job:
        return "", ""
    try:
        import psycopg2
        from psycopg2.extras import RealDictCursor
    except ImportError:
        return "", ""
    try:
        with psycopg2.connect(**_db_nesting(), cursor_factory=RealDictCursor) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT cliente, producto FROM public.erp_jobs
                    WHERE UPPER(TRIM(job_number)) = UPPER(TRIM(%s))
                    ORDER BY id DESC LIMIT 1
                    """,
                    (job,),
                )
                row = cur.fetchone()
                if row:
                    return (
                        str(row.get("cliente") or "").strip(),
                        str(row.get("producto") or "").strip(),
                    )
                cur.execute(
                    """
                    SELECT cliente, producto FROM public.diccionario_swo
                    WHERE UPPER(TRIM(job_numero)) = UPPER(TRIM(%s))
                    ORDER BY id DESC LIMIT 1
                    """,
                    (job,),
                )
                row = cur.fetchone()
                if row:
                    return (
                        str(row.get("cliente") or "").strip(),
                        str(row.get("producto") or "").strip(),
                    )
    except Exception:
        pass
    return "", ""


def resolver_cliente_producto(job: str, job_root: str = "") -> tuple[str, str, str]:
    """
    Devuelve (cliente, producto, fuente).

    Prioridad: env → contexto sesión → VSM jobs → parse ruta → ANS erp.
    """
    cli = _env("COTAS_DOSSIER_CLIENTE")
    prod = _env("COTAS_DOSSIER_PRODUCTO")
    if cli:
        return cli, prod, "env"

    ctx = cargar_contexto_dossier()
    if str(ctx.get("cliente") or "").strip():
        return (
            str(ctx.get("cliente") or "").strip(),
            str(ctx.get("producto") or "").strip(),
            "contexto",
        )

    vsm = _lookup_vsm_job(job)
    if vsm:
        return (
            str(vsm.get("client") or "").strip(),
            str(vsm.get("product") or "").strip(),
            "vsm",
        )

    root = (
        job_root
        or _env("COTAS_VSM_JOB_ROOT")
        or str(ctx.get("job_root") or "")
    )
    if root:
        prod2, cli2 = producto_cliente_desde_job_root(root)
        if cli2:
            return cli2, prod2, "ruta_job"

    cli3, prod3 = _lookup_ans_erp(job)
    if cli3:
        return cli3, prod3, "ans_erp"

    return "SIN_CLIENTE", "", "fallback"


def asegurar_contexto_desde_job(job: str, job_root: str = "") -> dict:
    """
    Resuelve cliente/producto + carpeta DOSSIER FILES\\JPGS y persiste sesión.

    Si no hay ruta accesible del JOB, abre el explorador (anti-crash / fail-soft).
    """
    job = str(job or "").strip()
    dest = resolver_destino_dossier(job, job_root, pedir_si_falta=True)
    root = dest.get("job_root") or job_root or _env("COTAS_VSM_JOB_ROOT")
    cli, prod, fuente = resolver_cliente_producto(job, root)
    # Si el picker dio ruta, preferir producto/cliente parseados de ella.
    if dest.get("fuente_ruta") == "picker" and root:
        prod2, cli2 = producto_cliente_desde_job_root(root)
        if cli2:
            cli, prod, fuente = cli2, prod2, "picker_ruta"
    guardar_contexto_dossier(
        job=job,
        cliente=cli,
        producto=prod,
        job_root=dest.get("job_root") or root or "",
        dossier_files=dest.get("dossier_files") or "",
        dossier_jpgs=dest.get("dossier_jpgs") or "",
    )
    print(
        f"[dossier] contexto: job={job!r} cliente={cli!r} "
        f"producto={prod!r} ({fuente})"
    )
    if dest.get("dossier_jpgs"):
        print(
            f"[dossier] inyección → {dest['dossier_jpgs']} "
            f"({dest.get('fuente_ruta')})"
        )
    else:
        print(
            "[dossier] AVISO: sin DOSSIER FILES\\JPGS mapeado; "
            "DB usará rutas locales hasta que se seleccione."
        )
    return {
        "job": job,
        "cliente": cli,
        "producto": prod,
        "job_root": dest.get("job_root") or root or "",
        "dossier_files": dest.get("dossier_files") or "",
        "dossier_jpgs": dest.get("dossier_jpgs") or "",
        "fuente": fuente,
        "fuente_ruta": dest.get("fuente_ruta") or "",
    }


# ---------------------------------------------------------------------------
# Tipología type / spoteos
# ---------------------------------------------------------------------------

_RE_HOLE = re.compile(r"HOLE\d{2}", re.I)
_RE_TYP = re.compile(r"(?:^|_)TYP(?:_|\d|$)|typ\+", re.I)
_RE_MEDIDA = re.compile(
    r"__(?P<med>LENGTH(?:_SIN_COTA)?|BROAD|WIDTH|THK|HEIGHT|LEG|OD|ID|"
    r"XCENTRO(?:_TYP)?|YCENTRO(?:_TYP)?|"
    r"HOLE\d{2}|FYP|TYP\+?)_",
    re.I,
)

_PROCESOS = (
    "Almacén",
    "Almacen",
    "Corte",
    "Doblado",
    "Maquinado",
    "Plasma",
    "Plasma Doblado",
    "Estañado",
    "Estanado",
    "Estañado Busbar",
    "Estanado Busbar",
    "SIN CLASIFICACION",
)

# Rutas lógicas nuevas (más específicas primero) + legacy Corte/*.
_CLASIF_RUTAS = (
    "Corte/Plasma y Laser/Corte metal",
    "Corte/Plasma y Laser/Corte Busbar",
    "Corte/Maquinado/Maquinados metal",
    "Corte/Maquinado/Corte Busbar",
    "Doblado/Metal",
    "Doblado/Busbar",
    "Estañado Busbar",
    # Legacy
    "Corte/Corte",
    "Corte/Doblado",
    "Corte/Estañado",
)

# Anidados legacy bajo Corte/ (migración / compat).
_CORTE_ANIDADOS = {
    "corte": "Corte",
    "doblado": "Doblado",
    "estañado": "Estañado",
    "estanado": "Estañado",
}


def clasificar_type_y_spoteos(
    nombre_archivo: str, cantidad_spoteos: int | None = None
) -> tuple[str, int]:
    """
    ``type`` + ``cantidad_spoteos`` (contrato dossier Cotas).

    - ``type`` = siempre ``TYP`` (cotas tipadas).
    - ``cantidad_spoteos`` = 1 si es cota normal / una sola aparición;
      N si hay N coincidencias TYP (misma pieza/foto se repite N veces
      en el dossier). El caller pasa N cuando lo conoce (qty ensamble,
      grupo de barrenos, etc.).
    """
    n = int(cantidad_spoteos) if cantidad_spoteos is not None else 1
    if n < 0:
        n = 0
    return "TYP", n


def proceso_desde_ruta(ruta: str) -> str:
    """
    Extrae clasificación desde JPGS/<proceso>/[anidados]/...

    Ejemplos (árbol acordado):
      .../JPGS/Corte/Plasma y Laser/Corte metal/<pieza>/...
        → Corte/Plasma y Laser/Corte metal
      .../JPGS/Corte/Maquinado/Corte Busbar/<pieza>/...
        → Corte/Maquinado/Corte Busbar
      .../JPGS/Doblado/Metal/<pieza>/... → Doblado/Metal
      .../JPGS/Estañado Busbar/... → Estañado Busbar
    Legacy: Corte/Corte, Corte/Doblado, Corte/Estañado.
    """
    try:
        parts = [p for p in os.path.normpath(ruta).replace("/", "\\").split("\\") if p]
        low = [p.casefold() for p in parts]

        def _es_archivo(cand: str) -> bool:
            return "." in cand and cand.lower().endswith(
                (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")
            )

        # Match longest known classification path anywhere in parts.
        for clasif in _CLASIF_RUTAS:
            segs = clasif.split("/")
            n = len(segs)
            for i in range(len(parts) - n + 1):
                if all(
                    parts[i + j].casefold() == segs[j].casefold()
                    for j in range(n)
                ):
                    # Canonical accents
                    if clasif.casefold() in (
                        "estañado busbar",
                        "estanado busbar",
                    ):
                        return "Estañado Busbar"
                    if clasif.casefold() == "corte/estañado":
                        return "Corte/Estañado"
                    return clasif

        canon = {}
        for proc in _PROCESOS:
            key = proc.casefold()
            if key in ("almacen", "almacén"):
                canon[key] = "Almacén"
            elif key in ("estanado", "estañado"):
                canon[key] = "Estañado"
            elif key in ("estanado busbar", "estañado busbar"):
                canon[key] = "Estañado Busbar"
            else:
                canon[key] = proc

        if "jpgs" in low:
            i = low.index("jpgs")
            if i + 1 < len(parts):
                cand = parts[i + 1]
                if _es_archivo(cand):
                    return ""
                hit = canon.get(cand.casefold())
                if hit == "Corte" and i + 2 < len(parts):
                    nest = parts[i + 2]
                    if not _es_archivo(nest):
                        # Plasma y Laser / Maquinado (árbol nuevo)
                        if nest.casefold() == "plasma y laser" and i + 3 < len(
                            parts
                        ):
                            leaf = parts[i + 3]
                            if leaf.casefold() in (
                                "corte metal",
                                "corte busbar",
                            ):
                                return f"Corte/Plasma y Laser/{leaf}"
                        if nest.casefold() == "maquinado" and i + 3 < len(
                            parts
                        ):
                            leaf = parts[i + 3]
                            if leaf.casefold() in (
                                "maquinados metal",
                                "corte busbar",
                            ):
                                return f"Corte/Maquinado/{leaf}"
                        nest_hit = _CORTE_ANIDADOS.get(nest.casefold())
                        if nest_hit:
                            return f"Corte/{nest_hit}"
                if hit == "Doblado" and i + 2 < len(parts):
                    nest = parts[i + 2]
                    if nest.casefold() in ("metal", "busbar") and not _es_archivo(
                        nest
                    ):
                        return f"Doblado/{nest}"
                if hit:
                    return hit
                return ""
        # Fallback: primer proceso conocido en la ruta.
        for idx, part in enumerate(parts):
            hit = canon.get(part.casefold())
            if hit == "Corte" and idx + 1 < len(parts):
                nest = parts[idx + 1]
                if nest.casefold() == "plasma y laser" and idx + 2 < len(parts):
                    leaf = parts[idx + 2]
                    if leaf.casefold() in ("corte metal", "corte busbar"):
                        return f"Corte/Plasma y Laser/{leaf}"
                if nest.casefold() == "maquinado" and idx + 2 < len(parts):
                    leaf = parts[idx + 2]
                    if leaf.casefold() in (
                        "maquinados metal",
                        "corte busbar",
                    ):
                        return f"Corte/Maquinado/{leaf}"
                nest_hit = _CORTE_ANIDADOS.get(nest.casefold())
                if nest_hit and not _es_archivo(nest):
                    return f"Corte/{nest_hit}"
            if hit == "Doblado" and idx + 1 < len(parts):
                nest = parts[idx + 1]
                if nest.casefold() in ("metal", "busbar") and not _es_archivo(
                    nest
                ):
                    return f"Doblado/{nest}"
            if hit and hit != "Corte":
                return hit
            if hit == "Corte":
                return "Corte"
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# DB NestingPro
# ---------------------------------------------------------------------------


def asegurar_tabla_cotas_dossier() -> bool:
    """CREATE IF NOT EXISTS (idempotente; espejo ANS)."""
    try:
        import psycopg2
    except ImportError:
        print("AVISO dossier: falta psycopg2.")
        return False
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS public.cotas_dossier (
                        id                  SERIAL PRIMARY KEY,
                        cliente             TEXT NOT NULL,
                        producto            TEXT NOT NULL DEFAULT '',
                        job                 TEXT NOT NULL,
                        type                TEXT NOT NULL DEFAULT '',
                        cantidad_spoteos    INTEGER NOT NULL DEFAULT 1
                            CHECK (cantidad_spoteos >= 0),
                        nombre_archivo      TEXT NOT NULL DEFAULT '',
                        ruta                TEXT NOT NULL DEFAULT '',
                        clasificacion       TEXT NOT NULL DEFAULT '',
                        seleccionadas       TEXT NOT NULL DEFAULT 'no',
                        created_at          TIMESTAMP NOT NULL DEFAULT NOW()
                    )
                    """
                )
                cur.execute(
                    """
                    ALTER TABLE public.cotas_dossier
                        ADD COLUMN IF NOT EXISTS clasificacion
                            TEXT NOT NULL DEFAULT ''
                    """
                )
                # Calidad cobre: primeras 2 XCENTRO + 2 YCENTRO desde (0,0).
                cur.execute(
                    """
                    ALTER TABLE public.cotas_dossier
                        ADD COLUMN IF NOT EXISTS seleccionadas
                            TEXT NOT NULL DEFAULT 'no'
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_job
                    ON public.cotas_dossier (job)
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_cliente_producto
                    ON public.cotas_dossier (cliente, producto)
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_type
                    ON public.cotas_dossier (type)
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_job_type
                    ON public.cotas_dossier (job, type)
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_clasificacion
                    ON public.cotas_dossier (clasificacion)
                    """
                )
                cur.execute(
                    """
                    CREATE INDEX IF NOT EXISTS idx_cotas_dossier_job_clasificacion
                    ON public.cotas_dossier (job, clasificacion)
                    """
                )
            conn.commit()
        return True
    except Exception as exc:
        print(f"AVISO dossier: no se pudo asegurar tabla ({exc})")
        return False


def _norm_seleccionadas(valor: str | None) -> str:
    v = str(valor or "").strip().casefold()
    if v in ("si", "sí", "yes", "true", "1", "s"):
        return "si"
    return "no"


def insertar_evidencia(
    *,
    cliente: str,
    producto: str,
    job: str,
    type_: str,
    cantidad_spoteos: int,
    nombre_archivo: str,
    ruta: str,
    clasificacion: str = "",
    seleccionadas: str = "no",
) -> int | None:
    """INSERT o UPDATE (misma job+nombre) y devuelve id, o None si falla."""
    try:
        import psycopg2
    except ImportError:
        return None
    clase = str(clasificacion or "").strip()
    if not clase:
        clase = proceso_desde_ruta(ruta) or ""
    sel = _norm_seleccionadas(seleccionadas)
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT id, ruta FROM public.cotas_dossier
                    WHERE job = %s AND nombre_archivo = %s
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (str(job or ""), str(nombre_archivo or "")),
                )
                prev = cur.fetchone()
                if prev:
                    eid = int(prev[0])
                    cur.execute(
                        """
                        UPDATE public.cotas_dossier
                        SET cliente = %s,
                            producto = %s,
                            type = %s,
                            cantidad_spoteos = %s,
                            ruta = %s,
                            clasificacion = %s,
                            seleccionadas = %s
                        WHERE id = %s
                        """,
                        (
                            str(cliente or "SIN_CLIENTE"),
                            str(producto or ""),
                            str(type_ or "TYP"),
                            max(0, int(cantidad_spoteos)),
                            str(ruta or ""),
                            clase,
                            sel,
                            eid,
                        ),
                    )
                    conn.commit()
                    return eid
                cur.execute(
                    """
                    INSERT INTO public.cotas_dossier
                        (cliente, producto, job, type, cantidad_spoteos,
                         nombre_archivo, ruta, clasificacion, seleccionadas)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                    """,
                    (
                        str(cliente or "SIN_CLIENTE"),
                        str(producto or ""),
                        str(job or ""),
                        str(type_ or "TYP"),
                        max(0, int(cantidad_spoteos)),
                        str(nombre_archivo or ""),
                        str(ruta or ""),
                        clase,
                        sel,
                    ),
                )
                row = cur.fetchone()
            conn.commit()
            return int(row[0]) if row else None
    except Exception as exc:
        print(f"AVISO dossier: INSERT/UPDATE falló ({exc})")
        return None


def recalcular_seleccionadas_en_carpeta(
    carpeta: str,
    *,
    job: str | None = None,
) -> int:
    """
    Recorre JPG de ``carpeta`` y actualiza ``seleccionadas`` (sí/no) por pieza.

    Pensado para el final de cada exportación GIGA: con el set completo de
    capturas por item cobre se marcan las 2 primeras XCENTRO y 2 YCENTRO.
    """
    if not dossier_habilitado():
        return 0
    root = os.path.abspath(str(carpeta or ""))
    if not os.path.isdir(root):
        return 0
    try:
        from cotas_seleccionadas_cobre import mapa_seleccionadas
    except Exception as exc:
        print(f"AVISO dossier: seleccionadas module ({exc})")
        return 0
    try:
        import psycopg2
    except ImportError:
        return 0

    por_dir: dict[str, list[str]] = {}
    try:
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                por_dir.setdefault(dirpath, []).append(fn)
    except Exception as exc:
        print(f"AVISO dossier: walk seleccionadas ({exc})")
        return 0

    ctx_job = str(job or cargar_contexto_dossier().get("job") or "").strip()
    n = 0
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                for dirpath, names in por_dir.items():
                    marks = mapa_seleccionadas(names)
                    for fn, flag in marks.items():
                        sel = _norm_seleccionadas(flag)
                        if ctx_job:
                            cur.execute(
                                """
                                UPDATE public.cotas_dossier
                                SET seleccionadas = %s
                                WHERE job = %s AND nombre_archivo = %s
                                """,
                                (sel, ctx_job, fn),
                            )
                        else:
                            cur.execute(
                                """
                                UPDATE public.cotas_dossier
                                SET seleccionadas = %s
                                WHERE nombre_archivo = %s
                                """,
                                (sel, fn),
                            )
                        n += int(cur.rowcount or 0)
            conn.commit()
    except Exception as exc:
        print(f"AVISO dossier: recalcular seleccionadas ({exc})")
        return 0
    if n:
        print(f"[dossier] seleccionadas actualizadas: {n} filas")
    return n


def _ya_registrada(job: str, ruta: str, nombre: str) -> bool:
    """True si ya existe misma ruta; False si solo hay otra ruta (se actualizará)."""
    try:
        import psycopg2
    except ImportError:
        return False
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT ruta FROM public.cotas_dossier
                    WHERE job = %s
                      AND (ruta = %s OR nombre_archivo = %s)
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (job, ruta, nombre),
                )
                row = cur.fetchone()
                if not row:
                    return False
                prev = str(row[0] or "")
                # Misma ruta corporativa/local → skip. Ruta distinta → upsert.
                return os.path.normcase(prev) == os.path.normcase(str(ruta or ""))
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Publicación anti-crash → DOSSIER FILES\\JPGS
# ---------------------------------------------------------------------------


def _copy_retry(src: str, dst: str, intentos: int = 3) -> bool:
    """Copia con reintentos (share/red). Nunca lanza."""
    last = None
    for i in range(max(1, intentos)):
        try:
            os.makedirs(os.path.dirname(dst) or dst, exist_ok=True)
            shutil.copy2(src, dst)
            if os.path.isfile(dst) and os.path.getsize(dst) > 0:
                return True
        except Exception as exc:
            last = exc
            time.sleep(0.4 * (i + 1))
    if last:
        print(f"AVISO dossier: copy falló {os.path.basename(src)} ({last})")
    return False


def _relativo_bajo_jpgs_local(ruta_jpg: str) -> str:
    """
    Relativo a conservar bajo ``JPGS\\``.

    Busca anclas ``PIEZAS_ACOTADAS``, ``BOARD``, ``JPGS`` o carpetas de proceso.
    """
    try:
        ruta = os.path.abspath(str(ruta_jpg or ""))
        parts = [p for p in ruta.replace("/", "\\").split("\\") if p]
        low = [p.casefold() for p in parts]
        for ancla in ("piezas_acotadas", "board", "jpgs"):
            if ancla in low:
                i = low.index(ancla)
                rel_parts = parts[i + 1 :]
                if rel_parts:
                    return "\\".join(rel_parts)
        # Fallback: proceso conocido hacia el final.
        procesos = {
            "almacén",
            "almacen",
            "corte",
            "doblado",
            "maquinado",
            "plasma",
            "plasma doblado",
            "sin clasificacion",
            "estañado",
            "estanado",
        }
        for idx, p in enumerate(low):
            if p in procesos:
                return "\\".join(parts[idx:])
        return os.path.basename(ruta)
    except Exception:
        return os.path.basename(str(ruta_jpg or "captura.jpg"))


def publicar_jpg_a_dossier(
    ruta_jpg: str,
    *,
    dossier_jpgs: str | None = None,
) -> str:
    """
    Copia un JPG al share ``DOSSIER FILES\\JPGS\\…``.

    Devuelve la ruta destino (UNC/local) o ``\"\"`` si falla (fail-soft).
    """
    if not dossier_publish_habilitado():
        return ""
    try:
        src = os.path.abspath(str(ruta_jpg or ""))
        if not src or not os.path.isfile(src):
            return ""
        ctx = cargar_contexto_dossier()
        dest_root = (
            dossier_jpgs
            or str(ctx.get("dossier_jpgs") or "").strip()
        )
        if not dest_root:
            dest = resolver_destino_dossier(
                str(ctx.get("job") or ""),
                str(ctx.get("job_root") or ""),
                pedir_si_falta=False,
            )
            dest_root = dest.get("dossier_jpgs") or ""
        if not dest_root:
            return ""
        rel = _relativo_bajo_jpgs_local(src)
        # Sin subcarpeta de proceso aún (JPG suelto pre-reorg): diferir publicación.
        if "\\" not in rel and "/" not in rel:
            return ""
        dst = os.path.join(dest_root, rel)
        if os.path.normcase(src) == os.path.normcase(os.path.abspath(dst)):
            return dst
        if _copy_retry(src, dst):
            return dst
    except Exception as exc:
        print(f"AVISO dossier: publicar_jpg ({exc})")
    return ""


def publicar_arbol_jpgs(
    carpeta_local: str,
    *,
    dossier_jpgs: str | None = None,
) -> tuple[str, int]:
    """
    Publica todo el árbol local de JPG al DOSSIER.

    Returns:
      ``(dossier_jpgs, n_copiados)``. Fail-soft.
    """
    if not dossier_habilitado() or not dossier_publish_habilitado():
        return "", 0
    n = 0
    try:
        root = os.path.abspath(str(carpeta_local or ""))
        if not os.path.isdir(root):
            return "", 0
        ctx = cargar_contexto_dossier()
        dest_root = dossier_jpgs or str(ctx.get("dossier_jpgs") or "").strip()
        if not dest_root:
            dest = resolver_destino_dossier(
                str(ctx.get("job") or ""),
                str(ctx.get("job_root") or ""),
                pedir_si_falta=True,
            )
            dest_root = dest.get("dossier_jpgs") or ""
            if dest_root:
                guardar_contexto_dossier(
                    job=str(ctx.get("job") or dest.get("job") or ""),
                    cliente=str(ctx.get("cliente") or ""),
                    producto=str(ctx.get("producto") or ""),
                    job_root=dest.get("job_root") or "",
                    dossier_files=dest.get("dossier_files") or "",
                    dossier_jpgs=dest_root,
                )
        if not dest_root:
            print("AVISO dossier: sin destino para publicar árbol JPG.")
            return "", 0
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                src = os.path.join(dirpath, fn)
                try:
                    if publicar_jpg_a_dossier(src, dossier_jpgs=dest_root):
                        n += 1
                except Exception as exc_f:
                    print(f"AVISO dossier: skip {fn} ({exc_f})")
                    continue
        print(f"[dossier] publicados {n} JPG → {dest_root}")
        return dest_root, n
    except Exception as exc:
        print(f"AVISO dossier: publicar_arbol ({exc})")
        return "", n


def publicar_y_sincronizar_dossier(
    carpeta_local: str,
    *,
    job: str | None = None,
) -> int:
    """
    Anti-crash: copia a ``DOSSIER FILES\\JPGS`` y registra en DB con esa ruta.

    Nunca aborta el acotado. Devuelve cantidad de evidencias tocadas.
    """
    if not dossier_habilitado():
        return 0
    try:
        dest_root, _n_pub = publicar_arbol_jpgs(carpeta_local)
        sync_root = dest_root if dest_root and _ruta_accesible(dest_root) else carpeta_local
        n = sincronizar_carpeta_jpgs(sync_root, job=job, solo_nuevos=True)
        # Recalcular otra vez sobre destino final (set completo cobre).
        try:
            recalcular_seleccionadas_en_carpeta(sync_root, job=job)
        except Exception:
            pass
        return n
    except Exception as exc:
        print(f"AVISO dossier: publicar_y_sincronizar ({exc})")
        try:
            return sincronizar_carpeta_jpgs(carpeta_local, job=job)
        except Exception:
            return 0


# ---------------------------------------------------------------------------
# API pública para generadores
# ---------------------------------------------------------------------------


def registrar_jpg(
    ruta_jpg: str,
    *,
    job: str | None = None,
    cantidad_spoteos: int | None = None,
    type_override: str | None = None,
    forzar: bool = False,
) -> bool:
    """
    Publica (si hay DOSSIER mapeado) y registra en DB.

    Fail-soft: nunca lanza hacia el caller. La ``ruta`` en DB es la del
    share corporativo cuando la copia funciona.
    """
    if not dossier_habilitado() and not forzar:
        return False
    try:
        ruta_local = os.path.abspath(str(ruta_jpg or ""))
        if not ruta_local or not os.path.isfile(ruta_local):
            return False
        nombre = os.path.basename(ruta_local)
        ctx = cargar_contexto_dossier()
        job_s = str(job or ctx.get("job") or "").strip()
        if not job_s:
            if "__" in nombre:
                job_s = nombre.split("__", 1)[0].strip()
        if not job_s:
            print(f"AVISO dossier: sin job para {nombre}")
            return False

        # Preferir ruta corporativa en DB.
        ruta_db = ruta_local
        publicada = publicar_jpg_a_dossier(ruta_local)
        if publicada:
            ruta_db = publicada

        cli, prod, _src = resolver_cliente_producto(
            job_s, str(ctx.get("job_root") or "")
        )
        if type_override:
            tipo = str(type_override)
            n_spot = 1 if cantidad_spoteos is None else max(0, int(cantidad_spoteos))
        else:
            tipo, n_spot = clasificar_type_y_spoteos(nombre, cantidad_spoteos)

        if _ya_registrada(job_s, ruta_db, nombre):
            return False

        sel_flag = "no"
        try:
            from cotas_seleccionadas_cobre import (
                hermanos_en_carpeta,
                seleccionada_para_captura,
            )

            hermanos = hermanos_en_carpeta(ruta_db)
            # Completar con carpeta local si el share aún no tiene todos.
            if os.path.normcase(ruta_db) != os.path.normcase(ruta_local):
                hermanos = sorted(
                    set(hermanos) | set(hermanos_en_carpeta(ruta_local))
                )
            sel_flag = seleccionada_para_captura(nombre, hermanos)
        except Exception:
            sel_flag = "no"

        asegurar_tabla_cotas_dossier()
        clase = proceso_desde_ruta(ruta_db) or proceso_desde_ruta(ruta_local)
        eid = insertar_evidencia(
            cliente=cli,
            producto=prod,
            job=job_s,
            type_=tipo,
            cantidad_spoteos=n_spot,
            nombre_archivo=nombre,
            ruta=ruta_db,
            clasificacion=clase,
            seleccionadas=sel_flag,
        )
        if eid:
            extra = f" proceso={clase} " if clase else ""
            sel_txt = f" seleccionadas={sel_flag}" if sel_flag else ""
            donde = "dossier" if publicada else "local"
            print(
                f"[dossier] id={eid}: {nombre} type={tipo} "
                f"spoteos={n_spot} {extra}cliente={cli}{sel_txt} ({donde})"
            )
            return True
    except Exception as exc:
        print(f"AVISO dossier: registrar_jpg ({exc})")
    return False


def sincronizar_carpeta_jpgs(
    carpeta: str,
    *,
    job: str | None = None,
    solo_nuevos: bool = True,
) -> int:
    """
    Recorre ``carpeta`` (p. ej. PIEZAS_ACOTADAS o DOSSIER FILES\\JPGS)
    e inserta/actualiza evidencias. Fail-soft.
    """
    if not dossier_habilitado():
        return 0
    n = 0
    root = os.path.abspath(str(carpeta or ""))
    if not os.path.isdir(root):
        return 0
    try:
        ctx_job = job or cargar_contexto_dossier().get("job")
        for dirpath, _dirs, files in os.walk(root):
            for fn in files:
                if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                ruta = os.path.join(dirpath, fn)
                try:
                    if solo_nuevos and ctx_job and _ya_registrada(
                        str(ctx_job), os.path.abspath(ruta), fn
                    ):
                        continue
                    if registrar_jpg(ruta, job=job):
                        n += 1
                except Exception as exc_f:
                    print(f"AVISO dossier: sync skip {fn} ({exc_f})")
                    continue
    except Exception as exc:
        print(f"AVISO dossier: sincronizar_carpeta ({exc})")
    # Con el set completo por carpeta: fijar Seleccionadas (cobre).
    try:
        recalcular_seleccionadas_en_carpeta(root, job=job)
    except Exception as exc_sel:
        print(f"AVISO dossier: post-sync seleccionadas ({exc_sel})")
    if n:
        print(f"[dossier] {n} evidencias registradas desde {root}")
    return n


def iniciar_sesion_dossier(nombre_job: str, job_root: str = "") -> dict:
    """Llamar al inicio del flujo de piezas/board (resuelve/picker DOSSIER)."""
    if not dossier_habilitado():
        return {}
    try:
        _lookup_vsm_job.cache_clear()
    except Exception:
        pass
    try:
        return asegurar_contexto_desde_job(nombre_job, job_root)
    except Exception as exc:
        print(f"AVISO dossier: iniciar_sesion ({exc})")
        return {"job": str(nombre_job or "").strip()}


# Smoke helpers -------------------------------------------------------------

def _smoke_self() -> int:
    assert clasificar_type_y_spoteos("a__LENGTH_1.jpg") == ("TYP", 1)
    assert clasificar_type_y_spoteos("a__HOLE01_2.jpg", 3) == ("TYP", 3)
    assert clasificar_type_y_spoteos("a__LENGTH_SIN_COTA_1.jpg") == ("TYP", 1)
    assert (
        proceso_desde_ruta(
            r"X:\JPGS\Corte\Plasma y Laser\Corte metal\P\a.jpg"
        )
        == "Corte/Plasma y Laser/Corte metal"
    )
    assert (
        proceso_desde_ruta(r"X:\JPGS\Corte\Maquinado\Corte Busbar\P\a.jpg")
        == "Corte/Maquinado/Corte Busbar"
    )
    assert proceso_desde_ruta(r"X:\JPGS\Doblado\Metal\P\a.jpg") == "Doblado/Metal"
    assert (
        proceso_desde_ruta(r"X:\JPGS\Estañado Busbar\a.jpg") == "Estañado Busbar"
    )
    # Legacy
    assert (
        proceso_desde_ruta(r"X:\JPGS\Corte\Doblado\P\a.jpg") == "Corte/Doblado"
    )
    assert proceso_desde_ruta(r"X:\JPGS\Corte\Estañado\a.jpg") == "Corte/Estañado"
    assert proceso_desde_ruta(r"X:\JPGS\Corte\Corte\P\a.jpg") == "Corte/Corte"
    info = normalizar_ruta_dossier(
        r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
        r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
        r"\9919-BOARD2_2\DOSSIER FILES"
    )
    assert info["job_root"].endswith("9919-BOARD2_2"), info
    assert info["dossier_jpgs"].endswith(
        os.path.join("DOSSIER FILES", "JPGS")
    ) or "JPGS" in info["dossier_jpgs"], info
    p, c = producto_cliente_desde_job_root(
        r"X:\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA\9919-BOARD2_2"
    )
    assert p == "ENCLOSURES NEMA 1" and c == "GIGA", (p, c)
    p2, c2 = producto_cliente_desde_job_root(
        r"X:\...\GIGA\9919-BOARD2_2\DOSSIER FILES\JPGS\Almacén"
    )
    assert c2 == "GIGA", (p2, c2)
    print("SMOKE cotas_dossier_registro OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(_smoke_self())
