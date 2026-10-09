"""
Plan de trabajo trazable para Abigail (OTC / --seleccion).

Inventario JSON + estructura de carpetas antes de acotar + publish por
pieza + registro DB por lote. Fail-soft: nunca tumba el acotado.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime, timezone
from typing import Any, Callable

PLAN_NOMBRE = "plan_trabajo.json"

# Contexto activo de la corrida (None = trazable off).
_CTX: dict[str, Any] | None = None
_LOG_FN: Callable[[str], None] | None = None


def _log(msg: str) -> None:
    print(msg)
    if _LOG_FN is not None:
        try:
            _LOG_FN(msg)
        except Exception:
            pass


def trazable_habilitado(*, con_seleccion: bool, es_board: bool) -> bool:
    """
    Default ON en tanque con ``--seleccion``; OFF en BOARD.
    Override: ``PIEZAS_TRAZABLE=0|1``.
    """
    raw = os.environ.get("PIEZAS_TRAZABLE", "").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    if raw in ("1", "true", "yes", "si", "on"):
        return True
    return bool(con_seleccion) and not bool(es_board)


def wipe_job_habilitado(*, trazable: bool, incremental: bool) -> bool:
    """Default ON si trazable y no incremental. Override ``PIEZAS_WIPE_JOB``."""
    if incremental:
        return False
    if os.environ.get("PIEZAS_FILTRO", "").strip():
        return False
    raw = os.environ.get("PIEZAS_WIPE_JOB", "").strip().lower()
    if raw in ("0", "false", "no", "off"):
        return False
    if raw in ("1", "true", "yes", "si", "on"):
        return True
    return bool(trazable)


def _clave(nombre: str) -> str:
    try:
        from generador_tanque_completo import _clave_pieza

        return str(_clave_pieza(nombre) or "").upper()
    except Exception:
        return re.sub(r"[^A-Z0-9]+", "", str(nombre or "").upper())


def _match_clave(a: str, b: str) -> bool:
    try:
        from generador_tanque_completo import _match_claves_pieza

        return bool(_match_claves_pieza(a, b))
    except Exception:
        return a == b or (a and b and (a in b or b in a))


def _invertir_mapa_cara(mapa_por_cara: dict | None) -> dict[str, str]:
    """clave_pieza → etiqueta cara (primera gana)."""
    out: dict[str, str] = {}
    for cara, piezas in (mapa_por_cara or {}).items():
        etiqueta = str(cara or "OTROS").strip().upper() or "OTROS"
        for p in piezas or []:
            k = _clave(p)
            if k and k not in out:
                out[k] = etiqueta
    return out


def _invertir_mapa_clase(mapa_cls: dict | None) -> dict[str, str]:
    """clave_pieza → clasificación canónica."""
    from generador_tanque_completo import (
        SUBCARPETA_SIN_CLASIFICAR,
        SUBCARPETAS_CLASIFICACION_PIEZAS,
    )

    validas = {c.casefold(): c for c in SUBCARPETAS_CLASIFICACION_PIEZAS}
    out: dict[str, str] = {}
    for clase, piezas in (mapa_cls or {}).items():
        cl_raw = str(clase or "").strip()
        if not cl_raw:
            continue
        if cl_raw.casefold() == SUBCARPETA_SIN_CLASIFICAR.casefold():
            dest = SUBCARPETA_SIN_CLASIFICAR
        elif cl_raw.casefold() in validas:
            dest = validas[cl_raw.casefold()]
        else:
            dest = cl_raw
        for p in piezas or []:
            k = _clave(p)
            if k and k not in out:
                out[k] = dest
    return out


def paquete_esperado(clase: str) -> list[str]:
    """Medidas mínimas esperadas por clasificación iProp."""
    cl = str(clase or "").casefold()
    if cl == "doblado" or cl == "plasma doblado":
        return ["LENGTH", "WIDTH", "THK", "HEIGHT", "LEG"]
    if cl in ("corte", "plasma", "maquinado", "piezas soldadas"):
        return ["LENGTH", "WIDTH", "THK"]
    return ["LENGTH", "WIDTH", "THK"]


def _medidas_en_jpgs(nombres: list[str]) -> set[str]:
    meds: set[str] = set()
    for fn in nombres:
        parts = os.path.splitext(os.path.basename(fn))[0].split("__")
        if len(parts) < 3:
            # legacy ITEM_LENGTH_1
            m = re.search(
                r"_(LENGTH|WIDTH|THK|HEIGHT|LEG|OD|ID|HOLE\d*|XMIN|YMIN|"
                r"XCENTRO|YCENTRO|CUT_LENGTH|CUT_WIDTH)_",
                os.path.splitext(os.path.basename(fn))[0],
                re.I,
            )
            if m:
                meds.add(m.group(1).upper())
            continue
        tok = parts[2]
        med = tok.split("_")[0].upper()
        if med:
            meds.add(med)
        # LENGTH_SIN_COTA
        if "SIN_COTA" in tok.upper():
            meds.add("SIN_COTA")
    return meds


def _pieza_completa(
    clase: str, jpgs: list[str]
) -> tuple[bool, list[str], list[str]]:
    """
    Returns ``(min_ok, faltan_min, extras_faltan)``.

    Mínimo: LENGTH+WIDTH+THK, o OD (barras/redondos).
    Extras (HEIGHT/LEG en doblado) solo informativos.
    """
    meds = _medidas_en_jpgs(jpgs)
    faltan_min: list[str] = []
    if "LENGTH" not in meds and "OD" not in meds:
        faltan_min.append("LENGTH")
    if "WIDTH" not in meds and "OD" not in meds:
        faltan_min.append("WIDTH")
    if "THK" not in meds and "OD" not in meds:
        faltan_min.append("THK")
    extras: list[str] = []
    for m in paquete_esperado(clase):
        if m in ("HEIGHT", "LEG") and m not in meds:
            extras.append(m)
    return (not faltan_min), faltan_min, extras


def ruta_plan(carpeta_tanque: str) -> str:
    return os.path.join(os.path.abspath(carpeta_tanque), PLAN_NOMBRE)


def construir_inventario(
    *,
    job: str,
    mapa_por_cara: dict | None,
    mapa_clasificacion: dict | None,
    catalogo: set[str] | None,
    nombres_almacen: set[str] | None = None,
) -> dict[str, Any]:
    """Arma el plan de trabajo (piezas a acotar)."""
    from generador_tanque_completo import (
        SUBCARPETA_OTROS_PIEZAS,
        SUBCARPETA_SIN_CLASIFICAR,
    )

    cara_por_clave = _invertir_mapa_cara(mapa_por_cara)
    clase_por_clave = _invertir_mapa_clase(mapa_clasificacion)
    alm_claves = {_clave(n) for n in (nombres_almacen or set()) if _clave(n)}

    # Universo de nombres: catálogo filtrado o unión de mapas.
    nombres: set[str] = set()
    if catalogo is not None:
        nombres = {str(n).strip() for n in catalogo if str(n).strip()}
    else:
        for piezas in (mapa_por_cara or {}).values():
            nombres.update(str(p).strip() for p in (piezas or []) if str(p).strip())
        for piezas in (mapa_clasificacion or {}).values():
            nombres.update(str(p).strip() for p in (piezas or []) if str(p).strip())

    piezas_out: list[dict[str, Any]] = []
    vistos: set[str] = set()
    for nombre in sorted(nombres, key=lambda s: s.upper()):
        k = _clave(nombre)
        if not k or k in vistos:
            continue
        if k in alm_claves:
            continue
        # Omitir Almacén por clase
        clase = clase_por_clave.get(k) or SUBCARPETA_SIN_CLASIFICAR
        if str(clase).casefold().startswith("almac"):
            continue
        vistos.add(k)
        cara = cara_por_clave.get(k) or SUBCARPETA_OTROS_PIEZAS
        piezas_out.append(
            {
                "item": nombre,
                "clave": k,
                "cara": cara,
                "clase": clase,
                "paquete_esperado": paquete_esperado(clase),
                "estado": "pendiente",
                "jpgs": [],
                "dir_local": "",
                "unc": "",
                "db_lote": None,
                "faltan": [],
                "error": "",
            }
        )

    return {
        "version": 1,
        "job": str(job or "").strip(),
        "creado": datetime.now(timezone.utc).isoformat(),
        "actualizado": datetime.now(timezone.utc).isoformat(),
        "piezas": piezas_out,
    }


def guardar_plan(carpeta_tanque: str, plan: dict) -> str:
    plan = dict(plan)
    plan["actualizado"] = datetime.now(timezone.utc).isoformat()
    ruta = ruta_plan(carpeta_tanque)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    tmp = ruta + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(plan, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ruta)
    return ruta


def cargar_plan(carpeta_tanque: str) -> dict | None:
    ruta = ruta_plan(carpeta_tanque)
    if not os.path.isfile(ruta):
        return None
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:
        _log(f"AVISO plan_trabajo: no se pudo leer ({exc})")
        return None


def _dir_pieza_local(carpeta_piezas: str, entry: dict) -> str:
    """Ruta local ``<CARA>/<clasificación anidada>/<PIEZA>/``."""
    from generador_tanque_completo import (
        SUBCARPETA_SIN_CLASIFICAR,
        _destino_dirs_clasificacion,
        _nombre_carpeta_pieza,
    )

    cara = str(entry.get("cara") or "OTROS").strip() or "OTROS"
    clase = str(entry.get("clase") or SUBCARPETA_SIN_CLASIFICAR)
    item = str(entry.get("item") or "PIEZA")
    cara_dir = os.path.join(carpeta_piezas, cara)
    fake = f"JOB__{_nombre_carpeta_pieza(item)}__LENGTH_1.00in.jpg"
    prev_prod = ""
    try:
        from creador_vistas import (
            configurar_producto_flujo,
            producto_flujo_actual,
        )

        prev_prod = str(producto_flujo_actual() or "")
        # Trazable Abigail OTC: árbol tanque (Corte → Plasma, no Doblado).
        if prev_prod.upper() != "BOARD":
            configurar_producto_flujo("TANQUE")
    except Exception:
        prev_prod = ""
    try:
        destino_dir, _clave, _ = _destino_dirs_clasificacion(
            cara_dir, clase, fake, None
        )
        return destino_dir
    except Exception:
        return os.path.join(
            cara_dir, clase, _nombre_carpeta_pieza(item)
        )
    finally:
        if prev_prod is not None:
            try:
                from creador_vistas import configurar_producto_flujo

                configurar_producto_flujo(prev_prod)
            except Exception:
                pass


def precrear_arbol(carpeta_piezas: str, plan: dict) -> int:
    """Crea carpetas de todas las piezas del plan. Devuelve cuantas."""
    n = 0
    for entry in plan.get("piezas") or []:
        d = _dir_pieza_local(carpeta_piezas, entry)
        entry["dir_local"] = d
        try:
            os.makedirs(d, exist_ok=True)
            n += 1
        except OSError as exc:
            _log(f"AVISO mkdir {d}: {exc}")
    # caras vacías también
    from generador_tanque_completo import (
        SUBCARPETAS_CARA_SELECCION,
        SUBCARPETA_OTROS_PIEZAS,
    )

    for cara in list(SUBCARPETAS_CARA_SELECCION) + [SUBCARPETA_OTROS_PIEZAS]:
        try:
            os.makedirs(os.path.join(carpeta_piezas, cara), exist_ok=True)
        except OSError:
            pass
    return n


def wipe_local_piezas(carpeta_piezas: str) -> None:
    """Vacía PIEZAS_ACOTADAS local (contenido)."""
    if not os.path.isdir(carpeta_piezas):
        os.makedirs(carpeta_piezas, exist_ok=True)
        return
    for nombre in list(os.listdir(carpeta_piezas)):
        ruta = os.path.join(carpeta_piezas, nombre)
        try:
            if os.path.isdir(ruta):
                shutil.rmtree(ruta, ignore_errors=True)
            else:
                os.remove(ruta)
        except OSError as exc:
            _log(f"AVISO wipe local {nombre}: {exc}")


def wipe_dossier_piezas_acotadas() -> None:
    """Vacía solo ``DOSSIER\\JPGS\\PIEZAS_ACOTADAS`` del job en sesión."""
    try:
        from cotas_dossier_registro import (
            cargar_contexto_dossier,
            dossier_habilitado,
            dossier_publish_habilitado,
        )

        if not dossier_habilitado() or not dossier_publish_habilitado():
            return
        ctx = cargar_contexto_dossier()
        root = str(ctx.get("dossier_jpgs") or "").strip()
        if not root:
            return
        pa = os.path.join(root, "PIEZAS_ACOTADAS")
        if not os.path.isdir(pa):
            return
        _log(f"[TRAZABLE] wipe dossier PIEZAS_ACOTADAS → {pa}")
        for nombre in list(os.listdir(pa)):
            ruta = os.path.join(pa, nombre)
            try:
                if os.path.isdir(ruta):
                    shutil.rmtree(ruta, ignore_errors=True)
                else:
                    os.remove(ruta)
            except OSError as exc:
                _log(f"AVISO wipe dossier {nombre}: {exc}")
    except Exception as exc:
        _log(f"AVISO wipe dossier: {exc}")


def buscar_entry(plan: dict, part_name: str) -> dict | None:
    """API pública: localiza entrada del plan por nombre de pieza."""
    return _buscar_entry(plan, part_name)


def _buscar_entry(plan: dict, part_name: str) -> dict | None:
    k = _clave(part_name)
    if not k:
        return None
    for entry in plan.get("piezas") or []:
        if entry.get("clave") == k:
            return entry
        if _match_clave(k, str(entry.get("clave") or "")):
            return entry
        if _match_clave(k, _clave(str(entry.get("item") or ""))):
            return entry
    return None


def _jpgs_sueltos_en_raiz(carpeta_piezas: str) -> list[str]:
    from generador_tanque_completo import STAGING_DESPLIEGUE, STAGING_ESTANIADO

    out: list[str] = []
    try:
        for nombre in os.listdir(carpeta_piezas):
            ruta = os.path.join(carpeta_piezas, nombre)
            if os.path.isfile(ruta) and nombre.lower().endswith((".jpg", ".jpeg", ".png")):
                out.append(ruta)
            elif os.path.isdir(ruta) and nombre in (
                STAGING_DESPLIEGUE,
                STAGING_ESTANIADO,
            ):
                for fn in os.listdir(ruta):
                    if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                        out.append(os.path.join(ruta, fn))
    except OSError:
        pass
    return out


def _jpgs_match_entry(jpgs: list[str], entry: dict) -> list[str]:
    from generador_tanque_completo import _extraer_pieza_de_jpg

    k_entry = str(entry.get("clave") or "")
    matched: list[str] = []
    for ruta in jpgs:
        try:
            item = _extraer_pieza_de_jpg(os.path.basename(ruta))
        except Exception:
            item = ""
        if _match_clave(_clave(item), k_entry) or _match_clave(
            _clave(item), _clave(str(entry.get("item") or ""))
        ):
            matched.append(ruta)
    return matched


def _mover_jpgs_a_dir(rutas: list[str], destino_dir: str) -> list[str]:
    from generador_tanque_completo import _mover_jpg_a_destino

    os.makedirs(destino_dir, exist_ok=True)
    finales: list[str] = []
    for ruta in rutas:
        nombre = os.path.basename(ruta)
        try:
            _mover_jpg_a_destino(ruta, destino_dir, nombre)
            dest = os.path.join(destino_dir, nombre)
            if os.path.isfile(dest):
                finales.append(dest)
            else:
                # _mover puede renombrar en colisión
                candidatos = [
                    os.path.join(destino_dir, f)
                    for f in os.listdir(destino_dir)
                    if f.lower().startswith(os.path.splitext(nombre)[0][:20].lower())
                ]
                if candidatos:
                    finales.append(sorted(candidatos, key=os.path.getmtime)[-1])
        except Exception as exc:
            _log(f"AVISO mover {nombre}: {exc}")
    return finales


def publicar_pieza_dir(dir_local: str) -> tuple[str, int]:
    """Copia JPGs de una carpeta de pieza al dossier. Returns (unc_dir, n)."""
    try:
        from cotas_dossier_registro import (
            dossier_habilitado,
            dossier_publish_habilitado,
            publicar_jpg_a_dossier,
            cargar_contexto_dossier,
        )

        if not dossier_habilitado() or not dossier_publish_habilitado():
            return "", 0
        if not os.path.isdir(dir_local):
            return "", 0
        n = 0
        unc_dir = ""
        for fn in os.listdir(dir_local):
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            src = os.path.join(dir_local, fn)
            dst = publicar_jpg_a_dossier(src)
            if dst:
                n += 1
                unc_dir = os.path.dirname(dst)
        return unc_dir, n
    except Exception as exc:
        _log(f"AVISO publicar_pieza: {exc}")
        return "", 0


def sincronizar_lote_dirs(dirs: list[str], *, job: str | None = None) -> int:
    """Registra en DB los JPG de las carpetas del lote."""
    try:
        from cotas_dossier_registro import (
            dossier_habilitado,
            registrar_jpg,
            recalcular_seleccionadas_en_carpeta,
            cargar_contexto_dossier,
        )

        if not dossier_habilitado():
            return 0
        n = 0
        job_s = job or str(cargar_contexto_dossier().get("job") or "")
        for d in dirs:
            if not d or not os.path.isdir(d):
                continue
            for fn in os.listdir(d):
                if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                ruta = os.path.join(d, fn)
                try:
                    # Preferir ruta UNC ya publicada si existe vía publish
                    from cotas_dossier_registro import publicar_jpg_a_dossier

                    pub = publicar_jpg_a_dossier(ruta)
                    target = pub or ruta
                    if registrar_jpg(target, job=job_s, forzar=True):
                        n += 1
                except Exception as exc_f:
                    _log(f"AVISO DB skip {fn}: {exc_f}")
            try:
                recalcular_seleccionadas_en_carpeta(d, job=job_s)
            except Exception:
                pass
        return n
    except Exception as exc:
        _log(f"AVISO sincronizar_lote: {exc}")
        return 0


def activar_contexto(
    *,
    plan: dict,
    carpeta_tanque: str,
    carpeta_piezas: str,
    log_fn: Callable[[str], None] | None = None,
) -> None:
    global _CTX, _LOG_FN
    _LOG_FN = log_fn
    _CTX = {
        "plan": plan,
        "carpeta_tanque": os.path.abspath(carpeta_tanque),
        "carpeta_piezas": os.path.abspath(carpeta_piezas),
        "job": str(plan.get("job") or ""),
    }
    _log(
        f"[TRAZABLE] activo — {len(plan.get('piezas') or [])} piezas en plan "
        f"({ruta_plan(carpeta_tanque)})"
    )


def desactivar_contexto() -> None:
    global _CTX, _LOG_FN
    _CTX = None
    _LOG_FN = None


def contexto_activo() -> dict | None:
    return _CTX


def on_lote_exportado(
    carpeta_salida: str,
    lote: list,
    idx_lote: int,
    total_pendientes: int,
    log_fn: Callable[[str], None] | None = None,
) -> None:
    """
    Tras exportar JPG de un lote COM: mover a estructura, publish UNC por
    pieza, registrar DB del lote.
    """
    global _LOG_FN
    if log_fn is not None:
        _LOG_FN = log_fn
    ctx = _CTX
    if not ctx:
        return
    plan = ctx.get("plan") or {}
    carpeta_tanque = ctx.get("carpeta_tanque") or ""
    carpeta_piezas = os.path.abspath(
        carpeta_salida or ctx.get("carpeta_piezas") or ""
    )
    t0 = time.time()
    sueltos = _jpgs_sueltos_en_raiz(carpeta_piezas)
    dirs_ok: list[str] = []
    n_ok = n_err = 0

    for _part_doc, part_name in lote:
        entry = _buscar_entry(plan, str(part_name))
        if entry is None:
            # pieza no estaba en plan (edge): crear OTROS/SIN CLASIFICACION
            from generador_tanque_completo import (
                SUBCARPETA_OTROS_PIEZAS,
                SUBCARPETA_SIN_CLASIFICAR,
            )

            entry = {
                "item": str(part_name),
                "clave": _clave(part_name),
                "cara": SUBCARPETA_OTROS_PIEZAS,
                "clase": SUBCARPETA_SIN_CLASIFICAR,
                "paquete_esperado": paquete_esperado(SUBCARPETA_SIN_CLASIFICAR),
                "estado": "pendiente",
                "jpgs": [],
                "dir_local": "",
                "unc": "",
                "db_lote": None,
                "faltan": [],
                "error": "",
            }
            plan.setdefault("piezas", []).append(entry)

        dest_dir = entry.get("dir_local") or _dir_pieza_local(carpeta_piezas, entry)
        entry["dir_local"] = dest_dir
        matched = _jpgs_match_entry(sueltos, entry)
        # también JPGs ya en dest (re-run)
        if os.path.isdir(dest_dir):
            for fn in os.listdir(dest_dir):
                if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    matched.append(os.path.join(dest_dir, fn))
        # dedupe paths
        seen = set()
        uniq = []
        for r in matched:
            nr = os.path.normcase(os.path.abspath(r))
            if nr in seen:
                continue
            seen.add(nr)
            uniq.append(r)
        matched = uniq

        movidos = []
        for r in matched:
            if os.path.normcase(os.path.dirname(os.path.abspath(r))) == os.path.normcase(
                os.path.abspath(dest_dir)
            ):
                movidos.append(r)
            else:
                movidos.extend(_mover_jpgs_a_dir([r], dest_dir))
                # quitar de sueltos
                try:
                    sueltos = [s for s in sueltos if os.path.normcase(s) != os.path.normcase(r)]
                except Exception:
                    pass

        nombres = [os.path.basename(p) for p in movidos]
        entry["jpgs"] = sorted(set(nombres))
        min_ok, faltan_min, extras = _pieza_completa(
            str(entry.get("clase") or ""), nombres
        )
        entry["faltan"] = list(faltan_min) + [f"{x}?" for x in extras]

        n_pub = 0
        if movidos:
            unc, n_pub = publicar_pieza_dir(dest_dir)
            entry["unc"] = unc or entry.get("unc") or ""
            if dest_dir not in dirs_ok:
                dirs_ok.append(dest_dir)

        if not movidos:
            entry["estado"] = "error"
            entry["error"] = "sin JPG exportados"
            n_err += 1
            _log(
                f"[PIEZA] item={entry.get('item')} cara={entry.get('cara')} "
                f"clase={entry.get('clase')} jpgs=0 unc=- estado=ERROR "
                f"faltan=sin_export"
            )
        elif min_ok:
            entry["estado"] = "ok"
            entry["error"] = ""
            n_ok += 1
            extra_txt = f" extras_faltan={extras}" if extras else ""
            _log(
                f"[PIEZA] item={entry.get('item')} cara={entry.get('cara')} "
                f"clase={entry.get('clase')} jpgs={len(nombres)} "
                f"unc={entry.get('unc') or '-'} estado=OK pub={n_pub}"
                f"{extra_txt}"
            )
        else:
            entry["estado"] = "error"
            entry["error"] = f"paquete incompleto: {faltan_min}"
            n_err += 1
            _log(
                f"[PIEZA] item={entry.get('item')} cara={entry.get('cara')} "
                f"clase={entry.get('clase')} jpgs={len(nombres)} "
                f"unc={entry.get('unc') or '-'} estado=ERROR "
                f"faltan={faltan_min} pub={n_pub}"
            )

    # DB del lote (piezas con carpeta poblada)
    db_n = 0
    if dirs_ok:
        db_n = sincronizar_lote_dirs(dirs_ok, job=str(ctx.get("job") or ""))
        for entry in plan.get("piezas") or []:
            if entry.get("dir_local") in dirs_ok:
                entry["db_lote"] = idx_lote

    elapsed = time.time() - t0
    try:
        guardar_plan(carpeta_tanque, plan)
        ctx["plan"] = plan
    except Exception as exc:
        _log(f"AVISO guardar plan: {exc}")

    _log(
        f"[LOTE {idx_lote}] piezas={len(lote)} ok={n_ok} error={n_err} "
        f"db_reg={db_n} elapsed={elapsed:.1f}s"
    )


def marcar_ok_existentes_incremental(plan: dict, carpeta_piezas: str) -> int:
    """Si incremental: marca ok piezas que ya tienen JPG en su carpeta."""
    n = 0
    for entry in plan.get("piezas") or []:
        d = entry.get("dir_local") or _dir_pieza_local(carpeta_piezas, entry)
        entry["dir_local"] = d
        if not os.path.isdir(d):
            continue
        jpgs = [
            f
            for f in os.listdir(d)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
        if not jpgs:
            continue
        min_ok, _faltan, _extras = _pieza_completa(
            str(entry.get("clase") or ""), jpgs
        )
        if min_ok:
            entry["estado"] = "ok"
            entry["jpgs"] = sorted(jpgs)
            n += 1
    return n


def resumen_plan(plan: dict) -> str:
    piezas = plan.get("piezas") or []
    ok = sum(1 for p in piezas if p.get("estado") == "ok")
    err = sum(1 for p in piezas if p.get("estado") == "error")
    pend = sum(1 for p in piezas if p.get("estado") == "pendiente")
    return f"plan: total={len(piezas)} ok={ok} error={err} pendiente={pend}"
