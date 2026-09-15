# -*- coding: utf-8 -*-
"""
Cierre post-flujo flat (colgado / interrumpido):

  1) Auditoría COM en Inventor (XY + HOLE + THK) → repara fallos
  2) Re-export JPG de hojas reparadas
  3) Publica SOLO piezas OK al share dossier Corte\\Corte (y JPGS\\Corte\\Corte)
  4) Re-corre piezas malas / faltantes hasta que pasen auditoría (max N pasadas)

Uso (machote + ensamble abiertos):
  python _cerrar_auditoria_publicar_retry.py
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import time

import pythoncom

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from rutas_arbol_giga import (  # noqa: E402
    ROOT_JPGS_BOARD2,
    catalogo_piezas_flat,
    dest_flat,
)

SHARE_JPGS_ROOT = ROOT_JPGS_BOARD2
SHARE_DOSSIER = os.path.join(
    SHARE_JPGS_ROOT, "Corte", "Plasma y Laser", "Corte metal"
)
SHARE_JPGS = SHARE_DOSSIER
JOB = "9919-Board 2"
MAX_RETRY = 3
_TOKENS = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")


def _piezas_catalogo(raiz: str) -> list[str]:
    if not os.path.isdir(raiz):
        return []
    out = []
    for n in sorted(os.listdir(raiz)):
        p = os.path.join(raiz, n)
        if not os.path.isdir(p):
            continue
        if n.upper().startswith("COPIA DE") or n.startswith("_"):
            continue
        out.append(n)
    return out


def _staging_dir() -> str:
    # Prefer Board 1 / Board 2 local outputs
    candidates = []
    jpg_root = os.path.join(ROOT, "JPG")
    if os.path.isdir(jpg_root):
        for d in os.listdir(jpg_root):
            st = os.path.join(jpg_root, d, "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE")
            if os.path.isdir(st):
                candidates.append(st)
    if not candidates:
        return ""
    # newest by mtime of folder or any file
    def _mtime(path):
        try:
            files = os.listdir(path)
            if not files:
                return os.path.getmtime(path)
            return max(os.path.getmtime(os.path.join(path, f)) for f in files)
        except Exception:
            return 0

    candidates.sort(key=_mtime, reverse=True)
    return candidates[0]


def _item_desde_fn(fn: str) -> str:
    parts = fn.split("__")
    if len(parts) >= 2:
        return parts[1]
    return ""


def _piezas_en_staging(staging: str) -> dict[str, list[str]]:
    """pieza -> lista de rutas jpg objetivo."""
    out: dict[str, list[str]] = {}
    if not staging or not os.path.isdir(staging):
        return out
    for fn in os.listdir(staging):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        up = fn.upper()
        if not any(t in up for t in _TOKENS):
            continue
        item = _item_desde_fn(fn)
        if not item:
            continue
        out.setdefault(item, []).append(os.path.join(staging, fn))
    return out


def _publicar_pieza(pieza: str, archivos: list[str], roots: list[str]) -> int:
    n = 0
    for root in roots:
        dst_dir = dest_flat(root, pieza)
        os.makedirs(dst_dir, exist_ok=True)
        # limpiar tokens viejos de esta pieza
        for fn in list(os.listdir(dst_dir)):
            up = fn.upper()
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if any(t in up for t in _TOKENS):
                try:
                    os.remove(os.path.join(dst_dir, fn))
                except OSError:
                    pass
        for src in archivos:
            fn = os.path.basename(src)
            try:
                shutil.copy2(src, os.path.join(dst_dir, fn))
                n += 1
            except OSError as e:
                print(f"  AVISO copy {fn}: {e}")
    return n


def _insert_db(rutas: list[str]) -> int:
    try:
        from cotas_dossier_registro import (
            clasificar_type_y_spoteos,
            proceso_desde_ruta,
            producto_cliente_desde_job_root,
            _db_nesting,
        )
        import psycopg2
    except Exception as e:
        print(f"  AVISO DB no disponible: {e}")
        return 0
    job_root = os.path.dirname(os.path.dirname(SHARE_JPGS_ROOT))
    prod, cli = producto_cliente_desde_job_root(job_root)
    if not cli:
        cli, prod = "GIGA", "ENCLOSURES NEMA 1"
    ins = 0
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                for ruta in rutas:
                    fn = os.path.basename(ruta)
                    tipo, n_spot = clasificar_type_y_spoteos(fn)
                    if "_TYP" in fn.upper():
                        n_spot = max(n_spot, 2)
                    clase = proceso_desde_ruta(ruta) or ""
                    cur.execute(
                        """
                        DELETE FROM public.cotas_dossier
                        WHERE job = %s AND nombre_archivo = %s
                        """,
                        (JOB, fn),
                    )
                    cur.execute(
                        """
                        INSERT INTO public.cotas_dossier
                            (cliente, producto, job, type, cantidad_spoteos,
                             nombre_archivo, ruta, clasificacion)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        """,
                        (
                            cli,
                            prod,
                            JOB,
                            tipo,
                            n_spot,
                            fn,
                            os.path.abspath(ruta),
                            clase,
                        ),
                    )
                    ins += 1
            conn.commit()
    except Exception as e:
        print(f"  AVISO DB insert: {e}")
        return 0
    return ins


def _clasificar_piezas_por_auditoria(inv, plano) -> tuple[list[str], list[str], list[str]]:
    """
    Returns (ok_bases, fail_bases, reparadas_nombres_hoja)
    basándose en hojas DESPLIEGUE del plano.
    """
    from auditoria_cotas_flat_com import (
        auditar_y_reparar_plano,
        _hoja_xy_falla,
        _hoja_thk_falla,
        _hoja_hole_falla,
        _es_hoja_thk,
    )

    audit_ok, reparadas = auditar_y_reparar_plano(inv, plano, max_pasadas=2)
    print(f"  audit_ok={audit_ok} reparadas={len(reparadas)}")

    # Re-export reparadas
    if reparadas:
        from generador_vistas import exportar_hojas_jpg

        permitidos = {str(r).upper() for r in reparadas}
        piezas_tocadas = set()
        for r in reparadas:
            u = str(r).upper()
            piezas_tocadas.add(re.sub(r"_DESPLIEGUE_.*$", "", u, flags=re.I))
        try:
            for i in range(1, int(plano.Sheets.Count) + 1):
                nom = str(plano.Sheets.Item(i).Name)
                up = nom.upper().rsplit(":", 1)[0]
                if "_DESPLIEGUE_" not in up:
                    continue
                for p in piezas_tocadas:
                    if p and p in up:
                        permitidos.add(up)
                        break
        except Exception:
            pass
        staging_parent = os.path.dirname(_staging_dir() or ROOT)
        carpeta = staging_parent if staging_parent else ROOT
        # Prefer PIEZAS_ACOTADAS
        if os.path.basename(carpeta) != "PIEZAS_ACOTADAS":
            # staging = .../PIEZAS_ACOTADAS/_STAGING_DESPLIEGUE
            st = _staging_dir()
            if st:
                carpeta = os.path.dirname(st)
        print(f"  re-export → {carpeta}")
        os.environ["SOLO_FLAT_CORTE"] = "1"
        exportar_hojas_jpg(
            inv,
            plano,
            carpeta_salida=carpeta,
            nombres_permitidos=permitidos,
            nombre_job=JOB,
        )

    # Clasificar piezas por estado final de hojas
    ok_set: set[str] = set()
    fail_set: set[str] = set()
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper().rsplit(":", 1)[0]
            if int(hoja.DrawingViews.Count) < 1:
                continue
            vista = hoja.DrawingViews.Item(1)
            pieza = re.sub(r"_DESPLIEGUE_.*$", "", up, flags=re.I)
            pieza = re.sub(r"_XCENTRO.*$|_YCENTRO.*$|_DIAMETRO.*$|_THK.*$|_LADO.*$", "", pieza)
            # mejor: quitar desde _DESPLIEGUE
            m = re.match(r"^(.*?)_DESPLIEGUE_", up)
            if m:
                pieza = m.group(1)
            falla = False
            if "XCENTRO" in up or "YCENTRO" in up:
                eje = "X" if "XCENTRO" in up else "Y"
                falla, _ = _hoja_xy_falla(hoja, vista, eje)
            elif _es_hoja_thk(up):
                falla, _ = _hoja_thk_falla(hoja, vista)
            elif "DIAMETRO_H" in up:
                falla, _ = _hoja_hole_falla(hoja, vista)
            else:
                continue
            if falla:
                fail_set.add(pieza)
            else:
                ok_set.add(pieza)
        except Exception:
            continue
    # Si falló alguna hoja de la pieza, no está OK
    ok_final = sorted(p for p in ok_set if p not in fail_set)
    fail_final = sorted(fail_set)
    return ok_final, fail_final, reparadas


def _match_pieza(nombre_hoja_pieza: str, catalogo: list[str]) -> str | None:
    u = nombre_hoja_pieza.upper()
    for p in catalogo:
        pu = p.upper()
        if pu == u or u.startswith(pu) or pu in u:
            return p
    return None


def _retry_piezas(inv, plano, ensamble, piezas: list[str], carpeta_piezas: str) -> bool:
    if not piezas:
        return True
    print(f"\n=== RETRY {len(piezas)} piezas malas/faltantes ===")
    for p in piezas:
        print(f"  - {p}")
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["PIEZAS_FILTRO"] = ",".join(piezas)
    os.environ["PIEZAS_INCREMENTAL"] = "0"
    os.environ["ENSAMBLES_INDEPENDIENTES"] = "0"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB

    from generador_vistas import ejecutar_flujo_desde_app
    import creador_vistas as _cv

    _cv.configurar_piezas_corte(set(piezas))
    return bool(
        ejecutar_flujo_desde_app(
            inv,
            ensamble,
            plano,
            carpeta_salida=carpeta_piezas,
            incremental=False,
            catalogo_piezas=set(piezas),
        )
    )


def main() -> int:
    print("=" * 62)
    print(" CIERRE: audit COM → publicar OK → retry malas")
    print("=" * 62)

    catalogo = catalogo_piezas_flat(SHARE_JPGS_ROOT)
    print(f"catalogo share: {len(catalogo)} piezas")
    staging = _staging_dir()
    print(f"staging: {staging}")
    por_pieza = _piezas_en_staging(staging)
    print(f"piezas con JPG en staging: {len(por_pieza)}")

    shares = [SHARE_JPGS_ROOT]

    pythoncom.CoInitialize()
    try:
        import win32com.client
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            print("ERROR: abre machote + ensamble Board")
            return 2
        try:
            h = _encontrar_hoja_machote(plano)
            if h is not None:
                h.Activate()
        except Exception:
            pass

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)
        os.makedirs(carpeta_piezas, exist_ok=True)
        os.environ["SOLO_FLAT_CORTE"] = "1"
        os.environ["COTAS_JOB_OVERRIDE"] = JOB

        pendientes_malas: list[str] = []
        # piezas del catalogo sin JPG
        for p in catalogo:
            if not any(p.upper() == k.upper() for k in por_pieza):
                pendientes_malas.append(p)

        for intento in range(1, MAX_RETRY + 1):
            print(f"\n######## PASADA AUDIT/PUBLISH {intento}/{MAX_RETRY} ########")
            ok_bases, fail_bases, _reps = _clasificar_piezas_por_auditoria(inv, plano)
            print(f"  hojas OK piezas≈{len(ok_bases)} FAIL≈{len(fail_bases)}")

            # Refrescar staging map
            staging = _staging_dir()
            por_pieza = _piezas_en_staging(staging)

            # Publicar solo piezas con JPG y sin fallos de auditoría
            fail_u = {f.upper() for f in fail_bases}
            pubs = 0
            rutas_db = []
            for item, archivos in sorted(por_pieza.items()):
                cat = _match_pieza(item, catalogo) or item
                # Si alguna hoja de esa pieza falló auditoría, no publicar aún
                if any(cat.upper() in f or f in cat.upper() for f in fail_u):
                    print(f"  SKIP publish (audit FAIL): {cat}")
                    if cat not in pendientes_malas:
                        pendientes_malas.append(cat)
                    continue
                # También si el nombre de hoja falla contiene el item
                if any(item.upper() in f for f in fail_u):
                    print(f"  SKIP publish (audit FAIL item): {item}")
                    if cat not in pendientes_malas:
                        pendientes_malas.append(cat)
                    continue
                n = _publicar_pieza(cat, archivos, shares)
                pubs += n
                # DB: rutas en dossier share
                for a in archivos:
                    rutas_db.append(
                        os.path.join(
                            dest_flat(SHARE_JPGS_ROOT, cat),
                            os.path.basename(a),
                        )
                    )
                print(f"  PUBLISH OK {cat}: {len(archivos)} jpg")

            print(f"  publicados archivos≈{pubs}")
            print(f"  DB inserts={_insert_db(rutas_db)}")

            # Retry set = fail + faltantes + aún no publicadas
            retry = []
            for p in fail_bases:
                m = _match_pieza(p, catalogo)
                if m and m not in retry:
                    retry.append(m)
                elif not m and p not in retry:
                    retry.append(p)
            for p in pendientes_malas:
                if p not in retry:
                    retry.append(p)
            for p in catalogo:
                if p not in por_pieza and p not in retry:
                    # sin match exacto de key
                    if not any(p.upper() == k.upper() for k in por_pieza):
                        retry.append(p)

            # Unique preserve order
            seen = set()
            retry2 = []
            for p in retry:
                u = p.upper()
                if u in seen:
                    continue
                seen.add(u)
                retry2.append(p)
            retry = retry2

            if not retry:
                print("\n=== TODAS LAS PIEZAS OK Y PUBLICADAS ===")
                return 0

            print(f"\nQuedan {len(retry)} por re-correr")
            pendientes_malas = list(retry)
            ok_flow = _retry_piezas(inv, plano, ensamble, retry, carpeta_piezas)
            print(f"  retry flujo ok={ok_flow}")
            time.sleep(1)

        print("\nAVISO: se agotaron reintentos; quedan malas:", pendientes_malas)
        return 3
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback

        traceback.print_exc()
        return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
