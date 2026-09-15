# -*- coding: utf-8 -*-
"""
Al terminar / ahora:
  1) Auditoría COM ESTRICTA (origen IL toca pieza, no AABB flotante)
  2) Repara fallos y re-exporta
  3) Publica SOLO piezas 100% OK → dossier Corte\\Corte
  4) Reintenta las malas; no sube nada dudoso
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
MAX_RETRY = 8  # audit → reparar → reflujo hasta que todo OK
_TOKENS = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")


def _catalogo(raiz: str) -> list[str]:
    if not os.path.isdir(raiz):
        return []
    return sorted(
        n
        for n in os.listdir(raiz)
        if os.path.isdir(os.path.join(raiz, n))
        and not n.upper().startswith("COPIA DE")
        and not n.startswith("_")
    )


def _staging() -> str:
    jpg = os.path.join(ROOT, "JPG")
    cands = []
    if os.path.isdir(jpg):
        for d in os.listdir(jpg):
            st = os.path.join(jpg, d, "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE")
            if os.path.isdir(st):
                cands.append(st)
    if not cands:
        return ""

    def mt(p):
        try:
            fs = os.listdir(p)
            if not fs:
                return os.path.getmtime(p)
            return max(os.path.getmtime(os.path.join(p, f)) for f in fs)
        except Exception:
            return 0

    return sorted(cands, key=mt, reverse=True)[0]


def _map_staging(st: str) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    if not st or not os.path.isdir(st):
        return out
    for fn in os.listdir(st):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        up = fn.upper()
        if not any(t in up for t in _TOKENS):
            continue
        parts = fn.split("__")
        if len(parts) < 2:
            continue
        out.setdefault(parts[1], []).append(os.path.join(st, fn))
    return out


def _completo(archivos: list[str]) -> bool:
    ups = [os.path.basename(a).upper() for a in archivos]
    return (
        any("XCENTRO" in u or "YCENTRO" in u for u in ups)
        and any("__THK_" in u for u in ups)
    )


def _pieza_de_hoja(up: str) -> str:
    m = re.match(r"^(.*?)_DESPLIEGUE_", up)
    return m.group(1) if m else up


def _auditar_piezas(plano) -> tuple[set[str], set[str]]:
    """
    Returns (ok_piezas, fail_piezas).
    Una pieza está OK solo si TODAS sus hojas XY/THK pasan.
    """
    from auditoria_cotas_flat_com import (
        _activar_hoja,
        _hoja_xy_falla,
        _hoja_thk_falla,
        _es_hoja_thk,
    )

    # Borrar Copia de
    for i in range(int(plano.Sheets.Count), 0, -1):
        try:
            h = plano.Sheets.Item(i)
            if str(h.Name).upper().startswith("COPIA DE"):
                h.Delete()
                print(f"  DEL basura {h.Name}")
        except Exception:
            pass

    vistos: dict[str, list[bool]] = {}
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper().rsplit(":", 1)[0]
            if "_DESPLIEGUE_" not in up:
                continue
            if int(hoja.DrawingViews.Count) < 1:
                continue
            _activar_hoja(hoja)
            vista = hoja.DrawingViews.Item(1)
            pieza = _pieza_de_hoja(up)
            ok = True
            motivo = ""
            if "XCENTRO" in up or "YCENTRO" in up:
                eje = "X" if "XCENTRO" in up else "Y"
                falla, motivo = _hoja_xy_falla(hoja, vista, eje)
                ok = not falla
            elif _es_hoja_thk(up):
                falla, motivo = _hoja_thk_falla(hoja, vista)
                ok = not falla
            else:
                continue
            vistos.setdefault(pieza, []).append(ok)
            tag = "OK  " if ok else "FAIL"
            print(f"  {tag} {up}" + (f": {motivo}" if not ok else ""))
        except Exception as e:
            print(f"  AVISO {e}")

    ok_p, fail_p = set(), set()
    for pieza, flags in vistos.items():
        if flags and all(flags):
            ok_p.add(pieza)
        else:
            fail_p.add(pieza)
    return ok_p, fail_p


def _publicar(pieza: str, archivos: list[str], roots: list[str]) -> list[str]:
    # Solo JPG con valor a 6 decimales (descartar 3)
    re6 = re.compile(r"\.\d{6}(?:_|\.|$)")
    archivos = [
        a
        for a in archivos
        if re6.search(os.path.basename(a))
        and any(t in os.path.basename(a).upper() for t in _TOKENS)
    ]
    out = []
    for root in roots:
        dst = dest_flat(root, pieza)
        os.makedirs(dst, exist_ok=True)
        for fn in list(os.listdir(dst)):
            if fn.lower().endswith((".jpg", ".jpeg", ".png")) and any(
                t in fn.upper() for t in _TOKENS
            ):
                try:
                    os.remove(os.path.join(dst, fn))
                except OSError:
                    pass
        for src in archivos:
            fn = os.path.basename(src)
            d = os.path.join(dst, fn)
            try:
                shutil.copy2(src, d)
                out.append(d)
            except OSError as e:
                print(f"  AVISO copy {fn}: {e}")
    return out


def _db(rutas: list[str]) -> int:
    try:
        from cotas_dossier_registro import (
            clasificar_type_y_spoteos,
            proceso_desde_ruta,
            producto_cliente_desde_job_root,
            _db_nesting,
        )
        import psycopg2
    except Exception as e:
        print(f"  AVISO DB: {e}")
        return 0
    job_root = os.path.dirname(os.path.dirname(SHARE_JPGS_ROOT))
    prod, cli = producto_cliente_desde_job_root(job_root)
    if not cli:
        cli, prod = "GIGA", "ENCLOSURES NEMA 1"
    n = 0
    try:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                for ruta in rutas:
                    fn = os.path.basename(ruta)
                    tipo, spot = clasificar_type_y_spoteos(fn)
                    if "_TYP" in fn.upper():
                        spot = max(spot, 2)
                    cur.execute(
                        "DELETE FROM public.cotas_dossier WHERE job=%s AND nombre_archivo=%s",
                        (JOB, fn),
                    )
                    cur.execute(
                        """
                        INSERT INTO public.cotas_dossier
                          (cliente,producto,job,type,cantidad_spoteos,
                           nombre_archivo,ruta,clasificacion)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        """,
                        (
                            cli,
                            prod,
                            JOB,
                            tipo,
                            spot,
                            fn,
                            os.path.abspath(ruta),
                            proceso_desde_ruta(ruta) or "",
                        ),
                    )
                    n += 1
            conn.commit()
    except Exception as e:
        print(f"  AVISO DB: {e}")
        return 0
    return n


def main() -> int:
    print("=" * 62)
    print(" AUDIT ESTRICTO → publicar SOLO OK → retry malas")
    print("=" * 62)
    cat = catalogo_piezas_flat(SHARE_JPGS_ROOT)
    shares = [SHARE_JPGS_ROOT]
    print(f"catalogo={len(cat)} destino=JPGS")

    pythoncom.CoInitialize()
    try:
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS
        from auditoria_cotas_flat_com import auditar_y_reparar_plano
        from generador_vistas import exportar_hojas_jpg
        import creador_vistas as cv
        from generador_vistas import ejecutar_flujo_desde_app

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            print("ERROR: machote + ensamble")
            return 2
        try:
            h = _encontrar_hoja_machote(plano)
            if h:
                h.Activate()
        except Exception:
            pass
        carpeta_piezas = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        os.makedirs(carpeta_piezas, exist_ok=True)
        os.environ["SOLO_FLAT_CORTE"] = "1"
        os.environ["COTAS_JOB_OVERRIDE"] = JOB
        # CRÍTICO: GIGA Board flat = mm (default cota_estilo es "in" de tanques)
        from cota_estilo import set_unidad_cota, get_unidad_cota
        from producto_tipo import aplicar_unidad_producto

        set_unidad_cota("mm")
        try:
            aplicar_unidad_producto(ensamble=ensamble)
        except Exception:
            pass
        set_unidad_cota("mm")  # re-fuerza por si clasificador falló
        print(f"  Unidades de cota FORZADAS: {get_unidad_cota()}")

        publicadas_total: set[str] = set()

        for pasada in range(1, MAX_RETRY + 1):
            print(f"\n######## PASADA {pasada}/{MAX_RETRY} ########")
            set_unidad_cota("mm")
            print("[1] reparar COM (origen IL + mm + 3 dec)…")
            ok_a, reps = auditar_y_reparar_plano(inv, plano, max_pasadas=2)
            print(f"  audit_repair ok={ok_a} reps={len(reps)}")
            if reps:
                permitidos = {r.upper() for r in reps}
                for i in range(1, int(plano.Sheets.Count) + 1):
                    try:
                        up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
                    except Exception:
                        continue
                    for r in list(permitidos):
                        base = re.sub(r"_DESPLIEGUE_.*$", "", r)
                        if base and base in up and "_DESPLIEGUE_" in up:
                            permitidos.add(up)
                exportar_hojas_jpg(
                    inv,
                    plano,
                    carpeta_salida=carpeta_piezas,
                    nombres_permitidos=permitidos,
                    nombre_job=JOB,
                )

            print("[2] clasificación ESTRICTA OK/FAIL…")
            set_unidad_cota("mm")
            ok_p, fail_p = _auditar_piezas(plano)
            print(f"  OK={len(ok_p)} FAIL={len(fail_p)}")
            if fail_p:
                print("  FAIL:", ", ".join(sorted(fail_p)[:20]))

            st = _staging()
            por = _map_staging(st)
            print(f"[3] staging piezas_jpg={len(por)}")

            pubs_db = []
            for item, archivos in sorted(por.items()):
                pieza = next((c for c in cat if c.upper() == item.upper()), item)
                # Solo si auditoría COM marca OK (todas las hojas XY/THK)
                match_ok = any(
                    pieza.upper() == o or o in pieza.upper() or pieza.upper() in o
                    for o in ok_p
                )
                match_fail = any(
                    pieza.upper() == f or f in pieza.upper() or pieza.upper() in f
                    for f in fail_p
                )
                if match_fail or not match_ok:
                    print(f"  NO SUBE (audit): {pieza}")
                    continue
                if not _completo(archivos):
                    print(f"  NO SUBE (incompleto): {pieza}")
                    continue
                if pieza in publicadas_total:
                    # Re-publicar si se re-exportó (mantener dossier al día)
                    pass
                rutas = _publicar(pieza, archivos, shares)
                pubs_db.extend(
                    [
                        r
                        for r in rutas
                        if SHARE_JPGS_ROOT.lower() in r.lower()
                        or SHARE_JPGS_ROOT in r
                    ]
                )
                pubs_db.extend(
                    [
                        os.path.join(
                            dest_flat(SHARE_JPGS_ROOT, pieza),
                            os.path.basename(a),
                        )
                        for a in archivos
                    ]
                )
                publicadas_total.add(pieza)
                print(f"  SUBE OK {pieza}: {len(archivos)} jpg → JPGS")

            print(f"[4] DB={_db(pubs_db)} acumulado_ok={len(publicadas_total)}/{len(cat)}")

            # ¿Catálogo completo y sin fails?
            faltan = [c for c in cat if c not in publicadas_total]
            if not fail_p and not faltan:
                print("\nLISTO: 100% piezas auditadas OK en JPGS")
                print(f"  {SHARE_JPGS_ROOT}")
                return 0

            # Retry = fails + faltantes del catálogo
            retry = []
            for f in sorted(fail_p):
                m = next(
                    (c for c in cat if c.upper() == f.upper() or f.upper() in c.upper()),
                    None,
                )
                if m and m not in retry:
                    retry.append(m)
            for c in faltan:
                if c not in retry:
                    retry.append(c)

            if not retry:
                print("\nLISTO: solo piezas auditadas OK en")
                print(f"  {SHARE_DOSSIER}")
                return 0

            if pasada >= MAX_RETRY:
                break

            print(f"[5] re-correr flujo de {len(retry)} malas/faltantes…")
            os.environ["PIEZAS_FILTRO"] = ",".join(retry)
            os.environ["PIEZAS_INCREMENTAL"] = "0"
            os.environ["SOLO_FLAT_CORTE"] = "1"
            set_unidad_cota("mm")
            cv.configurar_piezas_corte(set(retry))
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=False,
                catalogo_piezas=set(retry),
            )
            print("[6] post-flujo: re-auditar antes de siguiente pasada…")
            time.sleep(1)

        print(f"\nFIN parcial: publicadas={len(publicadas_total)}/{len(cat)}")
        print(f"Quedan fuera: {sorted(set(cat) - publicadas_total)[:30]}")
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
