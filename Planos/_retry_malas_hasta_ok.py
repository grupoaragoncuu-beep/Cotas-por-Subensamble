# -*- coding: utf-8 -*-
"""
Reintenta piezas malas/faltantes hasta que el dossier tenga XY+THK a 6 dec.

  1) Flujo flat solo para faltantes
  2) Auditoría COM + reparación XY/THK
  3) Export JPG
  4) Publicar SOLO OK con 6 decimales
  5) Repetir hasta 100% o MAX_PASADAS
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
    carpeta_pieza_flat,
    catalogo_piezas_flat,
    dest_flat,
)

SHARE_JPGS_ROOT = ROOT_JPGS_BOARD2
SHARE_DOSSIER = os.path.join(
    SHARE_JPGS_ROOT, "Corte", "Plasma y Laser", "Corte metal"
)
SHARE_JPGS = SHARE_DOSSIER
JOB = "9919-Board 2"
MAX_PASADAS = 5
_TOKENS = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")
_RE_6 = re.compile(r"\.\d{6}(?:_|\.|$)")


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


def _pieza_ok_en_share(pieza: str) -> bool:
    d = carpeta_pieza_flat(SHARE_JPGS_ROOT, pieza)
    if not os.path.isdir(d):
        return False
    xy = thk = False
    for fn in os.listdir(d):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        if not _RE_6.search(fn):
            continue
        up = fn.upper()
        if "XCENTRO" in up or "YCENTRO" in up or "XMIN" in up or "YMIN" in up:
            xy = True
        if "__THK_" in up:
            thk = True
    return xy and thk


def _faltantes(cat: list[str]) -> list[str]:
    return [c for c in cat if not _pieza_ok_en_share(c)]


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
        if not any(t in fn.upper() for t in _TOKENS):
            continue
        if not _RE_6.search(fn):
            continue
        parts = fn.split("__")
        if len(parts) < 2:
            continue
        out.setdefault(parts[1], []).append(os.path.join(st, fn))
    return out


def _completo(archivos: list[str]) -> bool:
    ups = [os.path.basename(a).upper() for a in archivos]
    return (
        any(
            t in u
            for u in ups
            for t in ("XCENTRO", "YCENTRO", "XMIN", "YMIN")
        )
        and any("__THK_" in u for u in ups)
    )


def _publicar(pieza: str, archivos: list[str], roots: list[str]) -> int:
    archivos = [a for a in archivos if _RE_6.search(os.path.basename(a))]
    if not archivos:
        return 0
    n = 0
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
            try:
                shutil.copy2(src, os.path.join(dst, os.path.basename(src)))
                n += 1
            except OSError as e:
                print(f"  AVISO copy: {e}")
    return n


def _auditar_pieza(plano, pieza: str) -> tuple[bool, str]:
    from auditoria_cotas_flat_com import (
        _activar_hoja,
        _hoja_xy_falla,
        _hoja_thk_falla,
        _es_hoja_thk,
        _reparar_hoja_xy,
        _forzar_mm_flat,
    )

    _forzar_mm_flat()
    up_p = pieza.upper()
    vistos = 0
    fallos = []
    inv = plano.Parent
    try:
        tg = inv.TransientGeometry
    except Exception:
        tg = None
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper().rsplit(":", 1)[0]
        except Exception:
            continue
        if up_p not in up or "_DESPLIEGUE_" not in up:
            continue
        if up.startswith("COPIA DE"):
            continue
        if int(hoja.DrawingViews.Count) < 1:
            continue
        _activar_hoja(hoja)
        vista = hoja.DrawingViews.Item(1)
        if "XCENTRO" in up or "YCENTRO" in up or "XMIN" in up or "YMIN" in up:
            eje = "X" if ("XCENTRO" in up or "XMIN" in up) else "Y"
            falla, motivo = _hoja_xy_falla(hoja, vista, eje)
            vistos += 1
            if falla and tg is not None:
                print(f"  reparar {up}: {motivo}")
                try:
                    _reparar_hoja_xy(hoja, vista, inv, tg, eje)
                    _activar_hoja(hoja)
                    vista = hoja.DrawingViews.Item(1)
                    falla, motivo = _hoja_xy_falla(hoja, vista, eje)
                except Exception as e:
                    falla, motivo = True, f"reparo falló: {e}"
            if falla:
                fallos.append(f"{up}: {motivo}")
        elif _es_hoja_thk(up):
            falla, motivo = _hoja_thk_falla(hoja, vista)
            vistos += 1
            if falla:
                fallos.append(f"{up}: {motivo}")
    if vistos == 0:
        return False, "sin hojas XY/THK"
    if fallos:
        return False, "; ".join(fallos[:3])
    return True, f"ok ({vistos} hojas)"


def main() -> int:
    print("=" * 62)
    print(" RETRY MALAS → reparar → publicar 6dec → hasta 100%")
    print("=" * 62)
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    from cota_estilo import set_unidad_cota

    set_unidad_cota("mm")

    cat = catalogo_piezas_flat(SHARE_JPGS_ROOT)
    shares = [SHARE_JPGS_ROOT]
    print(f"catalogo={len(cat)}")

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
        from generador_vistas import exportar_hojas_jpg, ejecutar_flujo_desde_app
        import creador_vistas as cv
        from producto_tipo import aplicar_unidad_producto
        from barrenos_xy_despliegue import acotar_barrenos_xy_despliegue

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            print("ERROR: machote + ensamble")
            return 2
        try:
            aplicar_unidad_producto(ensamble=ensamble)
        except Exception:
            pass
        set_unidad_cota("mm")
        try:
            h = _encontrar_hoja_machote(plano)
            if h:
                h.Activate()
        except Exception:
            pass
        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        os.makedirs(carpeta, exist_ok=True)

        for pasada in range(1, MAX_PASADAS + 1):
            faltan = _faltantes(cat)
            print(f"\n######## PASADA {pasada}/{MAX_PASADAS} faltan={len(faltan)} ########")
            if not faltan:
                print("LISTO: 100% piezas OK a 6 dec en dossier")
                print(f"  {SHARE_DOSSIER}")
                return 0
            print("  faltan:", ", ".join(faltan))

            print("[1] flujo flat solo malas…")
            os.environ["PIEZAS_FILTRO"] = ",".join(faltan)
            os.environ["PIEZAS_INCREMENTAL"] = "0"
            os.environ["SOLO_FLAT_CORTE"] = "1"
            set_unidad_cota("mm")
            cv.configurar_piezas_corte(set(faltan))
            try:
                ejecutar_flujo_desde_app(
                    inv,
                    ensamble,
                    plano,
                    carpeta_salida=carpeta,
                    incremental=False,
                    catalogo_piezas=set(faltan),
                )
            except Exception as e:
                print(f"  AVISO flujo: {e}")

            print("[2] re-acotar XY (incluye dx/dy < 0)…")
            set_unidad_cota("mm")
            try:
                # Solo frente de piezas faltantes si hay hojas
                acotar_barrenos_xy_despliegue(None)
            except Exception as e:
                print(f"  AVISO xy: {e}")

            print("[3] audit+reparar por pieza…")
            ok_p = []
            for pieza in faltan:
                bien, motivo = _auditar_pieza(plano, pieza)
                print(f"  {'OK' if bien else 'FAIL'} {pieza}: {motivo}")
                if bien:
                    ok_p.append(pieza)

            print("[4] export JPG de reparadas/faltantes…")
            permitidos = set()
            for i in range(1, int(plano.Sheets.Count) + 1):
                try:
                    up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
                except Exception:
                    continue
                if "_DESPLIEGUE_" not in up:
                    continue
                if not any(f.upper() in up for f in faltan):
                    continue
                if any(
                    t in up
                    for t in (
                        "XCENTRO",
                        "YCENTRO",
                        "XMIN",
                        "YMIN",
                        "_THK",
                        "_LADO",
                        "DIAMETRO_H",
                    )
                ):
                    permitidos.add(up)
            if permitidos:
                exportar_hojas_jpg(
                    inv,
                    plano,
                    carpeta_salida=carpeta,
                    nombres_permitidos=permitidos,
                    nombre_job=JOB,
                )

            print("[5] publicar OK 6dec…")
            por = _map_staging(_staging())
            pubs = 0
            for pieza in faltan:
                archivos = por.get(pieza) or next(
                    (v for k, v in por.items() if k.upper() == pieza.upper()),
                    [],
                )
                if not _completo(archivos):
                    print(f"  NO SUBE incompleto: {pieza}")
                    continue
                # Debe haber pasado audit en esta pasada o ya OK
                bien, motivo = _auditar_pieza(plano, pieza)
                if not bien:
                    print(f"  NO SUBE audit: {pieza}: {motivo}")
                    continue
                n = _publicar(pieza, archivos, shares)
                print(f"  SUBE {pieza}: {n} jpg")
                pubs += n

            quedan = _faltantes(cat)
            print(f"[6] publicados_files≈{pubs} quedan={len(quedan)}")
            if not quedan:
                print("\nLISTO: 100% OK")
                return 0
            time.sleep(1)

        print(f"\nFIN parcial: aún faltan {_faltantes(cat)}")
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
