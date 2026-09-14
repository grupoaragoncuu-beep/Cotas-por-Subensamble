# -*- coding: utf-8 -*-
"""
E2E directo (SIN recorrer el ensamble): solo GENE-FCU-5-118.

  1) Vistas DESPLIEGUE_FRENTE_1 + DESPLIEGUE_LADO
  2) X/Y+TYP (extremos ovalo segun orientacion)
  3) Ø por tipo (DIAMETRO_H / HOLE)
  4) THK en LADO
  5) Inventario de hojas + criterio PASS
"""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

PIEZA = "GENE-FCU-5-118"
kPartDocumentObject = 12290


def _find_part(inv):
    for i in range(1, int(inv.Documents.Count) + 1):
        raw = inv.Documents.Item(i)
        if int(raw.DocumentType) != kPartDocumentObject:
            continue
        if PIEZA.upper() in str(raw.DisplayName).upper():
            return win32com.client.CastTo(raw, "PartDocument")
    return None


def _limpiar_hojas_pieza(plano, pieza: str) -> int:
    pref = pieza.upper()
    n = 0
    for i in range(int(plano.Sheets.Count), 0, -1):
        try:
            h = plano.Sheets.Item(i)
            up = str(h.Name).upper()
            if pref not in up:
                continue
            if "MODELO" in up:
                continue
            h.Delete()
            n += 1
        except Exception:
            continue
    return n


def main() -> int:
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["PYTHONIOENCODING"] = "utf-8"
    os.environ["COTAS_JOB_OVERRIDE"] = "9919-Board 2"

    pythoncom.CoInitialize()
    from inventor_com import conectar_inventor
    from producto_tipo import aplicar_unidad_producto
    from cota_estilo import get_unidad_cota
    import creador_vistas as cv
    import barrenos_xy_despliegue as bx
    import diametro as dm
    import THK

    print("=" * 62)
    print(f" E2E DIRECTO SOLO: {PIEZA}")
    print("=" * 62)

    inv = conectar_inventor()
    tg = inv.TransientGeometry
    to = inv.TransientObjects
    plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")

    for i in range(1, int(inv.Documents.Count) + 1):
        d = inv.Documents.Item(i)
        if int(d.DocumentType) == 12291 and "BOARD" in str(d.DisplayName).upper():
            aplicar_unidad_producto(
                ensamble=win32com.client.CastTo(d, "AssemblyDocument")
            )
            break
    print("Unidad:", get_unidad_cota())

    part = _find_part(inv)
    if part is None:
        print(f"FAIL: abre/documenta {PIEZA}.ipt en Inventor")
        return 2
    print("Parte:", part.DisplayName)

    try:
        sm = win32com.client.CastTo(
            part.ComponentDefinition, "SheetMetalComponentDefinition"
        )
        if not sm.HasFlatPattern:
            sm.Unfold()
        if not sm.HasFlatPattern:
            print("FAIL: sin Flat Pattern")
            return 3
    except Exception as exc:
        print("FAIL sheet metal:", exc)
        return 3

    # Base sheet = machote
    base = None
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        if "MODELO" in str(h.Name).upper():
            base = h
            break
    if base is None:
        base = plano.Sheets.Item(1)

    print("[1] limpiar hojas viejas de la pieza...")
    print("   ", _limpiar_hojas_pieza(plano, PIEZA), "borradas")

    print("[2] crear vistas DESPLIEGUE FRENTE+LADO...")
    nombres = set()
    try:
        cv._crear_vistas_despliegue_corte(
            plano,
            base,
            part,
            PIEZA,
            tg,
            to,
            nombres,
        )
    except Exception as exc:
        print("FAIL vistas:", exc)
        import traceback

        traceback.print_exc()
        return 4
    print("   vistas:", sorted(nombres))
    time.sleep(0.4)

    # Confirmar frente existe
    frente = None
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        up = str(h.Name).upper()
        if PIEZA.upper() in up and "DESPLIEGUE_FRENTE_1" in up:
            if "XCENTRO" in up or "YCENTRO" in up or "DIAMETRO" in up:
                continue
            frente = str(h.Name).rsplit(":", 1)[0]
            break
    if not frente:
        print("FAIL: no hay DESPLIEGUE_FRENTE_1")
        return 5
    print("   frente=", frente)

    print("[3] barrenos X/Y + TYP...")
    xy = bx.acotar_barrenos_xy_despliegue(nombres_frente_ok=[frente])
    print(f"   hojas XY={len(xy)}")
    for n in xy:
        print("    ", n)

    print("[4] Ø por tipo de tamaño...")
    holes = dm.acotar_barrenos_placas(nombres_frente_ok=[frente])
    print(f"   hojas HOLE/DIAMETRO_H={len(holes)}")
    for n in holes:
        print("    ", n)

    print("[5] THK en DESPLIEGUE_LADO...")
    lados = []
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        up = str(h.Name).upper()
        if PIEZA.upper() in up and "DESPLIEGUE_LADO" in up and "THK" not in up:
            lados.append(str(h.Name).rsplit(":", 1)[0])
    thk_nombres = []
    try:
        thk_nombres = THK.acotar_thk(nombres_permitidos=lados or None) or []
    except Exception as exc:
        print("   AVISO THK:", exc)
    print(f"   THK creados/ref={len(thk_nombres) if thk_nombres else 0} lados={lados}")

    # Inventario final
    nx = ny = nh = nth = 0
    hojas = []
    for i in range(1, int(plano.Sheets.Count) + 1):
        nom = str(plano.Sheets.Item(i).Name)
        up = nom.upper()
        if PIEZA.upper() not in up:
            continue
        hojas.append(nom)
        if "XCENTRO" in up:
            nx += 1
        if "YCENTRO" in up:
            ny += 1
        if "DIAMETRO_H" in up:
            nh += 1
        if "THK" in up:
            nth += 1
        # THK.py a menudo acota IN-PLACE sobre DESPLIEGUE_LADO (sin renombrar)
        if "DESPLIEGUE_LADO" in up and "THK" not in up:
            try:
                h = plano.Sheets.Item(i)
                dims = h.DrawingDimensions.GeneralDimensions
                if int(dims.Count) >= 1:
                    nth += 1
            except Exception:
                pass

    print(f"\n=== HOJAS {PIEZA} ===")
    print(f"  XCENTRO={nx}  YCENTRO={ny}  DIAMETRO_H={nh}  THK={nth}  total={len(hojas)}")
    for n in sorted(hojas):
        print(f"   - {n}")

    # Criterio: >=4 X, >=14 Y, >=2 Ø (circulo+oval), >=1 THK
    ok = nx >= 4 and ny >= 14 and nh >= 2 and nth >= 1
    print(
        f"\nCRITERIOS: X>={4}->{nx}  Y>={14}->{ny}  "
        f"HOLE>={2}->{nh}  THK>={1}->{nth}"
    )
    if ok:
        print("PASS E2E DIRECTO — listo para flujo completo de todas las piezas")
        return 0
    print("FAIL E2E DIRECTO — revisar hojas arriba en Inventor")
    return 10


if __name__ == "__main__":
    raise SystemExit(main())
