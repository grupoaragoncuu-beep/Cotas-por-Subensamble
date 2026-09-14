# -*- coding: utf-8 -*-
"""
Validacion exacta GENE-FCU-5-118:
  espera 4 coincidencias X (TYP) y 15 coincidencias Y (TYP).
"""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

kPartDocumentObject = 12290
TARGET = "GENE-FCU-5-118"
EXPECT_X = 4
EXPECT_Y = 15
EXPECT_HOLES = 30


def _find_part(inv):
    for i in range(1, int(inv.Documents.Count) + 1):
        raw = inv.Documents.Item(i)
        if int(raw.DocumentType) != kPartDocumentObject:
            continue
        name = str(raw.DisplayName).upper()
        if TARGET in name:
            return win32com.client.CastTo(raw, "PartDocument")
    # abrir desde ensamble / busqueda en docs
    for i in range(1, int(inv.Documents.Count) + 1):
        raw = inv.Documents.Item(i)
        try:
            if int(raw.DocumentType) != 12291:
                continue
            asm = win32com.client.CastTo(raw, "AssemblyDocument")
            occs = asm.ComponentDefinition.Occurrences
            for j in range(1, int(occs.Count) + 1):
                occ = occs.Item(j)
                dn = str(occ.Name).upper()
                if TARGET not in dn:
                    continue
                try:
                    return win32com.client.CastTo(
                        occ.Definition.Document, "PartDocument"
                    )
                except Exception:
                    pass
        except Exception:
            continue
    return None


def main() -> int:
    pythoncom.CoInitialize()
    from inventor_com import conectar_inventor
    from cota_estilo import set_unidad_cota, get_unidad_cota
    from producto_tipo import aplicar_unidad_producto
    import creador_vistas as cv
    import barrenos_xy_despliegue as bx

    print(f"=== VALIDACION EXACTA {TARGET}: {EXPECT_X}X + {EXPECT_Y}Y TYP ===")
    inv = conectar_inventor()
    tg = inv.TransientGeometry
    to = inv.TransientObjects

    asm = None
    for i in range(1, int(inv.Documents.Count) + 1):
        d = inv.Documents.Item(i)
        if int(d.DocumentType) == 12291 and "BOARD" in str(d.DisplayName).upper():
            asm = win32com.client.CastTo(d, "AssemblyDocument")
            break
    if asm is not None:
        aplicar_unidad_producto(ensamble=asm)
    else:
        set_unidad_cota("mm")
    print("Unidad:", get_unidad_cota())

    part = _find_part(inv)
    if part is None:
        print(f"FAIL: no esta abierta {TARGET}")
        return 2
    print("USANDO:", part.DisplayName)

    try:
        sm = win32com.client.CastTo(
            part.ComponentDefinition, "SheetMetalComponentDefinition"
        )
        if not sm.HasFlatPattern:
            sm.Unfold()
    except Exception as exc:
        print("FAIL flat:", exc)
        return 3

    plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
    for i in range(plano.Sheets.Count, 0, -1):
        h = plano.Sheets.Item(i)
        if str(h.Name).upper().startswith("V118_"):
            try:
                h.Delete()
            except Exception:
                pass

    res_flat = cv.preparar_geometria_flat(part, True, to)
    if not res_flat:
        print("FAIL preparar flat")
        return 4
    _fp, caras, cuerpo = res_flat
    res_frente = cv.elegir_frente(caras)
    if not res_frente:
        print("FAIL frente")
        return 5
    frente_face, v_frente, area = res_frente
    res_lado = cv.elegir_lado(caras, frente_face, v_frente, area, True)
    v_lado = res_lado[1] if res_lado else cv.obtener_lado_fallback(tg, v_frente)
    tiene_guia, v_guia = cv.obtener_vector_guia_frente(frente_face, v_frente, tg)
    cx0, cy0, cz0 = cv.obtener_centro(part, cuerpo)
    up = v_guia if tiene_guia else v_lado
    cam = cv.crear_camara(part, tg, to, cx0, cy0, cz0, v_frente, up)

    base = plano.Sheets.Item(plano.Sheets.Count)
    hoja = base.CopyTo(plano)
    try:
        hoja.Name = "V118_DESPLIEGUE_FRENTE_1"
    except Exception:
        pass
    hoja.Activate()
    for i in range(hoja.DrawingViews.Count, 0, -1):
        try:
            hoja.DrawingViews.Item(i).Delete()
        except Exception:
            pass
    px, py = float(hoja.Width) * 0.5, float(hoja.Height) * 0.55
    vista = cv._crear_vista_base(
        hoja, part, tg, to, px, py, cam, use_flat_pattern_view=True
    )
    try:
        cv.escalar_vista(plano, vista, tg, px, py, 18.0, 22.0)
        vista.Update()
        inv.ActiveView.Update()
    except Exception:
        pass
    time.sleep(0.6)

    sil = bx._silueta_vista(vista)
    if not sil:
        print("FAIL silueta")
        return 6
    ox, _mx, oy, _my, span = sil
    print(f"Origen IL=({ox:.4f},{oy:.4f}) span={span:.4f}")
    print(f"Vista W={float(vista.Width):.2f} H={float(vista.Height):.2f} Scale={vista.Scale}")

    modelo = bx._centros_barrenos_modelo(vista, tg, sil)
    hlr = bx._centros_barrenos_hlr(vista)
    centros = bx._centros_barrenos(vista, tg)
    print(f"modelo={len(modelo)} hlr={len(hlr)} fused={len(centros)}")

    # listar X/Y unicos con alta precision
    xs = sorted({round(float(c["cx"]) - ox, 5) for c in centros})
    ys = sorted({round(float(c["cy"]) - oy, 5) for c in centros})
    print(f"X raw unicos (5dec)={len(xs)}  Y raw unicos={len(ys)}")
    for i, x in enumerate(xs, 1):
        n = sum(1 for c in centros if abs((float(c["cx"]) - ox) - x) < 1e-4)
        print(f"  Xraw{i}: {x:.5f} cm n~{n}")
    for i, y in enumerate(ys, 1):
        n = sum(1 for c in centros if abs((float(c["cy"]) - oy) - y) < 1e-4)
        print(f"  Yraw{i}: {y:.5f} cm n~{n}")

    mems_x, mems_y = [], []
    for c in centros:
        dx = float(c["cx"]) - ox
        dy = float(c["cy"]) - oy
        tam = float(c.get("tamaño") or 0.2)
        if dx >= 0.05:
            mems_x.append(
                {
                    "cx": float(c["cx"]),
                    "cy": float(c["cy"]),
                    "tamaño": tam,
                    "dist_hoja": dx,
                    "clave": bx._valor_desde_hoja(vista, hoja, ox, float(c["cx"])),
                    "tipo": c.get("tipo"),
                }
            )
        if dy >= 0.05:
            mems_y.append(
                {
                    "cx": float(c["cx"]),
                    "cy": float(c["cy"]),
                    "tamaño": tam,
                    "dist_hoja": dy,
                    "clave": bx._valor_desde_hoja(vista, hoja, oy, float(c["cy"])),
                    "tipo": c.get("tipo"),
                }
            )

    gx = bx._agrupar_coincidencias(mems_x, vista)
    gy = bx._agrupar_coincidencias(mems_y, vista)
    tol = bx._tol_coincidencia_hoja(vista, centros)
    print(f"tol={tol:.5f} cm hoja")
    print(f"GRUPOS X={len(gx)} Y={len(gy)} barrenos={len(centros)}")
    for i, g in enumerate(gx, 1):
        tipos = sorted({str(m.get("tipo") or "?") for m in g["miembros"]})
        print(
            f"  X{i}: dist={g['dist_hoja']:.5f} n={len(g['miembros'])} "
            f"TYP={g['typ']} clave={g['clave']} tipos={tipos}"
        )
    for i, g in enumerate(gy, 1):
        print(
            f"  Y{i}: dist={g['dist_hoja']:.5f} n={len(g['miembros'])} "
            f"TYP={g['typ']} clave={g['clave']}"
        )

    # Flujo real multi-hoja
    creadas = bx.acotar_barrenos_xy_despliegue(
        nombres_frente_ok=["V118_DESPLIEGUE_FRENTE_1"]
    )
    nx = sum(1 for n in creadas if "XCENTRO" in n.upper())
    ny = sum(1 for n in creadas if "YCENTRO" in n.upper())
    print(f"HOJAS flujo: total={len(creadas)} X={nx} Y={ny}")
    for n in creadas:
        print(" ", n)

    ok_count = len(centros) >= EXPECT_HOLES - 2  # margen 2 por HLR
    ok_x = len(gx) == EXPECT_X and all(g["typ"] for g in gx)
    ok_y = len(gy) == EXPECT_Y and all(g["typ"] for g in gy)
    ok_hojas = nx == EXPECT_X and ny == EXPECT_Y

    if ok_count and ok_x and ok_y and ok_hojas:
        print(
            f"PASS: {len(centros)} barrenos → {EXPECT_X} X TYP + {EXPECT_Y} Y TYP"
        )
        return 0

    print(
        f"FAIL: barrenos={len(centros)} (esperaba~{EXPECT_HOLES}) "
        f"X={len(gx)}/{EXPECT_X} typ_ok={ok_x} "
        f"Y={len(gy)}/{EXPECT_Y} typ_ok={ok_y} "
        f"hojas X={nx} Y={ny}"
    )
    return 10


if __name__ == "__main__":
    raise SystemExit(main())
