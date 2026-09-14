# -*- coding: utf-8 -*-
"""
Smoke COM correcto:
  - Vista FLAT (SheetMetalFoldedModel=False)
  - Cámara a la CARA PRINCIPAL (elegir_frente), NO al canto/THK
  - Cotas X/Y solo sobre esa cara
"""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

kPartDocumentObject = 12290


def _as_part(doc):
    return win32com.client.CastTo(doc, "PartDocument")


def _es_chapa_flatable(part) -> bool:
    try:
        sm = win32com.client.CastTo(
            part.ComponentDefinition, "SheetMetalComponentDefinition"
        )
        if not sm.HasFlatPattern:
            try:
                sm.Unfold()
            except Exception:
                return False
        return bool(sm.HasFlatPattern)
    except Exception:
        return False


def _buscar_chapa(inv):
    prefer = ("GENE-FCU", "GENE-BCU", "GENE-OP", "ABB", "GENE-BKT")
    ranked = []
    for i in range(1, int(inv.Documents.Count) + 1):
        raw = inv.Documents.Item(i)
        if int(raw.DocumentType) != kPartDocumentObject:
            continue
        try:
            part = _as_part(raw)
        except Exception:
            continue
        if not _es_chapa_flatable(part):
            continue
        name = str(part.DisplayName).upper()
        score = 1
        for idx, pref in enumerate(prefer):
            if pref in name:
                score += 1000 - idx * 10
                break
        ranked.append((score, part.DisplayName, part))
        if len(ranked) >= 40:
            break
    if not ranked:
        return None
    ranked.sort(key=lambda t: -t[0])
    print("Chapas flatables:")
    for r in ranked[:6]:
        print(f"  {r[1]} score={r[0]}")
    return ranked[0][2]


def _bbox_aspect_ok(vista) -> bool:
    """Cara principal: ambos lados del bbox 2D razonables (no canto fino)."""
    try:
        w = float(vista.Width)
        h = float(vista.Height)
        if w <= 0 or h <= 0:
            return False
        ratio = max(w, h) / min(w, h)
        # Canto/THK suele ser ratio enorme (>8). Cara flat típica < 6-8.
        print(f"  bbox vista W={w:.2f} H={h:.2f} ratio={ratio:.2f}")
        return ratio < 10.0 and min(w, h) > 0.8
    except Exception:
        return False


def main() -> int:
    pythoncom.CoInitialize()
    from inventor_com import conectar_inventor
    import creador_vistas as cv
    import barrenos_xy_despliegue as bx

    print("=== SMOKE COM: FRENTE FLAT (no THK/canto) ===")
    print("NOTA: esta prueba usa IN solo para demo; producción Board = mm")
    inv = conectar_inventor()
    from cota_estilo import set_unidad_cota

    set_unidad_cota("in")  # solo esta smoke
    tg = inv.TransientGeometry
    to = inv.TransientObjects

    try:
        plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
    except Exception as e:
        print("FAIL machote:", e)
        return 2

    part = _buscar_chapa(inv)
    if part is None:
        print("FAIL: no hay sheet-metal con Flat Pattern abierta")
        print("Abre una pieza flat de Corte (chapa) y reintenta.")
        return 3
    print("Pieza:", part.DisplayName)

    # Misma geometría/orientación que producción DESPLIEGUE_FRENTE_1
    res_flat = cv.preparar_geometria_flat(part, True, to)
    if not res_flat:
        print("FAIL: preparar_geometria_flat")
        return 4
    _fp, caras, cuerpo = res_flat
    res_frente = cv.elegir_frente(caras)
    if not res_frente:
        print("FAIL: elegir_frente")
        return 5
    frente_face, v_frente, _area = res_frente
    res_lado = cv.elegir_lado(caras, frente_face, v_frente, _area, True)
    if res_lado:
        v_lado = res_lado[1]
    else:
        v_lado = cv.obtener_lado_fallback(tg, v_frente)
    tiene_guia, v_guia = cv.obtener_vector_guia_frente(frente_face, v_frente, tg)
    cx, cy, cz = cv.obtener_centro(part, cuerpo)
    up_f = v_guia if tiene_guia else v_lado
    cam = cv.crear_camara(part, tg, to, cx, cy, cz, v_frente, up_f)

    # Limpiar smokes previos
    for i in range(plano.Sheets.Count, 0, -1):
        h = plano.Sheets.Item(i)
        if str(h.Name).upper().startswith("SMOKE_"):
            try:
                h.Delete()
            except Exception:
                pass

    base = plano.Sheets.Item(plano.Sheets.Count)
    hoja = base.CopyTo(plano)
    try:
        hoja.Name = "SMOKE_DESPLIEGUE_FRENTE_1"
    except Exception:
        pass
    hoja.Activate()
    for i in range(hoja.DrawingViews.Count, 0, -1):
        try:
            hoja.DrawingViews.Item(i).Delete()
        except Exception:
            pass

    px = float(hoja.Width) * 0.5
    py = float(hoja.Height) * 0.55
    vista = cv._crear_vista_base(
        hoja, part, tg, to, px, py, cam, use_flat_pattern_view=True
    )
    if vista is None:
        print("FAIL: AddBaseView flat")
        return 6
    try:
        cv.escalar_vista(plano, vista, tg, px, py, 18.0, 22.0)
    except Exception:
        pass
    try:
        vista.Update()
        inv.ActiveView.Update()
    except Exception:
        pass
    time.sleep(0.6)

    if not _bbox_aspect_ok(vista):
        print("FAIL: la vista sigue pareciendo canto/THK (bbox)")
        return 7

    try:
        ncurv = int(vista.DrawingCurves.Count)
    except Exception:
        ncurv = -1
    print(f"Vista flat frente OK curvas={ncurv}")

    sil = bx._silueta_vista(vista)
    if not sil:
        print("FAIL: sin silueta")
        return 8
    ox, _mx, oy, _my, span = sil
    print(f"Origen IL=({ox:.3f},{oy:.3f}) span={span:.3f}")

    centros = bx._centros_barrenos(vista, tg)
    print(
        f"Centros={len(centros)} fuente={centros[0].get('fuente') if centros else '-'}"
    )
    if not centros:
        print("FAIL: 0 barrenos en cara flat")
        return 9

    c0 = centros[0]
    cx2, cy2 = float(c0["cx"]), float(c0["cy"])
    bx._limpiar_dims_y_sketches(hoja)
    ok_x = bx._dibujar_cota_centro_sketch(
        hoja, vista, tg, inv, "X", ox, oy, cx2, cy2, "TEST_X", miembros_typ=centros
    )
    ok_y = bx._dibujar_cota_centro_sketch(
        hoja, vista, tg, inv, "Y", ox, oy, cx2, cy2, "TEST_Y", miembros_typ=centros
    )
    print("OK_X=", ok_x, "OK_Y=", ok_y)

    n_ln = n_tb = 0
    try:
        sketches = hoja.Sketches
    except Exception:
        sketches = hoja.DrawingSketches
    for i in range(1, int(sketches.Count) + 1):
        sk = sketches.Item(i)
        try:
            n_ln += int(sk.SketchLines.Count)
        except Exception:
            pass
        try:
            n_tb += int(sk.TextBoxes.Count)
        except Exception:
            pass
    print(f"Lineas={n_ln} TextBoxes={n_tb}")
    try:
        hoja.Activate()
        inv.ActiveView.Update()
    except Exception:
        pass

    if ok_x and ok_y and (n_ln >= 2 or n_tb >= 1):
        print("PASS: cotas en CARA FLAT FRENTE (no THK). Hoja SMOKE_DESPLIEGUE_FRENTE_1")
        return 0
    print("FAIL: sin cotas")
    return 10


if __name__ == "__main__":
    raise SystemExit(main())
