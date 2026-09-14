# -*- coding: utf-8 -*-
"""Diagnostico HLR/modelo de barrenos en hoja V118 o GENE-FCU-5-118 activa."""
from __future__ import annotations

import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pythoncom
import win32com.client

pythoncom.CoInitialize()
from inventor_com import conectar_inventor
import barrenos_xy_despliegue as bx
import diametro as dm

inv = conectar_inventor()
tg = inv.TransientGeometry
plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
hoja = None
for i in range(1, int(plano.Sheets.Count) + 1):
    h = plano.Sheets.Item(i)
    if "V118" in str(h.Name).upper() or "GENE-FCU-5-118" in str(h.Name).upper():
        hoja = h
        break
if hoja is None:
    hoja = plano.ActiveSheet
print("HOJA", hoja.Name)
vista = hoja.DrawingViews.Item(1)
print("Scale", vista.Scale, "W", vista.Width, "H", vista.Height)
sil = bx._silueta_vista(vista)
print("sil", sil)
min_tam = dm._min_tam_hoja(vista)
print("min_tam", min_tam)

ct_counts = Counter()
aspects = []
n = int(vista.DrawingCurves.Count)
print("curves", n)
for j in range(1, n + 1):
    try:
        c = vista.DrawingCurves.Item(j)
        ct = int(c.CurveType)
        ct_counts[ct] += 1
        caja = c.Evaluator2D.RangeBox
        w = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
        h = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
        maj, mi = max(w, h), min(w, h)
        if mi > 1e-9 and maj < 2.0:
            aspects.append((ct, round(w, 4), round(h, 4), round(maj / mi, 2)))
    except Exception:
        continue
print("CurveType counts:", dict(ct_counts))
print("small curves sample (ct,w,h,aspect):")
for a in sorted(aspects, key=lambda t: -t[3])[:40]:
    print(" ", a)

anillos = dm._anillos_en_vista(vista)
ext = dm._arcos_extremos_ranura(vista)
ran = dm._ranuras_en_vista(vista)
eli = bx._elipses_alargadas_hlr(vista)
modelo = bx._centros_barrenos_modelo(vista, tg, sil)
print("anillos", len(anillos), "extremos", len(ext), "ranuras", len(ran), "eli_alarg", len(eli), "modelo", len(modelo))

# modelo raw: count circle/ellipse edges without size filter
doc = bx._doc_vista(vista)
raw_c = raw_e = raw_skip = 0
esc = bx._escala_vista(vista)
lim = 0.48 * max(sil[4], 1e-9)
for body in bx._cuerpos_para_barrenos(doc):
    for j in range(1, int(body.Edges.Count) + 1):
        try:
            g = body.Edges.Item(j).Geometry
            if g is None:
                continue
            t = str(type(g)).upper()
            if "CIRCLE" in t:
                raw_c += 1
                r = float(g.Radius) * esc
                if r < 0.05 or r > lim:
                    raw_skip += 1
                    print(f"  CIRCLE skip r_hoja={r:.4f}")
            elif "ELLIPSE" in t:
                raw_e += 1
                try:
                    rmaj = float(getattr(g, "MajorRadius", 0) or 0)
                    rmin = float(getattr(g, "MinorRadius", 0) or 0)
                except Exception:
                    rmaj = rmin = 0
                r = min(rmaj, rmin) if rmin > 0 else rmaj
                rh = r * esc
                print(f"  ELLIPSE maj={rmaj:.4f} min={rmin:.4f} r_hoja={rh:.4f} skip={rh<0.05 or rh>lim}")
        except Exception:
            continue
print(f"raw edges circle={raw_c} ellipse={raw_e} skipped_size={raw_skip}")
