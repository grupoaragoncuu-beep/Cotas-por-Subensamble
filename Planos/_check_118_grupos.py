# -*- coding: utf-8 -*-
import importlib
import sys

import pythoncom
import win32com.client

sys.path.insert(0, ".")
pythoncom.CoInitialize()

from inventor_com import conectar_inventor
import diametro as dm
import barrenos_xy_despliegue as bx

importlib.reload(dm)
importlib.reload(bx)

inv = conectar_inventor()
tg = inv.TransientGeometry
plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
hoja = [
    plano.Sheets.Item(i)
    for i in range(1, plano.Sheets.Count + 1)
    if "V118" in plano.Sheets.Item(i).Name.upper()
][0]
vista = hoja.DrawingViews.Item(1)
sil = bx._silueta_vista(vista)
ox, _mx, oy, _my, _ = sil
centros = bx._centros_barrenos(vista, tg)
print(
    "fused",
    len(centros),
    "oval",
    sum(1 for c in centros if c.get("tipo") == "oval"),
)
mems_x, mems_y = [], []
for c in centros:
    dx = float(c["cx"]) - ox
    dy = float(c["cy"]) - oy
    tam = float(c.get("tamaño") or 0.2)
    if dx >= 0.05:
        mems_x.append(
            {
                "cx": c["cx"],
                "cy": c["cy"],
                "tamaño": tam,
                "dist_hoja": dx,
                "clave": bx._valor_desde_hoja(vista, hoja, ox, c["cx"]),
                "tipo": c.get("tipo"),
            }
        )
    if dy >= 0.05:
        mems_y.append(
            {
                "cx": c["cx"],
                "cy": c["cy"],
                "tamaño": tam,
                "dist_hoja": dy,
                "clave": bx._valor_desde_hoja(vista, hoja, oy, c["cy"]),
                "tipo": c.get("tipo"),
            }
        )
gx = bx._agrupar_coincidencias(mems_x, vista)
gy = bx._agrupar_coincidencias(mems_y, vista)
print("X", len(gx), "Y", len(gy), "tol", bx._tol_coincidencia_hoja(vista, centros))
for i, g in enumerate(gx, 1):
    tipos = sorted({str(m.get("tipo")) for m in g["miembros"]})
    print(
        f"  X{i} dist={g['dist_hoja']:.5f} n={len(g['miembros'])} "
        f"TYP={g['typ']} clave={g['clave']} tipos={tipos}"
    )
print(
    "Y levels",
    len(gy),
    "all_typ",
    all(g["typ"] for g in gy),
    "ns",
    [len(g["miembros"]) for g in gy],
)
from cota_estilo import texto_cota_dibujo

grupos = dm._agrupar_diametros_grupos(dm._barrenos_en_vista(vista))
print("Ø tipos", len(grupos))
for i, g in enumerate(grupos, 1):
    a = max(g, key=lambda x: x["tamaño"])
    print(
        f"  H{i:02d} n={len(g)} tam={a['tamaño']:.4f} "
        f"tipo={a.get('tipo')} txt={texto_cota_dibujo(a['tamaño'])}"
    )

ok = (
    len(centros) >= 28
    and len(gx) == 4
    and len(gy) == 15
    and all(g["typ"] for g in gx)
    and all(g["typ"] for g in gy)
)
print("PASS" if ok else "FAIL")
raise SystemExit(0 if ok else 10)
