# -*- coding: utf-8 -*-
import sys
import pythoncom
import win32com.client

sys.path.insert(0, ".")
pythoncom.CoInitialize()
from inventor_com import conectar_inventor

inv = conectar_inventor()
shown = 0
for i in range(1, int(inv.Documents.Count) + 1):
    raw = inv.Documents.Item(i)
    if int(raw.DocumentType) != 12290:
        continue
    try:
        part = win32com.client.CastTo(raw, "PartDocument")
    except Exception as e:
        print("cast fail", raw.DisplayName, e)
        continue
    name = part.DisplayName
    try:
        sub = str(part.SubType)
    except Exception:
        sub = "?"
    try:
        cdef = part.ComponentDefinition
        ctype = str(type(cdef))
    except Exception as e:
        print(name, "no cdef", e)
        continue
    sm = "SheetMetal" in ctype or "Sheet Metal" in sub
    flat = False
    try:
        flat = bool(cdef.HasFlatPattern)
    except Exception as e:
        flat = f"ERR:{e}"
    # count circles on body
    circ = 0
    try:
        for k in range(1, int(cdef.SurfaceBodies.Count) + 1):
            b = cdef.SurfaceBodies.Item(k)
            for j in range(1, min(int(b.Edges.Count), 400) + 1):
                g = b.Edges.Item(j).Geometry
                if g is not None and "CIRCLE" in str(type(g)).upper():
                    circ += 1
    except Exception:
        pass
    if sm or circ > 0 or "GENE" in name.upper() or "ABB" in name.upper() or "BCU" in name.upper():
        print(f"{name} | sub={sub[:40]} | ctype={ctype.split('.')[-1]} | flat={flat} | circ={circ}")
        shown += 1
    if shown >= 40:
        break
print("shown", shown)
