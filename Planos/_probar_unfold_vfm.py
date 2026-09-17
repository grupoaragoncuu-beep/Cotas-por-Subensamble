# -*- coding: utf-8 -*-
"""Prueba rápida: espesores + Unfold2 hasta bends>0 en una pieza VFM."""
from __future__ import annotations

import sys
from collections import Counter

import win32com.client

from inventor_com import conectar_inventor, localizar_documento

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

K_PLANE = (5890, 17921)
TARGET = "gene-vfm-20-101"


def log(m=""):
    print(m, flush=True)


def parallel_candidates(smd):
    body = smd.SurfaceBodies.Item(1)
    planes = []
    for i in range(1, body.Faces.Count + 1):
        f = body.Faces.Item(i)
        try:
            if int(f.SurfaceType) not in K_PLANE:
                continue
            n = f.Geometry.Normal
            pt = f.PointOnFace
            planes.append((n.X, n.Y, n.Z, pt.X, pt.Y, pt.Z, float(f.Evaluator.Area)))
        except Exception:
            continue
    dists = []
    for i, a in enumerate(planes):
        ax, ay, az, px, py, pz, aa = a
        for b in planes[i + 1 :]:
            bx, by, bz, qx, qy, qz, ba = b
            if abs(ax * bx + ay * by + az * bz + 1) > 0.05:
                continue
            dist = abs((qx - px) * ax + (qy - py) * ay + (qz - pz) * az)
            if 0.08 <= dist <= 2.6 and min(aa, ba) > 0.2:
                dists.append(round(dist, 4))
    cnt = Counter(round(d, 3) for d in dists)
    # Preferir espesores chicos con muchas repeticiones (chapa real)
    ranked = sorted(cnt.items(), key=lambda kv: (-kv[1], kv[0]))
    return ranked, planes


def plane_faces(smd, n=20):
    body = smd.SurfaceBodies.Item(1)
    out = []
    for i in range(1, body.Faces.Count + 1):
        f = body.Faces.Item(i)
        try:
            if int(f.SurfaceType) not in K_PLANE:
                continue
            out.append((float(f.Evaluator.Area), f))
        except Exception:
            continue
    out.sort(key=lambda x: -x[0])
    return out[:n]


def main():
    inv = conectar_inventor()
    raw = localizar_documento(inv, display_name="9919-Board 5.iam") or inv.ActiveDocument
    asm = win32com.client.CastTo(raw, "AssemblyDocument")
    leaves = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
    fp = None
    for i in range(1, int(leaves.Count) + 1):
        try:
            d = leaves.Item(i).Definition.Document
            if TARGET in (d.FullFileName or "").lower():
                fp = d.FullFileName
                break
        except Exception:
            continue
    if not fp:
        log("NO ENCONTRADA")
        return 1

    part = win32com.client.CastTo(inv.Documents.Open(fp, True), "PartDocument")
    try:
        part.Activate()
    except Exception:
        pass
    smd = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
    try:
        if smd.HasFlatPattern:
            smd.FlatPattern.Delete()
    except Exception:
        pass

    ranked, _ = parallel_candidates(smd)
    rb = smd.SurfaceBodies.Item(1).RangeBox
    dims = sorted(
        [
            abs(rb.MaxPoint.X - rb.MinPoint.X),
            abs(rb.MaxPoint.Y - rb.MinPoint.Y),
            abs(rb.MaxPoint.Z - rb.MinPoint.Z),
        ]
    )
    log(f"pieza={TARGET} dims={[round(x,3) for x in dims]}")
    log(f"distancias={ranked[:10]}")

    thks = []
    for d, n in ranked[:8]:
        thks.append(float(d))
    thks.append(float(round(dims[0], 3)))
    # calibres cobre típicos (cm)
    for t in (0.3175, 0.47625, 0.635, 0.9525, 1.27, 0.25, 0.3, 0.6, 0.9):
        thks.append(t)
    # únicos ordenados asc (probar chapa fina primero)
    seen = set()
    uniq = []
    for t in sorted(thks):
        k = round(t, 3)
        if k in seen or k < 0.08 or k > 2.6:
            continue
        seen.add(k)
        uniq.append(t)

    faces = plane_faces(smd, 25)
    log(f"caras_planas={len(faces)} thk_candidatos={ [round(t,3) for t in uniq[:12]] }")

    for thk in uniq[:12]:
        try:
            smd.UseSheetMetalStyleThickness = False
            smd.Thickness.Value = float(thk)
            try:
                smd.BendRadius.Value = max(thk, 0.1)
            except Exception:
                pass
            part.Update()
        except Exception as ex:
            log(f"  thk={thk:.4f} set FAIL {ex}")
            continue

        # refrescar caras tras Update
        faces = plane_faces(smd, 25)
        for area, face in faces:
            try:
                if smd.HasFlatPattern:
                    smd.FlatPattern.Delete()
                smd.Unfold2(face)
                if not smd.HasFlatPattern:
                    continue
                bends = 0
                try:
                    bends = int(smd.FlatPattern.FlatBendResults.Count)
                except Exception:
                    pass
                if bends > 0:
                    frb = smd.FlatPattern.Body.RangeBox
                    fd = sorted(
                        [
                            abs(frb.MaxPoint.X - frb.MinPoint.X),
                            abs(frb.MaxPoint.Y - frb.MinPoint.Y),
                            abs(frb.MaxPoint.Z - frb.MinPoint.Z),
                        ]
                    )
                    log(
                        f"OK REAL thk={thk:.4f} area={area:.2f} bends={bends} "
                        f"flat={[round(x,2) for x in fd]}"
                    )
                    part.Save()
                    part.Close(True)
                    return 0
                try:
                    smd.FlatPattern.Delete()
                except Exception:
                    pass
            except Exception:
                try:
                    if smd.HasFlatPattern:
                        smd.FlatPattern.Delete()
                except Exception:
                    pass
        log(f"  thk={thk:.4f} sin bends")

    log("FAIL sin unfold real")
    try:
        if smd.HasFlatPattern:
            smd.FlatPattern.Delete()
    except Exception:
        pass
    part.Close(True)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
