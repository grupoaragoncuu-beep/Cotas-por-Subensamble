# -*- coding: utf-8 -*-
"""Diagnóstico read-only: barrenos visibles en hojas FRENTE del machote."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client
from inventor_com import conectar_inventor
import diametro as d


def main():
    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        doc = None
        for i in range(1, int(inv.Documents.Count) + 1):
            try:
                x = inv.Documents.Item(i)
                if int(x.DocumentType) == 12292:
                    doc = win32com.client.CastTo(x, "DrawingDocument")
                    break
            except Exception:
                continue
        if doc is None:
            print("ERROR: sin drawing")
            return 1
        print(f"Plano: {doc.DisplayName} hojas={doc.Sheets.Count}", flush=True)
        n_frente = 0
        n_con_circ = 0
        n_con_oval = 0
        n_con_any = 0
        muestras = []
        for i in range(1, int(doc.Sheets.Count) + 1):
            hoja = doc.Sheets.Item(i)
            nom = str(hoja.Name)
            up = nom.upper()
            if "_FRENTE_1" not in up and "_FRENTE_2" not in up:
                continue
            if "DIAMETRO" in up or "HOLE" in up:
                continue
            n_frente += 1
            if hoja.DrawingViews.Count < 1:
                continue
            vista = hoja.DrawingViews.Item(1)
            try:
                sc = float(vista.Scale)
                nc = int(vista.DrawingCurves.Count)
            except Exception:
                sc, nc = -1, -1
            circ = d._anillos_en_vista(vista)
            oval = d._ranuras_en_vista(vista)
            allb = d._barrenos_en_vista(vista)
            # crudo: cuantos bbox casi cuadrados / alargados sin filtros duros
            crudo_c = crudo_o = 0
            tipos = {}
            try:
                for j in range(1, nc + 1):
                    c = vista.DrawingCurves.Item(j)
                    try:
                        ct = int(c.CurveType)
                        tipos[ct] = tipos.get(ct, 0) + 1
                    except Exception:
                        pass
                    try:
                        caja = c.Evaluator2D.RangeBox
                        w = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
                        h = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
                        if min(w, h) < 0.04:
                            continue
                        r = max(w, h) / max(min(w, h), 1e-9)
                        if r <= 1.25:
                            crudo_c += 1
                        elif 1.4 <= r <= 8:
                            crudo_o += 1
                    except Exception:
                        pass
            except Exception:
                pass
            if circ:
                n_con_circ += 1
            if oval:
                n_con_oval += 1
            if allb:
                n_con_any += 1
            if len(muestras) < 12:
                muestras.append(
                    {
                        "nom": nom[:60],
                        "scale": round(sc, 4),
                        "curves": nc,
                        "circ": len(circ),
                        "oval": len(oval),
                        "crudo_c": crudo_c,
                        "crudo_o": crudo_o,
                        "tipos": dict(sorted(tipos.items())),
                    }
                )
        print(
            f"FRENTE sheets={n_frente} | con_circ={n_con_circ} | "
            f"con_oval={n_con_oval} | con_barreno_filtrado={n_con_any}",
            flush=True,
        )
        for m in muestras:
            print(f"  {m}", flush=True)
        # hojas DIAMETRO ya creadas
        n_dh = 0
        for i in range(1, int(doc.Sheets.Count) + 1):
            if "DIAMETRO_H" in str(doc.Sheets.Item(i).Name).upper():
                n_dh += 1
        print(f"Hojas DIAMETRO_H* ya en plano: {n_dh}", flush=True)
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
