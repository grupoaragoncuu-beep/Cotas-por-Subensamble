# -*- coding: utf-8 -*-
"""
Exporta JPG de TODAS las hojas GENE-FCU-5-118 del plano activo
para validacion visual del usuario.
"""
from __future__ import annotations

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

PIEZA = "GENE-FCU-5-118"
OUT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "_validacion_capturas",
    PIEZA,
)


def main() -> int:
    pythoncom.CoInitialize()
    from inventor_com import conectar_inventor

    inv = conectar_inventor()
    plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
    os.makedirs(OUT, exist_ok=True)

    # Limpiar capturas previas de esta pieza
    for fn in os.listdir(OUT):
        try:
            os.remove(os.path.join(OUT, fn))
        except Exception:
            pass

    hojas = []
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        nom = str(h.Name)
        if PIEZA.upper() not in nom.upper():
            continue
        hojas.append(h)

    # Orden util: frente, x, y, hole, lado/thk
    def _key(h):
        u = str(h.Name).upper()
        if "FRENTE_1" in u and "XCENTRO" not in u and "YCENTRO" not in u and "DIAMETRO" not in u:
            return (0, u)
        if "XCENTRO" in u:
            return (1, u)
        if "YCENTRO" in u:
            return (2, u)
        if "DIAMETRO" in u:
            return (3, u)
        if "LADO" in u or "THK" in u:
            return (4, u)
        return (9, u)

    hojas.sort(key=_key)
    print(f"Exportando {len(hojas)} hojas → {OUT}")
    ok = 0
    for h in hojas:
        base = str(h.Name).rsplit(":", 1)[0]
        # nombre archivo seguro
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in base)
        path = os.path.join(OUT, f"{safe}.jpg")
        try:
            h.Activate()
            time.sleep(0.25)
            try:
                inv.ActiveView.Update()
            except Exception:
                pass
            time.sleep(0.15)
            # Fit para ver cota completa
            try:
                inv.ActiveView.Fit()
            except Exception:
                pass
            time.sleep(0.1)
            inv.ActiveView.Camera.SaveAsBitmap(path, 1920, 1080)
            time.sleep(0.1)
            if os.path.isfile(path) and os.path.getsize(path) > 1000:
                ok += 1
                print(f"  OK  {os.path.basename(path)}")
            else:
                print(f"  FAIL vacio {safe}")
        except Exception as exc:
            print(f"  FAIL {safe}: {exc}")

    print(f"\nLISTO: {ok}/{len(hojas)} capturas en:\n  {OUT}")
    # Abrir carpeta en Explorer
    try:
        os.startfile(OUT)
    except Exception:
        pass
    return 0 if ok >= 4 else 1


if __name__ == "__main__":
    raise SystemExit(main())
