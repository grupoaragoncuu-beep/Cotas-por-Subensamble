# -*- coding: utf-8 -*-
"""Re-acota Ø GENE-FCU-5-118 con precision 3 dec y reexporta JPG."""
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
    from producto_tipo import aplicar_unidad_producto
    from cota_estilo import (
        set_unidad_cota,
        get_unidad_cota,
        texto_cota_limpio,
        aplicar_estilo_cota,
    )
    import diametro as dm

    inv = conectar_inventor()
    plano = None
    for i in range(1, int(inv.Documents.Count) + 1):
        d = inv.Documents.Item(i)
        try:
            if int(d.DocumentType) == 12292:  # Drawing
                plano = win32com.client.CastTo(d, "DrawingDocument")
                break
        except Exception:
            continue
    if plano is None:
        try:
            plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
        except Exception:
            print("FAIL: no hay plano abierto")
            return 2
    try:
        plano.Activate()
    except Exception:
        pass

    for i in range(1, int(inv.Documents.Count) + 1):
        d = inv.Documents.Item(i)
        if int(d.DocumentType) == 12291 and "BOARD" in str(d.DisplayName).upper():
            aplicar_unidad_producto(
                ensamble=win32com.client.CastTo(d, "AssemblyDocument")
            )
            break
    else:
        set_unidad_cota("mm")
    print("Unidad:", get_unidad_cota())

    # smoke: 0.433 in = 1.09982 cm → mm
    demo = texto_cota_limpio(0.433 * 2.54)
    print(f"demo 0.433in → '{demo}' (debe ~10.998, NUNCA 11)")
    frente = None
    for i in range(1, int(plano.Sheets.Count) + 1):
        n = str(plano.Sheets.Item(i).Name)
        up = n.upper()
        if PIEZA.upper() in up and "DESPLIEGUE_FRENTE_1" in up:
            if "XCENTRO" in up or "YCENTRO" in up or "DIAMETRO" in up:
                continue
            frente = n.rsplit(":", 1)[0]
            break
    if not frente:
        print("FAIL: sin FRENTE_1")
        return 2
    print("frente", frente)

    # borrar DIAMETRO_H viejos de la pieza
    for i in range(int(plano.Sheets.Count), 0, -1):
        h = plano.Sheets.Item(i)
        up = str(h.Name).upper()
        if PIEZA.upper() in up and "DIAMETRO_H" in up:
            try:
                h.Delete()
            except Exception:
                pass

    creadas = dm.acotar_barrenos_placas(nombres_frente_ok=[frente])
    print("creadas", creadas)

    # Reaplicar estilo por si acaso + listar textos
    os.makedirs(OUT, exist_ok=True)
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        nom = str(h.Name)
        up = nom.upper()
        if PIEZA.upper() not in up or "DIAMETRO_H" not in up:
            continue
        h.Activate()
        time.sleep(0.2)
        try:
            dims = h.DrawingDimensions.GeneralDimensions
            for k in range(1, int(dims.Count) + 1):
                dim = dims.Item(k)
                try:
                    aplicar_estilo_cota(dim, hoja=h)
                except Exception:
                    pass
                try:
                    print("  texto:", dim.Text.Text, "model=", dim.ModelValue)
                except Exception:
                    pass
        except Exception as exc:
            print("  aviso dims:", exc)
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in nom.rsplit(":", 1)[0])
        path = os.path.join(OUT, f"{safe}.jpg")
        try:
            inv.ActiveView.Fit()
            time.sleep(0.1)
            inv.ActiveView.Camera.SaveAsBitmap(path, 1920, 1080)
            print("  JPG", path)
        except Exception as exc:
            print("  fail jpg", exc)

    # También re-estilar X/Y sketch? usuario dijo resto OK; solo Ø.
    print("LISTO precision Ø")
    try:
        os.startfile(OUT)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
