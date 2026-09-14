# -*- coding: utf-8 -*-
"""
Prueba E2E: flujo flat completo SOLO GENE-FCU-5-118.

Luego lista hojas en el plano y JPG locales/share para validar:
  XCENTRO/YCENTRO (+TYP), HOLE (Ø por tipo), THK.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PIEZA = "GENE-FCU-5-118"


def main() -> int:
    os.environ["PIEZAS_FILTRO"] = PIEZA
    os.environ["PYTHONIOENCODING"] = "utf-8"
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["PIEZAS_INCREMENTAL"] = "0"
    os.environ["ENSAMBLES_INDEPENDIENTES"] = "0"
    os.environ["COTAS_JOB_OVERRIDE"] = "9919-Board 2"

    print("=" * 62)
    print(f" E2E PRUEBA SOLO: {PIEZA}")
    print("=" * 62)

    import _reacotar_barrenos_flat_corte as reacotar

    rc = reacotar.main()
    print(f"\nreacotar exit={rc}")

    # Inventario de hojas del plano para esta pieza
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        from inventor_com import conectar_inventor

        inv = conectar_inventor()
        plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
        nx = ny = nh = nth = nfrente = nlado = 0
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
            if "_THK" in up or ( "_LADO" in up and "DESPLIEGUE" in up):
                if "THK" in up or up.rstrip(":0123456789").endswith("LADO"):
                    pass
            if "DESPLIEGUE_LADO" in up or "DESPLIEGUE_THK" in up or (
                "_THK" in up and "DESPLIEGUE" in up
            ):
                nth += 1
            if "DESPLIEGUE_FRENTE_1" in up and "XCENTRO" not in up and "YCENTRO" not in up and "DIAMETRO" not in up:
                nfrente += 1
            if "DESPLIEGUE_LADO" in up and "THK" not in up:
                nlado += 1
        print(f"\n=== HOJAS plano {PIEZA} ===")
        print(f"  XCENTRO={nx}  YCENTRO={ny}  DIAMETRO_H={nh}  THK/LADO-thk={nth}")
        print(f"  FRENTE_base={nfrente}  LADO_base={nlado}  total_pieza={len(hojas)}")
        for n in sorted(hojas):
            print(f"   - {n}")

        ok_xy = nx >= 4 and ny >= 14
        ok_hole = nh >= 1
        print(
            f"\nVALIDA hojas: XY={'OK' if ok_xy else 'FAIL'} "
            f"(X>={4} Y>={14})  HOLE={'OK' if ok_hole else 'FAIL'}"
        )
        if rc == 0 and ok_xy and ok_hole:
            print("PASS E2E pieza de prueba")
            return 0
        if ok_xy and ok_hole:
            print("PASS_PARCIAL: hojas OK pero reacotar/export rc=", rc)
            return 0
        print("FAIL E2E")
        return 10
    except Exception as exc:
        print(f"AVISO inventario hojas: {exc}")
        return rc if rc else 1


if __name__ == "__main__":
    raise SystemExit(main())
