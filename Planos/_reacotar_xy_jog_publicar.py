# -*- coding: utf-8 -*-
"""
Re-acota X/Y incluyendo barrenos a la izquierda/abajo del origen IL (jog/S),
re-exporta JPG y publica al dossier solo piezas OK.
"""
from __future__ import annotations

import os
import sys

import pythoncom

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

os.environ["SOLO_FLAT_CORTE"] = "1"
os.environ["COTAS_JOB_OVERRIDE"] = "9919-Board 2"


def main() -> int:
    from cota_estilo import set_unidad_cota

    set_unidad_cota("mm")
    pythoncom.CoInitialize()
    try:
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS
        from generador_vistas import exportar_hojas_jpg
        from barrenos_xy_despliegue import acotar_barrenos_xy_despliegue
        from producto_tipo import aplicar_unidad_producto

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            print("ERROR: machote + ensamble")
            return 2
        try:
            aplicar_unidad_producto(ensamble=ensamble)
        except Exception:
            pass
        set_unidad_cota("mm")
        try:
            h = _encontrar_hoja_machote(plano)
            if h:
                h.Activate()
        except Exception:
            pass

        print("[1] re-acotar XY (incluye barrenos dx/dy < 0)…")
        creadas = acotar_barrenos_xy_despliegue(None)
        print(f"    hojas tocadas/creadas≈{len(creadas) if creadas else 0}")

        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        os.makedirs(carpeta, exist_ok=True)
        print("[2] export JPG…")
        # Exportar solo hojas DESPLIEGUE X/Y/THK/HOLE
        permitidos = set()
        for i in range(1, int(plano.Sheets.Count) + 1):
            try:
                up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
            except Exception:
                continue
            if "_DESPLIEGUE_" not in up:
                continue
            if any(
                t in up
                for t in ("XCENTRO", "YCENTRO", "_THK", "_LADO", "DIAMETRO_H")
            ):
                permitidos.add(up)
        exportar_hojas_jpg(
            inv,
            plano,
            carpeta_salida=carpeta,
            nombres_permitidos=permitidos,
            nombre_job="9919-Board 2",
        )
        print("[3] publicar OK inmediato…")
    finally:
        pythoncom.CoUninitialize()

    # Publicar en proceso aparte limpio
    import subprocess

    r = subprocess.call(
        [sys.executable, "-u", os.path.join(ROOT, "_publicar_ok_inmediato.py")],
        cwd=ROOT,
        env={**os.environ, "SOLO_FLAT_CORTE": "1", "PYTHONUNBUFFERED": "1"},
    )
    return int(r)


if __name__ == "__main__":
    raise SystemExit(main())
