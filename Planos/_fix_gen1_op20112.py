# -*- coding: utf-8 -*-
"""Re-acota GEN1-OP-20-112: solo barrenos de loops interiores."""
from __future__ import annotations

import os
import re
import shutil
import sys

import pythoncom

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

SHARE = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\9919-BOARD2_2\Corte\Corte"
)
SHARE2 = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS\Corte\Corte"
)
JOB = "9919-Board 2"
PIEZA = "GEN1-OP-20-112"
RE6 = re.compile(r"\.\d{6}(?:_|\.|$)")
TOK = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")


def main() -> int:
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    from cota_estilo import set_unidad_cota

    set_unidad_cota("mm")
    pythoncom.CoInitialize()
    try:
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS
        from generador_vistas import exportar_hojas_jpg
        from barrenos_xy_despliegue import acotar_barrenos_xy_despliegue
        import diametro

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ens = _obtener_ensamble_principal(inv)
        if plano is None:
            print("ERROR: sin plano")
            return 2

        # Borrar hojas XY/HOLE viejas de esta pieza (conservar THK)
        for i in range(int(plano.Sheets.Count), 0, -1):
            try:
                h = plano.Sheets.Item(i)
                up = str(h.Name).upper()
            except Exception:
                continue
            if PIEZA.upper() not in up or "_DESPLIEGUE_" not in up:
                continue
            if any(t in up for t in ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "DIAMETRO_H")):
                try:
                    print("DEL", h.Name)
                    h.Delete()
                except Exception as e:
                    print("  no del", e)

        frentes = []
        for i in range(1, int(plano.Sheets.Count) + 1):
            up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
            if PIEZA.upper() in up and "_DESPLIEGUE_FRENTE_1" in up:
                frentes.append(up)
        print("frentes", frentes)
        set_unidad_cota("mm")
        xy = acotar_barrenos_xy_despliegue(frentes)
        print("XY creadas", xy)
        holes = diametro.acotar_barrenos_placas(frentes)
        print("HOLE creadas", holes)

        permit = set()
        for i in range(1, int(plano.Sheets.Count) + 1):
            up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
            if PIEZA.upper() not in up or "_DESPLIEGUE_" not in up:
                continue
            if any(
                t in up
                for t in ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "DIAMETRO_H", "_THK")
            ):
                permit.add(up)
        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ens), CARPETA_PIEZAS_ACOTADAS
        )
        exportar_hojas_jpg(
            inv, plano, carpeta_salida=carpeta, nombres_permitidos=permit, nombre_job=JOB
        )

        st = os.path.join(carpeta, "_STAGING_DESPLIEGUE")
        archs = []
        for fn in os.listdir(st):
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if PIEZA not in fn:
                continue
            if not any(t in fn.upper() for t in TOK):
                continue
            if not RE6.search(fn):
                continue
            if "_DESPLIEGUE_" in fn.upper() and "__LENGTH_" in fn.upper():
                continue
            if fn.split("__")[1] != PIEZA:
                continue
            archs.append(os.path.join(st, fn))
        ups = [os.path.basename(a).upper() for a in archs]
        ok = any("XCENTRO" in u or "YCENTRO" in u or "XMIN" in u for u in ups) and any(
            "__THK_" in u for u in ups
        )
        print("staging", len(archs), "ok", ok)
        for a in sorted(archs):
            print(" ", os.path.basename(a))
        if not ok:
            return 3
        for share in (SHARE, SHARE2):
            dst = os.path.join(share, PIEZA)
            os.makedirs(dst, exist_ok=True)
            for fn in list(os.listdir(dst)):
                if fn.lower().endswith((".jpg", ".jpeg", ".png")) and any(
                    t in fn.upper() for t in TOK
                ):
                    try:
                        os.remove(os.path.join(dst, fn))
                    except OSError:
                        pass
            for src in archs:
                shutil.copy2(src, os.path.join(dst, os.path.basename(src)))
        print("SUBE", PIEZA, len(archs))
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
