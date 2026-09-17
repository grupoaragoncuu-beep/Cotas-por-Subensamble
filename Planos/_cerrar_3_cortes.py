# -*- coding: utf-8 -*-
"""Cierra las 3 piezas con cortes internos no circulares (XMIN/YMIN + THK)."""
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
PIEZAS = ("GENE-GS-0820-708", "GENE-OP-10-116", "GENE-OP-10-117")
RE6 = re.compile(r"\.\d{6}(?:_|\.|$)")
TOK = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")


def _staging() -> str:
    return os.path.join(
        ROOT, "JPG", "9919-Board 1", "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE"
    )


def _completo(archs: list[str]) -> bool:
    ups = [os.path.basename(a).upper() for a in archs]
    xy = any(any(t in u for t in ("XCENTRO", "YCENTRO", "XMIN", "YMIN")) for u in ups)
    thk = any("__THK_" in u for u in ups)
    return xy and thk


def main() -> int:
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["PYTHONUNBUFFERED"] = "1"
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

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None:
            print("ERROR: sin plano")
            return 2

        # Solo FRENTE de las 3 piezas
        frentes = []
        for i in range(1, int(plano.Sheets.Count) + 1):
            h = plano.Sheets.Item(i)
            up = str(h.Name).upper().rsplit(":", 1)[0]
            if "_DESPLIEGUE_FRENTE_1" not in up:
                continue
            if any(p.upper() in up for p in PIEZAS):
                frentes.append(up)
        print("frentes:", frentes)
        set_unidad_cota("mm")
        creadas = acotar_barrenos_xy_despliegue(frentes)
        print("creadas XY:", len(creadas or []), creadas)

        permitidos = set()
        for i in range(1, int(plano.Sheets.Count) + 1):
            up = str(plano.Sheets.Item(i).Name).upper().rsplit(":", 1)[0]
            if not any(p.upper() in up for p in PIEZAS):
                continue
            if "_DESPLIEGUE_" not in up:
                continue
            if any(
                t in up
                for t in (
                    "XCENTRO",
                    "YCENTRO",
                    "XMIN",
                    "YMIN",
                    "_THK",
                    "_LADO",
                )
            ):
                permitidos.add(up)
        print("export permitidos:", sorted(permitidos))
        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        os.makedirs(carpeta, exist_ok=True)
        if permitidos:
            exportar_hojas_jpg(
                inv,
                plano,
                carpeta_salida=carpeta,
                nombres_permitidos=permitidos,
                nombre_job=JOB,
            )

        st = _staging()
        por: dict[str, list[str]] = {}
        for fn in os.listdir(st):
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if not any(t in fn.upper() for t in TOK):
                continue
            if not RE6.search(fn):
                continue
            p = fn.split("__")[1]
            if p in PIEZAS:
                por.setdefault(p, []).append(os.path.join(st, fn))

        for pieza in PIEZAS:
            archs = por.get(pieza) or []
            print(f"{pieza}: staging n={len(archs)} completo={_completo(archs)}")
            for a in archs:
                print(" ", os.path.basename(a))
            if not _completo(archs):
                print(f"NO SUBE incompleto: {pieza}")
                continue
            for share in (SHARE, SHARE2):
                dst = os.path.join(share, pieza)
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
            print(f"SUBE {pieza}: {len(archs)} jpg")
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
