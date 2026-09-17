# -*- coding: utf-8 -*-
"""Sube al dossier piezas con staging XY+THK a 6dec que aún faltan."""
from __future__ import annotations

import os
import re
import shutil

ROOT = os.path.dirname(os.path.abspath(__file__))
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
ST = os.path.join(
    ROOT, "JPG", "9919-Board 1", "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE"
)
RE6 = re.compile(r"\.\d{6}(?:_|\.|$)")
TOK = ("XCENTRO", "YCENTRO", "__THK_", "__HOLE")
FALTAN = {
    "ABB-42-BCK-726",
    "ABB-42-BCK-727",
    "ABB-42-BCK-728",
    "ABB-42-BCK-735",
    "GEN1-OP-20-112",
    "GENE-FCU-6-105",
    "GENE-FCU-6-106",
    "GENE-GS-0820-708",
    "GENE-OP-10-116",
    "GENE-OP-10-117",
    "GENE-OP-1020-115",
    "GENE-OP-1020-116",
}


def main() -> int:
    por: dict[str, list[str]] = {}
    for fn in os.listdir(ST):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        if not any(t in fn.upper() for t in TOK):
            continue
        if not RE6.search(fn):
            continue
        parts = fn.split("__")
        if len(parts) < 2:
            continue
        pieza = parts[1]
        if pieza not in FALTAN:
            continue
        por.setdefault(pieza, []).append(os.path.join(ST, fn))

    subidas = 0
    for pieza, archs in sorted(por.items()):
        ups = [os.path.basename(a).upper() for a in archs]
        if not (
            any("XCENTRO" in u or "YCENTRO" in u for u in ups)
            and any("__THK_" in u for u in ups)
        ):
            print(f"SKIP incompleto staging: {pieza}")
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
        subidas += 1
    print(f"done subidas={subidas}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
