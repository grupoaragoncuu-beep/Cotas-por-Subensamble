# -*- coding: utf-8 -*-
"""Publica las 3 piezas XMIN/YMIN+THK ya exportadas."""
from __future__ import annotations

import os
import re
import shutil

ROOT = os.path.dirname(os.path.abspath(__file__))
ST = os.path.join(
    ROOT, "JPG", "9919-Board 1", "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE"
)
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
PIEZAS = ("GENE-GS-0820-708", "GENE-OP-10-116", "GENE-OP-10-117")
RE6 = re.compile(r"\.\d{6}(?:_|\.|$)")
TOK = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")


def main() -> int:
    por: dict[str, list[str]] = {}
    for fn in os.listdir(ST):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        if not any(t in fn.upper() for t in TOK):
            continue
        if not RE6.search(fn):
            continue
        if "_DESPLIEGUE_" in fn.upper() and "__LENGTH_" in fn.upper():
            continue
        p = fn.split("__")[1]
        if p in PIEZAS:
            por.setdefault(p, []).append(os.path.join(ST, fn))

    for pieza in PIEZAS:
        archs = por.get(pieza) or []
        ups = [os.path.basename(a).upper() for a in archs]
        ok = any(
            any(t in u for t in ("XMIN", "YMIN", "XCENTRO", "YCENTRO")) for u in ups
        ) and any("__THK_" in u for u in ups)
        print(f"{pieza}: n={len(archs)} ok={ok}")
        for a in archs:
            print(" ", os.path.basename(a))
        if not ok:
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
        print(f"SUBE {pieza}")
    print("DONE")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
