# -*- coding: utf-8 -*-
"""Crea el arbol acordado bajo DOSSIER FILES\\JPGS. Uso: armar_arbol_dossier.py <ruta JPGS>."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cotas_dossier_registro import armar_arbol_dossier_jpgs


def main() -> int:
    if len(sys.argv) < 2 or not str(sys.argv[1]).strip():
        print("Falta la ruta JPGS")
        return 1
    n = armar_arbol_dossier_jpgs(sys.argv[1])
    print(n)
    return 0 if n else 1


if __name__ == "__main__":
    raise SystemExit(main())
