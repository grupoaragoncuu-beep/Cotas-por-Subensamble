# -*- coding: utf-8 -*-
"""Smoke: piezas_cobre + nomenclatura SIN_COTA."""
from __future__ import annotations

import sys

sys.path.insert(0, ".")
import piezas_cobre as pc
from nomenclatura_capturas import armar_nombre_captura_pieza


def main():
    assert pc.es_pieza_cobre("ABB-42-BCK-713")
    assert pc.es_pieza_cobre("GENE-FCU-5-102")
    assert pc.es_pieza_cobre("RLG-123-BAR")
    assert pc.es_pieza_cobre("NESTING_1.0_GENE-FCU-5-102")
    # Fuera de regla (aunque sean board):
    assert not pc.es_pieza_cobre("GEN1-EE-2011-T2")
    assert not pc.es_pieza_cobre("9919-F-4.3-5000A-2011")
    assert not pc.es_pieza_cobre("AcuCT-S650-Series-Mounting Bracket R")
    assert not pc.es_pieza_cobre("CT200")
    assert not pc.es_pieza_cobre("HW-BLT-ZN-TF-1024-375_93882A228")
    assert not pc.es_pieza_cobre("62201-1248-P18_405")
    assert pc.prefijo_cobre("GENE-BKT-101") == "GENE"
    assert pc.prefijo_cobre("RLG-01") == "RLG"
    assert pc.catalogo_prefijos() == ("ABB", "GENE", "RLG")
    assert pc.medida_con_sin_cota("LENGTH") == "LENGTH_SIN_COTA"
    from cota_estilo import set_unidad_cota

    set_unidad_cota("mm")
    nom = armar_nombre_captura_pieza(
        "9919-Board 1", "GENE-FCU-5-102", "LENGTH_SIN_COTA", "3.5"
    )
    assert "LENGTH_SIN_COTA_3.500mm" in nom, nom
    set_unidad_cota("in")
    nom_in = armar_nombre_captura_pieza(
        "62223-1246-A01", "62223-1248-P01", "LENGTH", "10.25"
    )
    assert "LENGTH_10.250in" in nom_in, nom_in
    print("OK cobre ABB/GENE/RLG + SIN_COTA ->", nom)
    print("OK tanque in ->", nom_in)
    print("SMOKE piezas_cobre PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
