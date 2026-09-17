# -*- coding: utf-8 -*-
"""Smoke: Seleccionadas cobre (2 primeras XCENTRO + 2 YCENTRO)."""
from cotas_seleccionadas_cobre import mapa_seleccionadas, parse_xycentro_captura

FILES = [
    "9919-Board 2__ABB-42-BCK-701__HOLE01_11.112500.jpg",
    "9919-Board 2__ABB-42-BCK-701__HOLE02_10.312400.jpg",
    "9919-Board 2__ABB-42-BCK-701__LENGTH_469.1.jpg",
    "9919-Board 2__ABB-42-BCK-701__THK_6.350000.jpg",
    "9919-Board 2__ABB-42-BCK-701__WIDTH_101.6.jpg",
    "9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_25.400000.jpg",
    "9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_76.200000.jpg",
    "9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_38.100000.jpg",
    "9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_443.731819.jpg",
    "9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_88.900000.jpg",
    # No cobre
    "9919-Board 2__GEN1-EE-2011-T2__XCENTRO_TYP_10.000000.jpg",
    "9919-Board 2__GEN1-EE-2011-T2__YCENTRO_TYP_20.000000.jpg",
]


def main():
    assert parse_xycentro_captura(FILES[5]) == (
        "ABB-42-BCK-701",
        "X",
        25.4,
    )
    marks = mapa_seleccionadas(FILES)
    assert marks["9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_25.400000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_76.200000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_38.100000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_88.900000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_443.731819.jpg"] == "no"
    assert marks["9919-Board 2__ABB-42-BCK-701__HOLE01_11.112500.jpg"] == "no"
    assert marks["9919-Board 2__ABB-42-BCK-701__LENGTH_469.1.jpg"] == "no"
    assert marks["9919-Board 2__GEN1-EE-2011-T2__XCENTRO_TYP_10.000000.jpg"] == "no"
    print("OK seleccionadas cobre:", sum(1 for v in marks.values() if v == "si"), "si")


if __name__ == "__main__":
    main()
