# -*- coding: utf-8 -*-
"""Smoke: Seleccionadas — metal todo SI; cobre LWT SI + 2 primeras XY."""
from cotas_seleccionadas_cobre import (
    mapa_seleccionadas,
    parse_xycentro_captura,
    es_medida_general,
    seleccionada_para_captura,
)

FILES_COBRE = [
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
    "9919-Board 11__GENE-FCU-5-118__XMIN_1.299000.jpg",
    "9919-Board 11__GENE-FCU-5-118__XMIN_3.000000.jpg",
    "9919-Board 11__GENE-FCU-5-118__YMIN_1.299000.jpg",
    "9919-Board 11__GENE-FCU-5-118__YMIN_2.500000.jpg",
    "9919-Board 11__GENE-FCU-5-118__YMIN_4.000000.jpg",
]

FILES_METAL = [
    "9919-Board 1__GENE-DF-10-124__LENGTH_847.70mm.jpg",
    "9919-Board 1__GENE-DF-10-124__WIDTH_441.73mm.jpg",
    "9919-Board 1__GENE-DF-10-124__THK_1.75mm.jpg",
    "9919-Board 1__GENE-DF-10-124__XMIN_15.88mm.jpg",
    "9919-Board 1__GENE-DF-10-124__XMIN_80.35mm.jpg",
    "9919-Board 1__GENE-DF-10-124__YMIN_21.42mm.jpg",
    "9919-Board 1__GENE-DF-10-124__CUT_WIDTH_TYP_31.00mm.jpg",
    "9919-Board 1__GENE-BKT-101__YMIN_TYP_19.88mm.jpg",
]


def main():
    assert parse_xycentro_captura(FILES_COBRE[5]) == (
        "ABB-42-BCK-701",
        "X",
        25.4,
    )
    assert es_medida_general(FILES_COBRE[2]) is True
    assert es_medida_general(FILES_COBRE[5]) is False

    marks = mapa_seleccionadas(FILES_COBRE)
    # Generales cobre → sí
    assert marks["9919-Board 2__ABB-42-BCK-701__LENGTH_469.1.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__WIDTH_101.6.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__THK_6.350000.jpg"] == "si"
    # XY primeras 2
    assert marks["9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_25.400000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__XCENTRO_TYP_76.200000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_38.100000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_88.900000.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__YCENTRO_TYP_443.731819.jpg"] == "no"
    # HOLE cobre → sí (diámetro = medida general)
    assert marks["9919-Board 2__ABB-42-BCK-701__HOLE01_11.112500.jpg"] == "si"
    assert marks["9919-Board 2__ABB-42-BCK-701__HOLE02_10.312400.jpg"] == "si"
    assert marks["9919-Board 11__GENE-FCU-5-118__YMIN_4.000000.jpg"] == "no"

    marks_m = mapa_seleccionadas(FILES_METAL)
    # Metal: TODO sí (incl. XY extra y CUT)
    for fn, flag in marks_m.items():
        assert flag == "si", (fn, flag)
    assert seleccionada_para_captura(FILES_METAL[3]) == "si"

    print(
        "OK seleccionadas:",
        "cobre_si=",
        sum(1 for v in marks.values() if v == "si"),
        "metal_si=",
        sum(1 for v in marks_m.values() if v == "si"),
    )


if __name__ == "__main__":
    main()
