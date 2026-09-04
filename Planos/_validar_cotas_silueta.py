"""
Validación offline del anclaje ALTO/LARGO_PATA a silueta exterior.

Caso P156 (canal C): si faltan pestañas en HLR o se elige el alma,
la cota cae en la tangencia del doblez (~10.44 vs ~11.02). La lógica
nueva debe:
  1) preferir pestañas horizontales exteriores,
  2) si no hay pestañas, usar arcos que SÍ tocan min/max global,
  3) rechazar cotas con ratio < 0.97 respecto a la silueta.

No requiere Inventor:
  python Planos/_validar_cotas_silueta.py
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from THK import (  # noqa: E402
    _elegir_curva_extrema_thk,
    _pares_borde_silueta,
    _span_bordes_seleccionados,
)


def _dato(nombre, minx, maxx, miny, maxy):
    return {
        "curve": nombre,
        "minx": float(minx),
        "maxx": float(maxx),
        "miny": float(miny),
        "maxy": float(maxy),
        "dx": abs(float(maxx) - float(minx)),
        "dy": abs(float(maxy) - float(miny)),
        "cx": (float(minx) + float(maxx)) * 0.5,
        "cy": (float(miny) + float(maxy)) * 0.5,
    }


def _canal_c_completo():
    """Canal C con pestañas exteriores visibles."""
    return [
        _dato("alma_vert", 0.0, 0.15, 1.0, 27.0),
        _dato("pestana_sup", 0.0, 5.0, 27.85, 28.0),
        _dato("pestana_inf", 0.0, 5.0, 0.0, 0.15),
        _dato("arco_sup", 0.0, 1.0, 26.5, 28.0),
        _dato("arco_inf", 0.0, 1.0, 0.0, 1.5),
        _dato("pata_sup_vert", 4.85, 5.0, 26.5, 28.0),
        _dato("pata_inf_vert", 4.85, 5.0, 0.0, 1.5),
    ]


def _canal_c_sin_pestanas_hlr():
    """
    Caso realista del bug: HLR no entrega la arista exterior de pestaña
    como curva usable; el alma termina en tangencia Y=1..27, y solo los
    arcos tocan Y=0 y Y=28.
    """
    return [
        _dato("alma_vert", 0.0, 0.15, 1.0, 27.0),
        _dato("arco_sup", 0.0, 1.0, 26.5, 28.0),
        _dato("arco_inf", 0.0, 1.0, 0.0, 1.5),
        _dato("pata_sup_vert", 4.85, 5.0, 26.5, 28.0),
        _dato("pata_inf_vert", 4.85, 5.0, 0.0, 1.5),
    ]


def _logica_vieja_span(datos):
    """max(maxy)/min(miny) + intent 'natural' en alma si gana el alma."""
    a = max(datos, key=lambda d: d["maxy"])
    b = min(datos, key=lambda d: d["miny"])
    # Si el par no cubre silueta (alma), la cota vieja queda en tangencia.
    y_sup = a["maxy"]
    y_inf = b["miny"]
    if str(a["curve"]).startswith("alma"):
        y_sup = 27.0
    if str(b["curve"]).startswith("alma"):
        y_inf = 1.0
    # Intent sin punto en arco: Inventor suele caer cerca de la tangencia.
    if "arco" in str(b["curve"]):
        y_inf = 1.0
    if "arco" in str(a["curve"]):
        y_sup = 27.0
    return a, b, abs(y_sup - y_inf)


def main():
    print("=" * 60)
    print(" VALIDACION silueta ALTO (caso P156 canal C)")
    print("=" * 60)

    completo = _canal_c_completo()
    span_real = 28.0
    tol = max(0.03, span_real * 0.02)

    # 1) Con pestañas: nueva elige pestañas.
    sup = _elegir_curva_extrema_thk(completo, "sup", tol)
    inf = _elegir_curva_extrema_thk(completo, "inf", tol)
    print(
        f"\n[1] Con pestanas: sup={sup['curve']} inf={inf['curve']}"
    )
    assert "pestana_sup" == str(sup["curve"])
    assert "pestana_inf" == str(inf["curve"])
    a, b, *_r, span = _pares_borde_silueta(completo, "V", tol)
    ratio = _span_bordes_seleccionados(a, b, "V") / span
    assert ratio >= 0.97, ratio
    print(f"    pares OK ratio={ratio:.3f}")

    # 2) Sin pestañas HLR: vieja falla (tangencia); nueva usa arcos extremos.
    parcial = _canal_c_sin_pestanas_hlr()
    a_old, b_old, span_old = _logica_vieja_span(parcial)
    ratio_old = span_old / span_real
    print(
        f"\n[2] Sin pestanas HLR - VIEJA: {a_old['curve']}/{b_old['curve']} "
        f"ratio={ratio_old:.3f}"
    )
    assert ratio_old < 0.97, "debe reproducir cota corta por tangencia"

    a2, b2, *_r2, span2 = _pares_borde_silueta(parcial, "V", tol)
    ratio2 = _span_bordes_seleccionados(a2, b2, "V") / span2
    print(
        f"    NUEVA: {a2['curve']}/{b2['curve']} ratio={ratio2:.3f}"
    )
    assert ratio2 >= 0.97, ratio2
    assert abs(a2["maxy"] - 28.0) <= tol
    assert abs(b2["miny"] - 0.0) <= tol

    # 3) Umbral rechaza el JPG malo 10.44 vs ~11.02 (ratio~0.947).
    ratio_jpg = 10.44 / 11.02
    print(f"\n[3] JPG malo P156_ALTO 10.44/11.02 ratio={ratio_jpg:.3f}")
    assert ratio_jpg < 0.97, "umbral 0.97 debe rechazar 10.44"

    # 4) Carpetas legacy no deben crearse en reorg (smoke import).
    from generador_tanque_completo import (  # noqa: E402
        SUBCARPETAS_CARA_LEGACY,
        SUBCARPETAS_CARA_SELECCION,
        _carpetas_cara_activas,
    )

    activas = _carpetas_cara_activas(
        {"SEGM1": set(), "TOP": set(), "BASE": set()}
    )
    print(f"\n[4] Carpetas activas: {activas}")
    for leg in SUBCARPETAS_CARA_LEGACY:
        assert leg not in activas, leg
    for seg in ("SEGM1", "TOP", "BASE", "OTROS"):
        assert seg in activas or seg in SUBCARPETAS_CARA_SELECCION
    assert "OTROS" in activas

    print("\nOK: validacion silueta + carpetas aprobada.")
    print(
        "NOTA: JPG en PIEZAS_ACOTADAS son de la corrida ANTERIOR; "
        "re-ejecutar COTAS_ILOGIC_ABIGAIL para materializar el fix."
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as err:
        print(f"\nFALLO: {err}")
        raise SystemExit(1)
