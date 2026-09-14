"""Smoke: cotas no deben solapar la silueta de la pieza."""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from cota_estilo import (
    OFFSET_FUERA_PIEZA_CM,
    bbox_intersecta,
    candidatos_texto_fuera_pieza,
    clearance_texto_cota_cm,
    empujar_punto_fuera_bbox,
    punto_dentro_bbox,
)
import cotas
import THK
import lineal_especial


def _assert(cond, msg):
    if not cond:
        raise AssertionError(msg)


def main():
    print("=" * 56)
    print(" SMOKE cota fuera de pieza (anti-solape)")
    print("=" * 56)

    _assert(OFFSET_FUERA_PIEZA_CM >= 2.5, "offset base insuficiente")
    _assert(clearance_texto_cota_cm() >= OFFSET_FUERA_PIEZA_CM, "clearance")
    _assert(cotas.OFFSET_COTA >= 2.5, f"cotas.OFFSET={cotas.OFFSET_COTA}")
    _assert(THK.OFFSET_COTA >= 2.5, f"THK.OFFSET={THK.OFFSET_COTA}")
    _assert(
        lineal_especial.OFFSET_COTA >= 2.5,
        f"lineal.OFFSET={lineal_especial.OFFSET_COTA}",
    )
    print("  OK offsets >= 2.5 cm")

    pieza = (10.0, 20.0, 5.0, 15.0)
    # Punto clampeado “hacia adentro” no debe quedar dentro.
    x, y = empujar_punto_fuera_bbox(15.0, 10.0, pieza, "izq")
    _assert(x < 10.0 and not punto_dentro_bbox(x, y, pieza), f"izq=({x},{y})")
    x2, y2 = empujar_punto_fuera_bbox(15.0, 10.0, pieza, "sup")
    _assert(y2 > 15.0 and not punto_dentro_bbox(x2, y2, pieza), f"sup=({x2},{y2})")
    print("  OK empujar fuera de bbox")

    cands_v = candidatos_texto_fuera_pieza(pieza, "V")
    _assert(len(cands_v) >= 4, f"cands V={len(cands_v)}")
    _assert(all(not punto_dentro_bbox(x, y, pieza) for x, y, _ in cands_v), "V fuera")
    cands_h = candidatos_texto_fuera_pieza(pieza, "H")
    _assert(all(not punto_dentro_bbox(x, y, pieza) for x, y, _ in cands_h), "H fuera")
    print("  OK candidatos texto fuera")

    # Caso Vantran: si el texto cae en el centro de la brida, intersecta.
    texto_mal = (14.0, 16.0, 9.0, 11.0)  # bbox texto sobre la pieza
    _assert(bbox_intersecta(pieza, texto_mal, holgura=0.2), "debe detectar solape")
    texto_ok = (5.0, 8.0, 8.0, 12.0)  # a la izquierda
    _assert(not bbox_intersecta(pieza, texto_ok, holgura=0.2), "no falso positivo")
    print("  OK detección de solape bbox")

    # Clamp con evitar_bbox no mete el punto en la pieza.
    class Hoja:
        Width = 40.0
        Height = 30.0

    class TG:
        @staticmethod
        def CreatePoint2d(x, y):
            return type("P", (), {"X": x, "Y": y})()

    # Forzar y=14 (dentro de pieza 5..15) como si el sheet hubiera empujado.
    pt = cotas._clampear_punto_hoja(
        Hoja(), TG(), 15.0, 14.0, margen=1.2, evitar_bbox=pieza
    )
    _assert(
        not punto_dentro_bbox(pt.X, pt.Y, pieza, holgura=0.15),
        f"clamp metió texto en pieza: ({pt.X},{pt.Y})",
    )
    print(f"  OK clamp evita pieza -> ({pt.X:.2f},{pt.Y:.2f})")

    print("\nTODAS LAS REVALIDACIONES OK")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except AssertionError as exc:
        print(f"\nFAIL: {exc}")
        sys.exit(1)
