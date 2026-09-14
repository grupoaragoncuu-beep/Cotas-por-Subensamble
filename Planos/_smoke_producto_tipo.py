# -*- coding: utf-8 -*-
"""Smoke offline: clasificador producto_tipo (sin Inventor)."""
from __future__ import annotations

import sys
from types import SimpleNamespace

sys.path.insert(0, ".")
import producto_tipo as pt


class FakeAsm:
    def __init__(self, display, path="", hijos=None):
        self.DisplayName = display
        self.FullFileName = path or f"C:\\tmp\\{display}.iam"
        self._hijos = hijos or []

    @property
    def ComponentDefinition(self):
        return SimpleNamespace(Occurrences=_Occs(self._hijos))


class _Occs:
    def __init__(self, hijos):
        self._hijos = hijos

    @property
    def Count(self):
        return len(self._hijos)

    def Item(self, i):
        return self._hijos[i - 1]


def _hijo(nombre, es_iam=False):
    doc = SimpleNamespace(
        FullFileName=f"C:\\x\\{nombre}.{'iam' if es_iam else 'ipt'}"
    )
    return SimpleNamespace(
        Name=f"{nombre}:1",
        Suppressed=False,
        Definition=SimpleNamespace(Document=doc),
    )


def main():
    giga_tank = FakeAsm("GIGA L3 TANK - GRUPO ARGA")
    info = pt.clasificar_producto(giga_tank, escanear_hijos=False)
    assert info["tipo"] == pt.TIPO_TANQUE, info
    assert info["familia"] == "GIGA", info
    assert pt.unidad_para_producto(info=info) == "mm", info
    print("OK giga tank (mm, no board):", info["motivo"])

    board = FakeAsm("9919-Board 1", r"C:\desk\9919-Board 1\9919-Board 1.iam")
    info = pt.clasificar_producto(board, escanear_hijos=False)
    assert info["tipo"] == pt.TIPO_BOARD, info
    assert pt.unidad_para_producto(info=info) == "mm", info
    print("OK board por nombre:", info["motivo"])

    vantran = FakeAsm("MODELO VANTRAN 251007")
    info = pt.clasificar_producto(vantran, escanear_hijos=False)
    assert info["tipo"] == pt.TIPO_TANQUE, info
    assert pt.unidad_para_producto(info=info) == "in", info
    print("OK vantran:", info["familia"])

    otc = FakeAsm("62201-1246-A01")
    info = pt.clasificar_producto(otc, escanear_hijos=False)
    assert info["tipo"] == pt.TIPO_TANQUE, info
    assert pt.unidad_para_producto(info=info) == "in", info
    print("OK otc:", info["motivo"])

    electric = FakeAsm(
        "PANEL X",
        hijos=[
            _hijo("GENE-BKT-101"),
            _hijo("GENE-DF-10-165"),
            _hijo("ABB-42-BCK-713"),
            _hijo("GEN1-EE-2011-T2", es_iam=True),
            _hijo("CITEL MS-Series SPD"),
            _hijo("91247A638_Screw"),
        ],
    )
    info = pt.clasificar_producto(electric, escanear_hijos=True)
    assert info["tipo"] == pt.TIPO_BOARD, info
    print("OK board por árbol:", info["motivo"])

    print("SMOKE producto_tipo PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
