# -*- coding: utf-8 -*-
"""Rutas del árbol acordado GIGA BOARD (flat / doblado / estañado)."""
from __future__ import annotations

import os

from piezas_cobre import es_pieza_cobre

DOSSIER_FILES_BOARD2 = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES"
)
ROOT_JPGS_BOARD2 = os.path.join(DOSSIER_FILES_BOARD2, "JPGS")
ROOT_TWIN_BOARD2 = os.path.join(DOSSIER_FILES_BOARD2, "9919-BOARD2_2")


def dest_flat(raiz: str, pieza: str) -> str:
    """Flat: cobre → Corte/Maquinado/Corte Busbar; metal → Plasma y Laser/Corte metal."""
    pieza = str(pieza or "").strip()
    if es_pieza_cobre(pieza):
        return os.path.join(raiz, "Corte", "Maquinado", "Corte Busbar", pieza)
    return os.path.join(raiz, "Corte", "Plasma y Laser", "Corte metal", pieza)


def dest_doblado(raiz: str, pieza: str) -> str:
    if es_pieza_cobre(pieza):
        return os.path.join(raiz, "Doblado", "Busbar", str(pieza).strip())
    return os.path.join(raiz, "Doblado", "Metal", str(pieza).strip())


def dest_estanado(raiz: str) -> str:
    """Solo ``Estañado Busbar`` (con ñ). Si existe legacy sin ñ, se usa al leer."""
    buen = os.path.join(raiz, "Estañado Busbar")
    mal = os.path.join(raiz, "Estanado Busbar")
    if os.path.isdir(buen):
        return buen
    if os.path.isdir(mal):
        return mal
    return buen


def roots_flat(raiz: str) -> list[str]:
    """Carpetas raíz donde viven piezas flat."""
    return [
        os.path.join(raiz, "Corte", "Plasma y Laser", "Corte metal"),
        os.path.join(raiz, "Corte", "Maquinado", "Corte Busbar"),
        # Legacy (migración incompleta)
        os.path.join(raiz, "Corte", "Corte"),
    ]


def catalogo_piezas_flat(raiz: str) -> list[str]:
    """Nombres de carpeta-pieza bajo destinos flat (+ legacy)."""
    seen: dict[str, str] = {}
    for base in roots_flat(raiz):
        if not os.path.isdir(base):
            continue
        for n in os.listdir(base):
            p = os.path.join(base, n)
            if not os.path.isdir(p):
                continue
            if n.upper().startswith("COPIA DE") or n.startswith("_"):
                continue
            key = n.upper()
            if key not in seen:
                seen[key] = n
    return sorted(seen.values(), key=lambda s: s.upper())


def carpeta_pieza_flat(raiz: str, pieza: str) -> str:
    """Carpeta existente de la pieza flat, o destino canónico si aún no existe."""
    for base in roots_flat(raiz):
        p = os.path.join(base, pieza)
        if os.path.isdir(p):
            return p
    return dest_flat(raiz, pieza)
