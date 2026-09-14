# -*- coding: utf-8 -*-
"""
Detección de piezas de COBRE (gabinetes / boards GIGA).

Regla de producto (GIGA): solo nombres que **inician** con
``ABB``, ``GENE`` o ``RLG`` (tras quitar ``NESTING_…_`` opcional).

Por pieza cobre se exporta **una** captura sin cota desde la cara de
mayor área (LENGTH / FRENTE_1): ``…__LENGTH_<v>.jpg`` +
``…__LENGTH_SIN_COTA_<v>.jpg``. El resto de cotas (WIDTH, THK, HOLE…)
no se duplican.
"""

from __future__ import annotations

import os
import re
from functools import lru_cache

# Únicos prefijos cobre admitidos en GIGA.
PREFIJOS_COBRE = (
    "ABB",
    "GENE",
    "RLG",
)

# Prefijo al inicio del token (ABB-42…, GENE-FCU…, RLG-…).
_RE_PREFIJO = re.compile(
    r"^(?:NESTING_[\d.]+_)?(?P<pre>"
    + "|".join(re.escape(p) for p in PREFIJOS_COBRE)
    + r")(?:[-_\s]|$)",
    re.IGNORECASE,
)

SUFIJO_SIN_COTA = "SIN_COTA"


def _nombre_base(nombre: str) -> str:
    s = str(nombre or "").strip()
    s = os.path.basename(s)
    # Solo extensiones CAD reales (splitext rompe NESTING_1.0_GENE-…).
    s = re.sub(
        r"\.(ipt|iam|idw|dwg|stp|step|jpg|jpeg|png)$",
        "",
        s,
        flags=re.IGNORECASE,
    )
    # Quita :N de ocurrencia Inventor.
    if ":" in s:
        s = s.split(":", 1)[0]
    return s.strip()


def prefijo_cobre(nombre: str) -> str | None:
    """Devuelve ``ABB`` / ``GENE`` / ``RLG`` o None."""
    base = _nombre_base(nombre)
    if not base:
        return None
    m = _RE_PREFIJO.match(base)
    if not m:
        return None
    return str(m.group("pre")).upper()


def es_pieza_cobre(nombre: str, material: str | None = None) -> bool:
    """
    True solo si el nombre inicia con ABB / GENE / RLG.

    ``material`` se ignora (la regla es por nomenclatura de archivo).
    """
    return prefijo_cobre(nombre) is not None


def medida_con_sin_cota(medida: str) -> str:
    """LENGTH → LENGTH_SIN_COTA (idempotente)."""
    m = str(medida or "LENGTH").upper().strip("_")
    if m.endswith("_" + SUFIJO_SIN_COTA) or m == SUFIJO_SIN_COTA:
        return m
    return f"{m}_{SUFIJO_SIN_COTA}"


@lru_cache(maxsize=1)
def catalogo_prefijos() -> tuple[str, ...]:
    return tuple(sorted(p.upper() for p in PREFIJOS_COBRE))
