# -*- coding: utf-8 -*-
"""
Detección de piezas de COBRE / busbar (gabinetes GIGA).

Fuente oficial del catálogo ``cobre_nesting_giga.txt``:
**todas las carpetas AutoDXF** bajo GIGA (y ``GIGA BOARD * - COBRE``),
regeneradas con ``.runtime/_rebuild_cobre_nesting_giga.py``.

En Inventor el material suele figurar como acero hasta que lo cambian a
cobre después; **no** usar material Inventor para clasificar cobre.
La verdad es el nesting AutoDXF (nombres ``*, Cobre, QTY…``).

Eso limita el flujo busbar (XY/HOLE/SIN_COTA/ESTANIADO / Corte Busbar)
a piezas con DXF de cobre real. El resto en BOARD = corte normal.
"""

from __future__ import annotations

import os
import re
from functools import lru_cache

# Prefijos históricos (solo documentación / fallback de parseo de nombre).
PREFIJOS_COBRE = (
    "ABB",
    "GENE",
    "RLG",
    "GE813",
)

_RE_PREFIJO = re.compile(
    r"^(?:NESTING_[\d.]+_)?(?P<pre>"
    + "|".join(re.escape(p) for p in PREFIJOS_COBRE)
    + r")(?:[-_\s]|$)",
    re.IGNORECASE,
)

# Quita basura de stem DXF: ", Cobre, QTY 4, Cal 0.25" / ", CU, QTY …"
_RE_SUFIJO_DXF = re.compile(
    r"\s*,\s*(?:Cobre|Copper|COBRE|CU)\b.*$",
    re.IGNORECASE,
)

SUFIJO_SIN_COTA = "SIN_COTA"

_CATALOGO_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "cobre_nesting_giga.txt"
)


def _nombre_base(nombre: str) -> str:
    s = str(nombre or "").strip()
    s = os.path.basename(s)
    s = re.sub(
        r"\.(ipt|iam|idw|dwg|stp|step|jpg|jpeg|png|dxf)$",
        "",
        s,
        flags=re.IGNORECASE,
    )
    if ":" in s:
        s = s.split(":", 1)[0]
    s = re.sub(r"^NESTING_[\d.]+_", "", s, flags=re.IGNORECASE)
    s = _RE_SUFIJO_DXF.sub("", s)
    if "," in s:
        s = s.split(",", 1)[0]
    return s.strip()


def nombre_catalogo_cobre(nombre: str) -> str:
    """Clave canónica para match contra el catálogo nesting."""
    return _nombre_base(nombre).upper()


@lru_cache(maxsize=1)
def catalogo_cobre_nesting() -> frozenset[str]:
    """
    Set UPPER de piezas busbar GIGA (AutoDXF nesting).

    Fuente: ``Planos/cobre_nesting_giga.txt`` (una pieza por línea).
    """
    path = _CATALOGO_FILE
    out: set[str] = set()
    try:
        with open(path, encoding="utf-8") as f:
            for ln in f:
                s = ln.strip()
                if not s or s.startswith("#"):
                    continue
                key = nombre_catalogo_cobre(s)
                if key:
                    out.add(key)
    except OSError:
        pass
    return frozenset(out)


def prefijo_cobre(nombre: str) -> str | None:
    """Devuelve prefijo ABB/GENE/RLG/GE813 si el nombre lo trae (informativo)."""
    base = _nombre_base(nombre)
    if not base:
        return None
    m = _RE_PREFIJO.match(base)
    if not m:
        return None
    return str(m.group("pre")).upper()


def es_pieza_cobre(nombre: str, material: str | None = None) -> bool:
    """
    True solo si la pieza está en el catálogo nesting cobre GIGA.

    ``material`` se ignora (la regla es por nombre = DXF AutoDXF).
    Env ``COTAS_COBRE_PREFIJO=1``: fallback legacy (cualquier ABB/GENE/RLG).
    """
    _ = material
    if os.environ.get("COTAS_COBRE_PREFIJO", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "si",
        "on",
    ):
        return prefijo_cobre(nombre) is not None

    cat = catalogo_cobre_nesting()
    if not cat:
        # Sin archivo: no inventar cobre por prefijo (evita saturar el flujo).
        return False
    return nombre_catalogo_cobre(nombre) in cat


def es_busbar_giga(nombre: str) -> bool:
    """Alias explícito del filtro nesting busbar."""
    return es_pieza_cobre(nombre)


def medida_con_sin_cota(medida: str) -> str:
    """LENGTH → LENGTH_SIN_COTA (idempotente)."""
    m = str(medida or "LENGTH").upper().strip("_")
    if m.endswith("_" + SUFIJO_SIN_COTA) or m == SUFIJO_SIN_COTA:
        return m
    return f"{m}_{SUFIJO_SIN_COTA}"


@lru_cache(maxsize=1)
def catalogo_prefijos() -> tuple[str, ...]:
    return tuple(sorted(p.upper() for p in PREFIJOS_COBRE))


def recargar_catalogo_cobre() -> int:
    """Invalida cache y relee el txt. Devuelve tamaño del catálogo."""
    catalogo_cobre_nesting.cache_clear()
    return len(catalogo_cobre_nesting())
