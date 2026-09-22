"""
Calidad cobre/busbar: flag ``Seleccionadas`` (sí/no) sobre cotas XY de barrenos.

Regla (GIGA / Abigail cobre):
  - Solo piezas ``es_pieza_cobre`` (ABB / GENE / RLG…).
  - Cotas al BORDE del barreno (XMIN/XMAX/YMIN/YMAX). También acepta
    XCENTRO/YCENTRO legado.
  - ``sí`` = las 2 primeras posiciones distintas en X y las 2 primeras en Y
    desde (0,0) (valores numéricos del nombre de captura, ascendente).
  - El resto de cotas de esa pieza → ``no``.
  - Piezas no-cobre → siempre ``no``.
"""

from __future__ import annotations

import os
import re

from piezas_cobre import es_pieza_cobre

try:
    from nomenclatura_capturas import SEP as _SEP
except Exception:
    _SEP = "__"

# Medida final: XMIN/XMAX/XCENTRO[_TYP]_{valor} (idem Y)
_RE_MEDIDA_XY = re.compile(
    r"^(?P<eje>X|Y)(?:MIN|MAX|CENTRO)(?:_TYP)?_(?P<val>-?\d+(?:\.\d+)?)$",
    re.IGNORECASE,
)

_VAL_TOL = 1e-4


def parse_xycentro_captura(nombre_archivo: str) -> tuple[str, str, float] | None:
    """
    Devuelve ``(item, 'X'|'Y', valor)`` si el JPG es cota XY de barreno
    (MIN/MAX/CENTRO).
    """
    base = os.path.splitext(os.path.basename(str(nombre_archivo or "")))[0]
    if not base:
        return None
    partes = base.split(_SEP)
    if len(partes) < 2:
        return None
    # {JOB}__{ITEM}__{MEDIDA_VALOR}  o  {ITEM}__{MEDIDA_VALOR}
    if len(partes) >= 3:
        item = str(partes[1] or "").strip()
        medida = str(partes[-1] or "").strip()
    else:
        item = str(partes[0] or "").strip()
        medida = str(partes[1] or "").strip()
    m = _RE_MEDIDA_XY.match(medida)
    if not m:
        return None
    eje = str(m.group("eje") or "").upper()
    if not item or eje not in ("X", "Y"):
        return None
    try:
        val = float(m.group("val"))
    except Exception:
        return None
    return item, eje, val


def _vals_cercanos(a: float, b: float) -> bool:
    return abs(float(a) - float(b)) <= _VAL_TOL


def mapa_seleccionadas(nombres_archivo: list[str] | tuple[str, ...]) -> dict[str, str]:
    """
    ``{basename_jpg: 'si'|'no'}`` para el conjunto dado (misma pieza o carpeta).
    """
    out: dict[str, str] = {}
    por_item: dict[str, dict[str, list[tuple[float, str]]]] = {}

    for raw in nombres_archivo or []:
        bn = os.path.basename(str(raw or ""))
        if not bn:
            continue
        parsed = parse_xycentro_captura(bn)
        if parsed is None:
            out[bn] = "no"
            continue
        item, eje, val = parsed
        if not es_pieza_cobre(item):
            out[bn] = "no"
            continue
        por_item.setdefault(item, {}).setdefault(eje, []).append((val, bn))

    for item, ejes in por_item.items():
        elegidos: set[tuple[str, float]] = set()
        for eje in ("X", "Y"):
            vals = sorted({v for v, _ in ejes.get(eje, [])})
            for v in vals[:2]:
                elegidos.add((eje, v))
        for eje, pares in ejes.items():
            for val, bn in pares:
                hit = any(
                    e == eje and _vals_cercanos(val, vv) for e, vv in elegidos
                )
                out[bn] = "si" if hit else "no"

    return out


def seleccionada_para_captura(
    nombre_archivo: str,
    hermanos: list[str] | tuple[str, ...] | None = None,
) -> str:
    """
    ``'si'`` / ``'no'`` para una captura.

    Si ``hermanos`` es None, solo puede marcar ``sí`` con lógica incompleta
    (sin hermanos → ``no`` salvo que se pase lista). Preferir pasar todos los
    JPG de la pieza/carpeta.
    """
    bn = os.path.basename(str(nombre_archivo or ""))
    if not bn:
        return "no"
    lista = list(hermanos) if hermanos is not None else [bn]
    if bn not in {os.path.basename(x) for x in lista}:
        lista.append(bn)
    return mapa_seleccionadas(lista).get(bn, "no")
