"""
Flag ``Seleccionadas`` (sí/no) sobre capturas JPG.

Reglas GIGA / Abigail:
  - **Metal (no cobre/busbar):** todas las cotas → ``sí``.
  - **Cobre/busbar:**
      - Medidas generales → ``sí``:
        LENGTH / WIDTH* (incl. WIDTH_TOTAL, WIDTH1..N) / THK / HEIGHT /
        LEG / WING / OD / ID / HOLE* (diámetros).
      - Barrenos XY (XMIN/XMAX/YMIN/YMAX/XCENTRO/YCENTRO ± TYP):
        ``sí`` = las 2 primeras posiciones distintas en X y las 2 en Y;
        el resto XY → ``no``.
      - CUT_* / otras no listadas → ``no``.
"""

from __future__ import annotations

import os
import re

try:
    from nomenclatura_capturas import SEP as _SEP
except Exception:
    _SEP = "__"

# Medida XY: XMIN/XMAX/XCENTRO[_TYP]_{valor}[mm|in] (idem Y)
_RE_MEDIDA_XY = re.compile(
    r"^(?P<eje>X|Y)(?:MIN|MAX|CENTRO)(?:_TYP)?_(?P<val>-?\d+(?:\.\d+)?)(?:mm|in)?$",
    re.IGNORECASE,
)

# Medidas generales de pieza (siempre sí en cobre; en metal ya va todo sí).
# Incluye Ø HOLE01..nn; la segregación cobre queda solo en XY.
_RE_MEDIDA_GENERAL = re.compile(
    r"^(?:"
    r"LENGTH|WIDTH(?:_TOTAL|[1-9])?|BROAD|THK|HEIGHT|LEG|WING\d{2}|ANGLE\d{2}|OD|ID|"
    r"HOLE\d{2}|"
    r"ANCHO|LARGO|ALTO|LADO|FRENTE_[12]|"
    r"DIAMETRO_(?:EXTERIOR|INTERIOR|H\d{2})"
    r")(?:_SIN_COTA)?_"
    r"-?\d+(?:\.\d+)?(?:mm|in|deg)?$",
    re.IGNORECASE,
)

_VAL_TOL = 1e-4


def _parts_captura(nombre_archivo: str) -> tuple[str, str] | None:
    """``(item, medida)`` desde ``JOB__ITEM__MEDIDA_val.jpg``."""
    base = os.path.splitext(os.path.basename(str(nombre_archivo or "")))[0]
    if not base or _SEP not in base:
        return None
    parts = base.split(_SEP)
    if len(parts) < 3:
        return None
    item = parts[1].strip()
    medida = parts[2].strip()
    if not item or not medida:
        return None
    return item, medida


def parse_xycentro_captura(nombre_archivo: str) -> tuple[str, str, float] | None:
    """
    Devuelve ``(item, 'X'|'Y', valor)`` si el JPG es cota XY de barreno
    (MIN/MAX/CENTRO).
    """
    parsed = _parts_captura(nombre_archivo)
    if parsed is None:
        return None
    item, medida = parsed
    m = _RE_MEDIDA_XY.match(medida)
    if not m:
        return None
    try:
        val = float(m.group("val"))
    except Exception:
        return None
    return item, m.group("eje").upper(), val


def es_medida_general(nombre_archivo: str) -> bool:
    parsed = _parts_captura(nombre_archivo)
    if parsed is None:
        return False
    return bool(_RE_MEDIDA_GENERAL.match(parsed[1]))


def _es_metal_item(item: str) -> bool:
    """True si la pieza NO es cobre/busbar (todo SI)."""
    try:
        from piezas_cobre import es_pieza_cobre

        return not bool(es_pieza_cobre(item))
    except Exception:
        # Sin catálogo: no asumir metal (conservador → lógica cobre)
        return False


def _vals_cercanos(a: float, b: float) -> bool:
    return abs(float(a) - float(b)) <= _VAL_TOL


def mapa_seleccionadas(nombres_archivo: list[str] | tuple[str, ...]) -> dict[str, str]:
    """
    ``{basename_jpg: 'si'|'no'}`` para el conjunto dado (misma pieza o carpeta).
    """
    out: dict[str, str] = {}
    # Cobre: acumular XY por item para elegir primeras 2 X / 2 Y
    por_item_xy: dict[str, dict[str, list[tuple[float, str]]]] = {}

    for raw in nombres_archivo or []:
        bn = os.path.basename(str(raw or ""))
        if not bn:
            continue
        parts = _parts_captura(bn)
        if parts is None:
            out[bn] = "no"
            continue
        item, _medida = parts

        # Metal: todas las cotas SI
        if _es_metal_item(item):
            out[bn] = "si"
            continue

        # Cobre — medidas generales siempre SI
        if es_medida_general(bn):
            out[bn] = "si"
            continue

        # Cobre — XY: diferir a agrupación
        parsed_xy = parse_xycentro_captura(bn)
        if parsed_xy is not None:
            _item, eje, val = parsed_xy
            por_item_xy.setdefault(item, {}).setdefault(eje, []).append((val, bn))
            continue

        # Cobre — HOLE / CUT / resto
        # (HOLE ya entró como medida general; CUT y desconocidos → no)
        out[bn] = "no"

    for item, ejes in por_item_xy.items():
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

    Preferir pasar todos los JPG de la pieza/carpeta como ``hermanos``
    (necesario para segregación XY en cobre).
    """
    bn = os.path.basename(str(nombre_archivo or ""))
    if not bn:
        return "no"
    # Metal / medida general: no hace falta hermanos
    parts = _parts_captura(bn)
    if parts is not None:
        item, _ = parts
        if _es_metal_item(item) or es_medida_general(bn):
            return "si"
    lista = list(hermanos) if hermanos is not None else [bn]
    if bn not in {os.path.basename(x) for x in lista}:
        lista.append(bn)
    return mapa_seleccionadas(lista).get(bn, "no")


def hermanos_en_carpeta(ruta_jpg_o_dir: str) -> list[str]:
    """Basenames JPG en la misma carpeta que ``ruta_jpg_o_dir``."""
    p = str(ruta_jpg_o_dir or "")
    if not p:
        return []
    carpeta = p if os.path.isdir(p) else os.path.dirname(os.path.abspath(p))
    if not carpeta or not os.path.isdir(carpeta):
        return []
    out: list[str] = []
    try:
        for fn in os.listdir(carpeta):
            if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                out.append(fn)
    except OSError:
        return []
    return sorted(out)
