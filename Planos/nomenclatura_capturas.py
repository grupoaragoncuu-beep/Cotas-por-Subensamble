"""
Nomenclatura unificada de capturas JPG (Abigail + Subensamble).

Formato:
  {JOB}__{ITEM}__{MEDIDA}_{VALOR}.jpg

- JOB    = nombre del ensamble (sin .iam)
- ITEM   = nombre de la pieza
- MEDIDA = LENGTH / WIDTH / THK / HEIGHT / LEG / OD / ID / HOLE## /
           (legacy: BROAD = WIDTH) /
           XMIN / XMAX / YMIN / YMAX / XCENTRO / YCENTRO (+ _TYP / _CONTACTO)
- VALOR  = valor numérico de la cota visible en la captura (p. ej. 82, 0.5, 3)
"""

from __future__ import annotations

import os
import re

# Separador entre JOB / ITEM / MEDIDA_VALOR (evita ambigüedad con _ en nombres).
SEP = "__"

# Token del valor en el nombre: enteros o decimales (82, 0.5, 3.536).
_RE_VALOR_NOMBRE = r"\d+(?:\.\d+)?"

# Orden: compuestos antes que simples.
_TIPOS_HOJA_A_EXPORT = (
    ("DIAMETRO_EXTERIOR", "OD"),
    ("DIAMETRO_INTERIOR", "ID"),
    ("DESPLIEGUE_DIAMETRO_EXTERIOR", "OD"),
    ("DESPLIEGUE_DIAMETRO_INTERIOR", "ID"),
    ("LARGO_PATA", "LEG"),
    ("DESPLIEGUE_XCENTRO_TYP", "XCENTRO_TYP"),
    ("DESPLIEGUE_YCENTRO_TYP", "YCENTRO_TYP"),
    ("DESPLIEGUE_XCENTRO", "XCENTRO"),
    ("DESPLIEGUE_YCENTRO", "YCENTRO"),
    ("DESPLIEGUE_XMIN_TYP", "XMIN_TYP"),
    ("DESPLIEGUE_YMIN_TYP", "YMIN_TYP"),
    ("DESPLIEGUE_XMIN", "XMIN"),
    ("DESPLIEGUE_YMIN", "YMIN"),
    ("DESPLIEGUE_CUT_LENGTH", "CUT_LENGTH"),
    ("DESPLIEGUE_CUT_WIDTH", "CUT_WIDTH"),
    ("XCENTRO_TYP", "XCENTRO_TYP"),
    ("YCENTRO_TYP", "YCENTRO_TYP"),
    ("XCENTRO", "XCENTRO"),
    ("YCENTRO", "YCENTRO"),
    ("XMIN_TYP", "XMIN_TYP"),
    ("YMIN_TYP", "YMIN_TYP"),
    ("XMIN", "XMIN"),
    ("YMIN", "YMIN"),
    ("CUT_LENGTH", "CUT_LENGTH"),
    ("CUT_WIDTH", "CUT_WIDTH"),
    ("DESPLIEGUE_ANCHO", "WIDTH"),
    ("DESPLIEGUE_LARGO", "LENGTH"),
    ("DESPLIEGUE_THK", "THK"),
    ("DESPLIEGUE_FRENTE_1", "LENGTH"),
    ("DESPLIEGUE_FRENTE_2", "WIDTH"),
    ("DESPLIEGUE_LADO", "THK"),
    ("ANCHO", "WIDTH"),
    ("LARGO", "LENGTH"),
    ("THK", "THK"),
    ("ALTO", "HEIGHT"),
    ("LADO", "THK"),
    ("FRENTE_1", "LENGTH"),
    ("FRENTE_2", "WIDTH"),
    ("ESTANIADO", "LENGTH"),
)

_RE_DIAMETRO_H = re.compile(r"^(?P<item>.+?)_DIAMETRO_H(?P<nn>\d{2})$", re.IGNORECASE)
_RE_SUFIJO_INV = re.compile(r"^(?P<base>.+?):(?P<num>\d+)$")
_RE_NUMERO_EN_TEXTO = re.compile(r"-?\d+(?:[.,]\d+)?")

# Tipos export Abigail (inglés) + legacy español.
# Opcional ``_SIN_COTA`` = 2ª captura cobre (misma vista sin dimensión).
_TIPOS_PIEZA_CORE = (
    "LENGTH",
    "WIDTH",
    "BROAD",  # legacy alias de WIDTH
    "THK",
    "HEIGHT",
    "LEG",
    "OD",
    "ID",
    "XCENTRO_TYP",
    "YCENTRO_TYP",
    "XCENTRO",
    "YCENTRO",
    "XMIN_TYP",
    "YMIN_TYP",
    "XMIN",
    "YMIN",
    "CUT_LENGTH",
    "CUT_WIDTH",
    "DIAMETRO_EXTERIOR",
    "DIAMETRO_INTERIOR",
    "LARGO_PATA",
    "ANCHO",
    "LARGO",
    "ALTO",
    "LADO",
    "FRENTE_1",
    "FRENTE_2",
)
_TIPOS_PIEZA_EXPORT = tuple(
    [t for core in _TIPOS_PIEZA_CORE for t in (core, f"{core}_SIN_COTA")]
    + [r"HOLE\d{2}", r"HOLE\d{2}_SIN_COTA", r"DIAMETRO_H\d{2}", r"DIAMETRO_H\d{2}_SIN_COTA"]
)

_RE_CAPTURA_PIEZA_NUEVA = re.compile(
    r"^(?P<job>.+?)"
    + re.escape(SEP)
    + r"(?P<item>.+?)"
    + re.escape(SEP)
    + r"(?P<tipo>"
    + "|".join(_TIPOS_PIEZA_EXPORT)
    + r")_(?P<num>"
    + _RE_VALOR_NOMBRE
    + r")$",
    re.IGNORECASE,
)

_RE_CAPTURA_PIEZA_LEGACY = re.compile(
    r"^(?P<item>.+?)_(?P<tipo>"
    + "|".join(_TIPOS_PIEZA_EXPORT)
    + r")(?:_(?P<num>"
    + _RE_VALOR_NOMBRE
    + r"))?$",
    re.IGNORECASE,
)

_RE_CAPTURA_REF_NUEVA = re.compile(
    r"^(?P<job>.+?)"
    + re.escape(SEP)
    + r"(?P<item>.+?)"
    + re.escape(SEP)
    + r"(?P<etiqueta>"
    r"(?:p\d+of\d+_)?"
    r"(?:"
    r"(?:XMIN|XMAX|YMIN|YMAX|XCENTRO|YCENTRO)(?:_CONTACTO)?(?:_TYP)?"
    r"|QTY\d+(?:of\d+)?"
    r")"
    r")_(?P<num>"
    + _RE_VALOR_NOMBRE
    + r")$",
    re.IGNORECASE,
)

_RE_CAPTURA_REF_LEGACY = re.compile(
    r"^\d{3}_"
    r"(?:p\d+of\d+_)?"
    r"(?:"
    r"QTY\d+(?:of\d+)?_"
    r"|"
    r"(?:XMIN|XMAX|YMIN|YMAX|XCENTRO|YCENTRO)(?:_CONTACTO)?_(?:TYP_)?"
    r")"
    r"(?P<item>.+)$",
    re.IGNORECASE,
)


def limpiar_token_archivo(nombre: str) -> str:
    """Quita caracteres inválidos en Windows; colapsa separador reservado.

    No usa ``os.path.splitext`` ni ``rstrip('.')`` a ciegas: nombres como
    ``PIPE FLANGE 0.250`` perderían el decimal (``.250`` = “extensión”).

    Sí quita ``.iam`` / ``.ipt`` embebidos en DisplayName de Inventor
    (``MODELO VANTRAN.iam (Estado de modelo1)``).
    """
    texto = str(nombre or "").strip()
    # Extensión Inventor al final O embebida antes de espacio/( 
    texto = re.sub(r"\.(iam|ipt|idw|dwg)(?=\s|\(|$)", "", texto, flags=re.I)
    low = texto.casefold()
    for ext in (".iam", ".ipt", ".idw", ".dwg"):
        if low.endswith(ext):
            texto = texto[: -len(ext)]
            break
    texto = texto.replace(SEP, "_")
    texto = re.sub(r'[<>:"/\\|?*]', "_", texto)
    texto = texto.strip().rstrip(" ")
    if texto.endswith(".") and (len(texto) < 2 or not texto[-2].isdigit()):
        texto = texto[:-1].rstrip(" ")
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto or "SIN_NOMBRE"


def formatear_valor_en_nombre(valor) -> str:
    """
    Valor de cota → sufijo de archivo.

    Exacto a 6 decimales: 10.998200 | 38.100000 | 25.400000
    """
    if valor is None:
        return "0.000000"
    if isinstance(valor, (int, float)):
        v = abs(float(valor))
    else:
        texto = str(valor).strip().replace(",", ".")
        m = _RE_NUMERO_EN_TEXTO.search(texto)
        if not m:
            return limpiar_token_archivo(texto) or "0.000000"
        try:
            v = abs(float(m.group(0).replace(",", ".")))
        except ValueError:
            return limpiar_token_archivo(m.group(0)) or "0.000000"

    return f"{v:.6f}"


def nombre_job_desde_ensamble(ensamble) -> str:
    """DisplayName del ensamble → token JOB."""
    if ensamble is None:
        return "JOB"
    if isinstance(ensamble, str):
        return limpiar_token_archivo(ensamble)
    try:
        nombre = str(getattr(ensamble, "DisplayName", None) or "JOB")
    except Exception:
        nombre = "JOB"
    return limpiar_token_archivo(nombre)


def _separar_sufijo_inventor(nombre: str):
    m = _RE_SUFIJO_INV.match(str(nombre or "").strip())
    if m:
        return m.group("base"), m.group("num")
    return str(nombre or "").strip(), None


def medida_export_desde_hoja(nombre_hoja: str):
    """
    De un nombre de hoja técnico → (item, medida_export, num_hoja_opcional).

    El num_hoja es solo respaldo; el nombre final debe usar el valor de cota.
    """
    base, num = _separar_sufijo_inventor(nombre_hoja)
    if not base:
        return "ITEM", "LENGTH", num

    m_h = _RE_DIAMETRO_H.match(base)
    if m_h:
        item = m_h.group("item").strip() or "ITEM"
        # Flat: PART_DESPLIEGUE_DIAMETRO_H01 → item sin _DESPLIEGUE
        if item.upper().endswith("_DESPLIEGUE"):
            item = item[: -len("_DESPLIEGUE")]
        return (
            item,
            f"HOLE{m_h.group('nn')}",
            num,
        )

    base_up = base.upper()
    # PART_DESPLIEGUE_XCENTRO_02 → quitar índice de hoja al final.
    base_match = re.sub(r"_\d{2}$", "", base, count=1)
    base_match_up = base_match.upper()
    for interno, export in _TIPOS_HOJA_A_EXPORT:
        token = "_" + interno
        for candidato, candidato_up in (
            (base, base_up),
            (base_match, base_match_up),
        ):
            idx = candidato_up.rfind(token)
            if idx >= 0 and idx + len(token) == len(candidato_up):
                item = candidato[:idx].strip() or "ITEM"
                if item.upper().endswith("_DESPLIEGUE"):
                    item = item[: -len("_DESPLIEGUE")]
                return item, export, num

    return base.strip() or "ITEM", "LENGTH", num


def armar_nombre_captura_pieza(job, item, medida, valor_cota) -> str:
    """{JOB}__{ITEM}__{MEDIDA}_{VALOR_COTA}"""
    j = limpiar_token_archivo(job)
    i = limpiar_token_archivo(item)
    m = limpiar_token_archivo(str(medida or "LENGTH")).upper()
    n_txt = formatear_valor_en_nombre(valor_cota)
    return f"{j}{SEP}{i}{SEP}{m}_{n_txt}"


def armar_nombre_captura_referencia(job, item, etiqueta, valor_cota, typ=False) -> str:
    """{JOB}__{ITEM}__{XMIN[_TYP]}_{VALOR_COTA}"""
    j = limpiar_token_archivo(job)
    i = limpiar_token_archivo(item)
    et = str(etiqueta or "XMIN").upper().strip("_")
    if typ and not et.endswith("_TYP") and "TYP" not in et:
        et = f"{et}_TYP"
    n_txt = formatear_valor_en_nombre(valor_cota)
    return f"{j}{SEP}{i}{SEP}{et}_{n_txt}"


# Sufijo que agrega aplanar/reorg ante colisión de nombres en la misma carpeta.
_RE_SUFIJO_DUP = re.compile(r"__(?:dup|v)\d+$", re.IGNORECASE)


def _base_sin_dup(nombre_archivo: str) -> str:
    """Basename sin extensión y sin ``__dupN`` (colisiones de aplanar)."""
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    return _RE_SUFIJO_DUP.sub("", base)


def extraer_item_de_captura_pieza(nombre_archivo: str) -> str:
    """ITEM desde JPG Abigail (formato nuevo o legacy)."""
    base = _base_sin_dup(nombre_archivo)
    m = _RE_CAPTURA_PIEZA_NUEVA.match(base)
    if m:
        return m.group("item")
    m = _RE_CAPTURA_PIEZA_LEGACY.match(base)
    if m:
        return m.group("item")
    return base


def extraer_item_de_captura_referencia(nombre_archivo: str) -> str:
    """ITEM desde JPG Subensamble (formato nuevo o legacy)."""
    base = _base_sin_dup(nombre_archivo)
    m = _RE_CAPTURA_REF_NUEVA.match(base)
    if m:
        return m.group("item")
    m = _RE_CAPTURA_REF_LEGACY.match(base)
    if m:
        return m.group("item")
    return base
