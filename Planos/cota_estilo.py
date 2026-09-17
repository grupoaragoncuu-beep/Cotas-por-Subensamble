import os
import re
import math

# Inventor usa cm para FontSize en estilos de dibujo.
# Base 0.18 → 0.225 (+25%) → 0.3375 (+50%) → 0.50625 (+50% adicional).
COTA_FONT_SIZE_CM = 0.50625
COTA_NAVY_RGB = (0, 0, 128)
COTA_BOLD = True
# Unidad de visualización: tanques Abigail = in; GIGA/BOARD = mm.
# Cambiar con ``set_unidad_cota("mm"|"in")`` al inicio del flujo.
# SOLO_FLAT_CORTE=1 fuerza mm siempre (GIGA Board flat).
COTA_UNIDAD = "in"
_UNIDAD_ACTIVA = "in"


def _env_fuerza_mm() -> bool:
    return os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "si",
        "on",
    )
# Holgura mínima entre el texto/línea de cota y la silueta de la pieza.
# Con fuente grande, 1.5 cm era insuficiente (Vantran PIPE FLANGE: texto
# encima del filete / cota horizontal atravesando el cuerpo).
OFFSET_FUERA_PIEZA_CM = 2.85
MARGEN_COTA_VS_PIEZA_CM = 0.45

_SIMBOLOS_TEXTO = re.compile(r"[Øø⌀°′″±]")
_UNIDADES_TEXTO = re.compile(
    r"(?:^|\s)(in|mm|cm|ft|m|pulg\.?)(?=\s|$)", re.IGNORECASE
)
_TAG_DIMENSION_VALUE = re.compile(r"<DimensionValue\s*/>", re.IGNORECASE)


def set_unidad_cota(unidad="in"):
    """
    Fija la unidad de texto de cotas para la corrida actual.

    ``in`` = tanques Vantran/OTC (default).
    ``mm`` = GIGA / BOARD / cobre de gabinete.
    """
    global COTA_UNIDAD, _UNIDAD_ACTIVA
    u = str(unidad or "in").strip().lower()
    if u in ("mm", "millimeter", "millimetre", "milimetro", "milímetro"):
        u = "mm"
    else:
        u = "in"
    COTA_UNIDAD = u
    _UNIDAD_ACTIVA = u
    return u


def get_unidad_cota():
    # Flat Corte/Corte GIGA: NUNCA pulgadas, aunque alguien resetee a "in".
    if _env_fuerza_mm():
        return "mm"
    return str(_UNIDAD_ACTIVA or COTA_UNIDAD or "in")


def _precision_dimension(dimension):
    """Precision de dibujo: 6 decimales exactos."""
    return _precision_default()


def _precision_default():
    """6 decimales exactos (sin recortar)."""
    return 6


def _strip_unidades(texto):
    """Quita tokens de unidad sin tocar el resto (p. ej. TYP)."""
    t = _UNIDADES_TEXTO.sub(" ", str(texto or ""))
    return re.sub(r"\s+", " ", t).strip()


def asegurar_unidad_cota(texto):
    """
    Garantiza sufijo de unidad activa (`` in`` o `` mm``).

    Ejemplos: ``4.25`` → ``4.25 in``; ``108.0 TYP`` → ``108.0 TYP mm``.
    """
    t = _SIMBOLOS_TEXTO.sub("", str(texto or "")).strip()
    t = _strip_unidades(t)
    if not t:
        return t
    return f"{t} {get_unidad_cota()}"


# Compat: nombre histórico usado en todo el repo.
asegurar_unidad_pulgadas = asegurar_unidad_cota


def texto_cota_limpio(valor, hoja=None, precision=None):
    """
    Convierte un valor numerico (cm de Inventor) a texto de cota sin unidad.

    Exacto desde DB Inventor (cm):
      mm = cm * 10
      in = cm / 2.54
    6 decimales (más exacto; no truncar a 3).
    """
    try:
        valor = abs(float(valor))
    except (TypeError, ValueError):
        return ""

    if precision is None:
        precision = _precision_default()
    else:
        try:
            precision = int(precision)
        except (TypeError, ValueError):
            precision = _precision_default()
    precision = max(6, precision)

    unidad = get_unidad_cota()
    try:
        if unidad == "mm":
            num = float(valor) * 10.0
        else:
            num = float(valor) / 2.54
        texto = f"{num:.{precision}f}"
    except Exception:
        if unidad == "mm":
            texto = f"{(float(valor) * 10.0):.{precision}f}"
        else:
            texto = f"{(float(valor) / 2.54):.{precision}f}"

    texto = _SIMBOLOS_TEXTO.sub("", texto)
    texto = _strip_unidades(texto)
    texto = re.sub(r"^[A-Za-z]+", "", texto).strip()

    if texto.startswith("-."):
        texto = "-0" + texto[1:]
    elif texto.startswith("."):
        texto = "0" + texto

    return texto


def texto_cota_dibujo(valor, hoja=None, precision=None):
    """Texto de cota listo para dibujar, con unidad activa (`` in`` / `` mm``)."""
    base = texto_cota_limpio(valor, hoja, precision)
    return asegurar_unidad_cota(base) if base else ""


def _texto_desde_dimension(dimension, hoja=None):
    precision = _precision_dimension(dimension)

    # Preferir ModelValue (cm API) → conversión según unidad activa (in/mm).
    try:
        return texto_cota_limpio(dimension.ModelValue, hoja, precision)
    except Exception:
        pass

    try:
        raw = str(dimension.Text.Text).strip()
        if raw:
            limpio = _SIMBOLOS_TEXTO.sub("", raw)
            limpio = _strip_unidades(limpio).strip()
            limpio = re.sub(r"^[A-Za-z]+", "", limpio).strip()
            if limpio and re.search(r"\d", limpio):
                # Si la unidad activa es mm y el texto venía en in del machote,
                # re-leer ModelValue; si no hay, dejar el número tal cual.
                if get_unidad_cota() == "mm":
                    try:
                        return texto_cota_limpio(
                            dimension.ModelValue, hoja, precision
                        )
                    except Exception:
                        pass
                return limpio
    except Exception:
        pass

    return ""


def _obtener_inv_app(hoja=None, inv_app=None):
    if inv_app is not None:
        return inv_app
    if hoja is None:
        return None
    try:
        return hoja.Parent.Application
    except Exception:
        return None


def _limpiar_prefijos_cota(dimension):
    """
    Quita prefijos/sufijos automáticos (Ø, R, etc.) del objeto de texto.
    """
    try:
        dimension.Text.Prefix = ""
    except Exception:
        pass

    try:
        dimension.Text.Suffix = ""
    except Exception:
        pass

    try:
        dimension.Text.PrefixSymbol = ""
    except Exception:
        pass


def aplicar_estilo_cota(dimension, inv_app=None, hoja=None, solo_color=False):
    """
    Aplica negrita, azul marino, tamaño mayor y texto numérico con `` in``.
    Usa solo texto literal (sin <DimensionValue/>) para evitar duplicados.

    solo_color=True: no toca HideValue ni FormattedText (evita romper
    cotas ordenadas dejando solo numeros flotantes sin lineas).
    """
    if dimension is None:
        return

    if hoja is None:
        try:
            hoja = dimension.Parent
        except Exception:
            hoja = None

    app = _obtener_inv_app(hoja, inv_app)

    if solo_color:
        if app is not None:
            try:
                r, g, b = COTA_NAVY_RGB
                color = app.TransientObjects.CreateColor(r, g, b)
                dimension.Text.Color = color
            except Exception:
                pass
        return

    texto = _texto_desde_dimension(dimension, hoja)
    if not texto:
        return
    texto = asegurar_unidad_cota(texto)

    _limpiar_prefijos_cota(dimension)

    # Forzar precision Inventor = 6
    try:
        dimension.Precision = 6
    except Exception:
        try:
            dimension.Precision = _precision_default()
        except Exception:
            pass

    try:
        dimension.HideValue = True
    except Exception:
        pass

    bold = "True" if COTA_BOLD else "False"
    formatted = (
        f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' Bold='{bold}'>"
        f"{texto}</StyleOverride>"
    )

    try:
        dimension.Text.FormattedText = formatted
    except Exception:
        try:
            dimension.Text.Text = texto
        except Exception:
            return

    # Seguridad: si quedó mezclado con <DimensionValue/>, forzar solo literal
    try:
        actual = str(dimension.Text.FormattedText)
        if _TAG_DIMENSION_VALUE.search(actual):
            dimension.HideValue = True
            dimension.Text.FormattedText = formatted
    except Exception:
        pass

    if app is None:
        return

    try:
        r, g, b = COTA_NAVY_RGB
        color = app.TransientObjects.CreateColor(r, g, b)
        dimension.Text.Color = color
    except Exception:
        pass


def aplicar_estilo_texto_cota(text_obj, texto, inv_app, vertical=False):
    """
    Mismo estilo para TextBoxes de sketch (arcos / caras / TYP).
    Conserva etiquetas (TYP, letras) y fuerza sufijo de unidad activa
    (`` in`` / `` mm``) cuando hay número.
    """
    if text_obj is None or not texto:
        return

    texto = _SIMBOLOS_TEXTO.sub("", str(texto)).strip()
    # Letras sueltas A/B/C de marcas TYP: sin unidad.
    if re.fullmatch(r"[A-Z]{1,2}", texto):
        pass
    elif re.search(r"\d", texto):
        texto = asegurar_unidad_cota(texto)
    if not texto:
        return

    bold = "True" if COTA_BOLD else "False"
    angle_attr = " Angle='90'" if vertical else ""
    formatted = (
        f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' Bold='{bold}'{angle_attr}>"
        f"{texto}</StyleOverride>"
    )

    try:
        text_obj.FormattedText = formatted
    except Exception:
        try:
            text_obj.Text = texto
        except Exception:
            pass

    try:
        text_obj.Style.Bold = COTA_BOLD
        text_obj.Style.FontSize = COTA_FONT_SIZE_CM
    except Exception:
        pass

    try:
        r, g, b = COTA_NAVY_RGB
        color = inv_app.TransientObjects.CreateColor(r, g, b)
        text_obj.Color = color
    except Exception:
        pass


def letra_typ_indice(indice):
    """0→A … 25→Z, 26→AA…"""
    i = max(0, int(indice))
    if i < 26:
        return chr(ord("A") + i)
    return letra_typ_indice(i // 26 - 1) + chr(ord("A") + (i % 26))


# Cotas TYP con letras A/B/C: ON por defecto.
# Solo cobre/busbar (ABB/GENE/RLG) las apaga en barrenos_xy_despliegue.
_ENV_TYP_LETRAS = "COTAS_TYP_LETRAS"


def typ_letras_habilitadas() -> bool:
    """True → dibujar A/B/C en marcas TYP. False en cobre/busbar."""
    v = str(os.environ.get(_ENV_TYP_LETRAS, "1") or "1").strip().lower()
    return v not in ("0", "false", "no", "off")


def set_typ_letras_habilitadas(activado: bool) -> None:
    """Activa/desactiva letras TYP (proceso actual)."""
    os.environ[_ENV_TYP_LETRAS] = "1" if activado else "0"


def clearance_texto_cota_cm(n_chars=8):
    """
    Separación mínima silueta → texto para que el número no monte la pieza.

    Escala con ``COTA_FONT_SIZE_CM`` (negrita navy).
    """
    n = max(4, int(n_chars or 8))
    return max(
        float(OFFSET_FUERA_PIEZA_CM),
        float(COTA_FONT_SIZE_CM) * n * 0.55 + 1.15,
    )


def _norm_bbox(bbox):
    """(minx, maxx, miny, maxy) desde tupla/lista/dict."""
    if bbox is None:
        return None
    if isinstance(bbox, dict):
        try:
            return (
                float(bbox["minx"]),
                float(bbox["maxx"]),
                float(bbox["miny"]),
                float(bbox["maxy"]),
            )
        except Exception:
            return None
    try:
        minx, maxx, miny, maxy = bbox
        return float(minx), float(maxx), float(miny), float(maxy)
    except Exception:
        return None


def bbox_intersecta(a, b, holgura=0.0):
    """True si dos bboxes 2D se solapan (con holgura expandiendo ``a``)."""
    aa = _norm_bbox(a)
    bb = _norm_bbox(b)
    if aa is None or bb is None:
        return False
    h = float(holgura or 0.0)
    ax0, ax1, ay0, ay1 = aa[0] - h, aa[1] + h, aa[2] - h, aa[3] + h
    bx0, bx1, by0, by1 = bb
    return not (ax1 < bx0 or bx1 < ax0 or ay1 < by0 or by1 < ay0)


def punto_dentro_bbox(x, y, bbox, holgura=0.0):
    b = _norm_bbox(bbox)
    if b is None:
        return False
    h = float(holgura or 0.0)
    return (
        b[0] - h <= float(x) <= b[1] + h
        and b[2] - h <= float(y) <= b[3] + h
    )


def empujar_punto_fuera_bbox(x, y, bbox, lado, clearance=None):
    """Empuja (x,y) fuera del bbox por el lado pedido."""
    b = _norm_bbox(bbox)
    if b is None:
        return float(x), float(y)
    clr = float(clearance if clearance is not None else OFFSET_FUERA_PIEZA_CM)
    minx, maxx, miny, maxy = b
    cx = 0.5 * (minx + maxx)
    cy = 0.5 * (miny + maxy)
    lado = str(lado or "").lower()
    if lado == "izq":
        return minx - clr, float(y) if y is not None else cy
    if lado == "der":
        return maxx + clr, float(y) if y is not None else cy
    if lado == "sup":
        return float(x) if x is not None else cx, maxy + clr
    if lado == "inf":
        return float(x) if x is not None else cx, miny - clr
    # auto: lado con más aire respecto al punto actual
    dist = {
        "izq": float(x) - minx,
        "der": maxx - float(x),
        "inf": float(y) - miny,
        "sup": maxy - float(y),
    }
    mejor = min(dist, key=dist.get)
    return empujar_punto_fuera_bbox(x, y, b, mejor, clr)


def dim_solapa_pieza(dim, pieza_bbox, holgura=None, solo_texto=False, n_chars=10):
    """
    True si la cota invade la silueta de la pieza.

    Por defecto prioriza el **texto** (visión/OCR). ``solo_texto=True``
    ignora el RangeBox completo de la cota (líderes/flechas hacia barrenos
    siempre cruzan la pieza y no deben contar como solape).
    """
    pieza = _norm_bbox(pieza_bbox)
    if dim is None or pieza is None:
        return False
    h = float(
        holgura if holgura is not None else MARGEN_COTA_VS_PIEZA_CM
    )
    # 1) Caja del texto (o estimación alrededor del origen).
    try:
        tb = None
        try:
            rb_t = dim.Text.RangeBox
            tb = (
                float(rb_t.MinPoint.X),
                float(rb_t.MaxPoint.X),
                float(rb_t.MinPoint.Y),
                float(rb_t.MaxPoint.Y),
            )
        except Exception:
            o = dim.Text.Origin
            ox, oy = float(o.X), float(o.Y)
            hw = float(COTA_FONT_SIZE_CM) * max(4, int(n_chars or 10)) * 0.38
            hh = float(COTA_FONT_SIZE_CM) * 0.95
            tb = (ox - hw, ox + hw, oy - hh, oy + hh)
        if tb is not None and bbox_intersecta(pieza, tb, holgura=h):
            return True
        if solo_texto:
            return False
    except Exception:
        if solo_texto:
            return False
    # 2) RangeBox completo (lineales / legado).
    if not solo_texto:
        try:
            rb = dim.RangeBox
            dim_bb = (
                float(rb.MinPoint.X),
                float(rb.MaxPoint.X),
                float(rb.MinPoint.Y),
                float(rb.MaxPoint.Y),
            )
            if bbox_intersecta(pieza, dim_bb, holgura=h):
                return True
        except Exception:
            pass
        try:
            o = dim.Text.Origin
            if punto_dentro_bbox(float(o.X), float(o.Y), pieza, holgura=h):
                return True
        except Exception:
            pass
    return False


def candidatos_texto_fuera_pieza(
    pieza_bbox, orientacion, clearance=None, extras=None
):
    """
    Lista de (x, y, lado) fuera de la silueta.

    ``orientacion``: ``V`` (cota vertical → texto a izq/der) o ``H``
    (cota horizontal → texto arriba/abajo).
    """
    b = _norm_bbox(pieza_bbox)
    if b is None:
        return []
    minx, maxx, miny, maxy = b
    cx = 0.5 * (minx + maxx)
    cy = 0.5 * (miny + maxy)
    clr0 = float(clearance if clearance is not None else clearance_texto_cota_cm())
    clearances = [clr0, clr0 + 1.2, clr0 + 2.4, clr0 + 3.6]
    if extras:
        clearances.extend(float(v) for v in extras)
    lados = ("izq", "der") if str(orientacion).upper().startswith("V") else ("sup", "inf")
    out = []
    for clr in clearances:
        for lado in lados:
            if lado == "izq":
                out.append((minx - clr, cy, lado))
            elif lado == "der":
                out.append((maxx + clr, cy, lado))
            elif lado == "sup":
                out.append((cx, maxy + clr, lado))
            else:
                out.append((cx, miny - clr, lado))
    return out


def asegurar_cota_fuera_pieza(dim, tg, pieza_bbox, orientacion, holgura=None):
    """
    Si la cota solapa la pieza, mueve ``Text.Origin`` fuera.

    Devuelve True si quedó libre (o ya lo estaba); False si sigue solapando.
    """
    if dim is None:
        return False
    if not dim_solapa_pieza(dim, pieza_bbox, holgura=holgura, solo_texto=True):
        return True
    for x, y, _lado in candidatos_texto_fuera_pieza(pieza_bbox, orientacion):
        try:
            dim.Text.Origin = tg.CreatePoint2d(float(x), float(y))
        except Exception:
            continue
        if not dim_solapa_pieza(dim, pieza_bbox, holgura=holgura, solo_texto=True):
            return True
    return not dim_solapa_pieza(
        dim, pieza_bbox, holgura=holgura, solo_texto=True
    )


def asegurar_cota_fuera_pieza_robusto(
    dim, tg, pieza_bbox, holgura=None, n_chars=10
):
    """
    Variante agresiva para Ø / arcos / barrenos: prueba H y V con holguras
    crecientes hasta que el texto no invada la silueta (visión OCR).
    """
    if dim is None:
        return False
    h0 = float(
        holgura if holgura is not None else max(MARGEN_COTA_VS_PIEZA_CM, 0.55)
    )
    if not dim_solapa_pieza(
        dim, pieza_bbox, holgura=h0, solo_texto=True, n_chars=n_chars
    ):
        return True
    base = clearance_texto_cota_cm(n_chars)
    extras_list = (
        None,
        (base + 1.5, base + 3.0),
        (base + 4.5, base + 6.0, base + 8.0),
    )
    for extras in extras_list:
        for orient in ("V", "H"):
            cands = candidatos_texto_fuera_pieza(
                pieza_bbox,
                orient,
                clearance=base if extras is None else None,
                extras=extras,
            )
            for x, y, _lado in cands:
                try:
                    dim.Text.Origin = tg.CreatePoint2d(float(x), float(y))
                except Exception:
                    continue
                if not dim_solapa_pieza(
                    dim,
                    pieza_bbox,
                    holgura=h0,
                    solo_texto=True,
                    n_chars=n_chars,
                ):
                    return True
            if asegurar_cota_fuera_pieza(dim, tg, pieza_bbox, orient, holgura=h0):
                return True
    return not dim_solapa_pieza(
        dim, pieza_bbox, holgura=h0, solo_texto=True, n_chars=n_chars
    )


def offset_letra_typ(cx, cy, radio, ocupados, sep_min=0.38):
    """
    Posición junto al círculo evitando solape con otras letras/centros.
    Prueba derecha → alrededor en octantes, luego radios mayores.
    """
    cx, cy, radio = float(cx), float(cy), float(radio)
    dirs = (
        (1.0, 0.0),
        (0.75, 0.75),
        (0.0, 1.0),
        (-0.75, 0.75),
        (-1.0, 0.0),
        (-0.75, -0.75),
        (0.0, -1.0),
        (0.75, -0.75),
    )
    for escala in (1.0, 1.35, 1.7, 2.1):
        off = (radio + 0.22) * escala
        for dx, dy in dirs:
            px = cx + dx * off
            py = cy + dy * off
            if all(
                math.hypot(px - ox, py - oy) >= sep_min for ox, oy in ocupados
            ):
                return px, py
    return cx + radio + 0.35, cy
