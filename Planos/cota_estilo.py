import re

# Inventor usa cm para FontSize en estilos de dibujo
COTA_FONT_SIZE_CM = 0.18
COTA_NAVY_RGB = (0, 0, 128)
COTA_BOLD = True

_SIMBOLOS_TEXTO = re.compile(r"[Øø⌀°′″±]")
_UNIDADES_TEXTO = re.compile(r"\s*(in|mm|cm|ft|m|pulg\.?)\s*$", re.IGNORECASE)
_TAG_DIMENSION_VALUE = re.compile(r"<DimensionValue\s*/>", re.IGNORECASE)


def _precision_dimension(dimension):
    try:
        return max(0, int(dimension.Precision))
    except Exception:
        return 3


def texto_cota_limpio(valor, hoja=None, precision=None):
    """
    Convierte un valor numérico a texto de cota sin símbolos ni unidades.
    """
    try:
        valor = abs(float(valor))
    except (TypeError, ValueError):
        return ""

    if precision is None:
        precision = 3

    try:
        if hoja is not None:
            doc = hoja.Parent
            uom = doc.UnitsOfMeasure
            s = str(uom.GetStringFromValue(valor, uom.LengthUnits))
            texto = s.split()[0].strip()
        else:
            texto = f"{valor:.{precision}f}".rstrip("0").rstrip(".")
    except Exception:
        texto = f"{valor:.{precision}f}".rstrip("0").rstrip(".")

    texto = _SIMBOLOS_TEXTO.sub("", texto)
    texto = _UNIDADES_TEXTO.sub("", texto)
    texto = re.sub(r"^[A-Za-z]+", "", texto).strip()

    if texto.startswith("-."):
        texto = "-0" + texto[1:]
    elif texto.startswith("."):
        texto = "0" + texto

    return texto


def _texto_desde_dimension(dimension, hoja=None):
    precision = _precision_dimension(dimension)

    try:
        raw = str(dimension.Text.Text).strip()
        if raw:
            limpio = _SIMBOLOS_TEXTO.sub("", raw)
            limpio = _UNIDADES_TEXTO.sub("", limpio).strip()
            limpio = re.sub(r"^[A-Za-z]+", "", limpio).strip()
            if limpio and re.search(r"\d", limpio):
                return limpio
    except Exception:
        pass

    try:
        return texto_cota_limpio(dimension.ModelValue, hoja, precision)
    except Exception:
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
    Aplica negrita, azul marino, tamaño mayor y texto numérico sin símbolos.
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

    _limpiar_prefijos_cota(dimension)

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


def aplicar_estilo_texto_cota(text_obj, texto, inv_app):
    """
    Mismo estilo para TextBoxes de sketch (arcos.py).
    """
    if text_obj is None or not texto:
        return

    texto = _SIMBOLOS_TEXTO.sub("", str(texto)).strip()
    texto = _UNIDADES_TEXTO.sub("", texto).strip()
    if not texto:
        return

    bold = "True" if COTA_BOLD else "False"
    formatted = (
        f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' Bold='{bold}'>"
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
