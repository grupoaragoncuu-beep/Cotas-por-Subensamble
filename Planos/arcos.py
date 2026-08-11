import math
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_texto_cota, texto_cota_limpio

EPS = 0.0001


def _bbox_curva(curva):
    try:
        caja = curva.Evaluator2D.RangeBox

        minx = float(caja.MinPoint.X)
        maxx = float(caja.MaxPoint.X)
        miny = float(caja.MinPoint.Y)
        maxy = float(caja.MaxPoint.Y)

        dx = abs(maxx - minx)
        dy = abs(maxy - miny)

        if dx < EPS and dy < EPS:
            return None

        return {
            "curve": curva,
            "minx": minx,
            "maxx": maxx,
            "miny": miny,
            "maxy": maxy,
            "dx": dx,
            "dy": dy
        }
    except:
        return None


def _obtener_curvas_validas(vista):
    datos = []
    for j in range(1, vista.DrawingCurves.Count + 1):
        curva = vista.DrawingCurves.Item(j)
        info = _bbox_curva(curva)
        if info:
            datos.append(info)
    return datos


def _bbox_global(datos):
    minx = min(d["minx"] for d in datos)
    maxx = max(d["maxx"] for d in datos)
    miny = min(d["miny"] for d in datos)
    maxy = max(d["maxy"] for d in datos)
    return minx, maxx, miny, maxy


def _esperado_modelo(vista, span_sheet):
    try:
        escala = float(vista.Scale)
        if abs(escala) < EPS:
            escala = 1.0
    except:
        escala = 1.0
    return abs(span_sheet / escala)


def _texto_valor(hoja, valor_modelo):
    try:
        uom = hoja.Parent.UnitsOfMeasure
        s = str(uom.GetStringFromValue(valor_modelo, uom.LengthUnits))
        texto = s.split()[0].strip()

        if texto.startswith("-."):
            texto = "-0" + texto[1:]
        elif texto.startswith("."):
            texto = "0" + texto

        return texto
    except:
        texto = f"{valor_modelo:.2f}".strip()

        if texto.startswith("-."):
            texto = "-0" + texto[1:]
        elif texto.startswith("."):
            texto = "0" + texto

        return texto


def _obtener_o_crear_sketch(hoja, nombre="__AUTO_ARCOS_SKETCH__"):
    sketch = None

    try:
        for i in range(1, hoja.Sketches.Count + 1):
            s = hoja.Sketches.Item(i)
            if str(s.Name).upper() == nombre.upper():
                sketch = s
                break
    except:
        pass

    if sketch is None:
        sketch = hoja.Sketches.Add()
        sketch.Name = nombre

    return sketch


def _limpiar_sketch(sketch):
    try:
        while sketch.SketchLines.Count > 0:
            sketch.SketchLines.Item(1).Delete()
    except:
        pass

    try:
        while sketch.TextBoxes.Count > 0:
            sketch.TextBoxes.Item(1).Delete()
    except:
        pass


def _crear_linea_horizontal(sketch, tg, x1, x2, y):
    p1 = tg.CreatePoint2d(x1, y)
    p2 = tg.CreatePoint2d(x2, y)
    return sketch.SketchLines.AddByTwoPoints(p1, p2)


def _crear_linea_vertical(sketch, tg, x, y1, y2):
    p1 = tg.CreatePoint2d(x, y1)
    p2 = tg.CreatePoint2d(x, y2)
    return sketch.SketchLines.AddByTwoPoints(p1, p2)


def _crear_texto(sketch, tg, inv_app, hoja, x, y, valor_modelo, rotacion_rad=0.0):
    texto = texto_cota_limpio(valor_modelo, hoja)
    if not texto:
        return None

    pt = tg.CreatePoint2d(x, y)
    tb = sketch.TextBoxes.AddFitted(pt, texto)
    try:
        tb.Rotation = rotacion_rad
    except:
        pass
    aplicar_estilo_texto_cota(tb, texto, inv_app)
    return tb


def _crear_cota_vertical_visual(hoja, vista, tg, datos, nombre_hoja, inv_app):
    sketch = None

    try:
        minx, maxx, miny, maxy = _bbox_global(datos)

        ancho = maxx - minx
        alto = maxy - miny

        if alto < EPS:
            return False

        valor_modelo = _esperado_modelo(vista, alto)

        sketch = _obtener_o_crear_sketch(hoja)

        try:
            sketch.Edit()
        except:
            pass

        _limpiar_sketch(sketch)

        x2 = minx - max(0.20, ancho * 0.06)
        x1 = x2 - max(0.40, ancho * 0.12)

        _crear_linea_horizontal(sketch, tg, x1, x2, miny)
        _crear_linea_horizontal(sketch, tg, x1, x2, maxy)
        _crear_linea_vertical(sketch, tg, x1, miny, maxy)

        _crear_texto(
            sketch,
            tg,
            inv_app,
            hoja,
            x1 - max(0.18, ancho * 0.04),
            (miny + maxy) / 2.0,
            valor_modelo,
            math.pi / 2.0
        )

        try:
            sketch.ExitEdit()
        except:
            pass

        print(f"DEBUG {nombre_hoja}: vertical_visual aceptada -> {texto_cota_limpio(valor_modelo, hoja)}")
        return True

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: error en vertical_visual -> {e}")

        try:
            if sketch is not None:
                sketch.ExitEdit()
        except:
            pass

        return False


def _crear_cota_horizontal_visual(hoja, vista, tg, datos, nombre_hoja, inv_app):
    sketch = None

    try:
        minx, maxx, miny, maxy = _bbox_global(datos)

        ancho = maxx - minx
        alto = maxy - miny

        if ancho < EPS:
            return False

        valor_modelo = _esperado_modelo(vista, ancho)

        sketch = _obtener_o_crear_sketch(hoja)

        try:
            sketch.Edit()
        except:
            pass

        _limpiar_sketch(sketch)

        y1 = maxy + max(0.20, alto * 0.06)
        y2 = y1 + max(0.40, alto * 0.12)

        _crear_linea_vertical(sketch, tg, minx, y1, y2)
        _crear_linea_vertical(sketch, tg, maxx, y1, y2)
        _crear_linea_horizontal(sketch, tg, minx, maxx, y2)

        _crear_texto(
            sketch,
            tg,
            inv_app,
            hoja,
            (minx + maxx) / 2.0,
            y2 + max(0.08, alto * 0.02),
            valor_modelo,
            0.0
        )

        try:
            sketch.ExitEdit()
        except:
            pass

        print(f"DEBUG {nombre_hoja}: horizontal_visual aceptada -> {texto_cota_limpio(valor_modelo, hoja)}")
        return True

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: error en horizontal_visual -> {e}")

        try:
            if sketch is not None:
                sketch.ExitEdit()
        except:
            pass

        return False


def acotar_arcos(hojas_objetivo):
    print("🌙 arcos.py: Iniciando módulo final para piezas con arcos...")

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return list(hojas_objetivo)

    tg = inv_app.TransientGeometry
    pendientes = []
    objetivo_set = set(h.upper() for h in hojas_objetivo)

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja = str(hoja.Name).upper()

        if nombre_hoja not in objetivo_set:
            continue

        if hoja.DrawingViews.Count == 0:
            pendientes.append(nombre_hoja)
            continue

        vista = hoja.DrawingViews.Item(1)
        datos = _obtener_curvas_validas(vista)

        if not datos:
            pendientes.append(nombre_hoja)
            continue

        ok = False

        if "_FRENTE_2" in nombre_hoja:
            ok = _crear_cota_vertical_visual(hoja, vista, tg, datos, nombre_hoja, inv_app)
        elif "_FRENTE_1" in nombre_hoja:
            ok = _crear_cota_horizontal_visual(hoja, vista, tg, datos, nombre_hoja, inv_app)

        if ok:
            print(f"✅ {nombre_hoja}: resuelta por arcos.py")
        else:
            print(f"⚠️ {nombre_hoja}: no pudo resolverse en arcos.py")
            pendientes.append(nombre_hoja)

    print("✅ arcos.py terminado.")
    return pendientes