import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota

kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

EPS = 0.0001
OFFSET_COTA = 1.5

FACTOR_VALIDACION_MIN = 0.80
FACTOR_VALIDACION_MAX = 1.15


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


def _puntos_de_curva(curva):
    """
    Para piezas inclinadas conviene usar vértices reales.
    Solo usamos StartPoint y EndPoint.
    """
    puntos = []
    usados = set()

    for attr in ("StartPoint", "MidPoint", "EndPoint"):
        try:
            p = getattr(curva, attr)
            if p:
                x = float(p.X)
                y = float(p.Y)
                key = (round(x, 6), round(y, 6))
                if key not in usados:
                    usados.add(key)
                    puntos.append((x, y, p))
        except:
            pass

    return puntos


def _obtener_curvas_con_puntos(vista):
    datos = []

    for j in range(1, vista.DrawingCurves.Count + 1):
        curva = vista.DrawingCurves.Item(j)
        info = _bbox_curva(curva)
        if not info:
            continue

        pts = _puntos_de_curva(curva)
        if len(pts) < 2:
            continue

        # descartamos curvas degeneradas
        x1, y1, _ = pts[0]
        x2, y2, _ = pts[-1]
        if abs(x1 - x2) < EPS and abs(y1 - y2) < EPS:
            continue

        info["pts"] = pts
        datos.append(info)

    return datos


def _recolectar_puntos(datos):
    puntos = []
    usados = set()

    for d in datos:
        for x, y, p in d["pts"]:
            key = (round(x, 6), round(y, 6))
            if key not in usados:
                usados.add(key)
                puntos.append({
                    "x": x,
                    "y": y,
                    "pt": p,
                    "curve": d["curve"]
                })

    return puntos


def _bbox_puntos(puntos):
    minx = min(p["x"] for p in puntos)
    maxx = max(p["x"] for p in puntos)
    miny = min(p["y"] for p in puntos)
    maxy = max(p["y"] for p in puntos)
    return minx, maxx, miny, maxy


def _esperado_modelo(vista, span_sheet):
    try:
        escala = float(vista.Scale)
        if abs(escala) < EPS:
            escala = 1.0
    except:
        escala = 1.0
    return abs(span_sheet / escala)


def _validar_dimension(dimension, esperado_modelo, nombre_hoja, eje):
    try:
        valor = abs(float(dimension.ModelValue))
    except:
        return True

    minimo = esperado_modelo * FACTOR_VALIDACION_MIN
    maximo = esperado_modelo * FACTOR_VALIDACION_MAX

    if valor < minimo or valor > maximo:
        try:
            dimension.Delete()
        except:
            pass

        print(
            f"⚠️ {nombre_hoja}: cota especial {eje} descartada "
            f"(valor={valor:.3f}, esperado≈{esperado_modelo:.3f})"
        )
        return False

    return True

def _tiene_extremo_curvo(datos, lado, tol=0.001):
    """
    Detecta si el extremo superior o inferior real depende de un arco/redondeo
    y no de una recta horizontal clara.

    lado = "sup" o "inf"
    """
    if not datos:
        return False

    try:
        if lado == "sup":
            extremo = max(d["maxy"] for d in datos)
            candidatos = [d for d in datos if abs(d["maxy"] - extremo) <= tol]
        else:
            extremo = min(d["miny"] for d in datos)
            candidatos = [d for d in datos if abs(d["miny"] - extremo) <= tol]

        if not candidatos:
            return False

        # Si en el extremo hay una recta horizontal clara, no mandar a arcos.py
        for d in candidatos:
            dx = float(d["dx"])
            dy = float(d["dy"])

            if dx > 0.20 and dy <= 0.05:
                return False

        # Si no hay recta horizontal clara y sí hay curva/redondeo, mandar a arcos.py
        for d in candidatos:
            dx = float(d["dx"])
            dy = float(d["dy"])

            if dx > 0.05 and dy > 0.05:
                return True

        return False

    except:
        return False

def _tiene_curvas_reales(datos):
    """
    Detecta si en la hoja hay curvas/redondeos reales.
    """
    try:
        for d in datos:
            dx = float(d["dx"])
            dy = float(d["dy"])

            if dx > 0.05 and dy > 0.05:
                return True

        return False

    except:
        return False

def _crear_intent_punto(hoja, item):
    try:
        return hoja.CreateGeometryIntent(item["curve"], item["pt"])
    except:
        return None


def _crear_cota_horizontal_por_puntos(hoja, vista, tg, datos, puntos, nombre_hoja):
    minx = min(d["minx"] for d in datos)
    maxx = max(d["maxx"] for d in datos)
    miny = min(d["miny"] for d in datos)
    maxy = max(d["maxy"] for d in datos)

    izq = min(puntos, key=lambda p: abs(p["x"] - minx))
    der = min(puntos, key=lambda p: abs(p["x"] - maxx))

    if abs(izq["x"] - der["x"]) < EPS:
        return False

    int_izq = _crear_intent_punto(hoja, izq)
    int_der = _crear_intent_punto(hoja, der)

    if not int_izq or not int_der:
        return False

    try:
        pt_texto = tg.CreatePoint2d((minx + maxx) / 2.0, maxy + OFFSET_COTA)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_izq, int_der, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        esperado = _esperado_modelo(vista, maxx - minx)
        return _validar_dimension(dim, esperado, nombre_hoja, "horizontal")

    except:
        return False


def _crear_cota_vertical_por_puntos(hoja, vista, tg, puntos, nombre_hoja):
    minx, maxx, miny, maxy = _bbox_puntos(puntos)

    inf = min(puntos, key=lambda p: p["y"])
    sup = max(puntos, key=lambda p: p["y"])

    if abs(inf["y"] - sup["y"]) < EPS:
        return False

    int_inf = _crear_intent_punto(hoja, inf)
    int_sup = _crear_intent_punto(hoja, sup)

    if not int_inf or not int_sup:
        return False

    try:
        pt_texto = tg.CreatePoint2d(minx - OFFSET_COTA, (miny + maxy) / 2.0)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inf, int_sup, kVerticalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        esperado = _esperado_modelo(vista, maxy - miny)
        return _validar_dimension(dim, esperado, nombre_hoja, "vertical")

    except:
        return False


def acotar_especiales(hojas_objetivo):
    print("🧩 Iniciando módulo lineal_especial.py...")

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return hojas_objetivo[:]

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
        datos = _obtener_curvas_con_puntos(vista)

        if not datos:
            pendientes.append(nombre_hoja)
            continue

        puntos = _recolectar_puntos(datos)
        if len(puntos) < 2:
            pendientes.append(nombre_hoja)
            continue

        ok = False

        if "_FRENTE_1" in nombre_hoja:
            ok = _crear_cota_horizontal_por_puntos(hoja, vista, tg, datos, puntos, nombre_hoja)

        elif "_FRENTE_2" in nombre_hoja:
           print(f"🌙 {nombre_hoja}: _FRENTE_2 especial se deja para arcos.py")
           ok = False
           
        if ok:
            print(f"✅ {nombre_hoja}: resuelta por lineal_especial.py")
        else:
            print(f"⚠️ {nombre_hoja}: no pudo resolverse en lineal_especial.py")
            pendientes.append(nombre_hoja)

    print("✅ lineal_especial.py terminado.")
    return pendientes