import os
import win32com.client
import diametro
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota
import lineal_especial
import arcos

kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

EPS_GEOM = 0.0001
TOL_EXTREMO_RATIO = 0.01
DOMINANCIA_RECTA = 2.5
OFFSET_COTA = 1.5

FACTOR_VALIDACION_MIN = 0.85
FACTOR_VALIDACION_MAX = 1.15
RUTA_HOJAS_DIAMETRO = r"C:\Temp\hojas_para_diametro.txt"

# =========================================================
# UTILIDADES GENERALES
# =========================================================
def _bbox_curva(curva):
    try:
        caja = curva.Evaluator2D.RangeBox

        minx = float(caja.MinPoint.X)
        maxx = float(caja.MaxPoint.X)
        miny = float(caja.MinPoint.Y)
        maxy = float(caja.MaxPoint.Y)

        dx = abs(maxx - minx)
        dy = abs(maxy - miny)

        if dx < EPS_GEOM and dy < EPS_GEOM:
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
        if abs(escala) < EPS_GEOM:
            escala = 1.0
    except:
        escala = 1.0
    return abs(span_sheet / escala)


def _validar_dimension(dimension, esperado_modelo, nombre_hoja, eje):
    """
    Si la cota creada sale muy distinta al tamaño esperado, la borra.
    Esto evita casos como la cota gigante de tus imágenes 1 y 2.
    """
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
            f"⚠️ {nombre_hoja}: cota {eje} descartada "
            f"(valor={valor:.3f}, esperado≈{esperado_modelo:.3f})"
        )
        return False

    return True

def _guardar_hojas_para_diametro(hojas):
    """
    Guarda en C:\\Temp la lista de hojas que fueron clasificadas
    para diametro.py.
    """
    try:
        os.makedirs(r"C:\Temp", exist_ok=True)

        unicas = []
        vistos = set()

        for h in hojas:
            hu = str(h).upper()
            if hu not in vistos:
                vistos.add(hu)
                unicas.append(hu)

        with open(RUTA_HOJAS_DIAMETRO, "w", encoding="utf-8") as f:
            for h in unicas:
                f.write(h + "\n")

        print(f"📝 Lista de hojas para diámetro guardada en: {RUTA_HOJAS_DIAMETRO}")

    except Exception as e:
        print(f"⚠️ No se pudo guardar la lista de hojas para diámetro: {e}")

def _puntos_clave_curva(curva):
    """
    Obtiene Start/Mid/End si existen, para crear intents más seguros.
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


def _crear_intent_seguro(hoja, dato, lado):
    """
    Primero intenta crear intent sobre un punto real de la curva.
    Si no puede, usa la curva completa.
    """
    curva = dato["curve"]
    puntos = _puntos_clave_curva(curva)

    if puntos:
        try:
            if lado == "izq":
                p = min(puntos, key=lambda t: t[0])[2]
            elif lado == "der":
                p = max(puntos, key=lambda t: t[0])[2]
            elif lado == "inf":
                p = min(puntos, key=lambda t: t[1])[2]
            else:  # sup
                p = max(puntos, key=lambda t: t[1])[2]

            return hoja.CreateGeometryIntent(curva, p)
        except:
            pass

    try:
        return hoja.CreateGeometryIntent(curva)
    except:
        return None


def _es_recta_dominante(d, lado):
    if lado in ("izq", "der"):
        return d["dy"] >= max(EPS_GEOM, d["dx"] * DOMINANCIA_RECTA)
    else:
        return d["dx"] >= max(EPS_GEOM, d["dy"] * DOMINANCIA_RECTA)


def _hay_suficiente_geometria_lineal(datos):
    rectas = 0
    for d in datos:
        if d["dx"] >= max(EPS_GEOM, d["dy"] * DOMINANCIA_RECTA):
            rectas += 1
        elif d["dy"] >= max(EPS_GEOM, d["dx"] * DOMINANCIA_RECTA):
            rectas += 1
    return rectas >= 2


def _elegir_curva_extrema(datos, lado, tol):
    minx, maxx, miny, maxy = _bbox_global(datos)

    if lado == "izq":
        objetivo = minx
        candidatos = [d for d in datos if abs(d["minx"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["dy"], d["dx"]))

    elif lado == "der":
        objetivo = maxx
        candidatos = [d for d in datos if abs(d["maxx"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["dy"], d["dx"]))

    elif lado == "inf":
        objetivo = miny
        candidatos = [d for d in datos if abs(d["miny"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["dx"], d["dy"]))

    elif lado == "sup":
        objetivo = maxy
        candidatos = [d for d in datos if abs(d["maxy"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["dx"], d["dy"]))

    return None

def _es_curva_redondeada(d):
    """
    Detecta curvas con presencia real de arco/redondeo.
    No son rectas dominantes puras.
    """
    try:
        return d["dx"] > 0.20 and d["dy"] > 0.20
    except:
        return False

# =========================================================
# MÉTODO MEJORADO
# =========================================================
def _crear_cota_horizontal_mejorada(hoja, vista, tg, datos, nombre_hoja):
    minx, maxx, miny, maxy = _bbox_global(datos)
    ancho_sheet = maxx - minx

    if ancho_sheet < EPS_GEOM:
        return False

    tol = max(0.02, max(vista.Width, vista.Height) * TOL_EXTREMO_RATIO)

    curva_izq = _elegir_curva_extrema(datos, "izq", tol)
    curva_der = _elegir_curva_extrema(datos, "der", tol)

    if not curva_izq or not curva_der:
        return False

    int_izq = _crear_intent_seguro(hoja, curva_izq, "izq")
    int_der = _crear_intent_seguro(hoja, curva_der, "der")

    if not int_izq or not int_der:
        return False

    try:
        pt_texto = tg.CreatePoint2d((minx + maxx) / 2.0, maxy + OFFSET_COTA)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_izq, int_der, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        esperado = _esperado_modelo(vista, ancho_sheet)
        return _validar_dimension(dim, esperado, nombre_hoja, "horizontal")

    except:
        return False


def _crear_cota_vertical_mejorada(hoja, vista, tg, datos, nombre_hoja):
    minx, maxx, miny, maxy = _bbox_global(datos)
    alto_sheet = maxy - miny

    if alto_sheet < EPS_GEOM:
        return False

    tol = max(0.02, max(vista.Width, vista.Height) * TOL_EXTREMO_RATIO)

    curva_inf = _elegir_curva_extrema(datos, "inf", tol)
    curva_sup = _elegir_curva_extrema(datos, "sup", tol)

    if not curva_inf or not curva_sup:
        return False

    # Si arriba o abajo manda un arco/redondeo,
    # mejor dejar esta hoja para lineal_especial/arcos.py
    # y que la cota salga visual, no agarrada al centro/tangencia.
    if _es_curva_redondeada(curva_inf) or _es_curva_redondeada(curva_sup):
        return False

    int_inf = _crear_intent_seguro(hoja, curva_inf, "inf")
    int_sup = _crear_intent_seguro(hoja, curva_sup, "sup")

    if not int_inf or not int_sup:
        return False

    try:
        pt_texto = tg.CreatePoint2d(minx - OFFSET_COTA, (miny + maxy) / 2.0)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inf, int_sup, kVerticalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        esperado = _esperado_modelo(vista, alto_sheet)
        return _validar_dimension(dim, esperado, nombre_hoja, "vertical")

    except:
        return False


# =========================================================
# MÉTODO LEGACY DE RESCATE
# =========================================================
def _clasificar_legacy(datos):
    lineas_verticales = []
    lineas_horizontales = []

    for d in datos:
        if d["dy"] >= d["dx"]:
            lineas_verticales.append((d["curve"], d["minx"], d["maxx"]))
        else:
            lineas_horizontales.append((d["curve"], d["miny"], d["maxy"]))

    return lineas_verticales, lineas_horizontales


def _crear_cota_horizontal_legacy(hoja, vista, tg, datos, nombre_hoja):
    lineas_verticales, _ = _clasificar_legacy(datos)
    if not lineas_verticales:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    esperado = _esperado_modelo(vista, maxx - minx)

    try:
        lin_izq = min(lineas_verticales, key=lambda x: x[1])[0]
        lin_der = max(lineas_verticales, key=lambda x: x[2])[0]

        if lin_izq == lin_der:
            return False

        int_izq = hoja.CreateGeometryIntent(lin_izq)
        int_der = hoja.CreateGeometryIntent(lin_der)

        pt_texto = tg.CreatePoint2d(vista.Position.X, vista.Position.Y + (vista.Height / 2.0) + OFFSET_COTA)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_izq, int_der, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        return _validar_dimension(dim, esperado, nombre_hoja, "horizontal_legacy")

    except:
        return False


def _crear_cota_vertical_legacy(hoja, vista, tg, datos, nombre_hoja):
    _, lineas_horizontales = _clasificar_legacy(datos)
    if not lineas_horizontales:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    esperado = _esperado_modelo(vista, maxy - miny)

    try:
        lin_inf = min(lineas_horizontales, key=lambda x: x[1])[0]
        lin_sup = max(lineas_horizontales, key=lambda x: x[2])[0]

        if lin_inf == lin_sup:
            return False

        int_inf = hoja.CreateGeometryIntent(lin_inf)
        int_sup = hoja.CreateGeometryIntent(lin_sup)

        pt_texto = tg.CreatePoint2d(vista.Position.X - (vista.Width / 2.0) - OFFSET_COTA, vista.Position.Y)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inf, int_sup, kVerticalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        return _validar_dimension(dim, esperado, nombre_hoja, "vertical_legacy")

    except:
        return False


# =========================================================
# FUNCIÓN PRINCIPAL
# =========================================================
def acotar_planos():
    print("📐 Iniciando módulo de cotas lineales (mejorado + rescate + especiales)...")

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return

    tg = inv_app.TransientGeometry

    hojas_para_diametro = []
    hojas_para_lineal_especial = []

    # Limpiar archivo temporal de diámetro al inicio de cada corrida
    _guardar_hojas_para_diametro([])

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja = str(hoja.Name).upper()

        if "_LADO" in nombre_hoja:
            print(f"⏭️ {nombre_hoja}: omitida por regla _LADO.")
            continue

        if hoja.DrawingViews.Count == 0:
            print(f"⏭️ {nombre_hoja}: sin vistas.")
            continue

        vista = hoja.DrawingViews.Item(1)
        datos = _obtener_curvas_validas(vista)

        if not datos:
            print(f"⚠️ {nombre_hoja}: sin curvas válidas.")
            hojas_para_diametro.append(nombre_hoja)
            continue

        # Si de plano casi no hay líneas, lo mandamos a círculos
        if not _hay_suficiente_geometria_lineal(datos):
            print(f"⚠️ {nombre_hoja}: muy poca geometría lineal, se manda a diametro.py")
            hojas_para_diametro.append(nombre_hoja)
            continue

        cota_ok = False

        if "_FRENTE_1" in nombre_hoja:
            # 1) Método mejorado
            cota_ok = _crear_cota_horizontal_mejorada(hoja, vista, tg, datos, nombre_hoja)

            # 2) Si falla, rescate legacy
            if not cota_ok:
                print(f"↩️ {nombre_hoja}: intentando rescate legacy horizontal...")
                cota_ok = _crear_cota_horizontal_legacy(hoja, vista, tg, datos, nombre_hoja)

            # 3) Si sigue fallando, pasa a especiales
            if not cota_ok:
                print(f"🧩 {nombre_hoja}: pasa a lineal_especial.py")
                hojas_para_lineal_especial.append(nombre_hoja)

        elif "_FRENTE_2" in nombre_hoja:
            # 1) Método mejorado
            cota_ok = _crear_cota_vertical_mejorada(hoja, vista, tg, datos, nombre_hoja)

            # 2) Si falla, rescate legacy
            if not cota_ok:
                print(f"↩️ {nombre_hoja}: intentando rescate legacy vertical...")
                cota_ok = _crear_cota_vertical_legacy(hoja, vista, tg, datos, nombre_hoja)

            # 3) Si sigue fallando, pasa a especiales
            if not cota_ok:
                print(f"🧩 {nombre_hoja}: pasa a lineal_especial.py")
                hojas_para_lineal_especial.append(nombre_hoja)

        else:
            print(f"⏭️ {nombre_hoja}: no contiene _FRENTE_1 ni _FRENTE_2.")

        print("✅ Etapa principal lineal finalizada.")

    # =====================================================
    # ETAPA ESPECIAL LINEAL
    # =====================================================
    hojas_no_resueltas = []

    if hojas_para_lineal_especial:
        print(f"\n🔧 Llamando a lineal_especial.py para {len(hojas_para_lineal_especial)} hojas...")
        hojas_no_resueltas = lineal_especial.acotar_especiales(hojas_para_lineal_especial)

    # =====================================================
    # ETAPA ARCOS
    # Solo procesa lo que lineal_especial no pudo resolver
    # =====================================================
    hojas_no_resueltas_arcos = []

    if hojas_no_resueltas:
        print(f"\n🌙 Llamando a arcos.py para {len(hojas_no_resueltas)} hojas...")
        hojas_no_resueltas_arcos = arcos.acotar_arcos(hojas_no_resueltas)

    # =====================================================
    # ETAPA CÍRCULOS
    # Solo procesa lo que desde el inicio fue clasificado como circular
    # =====================================================
    hojas_no_resueltas_diametro = []

    # Guardar siempre la lista definitiva de hojas para diámetro
    _guardar_hojas_para_diametro(hojas_para_diametro)

    if hojas_para_diametro:
        print(f"\n🔄 Llamando a diametro.py para {len(hojas_para_diametro)} hojas...")
        hojas_no_resueltas_diametro = diametro.acotar_diametros(hojas_para_diametro)

    # =====================================================
    # REPORTE FINAL
    # =====================================================
    pendientes_finales = []
    vistos = set()

    for h in hojas_no_resueltas_arcos + hojas_no_resueltas_diametro:
        hu = str(h).upper()
        if hu not in vistos:
            vistos.add(hu)
            pendientes_finales.append(hu)

    if pendientes_finales:
        print("\n⚠️ Hojas lineales no resueltas automáticamente:")
        for h in pendientes_finales:
            print(f"   - {h}")

    print("\n🏁 Proceso terminado.")

if __name__ == "__main__":
    acotar_planos()