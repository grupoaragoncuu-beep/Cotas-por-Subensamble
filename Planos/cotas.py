import os
import win32com.client
import diametro
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota
import lineal_especial
import arcos
from rutas_runtime import ruta_hojas_diametro

kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

EPS_GEOM = 0.0001
TOL_EXTREMO_RATIO = 0.01
DOMINANCIA_RECTA = 2.5
OFFSET_COTA = 1.5

FACTOR_VALIDACION_MIN = 0.85
FACTOR_VALIDACION_MAX = 1.15
# Portable: Planos/.runtime/ (antes C:\Temp\...)
RUTA_HOJAS_DIAMETRO = ruta_hojas_diametro()


# Log detallado opt-in para diagnosticar por qué una hoja termina sin cota.
_COTAS_LOG = os.environ.get("COTAS_LOG", "").strip().lower() in ("1", "true", "yes", "on")


def _dbg(msg):
    if _COTAS_LOG:
        try:
            print(f"[COTAS_LOG] {msg}")
        except Exception:
            pass


def _base_hoja(nombre):
    """
    Devuelve la base del nombre de hoja sin el sufijo `:N` que Inventor
    agrega cuando el nombre ya existe.

        '62176-1247-P01_FRENTE_1:54' -> '62176-1247-P01_FRENTE_1'
    """
    if not nombre:
        return nombre
    partes = str(nombre).rsplit(":", 1)
    if len(partes) == 2 and partes[1].isdigit():
        return partes[0]
    return str(nombre)


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2):
    """
    Fuerza el Point2d de texto de cota a caer dentro del rectángulo físico de
    la hoja de Inventor con un margen mínimo. Sin esto, cotas colocadas
    cerca del borde quedan fuera del rectángulo que la cámara exporta como
    JPG (Inventor las acepta pero el bitmap no las incluye) y el resultado
    es la clásica captura sin cota.

    El margen por defecto (1.2 cm) da espacio al número + flecha para que
    quepan enteros dentro del sheet, evitando que el recorte del JPG los
    corte.
    """
    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        x = max(margen, min(sheet_w - margen, x))
        y = max(margen, min(sheet_h - margen, y))
    except Exception:
        pass
    return tg.CreatePoint2d(x, y)

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
        _dbg(f"{nombre_hoja}: {eje} sin ModelValue accesible, se acepta")
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

    _dbg(f"{nombre_hoja}: {eje} OK valor={valor:.4f}cm esperado={esperado_modelo:.4f}cm")
    return True

def _guardar_hojas_para_diametro(hojas):
    """
    Guarda en Planos/.runtime/ la lista de hojas clasificadas para diametro.py.
    """
    try:
        destino = ruta_hojas_diametro()
        os.makedirs(os.path.dirname(destino), exist_ok=True)

        unicas = []
        vistos = set()

        for h in hojas:
            hu = str(h).upper()
            if hu not in vistos:
                vistos.add(hu)
                unicas.append(hu)

        with open(destino, "w", encoding="utf-8") as f:
            for h in unicas:
                f.write(h + "\n")

        print(f"📝 Lista de hojas para diámetro guardada en: {destino}")

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
        pt_texto = _clampear_punto_hoja(
            hoja, tg, (minx + maxx) / 2.0, maxy + OFFSET_COTA
        )
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
        pt_texto = _clampear_punto_hoja(
            hoja, tg, minx - OFFSET_COTA, (miny + maxy) / 2.0
        )
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

        pt_texto = _clampear_punto_hoja(
            hoja, tg,
            vista.Position.X,
            vista.Position.Y + (vista.Height / 2.0) + OFFSET_COTA,
        )
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

        pt_texto = _clampear_punto_hoja(
            hoja, tg,
            vista.Position.X - (vista.Width / 2.0) - OFFSET_COTA,
            vista.Position.Y,
        )
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
def acotar_planos(nombres_permitidos=None, reset_diametro=True):
    """
    Aplica cotas lineales sobre las hojas del machote.

    Parametros
    ----------
    nombres_permitidos : set[str] | None
        Si se provee, solo se procesan hojas cuyo nombre (upper) esté en el
        set. Útil para procesar por lotes (modo D).
    reset_diametro : bool
        Si True, borra el archivo temporal de hojas para diámetro al inicio.
        En lotes >= 2 debe pasarse False para no perder el mapeo previo.
    """
    print("📐 Iniciando módulo de cotas lineales (mejorado + rescate + especiales)...")
    if _COTAS_LOG:
        print("[COTAS_LOG] modo diagnóstico ACTIVO (COTAS_LOG=1)")

    permitidos_up = None
    if nombres_permitidos is not None:
        permitidos_up = {str(x).upper() for x in nombres_permitidos}
        print(f"  Modo lote: {len(permitidos_up)} hojas permitidas")
        _dbg(
            "primeros permitidos: "
            + ", ".join(sorted(list(permitidos_up))[:5])
            + (" ..." if len(permitidos_up) > 5 else "")
        )

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return

    tg = inv_app.TransientGeometry

    hojas_para_diametro = []
    hojas_para_lineal_especial = []

    if reset_diametro:
        _guardar_hojas_para_diametro([])
        # También resetear la lista de piezas cilíndricas sólidas para que
        # no arrastre nombres detectados en corridas previas del mismo día.
        try:
            ruta_solidas = getattr(diametro, "RUTA_PIEZAS_SOLIDAS", None)
            if ruta_solidas and os.path.exists(ruta_solidas):
                os.remove(ruta_solidas)
        except Exception:
            pass

    contadores = {
        "visitadas": 0,
        "filtradas_por_permitidos": 0,
        "frente1_ok": 0,
        "frente1_legacy": 0,
        "frente1_a_especiales": 0,
        "frente2_ok": 0,
        "frente2_legacy": 0,
        "frente2_a_especiales": 0,
        "descartada_por_curvas_vacias": 0,
        "descartada_por_poca_geom": 0,
        "sin_frente_match": 0,
        "excepciones": 0,
    }

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_completo = str(hoja.Name)
        nombre_hoja = nombre_completo.upper()
        # Inventor agrega ":N" a nombres duplicados. Comparamos por base.
        base_up = _base_hoja(nombre_completo).upper()

        if permitidos_up is not None and base_up not in permitidos_up:
            contadores["filtradas_por_permitidos"] += 1
            continue

        contadores["visitadas"] += 1
        _dbg(f"visita: {nombre_hoja} (base={base_up})")

        if (
            "_LADO" in nombre_hoja
            or "_ALTO" in nombre_hoja
            or "_LARGO_PATA" in nombre_hoja
        ):
            print(f"⏭️ {nombre_hoja}: omitida por regla _LADO/_ALTO/_LARGO_PATA.")
            continue

        if hoja.DrawingViews.Count == 0:
            print(f"⏭️ {nombre_hoja}: sin vistas.")
            continue

        vista = hoja.DrawingViews.Item(1)
        datos = _obtener_curvas_validas(vista)

        if not datos:
            print(f"⚠️ {nombre_hoja}: sin curvas válidas.")
            contadores["descartada_por_curvas_vacias"] += 1
            hojas_para_diametro.append(nombre_hoja)
            continue

        # Si de plano casi no hay líneas, lo mandamos a círculos
        if not _hay_suficiente_geometria_lineal(datos):
            print(f"⚠️ {nombre_hoja}: muy poca geometría lineal, se manda a diametro.py")
            contadores["descartada_por_poca_geom"] += 1
            hojas_para_diametro.append(nombre_hoja)
            continue

        cota_ok = False
        _dbg(f"  {nombre_hoja}: {len(datos)} curvas válidas")

        if "_FRENTE_1" in nombre_hoja:
            try:
                cota_ok = _crear_cota_horizontal_mejorada(hoja, vista, tg, datos, nombre_hoja)
            except Exception as e:
                contadores["excepciones"] += 1
                _dbg(f"  excepción en _crear_cota_horizontal_mejorada: {e}")

            if cota_ok:
                contadores["frente1_ok"] += 1
            else:
                print(f"↩️ {nombre_hoja}: intentando rescate legacy horizontal...")
                try:
                    cota_ok = _crear_cota_horizontal_legacy(hoja, vista, tg, datos, nombre_hoja)
                except Exception as e:
                    contadores["excepciones"] += 1
                    _dbg(f"  excepción en _crear_cota_horizontal_legacy: {e}")
                if cota_ok:
                    contadores["frente1_legacy"] += 1

            if not cota_ok:
                print(f"🧩 {nombre_hoja}: pasa a lineal_especial.py")
                contadores["frente1_a_especiales"] += 1
                hojas_para_lineal_especial.append(nombre_hoja)

        elif "_FRENTE_2" in nombre_hoja:
            try:
                cota_ok = _crear_cota_vertical_mejorada(hoja, vista, tg, datos, nombre_hoja)
            except Exception as e:
                contadores["excepciones"] += 1
                _dbg(f"  excepción en _crear_cota_vertical_mejorada: {e}")

            if cota_ok:
                contadores["frente2_ok"] += 1
            else:
                print(f"↩️ {nombre_hoja}: intentando rescate legacy vertical...")
                try:
                    cota_ok = _crear_cota_vertical_legacy(hoja, vista, tg, datos, nombre_hoja)
                except Exception as e:
                    contadores["excepciones"] += 1
                    _dbg(f"  excepción en _crear_cota_vertical_legacy: {e}")
                if cota_ok:
                    contadores["frente2_legacy"] += 1

            if not cota_ok:
                print(f"🧩 {nombre_hoja}: pasa a lineal_especial.py")
                contadores["frente2_a_especiales"] += 1
                hojas_para_lineal_especial.append(nombre_hoja)

        else:
            print(f"⏭️ {nombre_hoja}: no contiene _FRENTE_1 ni _FRENTE_2.")
            contadores["sin_frente_match"] += 1

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

    if _COTAS_LOG:
        print("[COTAS_LOG] resumen del lote:")
        for k, v in contadores.items():
            print(f"  {k}: {v}")

    print("\n🏁 Proceso terminado.")

if __name__ == "__main__":
    acotar_planos()