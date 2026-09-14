import os
import win32com.client
import diametro
from inventor_com import conectar_inventor
from cota_estilo import (
    OFFSET_FUERA_PIEZA_CM,
    aplicar_estilo_cota,
    asegurar_cota_fuera_pieza_robusto,
    candidatos_texto_fuera_pieza,
    clearance_texto_cota_cm,
    dim_solapa_pieza,
)
import lineal_especial
import arcos
from rutas_runtime import ruta_hojas_diametro
kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

EPS_GEOM = 0.0001
TOL_EXTREMO_RATIO = 0.01
DOMINANCIA_RECTA = 2.5
# Separación silueta → texto (antes 1.5: solapaba bridas Vantran).
OFFSET_COTA = float(OFFSET_FUERA_PIEZA_CM)

FACTOR_VALIDACION_MIN = 0.97
FACTOR_VALIDACION_MAX = 1.05
# Portable: Planos/.runtime/ (antes C:\Temp\...)
RUTA_HOJAS_DIAMETRO = ruta_hojas_diametro()

# PointIntentEnum (Inventor) — cuadrantes de círculo/arco.
# CreateGeometryIntent(arco) o Point2d en arco → CENTRO (bug Jacking Pad).
_CIRCULAR_POINT_INTENT = {
    "izq": "kCircularLeftPointIntent",
    "der": "kCircularRightPointIntent",
    "sup": "kCircularTopPointIntent",
    "inf": "kCircularBottomPointIntent",
}
_CIRCULAR_POINT_INTENT_FALLBACK = {
    "izq": 57862,
    "der": 57863,
    "sup": 57864,
    "inf": 57865,
}

_COTAS_LOG = os.environ.get("COTAS_LOG", "").strip().lower() in (
    "1", "true", "yes", "on"
)


def _dbg(msg):
    if _COTAS_LOG:
        try:
            print(f"[COTAS_LOG] {msg}")
        except Exception:
            pass


def _point_intent_enum(nombre):
    try:
        return getattr(win32com.client.constants, nombre)
    except Exception:
        return _CIRCULAR_POINT_INTENT_FALLBACK.get(nombre)


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


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2, evitar_bbox=None):
    """
    Fuerza el Point2d de texto dentro del sheet, sin meterlo en la pieza.

    Si ``evitar_bbox`` (silueta) está dado y el clamp caería dentro, se
    empuja al borde exterior más cercano. Evita el bug Vantran: cota
    horizontal/vertical atravesando el cuerpo al clampear contra el sheet.
    """
    from cota_estilo import empujar_punto_fuera_bbox, punto_dentro_bbox

    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        x = max(margen, min(sheet_w - margen, float(x)))
        y = max(margen, min(sheet_h - margen, float(y)))
    except Exception:
        x, y = float(x), float(y)
    if evitar_bbox is not None and punto_dentro_bbox(
        x, y, evitar_bbox, holgura=0.2
    ):
        # Preferir el lado con más aire hacia el borde del sheet.
        try:
            sheet_w = float(hoja.Width)
            sheet_h = float(hoja.Height)
            minx, maxx, miny, maxy = (
                float(evitar_bbox[0]),
                float(evitar_bbox[1]),
                float(evitar_bbox[2]),
                float(evitar_bbox[3]),
            )
            aire = {
                "izq": minx - margen,
                "der": sheet_w - margen - maxx,
                "inf": miny - margen,
                "sup": sheet_h - margen - maxy,
            }
            lado = max(aire, key=aire.get)
        except Exception:
            lado = "auto"
        clr = clearance_texto_cota_cm()
        x, y = empujar_punto_fuera_bbox(x, y, evitar_bbox, lado, clr)
        try:
            sheet_w = float(hoja.Width)
            sheet_h = float(hoja.Height)
            x = max(margen, min(sheet_w - margen, x))
            y = max(margen, min(sheet_h - margen, y))
        except Exception:
            pass
        # Si el clamp volvió a meter el punto, priorizar fuera de pieza.
        if punto_dentro_bbox(x, y, evitar_bbox, holgura=0.15):
            x, y = empujar_punto_fuera_bbox(x, y, evitar_bbox, lado, clr)
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

def _puntos_clave_curva(curva, solo_extremos=False):
    """
    Start/End (y Mid solo si no es arco).

    En curvas circulares Inventor a veces expone MidPoint = centro; usarlo
    en AddLinear ancla la cota al centro (Solera Jacking Pad).
    """
    puntos = []
    usados = set()
    attrs = ("StartPoint", "EndPoint") if solo_extremos else (
        "StartPoint", "MidPoint", "EndPoint"
    )

    for attr in attrs:
        try:
            p = getattr(curva, attr)
            if p:
                x = float(p.X)
                y = float(p.Y)
                key = (round(x, 6), round(y, 6))
                if key not in usados:
                    usados.add(key)
                    puntos.append((x, y, p))
        except Exception:
            pass

    return puntos


def _intent_cuadrante_circular(hoja, dato, lado):
    """GeometryIntent en cuadrante L/R/T/B del arco (no centro)."""
    nombre = _CIRCULAR_POINT_INTENT.get(lado)
    if not nombre:
        return None
    codigo = _point_intent_enum(nombre)
    if codigo is None:
        return None
    try:
        return hoja.CreateGeometryIntent(dato["curve"], codigo)
    except Exception:
        return None


def _intent_recta_en_extremo(hoja, tg, datos, lado, tol):
    """
    Ancla en una RECTA cercana al extremo global (evita arcos).
    """
    minx, maxx, miny, maxy = _bbox_global(datos)
    rectas = [d for d in datos if not _es_curva_redondeada(d)]
    if not rectas:
        return None

    if lado == "izq":
        cands = [d for d in rectas if abs(d["minx"] - minx) <= tol * 4]
        if not cands:
            cands = sorted(rectas, key=lambda d: d["minx"])[:3]
        dato = min(cands, key=lambda d: d["minx"])
        return _crear_intent_seguro(hoja, dato, "izq", tg=tg, forzar_lineal=True)
    if lado == "der":
        cands = [d for d in rectas if abs(d["maxx"] - maxx) <= tol * 4]
        if not cands:
            cands = sorted(rectas, key=lambda d: -d["maxx"])[:3]
        dato = max(cands, key=lambda d: d["maxx"])
        return _crear_intent_seguro(hoja, dato, "der", tg=tg, forzar_lineal=True)
    if lado == "inf":
        cands = [d for d in rectas if abs(d["miny"] - miny) <= tol * 4]
        if not cands:
            cands = sorted(rectas, key=lambda d: d["miny"])[:3]
        dato = min(cands, key=lambda d: d["miny"])
        return _crear_intent_seguro(hoja, dato, "inf", tg=tg, forzar_lineal=True)
    cands = [d for d in rectas if abs(d["maxy"] - maxy) <= tol * 4]
    if not cands:
        cands = sorted(rectas, key=lambda d: -d["maxy"])[:3]
    dato = max(cands, key=lambda d: d["maxy"])
    return _crear_intent_seguro(hoja, dato, "sup", tg=tg, forzar_lineal=True)


def _crear_intent_seguro(hoja, dato, lado, tg=None, forzar_lineal=False):
    """
    GeometryIntent para cotas LINEALES.

    En arcos: usar cuadrante circular o recta vecina — NUNCA
    CreateGeometryIntent(arco) ni MidPoint (caen al centro).
    """
    curva = dato["curve"]

    if _es_curva_redondeada(dato) and not forzar_lineal:
        intent = _intent_cuadrante_circular(hoja, dato, lado)
        if intent is not None:
            return intent
        return None

    puntos = _puntos_clave_curva(curva, solo_extremos=True)
    if puntos:
        try:
            if lado == "izq":
                p = min(puntos, key=lambda t: t[0])[2]
            elif lado == "der":
                p = max(puntos, key=lambda t: t[0])[2]
            elif lado == "inf":
                p = min(puntos, key=lambda t: t[1])[2]
            else:
                p = max(puntos, key=lambda t: t[1])[2]
            return hoja.CreateGeometryIntent(curva, p)
        except Exception:
            pass

    if tg is not None and not _es_curva_redondeada(dato):
        if lado == "izq":
            x, y = dato["minx"], (dato["miny"] + dato["maxy"]) * 0.5
        elif lado == "der":
            x, y = dato["maxx"], (dato["miny"] + dato["maxy"]) * 0.5
        elif lado == "inf":
            x, y = (dato["minx"] + dato["maxx"]) * 0.5, dato["miny"]
        else:
            x, y = (dato["minx"] + dato["maxx"]) * 0.5, dato["maxy"]
        try:
            pt = tg.CreatePoint2d(x, y)
            return hoja.CreateGeometryIntent(curva, pt)
        except Exception:
            pass

    if _es_curva_redondeada(dato):
        return None
    try:
        return hoja.CreateGeometryIntent(curva)
    except Exception:
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
        return min(base, key=lambda d: (d["minx"], -d["dy"]))

    elif lado == "der":
        objetivo = maxx
        candidatos = [d for d in datos if abs(d["maxx"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["maxx"], d["dy"]))

    elif lado == "inf":
        objetivo = miny
        candidatos = [d for d in datos if abs(d["miny"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return min(base, key=lambda d: (d["miny"], -d["dx"]))

    elif lado == "sup":
        objetivo = maxy
        candidatos = [d for d in datos if abs(d["maxy"] - objetivo) <= tol]
        rectos = [d for d in candidatos if _es_recta_dominante(d, lado)]
        base = rectos if rectos else candidatos
        if not base:
            return None
        return max(base, key=lambda d: (d["maxy"], d["dx"]))

    return None

def _es_curva_redondeada(d):
    """
    Detecta curvas con presencia real de arco/redondeo.
    No son rectas dominantes puras.
    """
    try:
        return d["dx"] > 0.20 and d["dy"] > 0.20
    except Exception:
        return False

# =========================================================
# MÉTODO MEJORADO
# =========================================================
def _silueta_punta_redondeada(datos):
    """True si el extremo izquierdo es un arco/redondeo (Jacking Pad, etc.)."""
    if not datos:
        return False
    minx, maxx, miny, maxy = _bbox_global(datos)
    tol = max(0.03, (maxx - minx) * 0.08)
    for d in datos:
        if _es_curva_redondeada(d) and abs(float(d["minx"]) - minx) <= tol:
            return True
    return False


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

    # Punta redondeada (Jacking Pad): SIEMPRE cuadrante izq o recta de
    # respaldo. Nunca CreateGeometryIntent(arco) ni Point2d suelto en arco
    # (caen al centro → LARGO 1.33 en vez de 1.64).
    if _es_curva_redondeada(curva_izq):
        int_izq = _intent_cuadrante_circular(hoja, curva_izq, "izq")
        if int_izq is None:
            int_izq = _intent_recta_en_extremo(hoja, tg, datos, "izq", tol)
        if int_izq is None:
            return False
    else:
        int_izq = _crear_intent_seguro(hoja, curva_izq, "izq", tg=tg)
        if int_izq is None:
            int_izq = _intent_recta_en_extremo(hoja, tg, datos, "izq", tol)

    if _es_curva_redondeada(curva_der):
        int_der = _intent_cuadrante_circular(hoja, curva_der, "der")
        if int_der is None:
            int_der = _intent_recta_en_extremo(hoja, tg, datos, "der", tol)
        if int_der is None:
            return False
    else:
        int_der = _crear_intent_seguro(hoja, curva_der, "der", tg=tg)
        if int_der is None:
            int_der = _intent_recta_en_extremo(hoja, tg, datos, "der", tol)

    if not int_izq or not int_der:
        return False

    pieza_bb = (minx, maxx, miny, maxy)
    clr = clearance_texto_cota_cm()
    try:
        dim = None
        for x, y, _lado in candidatos_texto_fuera_pieza(
            pieza_bb, "H", clearance=clr
        ):
            pt_texto = _clampear_punto_hoja(
                hoja, tg, x, y, evitar_bbox=pieza_bb
            )
            try:
                dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                    pt_texto, int_izq, int_der, kHorizontalDimensionType
                )
            except Exception:
                continue
            aplicar_estilo_cota(dim_test, hoja=hoja)
            asegurar_cota_fuera_pieza_robusto(dim_test, tg, pieza_bb, n_chars=10)
            if dim_solapa_pieza(dim_test, pieza_bb):
                try:
                    dim_test.Delete()
                except Exception:
                    pass
                continue
            dim = dim_test
            break
        if dim is None:
            return False

        esperado = _esperado_modelo(vista, ancho_sheet)
        return _validar_dimension(dim, esperado, nombre_hoja, "horizontal")

    except Exception:
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

    # Extremos con arco/redondeo: cuadrante o recta vecina (mismo criterio
    # que horizontal). Antes se abortaba y fallaba ANCHO en soleras/punta.
    if _es_curva_redondeada(curva_inf):
        int_inf = _intent_cuadrante_circular(hoja, curva_inf, "inf")
        if int_inf is None:
            int_inf = _intent_recta_en_extremo(hoja, tg, datos, "inf", tol)
    else:
        int_inf = _crear_intent_seguro(hoja, curva_inf, "inf", tg=tg)

    if _es_curva_redondeada(curva_sup):
        int_sup = _intent_cuadrante_circular(hoja, curva_sup, "sup")
        if int_sup is None:
            int_sup = _intent_recta_en_extremo(hoja, tg, datos, "sup", tol)
    else:
        int_sup = _crear_intent_seguro(hoja, curva_sup, "sup", tg=tg)

    if not int_inf or not int_sup:
        return False

    pieza_bb = (minx, maxx, miny, maxy)
    clr = clearance_texto_cota_cm()
    try:
        dim = None
        for x, y, _lado in candidatos_texto_fuera_pieza(
            pieza_bb, "V", clearance=clr
        ):
            pt_texto = _clampear_punto_hoja(
                hoja, tg, x, y, evitar_bbox=pieza_bb
            )
            try:
                dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                    pt_texto, int_inf, int_sup, kVerticalDimensionType
                )
            except Exception:
                continue
            aplicar_estilo_cota(dim_test, hoja=hoja)
            asegurar_cota_fuera_pieza_robusto(dim_test, tg, pieza_bb, n_chars=10)
            if dim_solapa_pieza(dim_test, pieza_bb):
                try:
                    dim_test.Delete()
                except Exception:
                    pass
                continue
            dim = dim_test
            break
        if dim is None:
            return False

        esperado = _esperado_modelo(vista, alto_sheet)
        return _validar_dimension(dim, esperado, nombre_hoja, "vertical")

    except Exception:
        return False


def _crear_cota_vertical_solo_rectas(hoja, vista, tg, datos, nombre_hoja):
    """
    ANCHO en punta redondeada (Jacking Pad): solo rectas en borde
    superior/inferior — evita arcos de la punta que invalidan el span.
    """
    minx, maxx, miny, maxy = _bbox_global(datos)
    alto_sheet = maxy - miny
    if alto_sheet < EPS_GEOM:
        return False

    tol = max(0.02, max(vista.Width, vista.Height) * TOL_EXTREMO_RATIO)
    rectas = [d for d in datos if not _es_curva_redondeada(d)]
    if not rectas:
        return False

    tops = [d for d in rectas if d["maxy"] >= maxy - tol]
    bots = [d for d in rectas if d["miny"] <= miny + tol]
    if not tops or not bots:
        return False

    curva_sup = max(tops, key=lambda d: (d["dx"], d["dy"]))
    curva_inf = max(bots, key=lambda d: (d["dx"], d["dy"]))
    if curva_sup.get("curve") is curva_inf.get("curve"):
        return False

    int_sup = _crear_intent_seguro(hoja, curva_sup, "sup", tg=tg)
    int_inf = _crear_intent_seguro(hoja, curva_inf, "inf", tg=tg)
    if not int_sup or not int_inf:
        int_sup = _intent_recta_en_extremo(hoja, tg, datos, "sup", tol)
        int_inf = _intent_recta_en_extremo(hoja, tg, datos, "inf", tol)
    if not int_sup or not int_inf:
        return False

    try:
        pieza_bb = (minx, maxx, miny, maxy)
        dim = None
        for x, y, _lado in candidatos_texto_fuera_pieza(
            pieza_bb, "V", clearance=clearance_texto_cota_cm()
        ):
            pt_texto = _clampear_punto_hoja(
                hoja, tg, x, y, evitar_bbox=pieza_bb
            )
            try:
                dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                    pt_texto, int_inf, int_sup, kVerticalDimensionType
                )
            except Exception:
                continue
            aplicar_estilo_cota(dim_test, hoja=hoja)
            asegurar_cota_fuera_pieza_robusto(dim_test, tg, pieza_bb, n_chars=10)
            if dim_solapa_pieza(dim_test, pieza_bb):
                try:
                    dim_test.Delete()
                except Exception:
                    pass
                continue
            dim = dim_test
            break
        if dim is None:
            return False
        esperado = _esperado_modelo(vista, alto_sheet)
        return _validar_dimension(dim, esperado, nombre_hoja, "vertical_rectas")
    except Exception:
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
            evitar_bbox=(minx, maxx, miny, maxy),
        )
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_izq, int_der, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)
        asegurar_cota_fuera_pieza_robusto(
            dim, tg, (minx, maxx, miny, maxy), n_chars=10
        )

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
            evitar_bbox=(minx, maxx, miny, maxy),
        )
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inf, int_sup, kVerticalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)
        asegurar_cota_fuera_pieza_robusto(
            dim, tg, (minx, maxx, miny, maxy), n_chars=10
        )

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
        return []

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
    hojas_frente_ok = []
    solo_barrenos_thk = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    if solo_barrenos_thk:
        print(
            "SOLO_FLAT_CORTE: no LENGTH/WIDTH; "
            "barrenos X/Y+TYP + Ø por tipo + THK (LADO)."
        )

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_completo = str(hoja.Name)
        nombre_hoja = nombre_completo.upper()
        # Inventor agrega ":N" a nombres duplicados. Comparamos por base.
        base_up = _base_hoja(nombre_completo).upper()

        if permitidos_up is not None and base_up not in permitidos_up:
            contadores["filtradas_por_permitidos"] += 1
            continue

        # Modo rápido flat: omitir LARGO/ANCHO en DESPLIEGUE_FRENTE_*.
        # Los barrenos X/Y se agregan al final; el espesor va por THK.py.
        if solo_barrenos_thk and "_DESPLIEGUE_" in nombre_hoja and (
            "_FRENTE_1" in nombre_hoja or "_FRENTE_2" in nombre_hoja
        ):
            contadores["visitadas"] += 1
            _dbg(f"  skip lineal (SOLO_FLAT): {nombre_completo}")
            continue

        contadores["visitadas"] += 1
        _dbg(f"visita: {nombre_hoja} (base={base_up})")

        if (
            "_ESTANIADO" in nombre_hoja
            or "_LADO" in nombre_hoja
            or "_ALTO" in nombre_hoja
            or "_LARGO_PATA" in nombre_hoja
        ):
            print(
                f"⏭️ {nombre_hoja}: omitida por regla "
                f"_ESTANIADO/_LADO/_ALTO/_LARGO_PATA."
            )
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

        # LARGO (FRENTE_1) = eje MAYOR; ANCHO (FRENTE_2) = eje menor.
        # Excepción: punta redondeada (Jacking Pad) → LARGO = tip→base (H)
        # como ANCHO ya hacía bien; nunca dejar que caiga a intent de centro.
        minx_b, maxx_b, miny_b, maxy_b = _bbox_global(datos)
        punta = _silueta_punta_redondeada(datos)
        if punta:
            eje_mayor_horizontal = True
        else:
            eje_mayor_horizontal = (maxx_b - minx_b) >= (maxy_b - miny_b)

        if "_FRENTE_1" in nombre_hoja:
            try:
                if eje_mayor_horizontal:
                    cota_ok = _crear_cota_horizontal_mejorada(
                        hoja, vista, tg, datos, nombre_hoja
                    )
                else:
                    cota_ok = _crear_cota_vertical_mejorada(
                        hoja, vista, tg, datos, nombre_hoja
                    )
            except Exception as e:
                contadores["excepciones"] += 1
                _dbg(f"  excepción en cota mayor FRENTE_1: {e}")

            # Rescate tip→base si el eje mayor falló (evita arcos→centro).
            if not cota_ok and not eje_mayor_horizontal:
                try:
                    cota_ok = _crear_cota_horizontal_mejorada(
                        hoja, vista, tg, datos, nombre_hoja
                    )
                except Exception:
                    pass

            if cota_ok:
                contadores["frente1_ok"] += 1
                hojas_frente_ok.append(nombre_completo)
            else:
                # En punta redondeada NO usar legacy (CreateGeometryIntent
                # sobre arco → centro). Mejor lineal_especial / arcos visual.
                if not punta:
                    print(f"↩️ {nombre_hoja}: intentando rescate legacy (eje mayor)...")
                    try:
                        if eje_mayor_horizontal:
                            cota_ok = _crear_cota_horizontal_legacy(
                                hoja, vista, tg, datos, nombre_hoja
                            )
                        else:
                            cota_ok = _crear_cota_vertical_legacy(
                                hoja, vista, tg, datos, nombre_hoja
                            )
                    except Exception as e:
                        contadores["excepciones"] += 1
                        _dbg(f"  excepción en legacy FRENTE_1: {e}")
                    if cota_ok:
                        contadores["frente1_legacy"] += 1
                        hojas_frente_ok.append(nombre_completo)

            if not cota_ok:
                print(f"🧩 {nombre_hoja}: pasa a lineal_especial.py")
                contadores["frente1_a_especiales"] += 1
                hojas_para_lineal_especial.append(nombre_hoja)

        elif "_FRENTE_2" in nombre_hoja:
            try:
                if eje_mayor_horizontal:
                    cota_ok = _crear_cota_vertical_mejorada(
                        hoja, vista, tg, datos, nombre_hoja
                    )
                else:
                    cota_ok = _crear_cota_horizontal_mejorada(
                        hoja, vista, tg, datos, nombre_hoja
                    )
            except Exception as e:
                contadores["excepciones"] += 1
                _dbg(f"  excepción en cota menor FRENTE_2: {e}")

            if cota_ok:
                contadores["frente2_ok"] += 1
                hojas_frente_ok.append(nombre_completo)
            else:
                # Jacking Pad / punta: legacy SÍ para ANCHO (eje menor con
                # rectas H). Antes `if not punta` bloqueaba FRENTE_2 y
                # solo quedaban LARGO+THK (faltaba ANCHO_44).
                print(f"↩️ {nombre_hoja}: intentando rescate legacy (eje menor)...")
                try:
                    if eje_mayor_horizontal:
                        cota_ok = _crear_cota_vertical_legacy(
                            hoja, vista, tg, datos, nombre_hoja
                        )
                    else:
                        cota_ok = _crear_cota_horizontal_legacy(
                            hoja, vista, tg, datos, nombre_hoja
                        )
                except Exception as e:
                    contadores["excepciones"] += 1
                    _dbg(f"  excepción en legacy FRENTE_2: {e}")
                if cota_ok:
                    contadores["frente2_legacy"] += 1
                    hojas_frente_ok.append(nombre_completo)

            # Rescate final punta: forzar intents solo en rectas de borde.
            if not cota_ok and punta:
                try:
                    cota_ok = _crear_cota_vertical_solo_rectas(
                        hoja, vista, tg, datos, nombre_hoja
                    )
                    if cota_ok:
                        contadores["frente2_ok"] += 1
                        print(
                            f"✅ {nombre_hoja}: ANCHO por rectas de borde "
                            f"(punta redondeada)"
                        )
                except Exception as e:
                    _dbg(f"  excepción rescate rectas FRENTE_2: {e}")

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

    # Barrenos flat: X/Y+TYP (posición) + Ø por TIPO de tamaño (DIAMETRO_Hnn).
    # Doblado: solo Ø por tipo (como antes).
    hojas_extra_barrenos = []
    try:
        import barrenos_xy_despliegue

        print("\nBarrenos flat X/Y+TYP (todas las DESPLIEGUE_FRENTE_1)...")
        resultado_xy = barrenos_xy_despliegue.acotar_barrenos_xy_despliegue(None)
        if resultado_xy:
            hojas_extra_barrenos.extend(resultado_xy)
            print(f"  -> {len(resultado_xy)} hojas XCENTRO/YCENTRO creadas")
        else:
            print("  -> 0 hojas XY (sin barrenos detectados)")
    except Exception as e:
        print(f"AVISO: barrenos X/Y despliegue fallo: {e}")

    # Ø: una hoja por cada tamaño distinto (circulo 11 + oval 13 = 2 hojas).
    try:
        frentes_diam = []
        for i in range(1, int(plano.Sheets.Count) + 1):
            try:
                h = plano.Sheets.Item(i)
            except Exception:
                continue
            nu = str(h.Name).upper()
            if "_DIAMETRO_" in nu or "_XCENTRO" in nu or "_YCENTRO" in nu:
                continue
            if "_FRENTE_1" not in nu and "_FRENTE_2" not in nu:
                continue
            # En modo SOLO_FLAT_CORTE solo flat; si no, flat + doblado.
            if solo_barrenos_thk and "_DESPLIEGUE_" not in nu:
                continue
            frentes_diam.append(str(h.Name).rsplit(":", 1)[0])
        if frentes_diam:
            print(
                f"\nBarrenos Ø por tipo de tamaño en "
                f"{len(frentes_diam)} hojas FRENTE..."
            )
            resultado_barrenos = diametro.acotar_barrenos_placas(frentes_diam)
            if resultado_barrenos:
                hojas_extra_barrenos.extend(resultado_barrenos)
                print(f"  -> {len(resultado_barrenos)} hojas DIAMETRO_H* creadas")
    except Exception as e:
        print(f"AVISO: barrenos Ø por tipo fallo: {e}")

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
    return hojas_extra_barrenos


if __name__ == "__main__":
    acotar_planos()