import math
import os
import re
import time
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import (
    OFFSET_FUERA_PIEZA_CM,
    aplicar_estilo_cota,
    asegurar_cota_fuera_pieza,
    asegurar_cota_fuera_pieza_robusto,
    candidatos_texto_fuera_pieza,
    clearance_texto_cota_cm,
    dim_solapa_pieza,
)

# Import diferido para evitar imports circulares (creador_vistas es cliente
# de THK.py en el flujo por lotes, pero aquí sólo lo usamos como utilidad
# para reciclar la limpieza de border/titleblock probada en el flujo normal).
try:
    import creador_vistas as _creador_vistas
except Exception:
    _creador_vistas = None


# Log detallado sólo cuando THK_LOG=1 (o similar). Silencia por defecto para
# no ensuciar la corrida normal, pero permite diagnosticar hojas que quedan
# sin cotas.
_THK_LOG = os.environ.get("THK_LOG", "").strip().lower() in ("1", "true", "yes", "on")


def _dbg(msg):
    if _THK_LOG:
        try:
            print(f"[THK_LOG] {msg}")
        except Exception:
            pass


# Registro global de hojas que quedaron sin cota THK a lo largo de una
# corrida. Cada llamada a ``acotar_thk`` agrega sus pendientes aquí en vez
# de sobrescribir, para que el flujo por lotes pueda consolidar el total al
# final. Usar ``reset_pendientes_thk()`` al inicio de un flujo.
LAST_PENDIENTES_THK: list = []
# Hojas ``_LADO`` donde se acotó Ø de barra/pin sólido (no espesor).
# El renombrado las convierte a ``_DIAMETRO_EXTERIOR`` en vez de ``_THK``.
LAST_OD_SOLID_LADO: list = []


def reset_pendientes_thk():
    """Vacía el registro global de pendientes THK. Llamar al inicio del flujo."""
    LAST_PENDIENTES_THK.clear()
    LAST_OD_SOLID_LADO.clear()


def _base_hoja(nombre):
    """
    Devuelve la base del nombre de hoja sin el sufijo `:N` que Inventor
    agrega cuando el nombre ya existe (p.ej. 'X_LADO:12' -> 'X_LADO').
    """
    if not nombre:
        return nombre
    partes = str(nombre).rsplit(":", 1)
    if len(partes) == 2 and partes[1].isdigit():
        return partes[0]
    return str(nombre)

kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

# Enums de Inventor para crear vistas nuevas (usados por _crear_hoja_alto tras
# reemplazar el `CopyTo` que causaba errores COM en la hoja duplicada).
kArbitraryViewOrientation = 10763
kDefaultViewOrientation = 10753
kHiddenLineRemovedDrawingViewStyle = 32258

EPS = 0.0001
OFFSET_COTA = float(OFFSET_FUERA_PIEZA_CM)

# Inventor Curve2dTypeEnum (parcial)
kCircularArcCurve2d = 5121
kCircleCurve2d = 5122

# Inventor internamente trabaja en cm
IN_TO_CM = 2.54
TOL_IN = 0.005
TOL_CM = TOL_IN * IN_TO_CM

ALLOWED_IN = [
    0.06,
    0.07,
    0.105,
    0.119,
    0.125,
    0.187,
    0.1875,
    0.25,
    0.3125,
    0.375,
    0.38,
    0.5,
    0.625,
    0.75,
    0.875,
    1.0,
    1.06,
    1.07,
    1.105,
    1.119,
    1.187,
    1.25,
    1.3125,
    1.375,
    1.5,
    1.625,
    1.75,
    1.875,
    2.0,
    2.157,
    2.25,
    2.5,
    3.0,
]
ALLOWED_CM = [x * IN_TO_CM for x in ALLOWED_IN]


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

        info = {
            "curve": curva,
            "minx": minx,
            "maxx": maxx,
            "miny": miny,
            "maxy": maxy,
            "dx": dx,
            "dy": dy,
            "cx": (minx + maxx) / 2.0,
            "cy": (miny + maxy) / 2.0,
            "curve_type": 0,
            "es_arco": False,
            "radius": None,
        }

        try:
            info["curve_type"] = int(curva.CurveType)
        except Exception:
            pass

        # Intentar radio/centro reales del arco/círculo 2D.
        for attr in ("Curve2d", "Geometry"):
            try:
                geom = getattr(curva, attr, None)
                if geom is None:
                    continue
                centro = geom.Center
                radio = float(geom.Radius)
                if radio > EPS:
                    info["cx"] = float(centro.X)
                    info["cy"] = float(centro.Y)
                    info["radius"] = radio
                    info["es_arco"] = True
                    break
            except Exception:
                continue

        if not info["es_arco"]:
            if info["curve_type"] in (kCircularArcCurve2d, kCircleCurve2d):
                info["es_arco"] = True
                info["radius"] = max(dx, dy) / 2.0
            else:
                lado = max(dx, dy)
                if lado >= 0.05 and abs(dx - dy) <= (lado * 0.12):
                    info["es_arco"] = True
                    info["radius"] = lado / 2.0

        return info
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


def _snap_a_catalogo(valor_cm):
    mejor = None
    mejor_diff = None

    for permitido_cm, permitido_in in zip(ALLOWED_CM, ALLOWED_IN):
        diff = abs(valor_cm - permitido_cm)
        if mejor is None or diff < mejor_diff:
            mejor = (permitido_cm, permitido_in)
            mejor_diff = diff

    if mejor is None:
        return None

    if mejor_diff <= TOL_CM:
        return {
            "valor_cm": mejor[0],
            "valor_in": mejor[1],
            "diff_cm": mejor_diff,
            "desde_catalogo": True,
        }

    return None


def _snap_o_medido(valor_cm):
    """
    Prefiere catálogo; si no hay match, usa el valor medido real.

    Evita dejar THK vacío en placas/canales cuyo espesor no está en gauge
    estándar (p. ej. 2.157 in).
    """
    snap = _snap_a_catalogo(valor_cm)
    if snap:
        return snap
    return {
        "valor_cm": float(valor_cm),
        "valor_in": float(valor_cm) / IN_TO_CM,
        "diff_cm": 0.0,
        "desde_catalogo": False,
    }


def _crear_intent_punto2d(hoja, tg, curva, x, y):
    try:
        pt = tg.CreatePoint2d(x, y)
        return hoja.CreateGeometryIntent(curva, pt)
    except:
        return None


def _puntos_clave_curva_thk(curva):
    puntos = []
    usados = set()
    for attr in ("StartPoint", "MidPoint", "EndPoint"):
        try:
            p = getattr(curva, attr)
            if p is None:
                continue
            x, y = float(p.X), float(p.Y)
            key = (round(x, 6), round(y, 6))
            if key in usados:
                continue
            usados.add(key)
            puntos.append((x, y, p))
        except Exception:
            continue
    return puntos


def _es_recta_dominante_thk(d, lado):
    if lado in ("izq", "der"):
        return d["dy"] >= max(EPS, d["dx"] * 2.0)
    return d["dx"] >= max(EPS, d["dy"] * 2.0)


def _elegir_curva_extrema_thk(datos, lado, tol):
    """
    Extremo REAL de silueta (incluye filo exterior del doblez).

    Prefiere rectas alineadas al borde (horizontales arriba/abajo,
    verticales izq/der) para no anclar en la tangencia del radio.
    """
    minx, maxx, miny, maxy = _bbox_global(datos)
    if lado == "izq":
        objetivo = minx
        cands = [d for d in datos if abs(d["minx"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante_thk(d, lado)]
        base = rectos if rectos else cands
        if not base:
            return None
        return max(base, key=lambda d: (d["dy"], d["dx"]))
    if lado == "der":
        objetivo = maxx
        cands = [d for d in datos if abs(d["maxx"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante_thk(d, lado)]
        base = rectos if rectos else cands
        if not base:
            return None
        return max(base, key=lambda d: (d["dy"], d["dx"]))
    if lado == "inf":
        objetivo = miny
        cands = [d for d in datos if abs(d["miny"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante_thk(d, lado)]
        base = rectos if rectos else cands
        if not base:
            return None
        return max(base, key=lambda d: (d["dx"], d["dy"]))
    if lado == "sup":
        objetivo = maxy
        cands = [d for d in datos if abs(d["maxy"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante_thk(d, lado)]
        base = rectos if rectos else cands
        if not base:
            return None
        return max(base, key=lambda d: (d["dx"], d["dy"]))
    return None


def _intent_en_extremo(hoja, tg, dato, lado):
    """
    GeometryIntent anclado al punto extremo real de la curva.

    Importante: en filetes/arcos parciales (p. ej. PIPE FLANE 0.375) el
    ``kCircularTopPointIntent`` apunta al cuadrante del círculo completo,
    que queda FUERA de la silueta → cota flotante y valor inflado (0.75
    en vez de ~0.50). Por eso Start/End del arco van PRIMERO; el cuadrante
    circular solo se usa como respaldo cuando Start/End no bastan.
    """
    curva = dato["curve"]

    # 1) Start/End — MidPoint en círculos puede ser el centro.
    puntos = []
    for attr in ("StartPoint", "EndPoint"):
        try:
            p = getattr(curva, attr)
            if p is None:
                continue
            puntos.append((float(p.X), float(p.Y), p))
        except Exception:
            continue

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

    # 2) Punto 2D en el extremo del bbox de la curva (seguro en filetes).
    if lado == "izq":
        x, y = dato["minx"], (dato["miny"] + dato["maxy"]) * 0.5
    elif lado == "der":
        x, y = dato["maxx"], (dato["miny"] + dato["maxy"]) * 0.5
    elif lado == "inf":
        x, y = (dato["minx"] + dato["maxx"]) * 0.5, dato["miny"]
    else:
        x, y = (dato["minx"] + dato["maxx"]) * 0.5, dato["maxy"]
    intent = _crear_intent_punto2d(hoja, tg, curva, x, y)
    if intent is not None:
        return intent

    # 3) Cuadrante circular solo si dx≈dy (círculo casi completo).
    if abs(dato.get("dx", 0) - dato.get("dy", 0)) <= max(
        dato.get("dx", 0), dato.get("dy", 0)
    ) * 0.25 and min(dato.get("dx", 0), dato.get("dy", 0)) > 0.15:
        nombres = {
            "izq": "kCircularLeftPointIntent",
            "der": "kCircularRightPointIntent",
            "sup": "kCircularTopPointIntent",
            "inf": "kCircularBottomPointIntent",
        }
        fallback = {"izq": 57862, "der": 57863, "sup": 57864, "inf": 57865}
        try:
            codigo = getattr(
                win32com.client.constants, nombres[lado], fallback[lado]
            )
        except Exception:
            codigo = fallback.get(lado)
        if codigo is not None:
            try:
                return hoja.CreateGeometryIntent(curva, codigo)
            except Exception:
                pass

    try:
        return hoja.CreateGeometryIntent(curva)
    except Exception:
        return None


def _validar_span_cota(dimension, esperado_sheet, vista, nombre_hoja, etiqueta):
    """
    Rechaza cotas ancladas a tangencia de doblez (típicamente 2–8% cortas).
    """
    try:
        valor = abs(float(dimension.ModelValue))
    except Exception:
        return True
    esperado = _esperado_modelo(vista, esperado_sheet)
    if esperado <= EPS:
        return True
    ratio = valor / esperado
    if ratio < 0.97 or ratio > 1.05:
        try:
            dimension.Delete()
        except Exception:
            pass
        print(
            f"⚠️ {nombre_hoja}: {etiqueta} descartada "
            f"(valor={valor / IN_TO_CM:.3f} in, "
            f"silueta≈{esperado / IN_TO_CM:.3f} in, ratio={ratio:.3f})"
        )
        return False
    return True


def _es_circular_aprox(d):
    lado = max(d["dx"], d["dy"])
    if lado < 0.05:
        return False
    return abs(d["dx"] - d["dy"]) <= (lado * 0.12)


def _buscar_circulos(datos):
    return [d for d in datos if _es_circular_aprox(d)]


def _buscar_arcos(datos):
    return [d for d in datos if d.get("es_arco")]


def _es_perfil_semicircular(datos):
    """
    Media caña / contour flange: varios arcos abiertos que NO llenan un
    círculo completo de la envolvente (P29).
    """
    arcos = _buscar_arcos(datos)
    if len(arcos) < 2:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    mayor = max(w, h)
    if mayor <= EPS:
        return False

    # Círculo cerrado completo → lo manejan los resolvers circular_*
    circulos = _buscar_circulos(datos)
    if circulos and abs(w - h) <= mayor * 0.15:
        outer = max(circulos, key=lambda d: d["dx"])
        if abs(outer["dx"] - mayor) <= mayor * 0.12:
            return False

    return True


def _pares_arcos_concentricos(arcos):
    """Busca pares de arcos concéntricos para espesor radial de chapa curva."""
    pares = []
    for i in range(len(arcos)):
        for j in range(i + 1, len(arcos)):
            a = arcos[i]
            b = arcos[j]
            ra = a.get("radius")
            rb = b.get("radius")
            if ra is None or rb is None:
                continue
            if abs(ra - rb) <= max(0.01, min(ra, rb) * 0.05):
                continue
            center_tol = max(0.05, max(ra, rb) * 0.08)
            if abs(a["cx"] - b["cx"]) > center_tol:
                continue
            if abs(a["cy"] - b["cy"]) > center_tol:
                continue
            gap = abs(ra - rb)
            outer, inner = (a, b) if ra > rb else (b, a)
            pares.append({
                "outer": outer,
                "inner": inner,
                "gap_sheet": gap,
            })
    return pares


def _resolver_semicircular(hoja, vista, tg, datos, nombre_hoja):
    """
    THK para perfiles semicirculares / media caña.

    1) Intenta pares lineales (puntas de las patas).
    2) Si no, espesor radial entre arcos concéntricos.
    3) Si no, rescate bbox menor (solo si el perfil es claramente alargado).
    """
    print(f"🌙 {nombre_hoja}: perfil semicircular detectado")

    ok, meta = _resolver_prismatico(hoja, vista, tg, datos, nombre_hoja)
    if ok:
        return ok, meta

    pares = _pares_arcos_concentricos(_buscar_arcos(datos))
    if pares:
        # Preferir el menor gap usable (= espesor de chapa)
        ranqueados = []
        for p in pares:
            valor_cm = _esperado_modelo(vista, p["gap_sheet"])
            if valor_cm > 12.0 * IN_TO_CM:
                continue
            if valor_cm < 0.02 * IN_TO_CM:
                continue
            p = dict(p)
            p["valor_cm"] = valor_cm
            p["snap"] = _snap_o_medido(valor_cm)
            ranqueados.append(p)

        if ranqueados:
            ranqueados.sort(
                key=lambda x: (
                    0 if x["snap"].get("desde_catalogo") else 1,
                    x["valor_cm"],
                )
            )
            mejor = ranqueados[0]
            try:
                outer = mejor["outer"]
                inner = mejor["inner"]
                # Medir en la dirección horizontal desde el centro común.
                cx = (outer["cx"] + inner["cx"]) / 2.0
                cy = (outer["cy"] + inner["cy"]) / 2.0
                ro = float(outer["radius"])
                ri = float(inner["radius"])
                int_o = _crear_intent_punto2d(
                    hoja, tg, outer["curve"], cx + ro, cy
                )
                int_i = _crear_intent_punto2d(
                    hoja, tg, inner["curve"], cx + ri, cy
                )
                if int_o and int_i:
                    pieza_bb = _bbox_global(datos)
                    clr = clearance_texto_cota_cm(10) + OFFSET_COTA
                    from cota_estilo import empujar_punto_fuera_bbox

                    try:
                        sheet_w = float(hoja.Width)
                        sheet_h = float(hoja.Height)
                        minx, maxx, miny, maxy = pieza_bb
                        aire = {
                            "der": sheet_w - maxx,
                            "izq": minx,
                            "sup": sheet_h - maxy,
                            "inf": miny,
                        }
                        lado = max(aire, key=aire.get)
                    except Exception:
                        lado = "der"
                    tx, ty = empujar_punto_fuera_bbox(
                        cx, cy, pieza_bb, lado, clr
                    )
                    pt = _clampear_punto_hoja(
                        hoja, tg, tx, ty, evitar_bbox=pieza_bb
                    )
                    dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                        pt, int_i, int_o, kHorizontalDimensionType
                    )
                    aplicar_estilo_cota(dim, hoja=hoja)
                    asegurar_cota_fuera_pieza_robusto(
                        dim, tg, pieza_bb, holgura=0.6, n_chars=12
                    )
                    origen = (
                        "catálogo" if mejor["snap"].get("desde_catalogo") else "medido"
                    )
                    print(
                        f"✅ {nombre_hoja}: THK semicircular = "
                        f"{mejor['snap']['valor_in']:.4f} in ({origen})"
                    )
                    return True, {
                        "gap_sheet": mejor["gap_sheet"],
                        "valor_cm": mejor["valor_cm"],
                        "valor_in": mejor["snap"]["valor_in"],
                        "semicircular": True,
                    }
            except Exception as e:
                print(f"⚠️ {nombre_hoja}: fallo cota radial semicircular -> {e}")

    print(f"⚠️ {nombre_hoja}: no se pudo resolver THK semicircular.")
    return False, None


def _es_vista_cara_plana(datos):
    """
    Vista que mira la cara grande (p. ej. placa con agujero pasado): no es útil
    para THK porque no expone el espesor lateral.

    Restricciones (todas deben cumplirse para marcar como cara plana):

    1. Debe existir al menos un AGUJERO CIRCULAR INTERIOR — círculo cuyo bbox
       esté claramente dentro del bbox global (no pegado al borde). Esto
       descarta los redondeos de esquina de placas y perfiles, que técnicamente
       son arcos pero NO son "agujeros pasados".
    2. Aspect ratio del bbox global < 1.8 (placa casi cuadrada). Perfiles U/L
       vistos de canto son alargados y no cumplen esto.
    3. Silueta exterior con los 4 lados formados por segmentos rectos que
       cubran cada uno >=40% del ancho/alto. Una U tiene solo 2 o 3 lados
       cerrados.
    """
    if not datos:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    menor = min(w, h)
    if menor <= EPS:
        return False

    # Regla 1: al menos un agujero INTERIOR (no redondeo de esquina).
    circulos = _buscar_circulos(datos)
    margen_interior = max(0.20, menor * 0.08)  # separación mínima del borde
    agujeros_interiores = []
    for c in circulos:
        diam = max(c["dx"], c["dy"])
        if diam < 0.25:
            continue
        cx = c.get("cx")
        cy = c.get("cy")
        radio = diam / 2.0
        if cx is None or cy is None:
            continue
        if (cx - radio) < minx + margen_interior:
            continue
        if (cx + radio) > maxx - margen_interior:
            continue
        if (cy - radio) < miny + margen_interior:
            continue
        if (cy + radio) > maxy - margen_interior:
            continue
        agujeros_interiores.append(c)

    if not agujeros_interiores:
        return False

    # Regla 2: aspect ratio.
    aspect = max(w, h) / menor
    if aspect >= 1.8:
        return False

    # Regla 3: los 4 lados del bbox exterior cerrados con rectos.
    tol_h = max(0.1, h * 0.05)
    tol_w = max(0.1, w * 0.05)

    def _hay_lado_horizontal(y_lado):
        for d in datos:
            if d["dy"] < max(0.05, d["dx"] * 0.2):
                y_prom = (d["miny"] + d["maxy"]) / 2.0
                if abs(y_prom - y_lado) <= tol_h:
                    if (d["maxx"] - d["minx"]) >= w * 0.40:
                        return True
        return False

    def _hay_lado_vertical(x_lado):
        for d in datos:
            if d["dx"] < max(0.05, d["dy"] * 0.2):
                x_prom = (d["minx"] + d["maxx"]) / 2.0
                if abs(x_prom - x_lado) <= tol_w:
                    if (d["maxy"] - d["miny"]) >= h * 0.40:
                        return True
        return False

    lados_cerrados = 0
    if _hay_lado_horizontal(miny):
        lados_cerrados += 1
    if _hay_lado_horizontal(maxy):
        lados_cerrados += 1
    if _hay_lado_vertical(minx):
        lados_cerrados += 1
    if _hay_lado_vertical(maxx):
        lados_cerrados += 1

    return lados_cerrados >= 4


def _espesor_chapa_desde_vista(vista):
    """Lee Thickness del SheetMetalComponentDefinition referenciado por la vista."""
    try:
        doc = None
        try:
            doc = vista.ReferencedDocumentDescriptor.ReferencedDocument
        except Exception:
            try:
                doc = vista.ReferencedFile.DocumentDescriptor.ReferencedDocument
            except Exception:
                doc = None
        if doc is None:
            return None
        sm_def = win32com.client.CastTo(
            doc.ComponentDefinition, "SheetMetalComponentDefinition"
        )
        # Thickness.Value está en cm internos de Inventor
        return float(sm_def.Thickness.Value)
    except Exception:
        return None


def _dimensiones_bbox_3d(vista):
    """Devuelve [dx, dy, dz] del bbox 3D del modelo referenciado, en cm."""
    try:
        doc = None
        try:
            doc = vista.ReferencedDocumentDescriptor.ReferencedDocument
        except Exception:
            try:
                doc = vista.ReferencedFile.DocumentDescriptor.ReferencedDocument
            except Exception:
                doc = None
        if doc is None:
            return None
        rb = doc.ComponentDefinition.RangeBox
        return [
            abs(float(rb.MaxPoint.X) - float(rb.MinPoint.X)),
            abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y)),
            abs(float(rb.MaxPoint.Z) - float(rb.MinPoint.Z)),
        ]
    except Exception:
        return None


def _espesor_desde_bbox_3d(vista):
    """
    Fallback para partes no sheet metal: usa la dimensión MÁS PEQUEÑA del
    bbox 3D de la pieza como espesor. En cm de Inventor.

    OJO: en perfiles U/L (Parking) el menor del bbox suele ser el ALTO del
    doblez (p. ej. 0.875 in), NO el canto de pared. No usar solo para
    etiquetar THK sin validar con ``_espesor_thk_validado``.
    """
    dims = _dimensiones_bbox_3d(vista)
    if not dims:
        return None
    dims.sort()
    menor = dims[0]
    if menor <= EPS:
        return None
    return menor


def _thk_duplica_referencia(thk_cm, ref_cm, tol_rel=0.12):
    """True si thk ≈ ref (p. ej. bbox menor = alto patas→base)."""
    if thk_cm is None or ref_cm is None:
        return False
    if float(thk_cm) <= EPS or float(ref_cm) <= EPS:
        return False
    return abs(float(thk_cm) - float(ref_cm)) / float(ref_cm) <= tol_rel


def _espesor_pared_desde_curvas(vista, datos, alto_cm=None):
    """
    Candidato de pared (gap fino) desde curvas 2D, descartando gaps ≈ ALTO.
    """
    if not datos:
        return None
    candidatos = _buscar_candidatos_lineales(datos)
    if not candidatos:
        return None
    vals = []
    for c in candidatos:
        try:
            v = _esperado_modelo(vista, c["gap_sheet"])
        except Exception:
            continue
        if v is None or v <= EPS:
            continue
        if alto_cm and _thk_duplica_referencia(v, alto_cm, tol_rel=0.40):
            continue
        # Espesores de chapa/barra típicos; evita anchos de placa.
        if v > 1.25 * IN_TO_CM:
            continue
        vals.append(float(v))
    if not vals:
        return None
    return min(vals)


def _espesor_thk_validado(vista, alto_cm=None, datos=None, nombre_hoja=None):
    """
    Espesor usable para etiquetar/exportar como THK.

    Prioridad BOARD/GIGA (nombre pieza completo):
      1) Sheet Metal Thickness (canónico en chapa)
      2) Gap fino de curvas 2D (pared)
      3) Bbox 3D menor si no duplica ALTO

    Prioridad tanque:
      1) Gap fino de curvas 2D (pared)
      2) Thickness de Sheet Metal (si no duplica el ALTO)
      3) Bbox 3D menor SOLO si es claramente más fino que el ALTO

    Si no hay ancla confiable → None (mejor omitir THK que confundir).
    """
    parking = _nombre_parece_parking_u(nombre_hoja)

    board_giga = False
    try:
        from creador_vistas import get_nombre_pieza_completo

        board_giga = bool(get_nombre_pieza_completo())
    except Exception:
        board_giga = False

    if board_giga:
        chapa = _espesor_chapa_desde_vista(vista)
        if chapa is not None and chapa > EPS:
            if alto_cm and _thk_duplica_referencia(chapa, alto_cm):
                print(
                    f"⚠️ {nombre_hoja or 'hoja'}: Thickness BOARD "
                    f"≈ ALTO; se intenta pared/bbox."
                )
            else:
                return float(chapa), "sheet_metal_thickness_board"
        pared = _espesor_pared_desde_curvas(vista, datos, alto_cm=alto_cm)
        if pared is not None and pared > EPS:
            return float(pared), "curva_pared_board"
        bbox = _espesor_desde_bbox_3d(vista)
        if bbox is not None and bbox > EPS:
            if not (
                alto_cm
                and (
                    _thk_duplica_referencia(bbox, alto_cm)
                    or float(bbox) >= float(alto_cm) * 0.45
                )
            ):
                return float(bbox), "bbox_3d_board"
        return None, None

    pared = _espesor_pared_desde_curvas(vista, datos, alto_cm=alto_cm)
    if pared is not None and pared > EPS:
        return float(pared), "curva_pared"

    chapa = _espesor_chapa_desde_vista(vista)
    if chapa is not None and chapa > EPS:
        if alto_cm and _thk_duplica_referencia(chapa, alto_cm):
            print(
                f"⚠️ {nombre_hoja or 'hoja'}: Thickness de chapa "
                f"({chapa / IN_TO_CM:.4f} in) ≈ ALTO; no se usa como THK."
            )
        else:
            return float(chapa), "sheet_metal_thickness"

    if parking:
        # En U Parking el bbox menor = alto del doblez con mucha frecuencia.
        return None, None

    bbox = _espesor_desde_bbox_3d(vista)
    if bbox is None or bbox <= EPS:
        return None, None
    if alto_cm and (
        _thk_duplica_referencia(bbox, alto_cm)
        or float(bbox) >= float(alto_cm) * 0.45
    ):
        return None, None
    return float(bbox), "bbox_3d_menor_validado"


def _alinear_cota_thk_a_chapa_board(hoja, vista, nombre_hoja, meta=None):
    """
    En BOARD: si hay Sheet Metal Thickness y la cota dibujada difiere,
    fuerza el texto visible al Thickness (LADO solo ancla visual).
    """
    try:
        from creador_vistas import get_nombre_pieza_completo

        if not get_nombre_pieza_completo():
            return False
    except Exception:
        return False

    chapa = _espesor_chapa_desde_vista(vista)
    if chapa is None or chapa <= EPS:
        return False

    try:
        from cota_estilo import texto_cota_dibujo

        texto = texto_cota_dibujo(chapa, hoja)
    except Exception:
        texto = f"{chapa / IN_TO_CM:.4f}"

    if not texto:
        return False

    cambiado = False
    try:
        dims = hoja.DrawingDimensions.GeneralDimensions
        for di in range(1, dims.Count + 1):
            dim = dims.Item(di)
            try:
                mv = float(dim.ModelValue)
            except Exception:
                mv = None
            if mv is not None and abs(mv - chapa) / max(chapa, EPS) <= 0.03:
                continue
            try:
                dim.HideValue = True
            except Exception:
                pass
            bold = "True"
            try:
                from cota_estilo import COTA_BOLD, COTA_FONT_SIZE_CM

                bold = "True" if COTA_BOLD else "False"
                formatted = (
                    f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' "
                    f"Bold='{bold}'>{texto}</StyleOverride>"
                )
            except Exception:
                formatted = texto
            try:
                dim.Text.FormattedText = formatted
                cambiado = True
            except Exception:
                try:
                    dim.Text.Text = texto
                    cambiado = True
                except Exception:
                    pass
    except Exception:
        return False

    if cambiado:
        if isinstance(meta, dict):
            meta["valor_cm"] = float(chapa)
            meta["gap_sheet"] = float(chapa)
            meta["origen_thk"] = "sheet_metal_thickness_board_forced"
        print(
            f"↩️ {nombre_hoja}: THK texto alineado a Sheet Metal "
            f"Thickness ({chapa / IN_TO_CM:.4f} in)"
        )
    return cambiado


def _forzar_cota_thk_desde_modelo(hoja, tg, vista, nombre_hoja, alto_cm=None, datos=None):
    """
    Fallback tipográfico SOLO con espesor validado (Sheet Metal o pared 2D).

    Ya no publica ``THK = …`` desde bbox 3D crudo: en perfiles U eso
    etiquetaba el ALTO (0.875) como THK y confundía las capturas.

    Retorna (True, valor_cm) si logró agregar la nota, (False, None) si no.
    """
    valor_cm, origen = _espesor_thk_validado(
        vista, alto_cm=alto_cm, datos=datos, nombre_hoja=nombre_hoja
    )
    if valor_cm is None or valor_cm <= EPS:
        print(
            f"⚠️ {nombre_hoja}: sin THK validado (ni Sheet Metal Thickness "
            f"ni pared 2D); no se etiqueta como THK."
        )
        return False, None

    try:
        from cota_estilo import texto_cota_dibujo

        texto = f"THK = {texto_cota_dibujo(valor_cm)}"
    except Exception:
        valor_in = valor_cm / IN_TO_CM
        texto = f"THK = {valor_in:.3f} in"

    # Posición de la nota: al lado derecho de la vista, cerca de la esquina
    # superior. Se clampea al sheet para no salirse.
    try:
        left = float(vista.Left)
        top = float(vista.Top)
        width = float(vista.Width)
        pt_x = left + width + 1.0
        pt_y = top - 0.6
    except Exception:
        pt_x, pt_y = 5.0, 5.0

    pt = _clampear_punto_hoja(hoja, tg, pt_x, pt_y)

    try:
        gn = hoja.DrawingNotes.GeneralNotes.AddFitted(pt, texto)
        # Misma legibilidad que las cotas (+25% fuente vía cota_estilo).
        try:
            from cota_estilo import COTA_FONT_SIZE_CM, COTA_BOLD, COTA_NAVY_RGB
            bold = "True" if COTA_BOLD else "False"
            gn.FormattedText = (
                f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' Bold='{bold}'>"
                f"{texto}</StyleOverride>"
            )
            app_nota = None
            try:
                app_nota = conectar_inventor()
            except Exception:
                app_nota = None
            if app_nota is not None:
                r, g, b = COTA_NAVY_RGB
                gn.Color = app_nota.TransientObjects.CreateColor(r, g, b)
        except Exception:
            pass
    except Exception as exc:
        print(
            f"⚠️ {nombre_hoja}: no se pudo crear nota THK forzada ({exc})."
        )
        return False, None

    print(
        f"↩️ {nombre_hoja}: THK forzado desde modelo = {texto} "
        f"(origen={origen})"
    )
    return True, valor_cm


def _forzar_nota_dimension_individual(
    hoja, tg, vista, nombre_hoja, etiqueta, valor_cm
):
    """
    Coloca UNA sola ``GeneralNote`` centrada bajo la vista, con el formato
    ``ETIQUETA = X.XXX in|mm``. Se usa como fallback cuando la cota geométrica
    en las hojas _ALTO / _LARGO_PATA no cabe o falla.

    Retorna True si logró añadir la nota, False si no.
    """
    if valor_cm is None or valor_cm <= EPS:
        return False

    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
    except Exception:
        sheet_w = 40.0
        sheet_h = 30.0

    # Posición: DEBAJO del centro de la vista, con margen a los bordes.
    try:
        left = float(vista.Left)
        top = float(vista.Top)
        width = float(vista.Width)
        height = float(vista.Height)
        pt_x = left + width / 2.0
        pt_y = top - height - 1.2
    except Exception:
        pt_x = sheet_w / 2.0
        pt_y = sheet_h * 0.15

    pt = _clampear_punto_hoja(hoja, tg, pt_x, pt_y, margen=1.5)
    try:
        from cota_estilo import texto_cota_dibujo

        texto = f"{etiqueta} = {texto_cota_dibujo(valor_cm)}"
    except Exception:
        valor_in = valor_cm / IN_TO_CM
        texto = f"{etiqueta} = {valor_in:.3f} in"

    try:
        hoja.DrawingNotes.GeneralNotes.AddFitted(pt, texto)
    except Exception as exc:
        print(f"⚠️ {nombre_hoja}: no se pudo crear nota {etiqueta} ({exc}).")
        return False

    print(f"↩️ {nombre_hoja}: {etiqueta} forzado desde bbox 3D = {texto}")
    return True


def _forzar_notas_perfil_desde_modelo(hoja, tg, vista, nombre_hoja, thk_sheet=None):
    """
    Fallback cuando ``_crear_hoja_alto`` falla al crear la hoja extra:
    coloca notas de texto con las dimensiones transversales del perfil
    (ALTO cuerpo y LARGO pata) leídas del bbox 3D del modelo directamente
    sobre la MISMA hoja de LADO/THK.

    Aplica sólo a piezas con aspect ratio de perfil (una dimensión >> las
    otras dos: mayor/medio >= 3). Así garantizamos que aunque la creación
    de la hoja extra falle con COM error, el JPG del _LADO/_THK muestre
    los tres números necesarios para verificar bend deduction (THK, ALTO,
    LARGO_PATA).

    Devuelve True si logró añadir al menos una nota, False si no aplica.
    """
    dims = _dimensiones_bbox_3d(vista)
    if not dims or len(dims) < 3:
        return False

    dims_sorted = sorted(dims)  # [menor, medio, mayor]
    menor_in = dims_sorted[0] / IN_TO_CM
    medio_in = dims_sorted[1] / IN_TO_CM
    mayor_in = dims_sorted[2] / IN_TO_CM

    # Sólo perfiles largos (viga/canal/escuadra). Placas y barras no aplican.
    if medio_in <= EPS or (mayor_in / medio_in) < 3.0:
        return False

    alto_perfil_in = medio_in
    largo_pata_in = menor_in

    # Si el "largo de pata" coincide con el espesor de chapa, es una placa
    # de canto, no un perfil doblado → no publicar cifra redundante.
    if thk_sheet is not None and thk_sheet > EPS:
        thk_in = thk_sheet / IN_TO_CM
        if abs(largo_pata_in - thk_in) <= max(0.03, thk_in * 0.20):
            return False

    textos = [
        f"ALTO ~ {alto_perfil_in:.3f} in",
        f"LARGO_PATA ~ {largo_pata_in:.3f} in",
    ]

    try:
        left = float(vista.Left)
        top = float(vista.Top)
        height = float(vista.Height)
        pt_x = left + 1.0
        pt_y = top - height - 1.5
    except Exception:
        pt_x, pt_y = 3.0, 3.0

    exitos = 0
    for i, t in enumerate(textos):
        pt = _clampear_punto_hoja(hoja, tg, pt_x, pt_y - i * 1.0)
        try:
            hoja.DrawingNotes.GeneralNotes.AddFitted(pt, t)
            exitos += 1
        except Exception:
            continue

    if exitos > 0:
        print(
            f"↩️ {nombre_hoja}: notas perfil forzadas desde modelo "
            f"(ALTO={alto_perfil_in:.3f}in, PATA={largo_pata_in:.3f}in)"
        )
        return True
    return False


def _es_tubo_rectangular_hueco(datos):
    """
    Detecta un perfil de tubo rectangular/cuadrado hueco (HSS).

    Silueta: rectángulo exterior + rectángulo interior más pequeño con la
    misma proporción y centrados. Devuelve un dict con las envolventes
    exterior/interior o ``None`` si no aplica.

    Restricciones estrictas para evitar falsos positivos como:
    - Placas con agujeros circulares (los círculos son "interiores" pero NO
      forman un rectángulo).
    - Perfiles L (tienen contorno interior pero es un ángulo, no un rectángulo
      completo).
    - Piezas con features arbitrarios cerca del centro.

    Requisitos que TODOS deben cumplirse:
    1. Los 4 espesores de pared deben existir y ser positivos.
    2. Los espesores deben ser consistentes (±25% del promedio).
    3. El contorno interior debe estar formado por al menos 2 líneas rectas
       horizontales y 2 rectas verticales que cubran >=60% del perímetro
       interior esperado (esto descarta el contorno interior formado por
       círculos aislados, que son placas con agujeros, no HSS).
    4. El interior NO debe estar dominado por círculos (una placa con
       barrenos tiene círculos como "interiores"; los círculos no son un
       rectángulo).
    """
    if not datos:
        return None
    minx_g, maxx_g, miny_g, maxy_g = _bbox_global(datos)
    w_g = maxx_g - minx_g
    h_g = maxy_g - miny_g
    if w_g <= EPS or h_g <= EPS:
        return None

    margen_x = max(0.05, w_g * 0.05)
    margen_y = max(0.05, h_g * 0.05)

    interiores = []
    for d in datos:
        toca_borde = (
            abs(d["minx"] - minx_g) <= margen_x * 0.5
            or abs(d["maxx"] - maxx_g) <= margen_x * 0.5
            or abs(d["miny"] - miny_g) <= margen_y * 0.5
            or abs(d["maxy"] - maxy_g) <= margen_y * 0.5
        )
        if not toca_borde:
            interiores.append(d)

    if len(interiores) < 4:
        # HSS de verdad tiene al menos 4 líneas rectas interiores (los 4 lados
        # del hueco). Menos que eso probablemente sea un feature aislado.
        return None

    # ---- Descartar cuando el "interior" es mayoritariamente CIRCULAR ----
    # Placas con barrenos: los círculos aparecen como interiores pero NO
    # forman un contorno rectangular. Si más de 40% del interior son curvas
    # cerradas / arcos con dx≈dy, es una placa con agujeros y NO un HSS.
    circulos_interiores = 0
    for d in interiores:
        # Un círculo/arco tiene dx y dy similares (no es una línea recta).
        if d["dx"] > EPS and d["dy"] > EPS:
            razon = min(d["dx"], d["dy"]) / max(d["dx"], d["dy"])
            if razon >= 0.5:  # forma cuasi circular
                circulos_interiores += 1
    if circulos_interiores >= max(1, len(interiores) * 0.4):
        return None

    minx_i = min(d["minx"] for d in interiores)
    maxx_i = max(d["maxx"] for d in interiores)
    miny_i = min(d["miny"] for d in interiores)
    maxy_i = max(d["maxy"] for d in interiores)
    w_i = maxx_i - minx_i
    h_i = maxy_i - miny_i
    if w_i <= EPS or h_i <= EPS:
        return None

    if w_i >= w_g - margen_x or h_i >= h_g - margen_y:
        return None

    # Los 4 espesores DEBEN existir (perfil L / U tendría uno o dos en cero
    # porque el interior está pegado a un borde).
    thk_izq = minx_i - minx_g
    thk_der = maxx_g - maxx_i
    thk_inf = miny_i - miny_g
    thk_sup = maxy_g - maxy_i
    espesores = [thk_izq, thk_der, thk_inf, thk_sup]
    tol_pared = max(0.05, min(w_g, h_g) * 0.02)
    if any(t <= tol_pared for t in espesores):
        # Al menos una pared "no existe" -> es un L o una U, no un HSS.
        return None

    thk = sum(espesores) / 4.0
    if thk <= EPS:
        return None
    if any(abs(e - thk) > thk * 0.35 for e in espesores):
        return None

    # ---- Verificar contorno interior RECTANGULAR ----
    # Debe haber líneas rectas horizontales cerca de miny_i y maxy_i, y
    # verticales cerca de minx_i y maxx_i. Cada lado del rectángulo interior
    # debe cubrir >=50% del ancho/alto interior con líneas rectas reales.
    tol_i_h = max(0.05, h_i * 0.1)
    tol_i_v = max(0.05, w_i * 0.1)

    def _cobertura_horizontal(y_target):
        cubierto = 0.0
        for d in interiores:
            if d["dy"] >= max(0.05, d["dx"] * 0.2):
                continue  # no es horizontal
            y_prom = (d["miny"] + d["maxy"]) / 2.0
            if abs(y_prom - y_target) <= tol_i_h:
                cubierto += (d["maxx"] - d["minx"])
        return cubierto

    def _cobertura_vertical(x_target):
        cubierto = 0.0
        for d in interiores:
            if d["dx"] >= max(0.05, d["dy"] * 0.2):
                continue
            x_prom = (d["minx"] + d["maxx"]) / 2.0
            if abs(x_prom - x_target) <= tol_i_v:
                cubierto += (d["maxy"] - d["miny"])
        return cubierto

    cob_min = 0.5
    if _cobertura_horizontal(miny_i) < w_i * cob_min:
        return None
    if _cobertura_horizontal(maxy_i) < w_i * cob_min:
        return None
    if _cobertura_vertical(minx_i) < h_i * cob_min:
        return None
    if _cobertura_vertical(maxx_i) < h_i * cob_min:
        return None

    return {
        "gap_sheet": thk,
        "bbox_ext": (minx_g, maxx_g, miny_g, maxy_g),
        "bbox_int": (minx_i, maxx_i, miny_i, maxy_i),
    }


def _resolver_rectangular_hollow(hoja, vista, tg, datos, nombre_hoja):
    """
    THK sobre un tubo rectangular hueco. Dibuja UNA cota entre la pared
    exterior y la interior del mismo lado. Intenta en este orden: pared
    izquierda (cota horizontal), pared inferior (cota vertical), derecha,
    superior. La primera que consiga par exterior+interior gana.

    Deja print()s siempre visibles (sin depender de _THK_LOG) porque este
    resolver está históricamente bugueado y el usuario necesita ver por qué
    falla cuando falla.
    """
    envolvente = _es_tubo_rectangular_hueco(datos)
    if not envolvente:
        print(f"⚠️ {nombre_hoja}: _es_tubo_rectangular_hueco devolvió None — no se detectó HSS.")
        return False, None

    thk_sheet = envolvente["gap_sheet"]
    valor_cm = _esperado_modelo(vista, thk_sheet)
    snap = _snap_o_medido(valor_cm)

    minx_g, maxx_g, miny_g, maxy_g = envolvente["bbox_ext"]
    minx_i, maxx_i, miny_i, maxy_i = envolvente["bbox_int"]

    # Tolerancia amplia: HSS con esquinas redondeadas puede tener las líneas
    # rectas un poco desplazadas del bbox. Antes usábamos 0.03..thk*0.35 y
    # fallaba silenciosamente. Ampliamos a max(0.30, thk*1.0).
    tol = max(0.30, thk_sheet * 1.0)

    horiz = [d for d in datos if d["dy"] < max(0.05, d["dx"] * 0.2)]
    vert = [d for d in datos if d["dx"] < max(0.05, d["dy"] * 0.2)]

    def _mejor_horiz(y_target):
        """Línea horizontal (dy≈0) más cercana a y_target y con buena longitud."""
        candidatos = [
            d for d in horiz
            if abs(((d["miny"] + d["maxy"]) / 2.0) - y_target) <= tol
        ]
        if not candidatos:
            return None
        # Preferir la línea más larga (más contorno).
        candidatos.sort(key=lambda d: -(d["maxx"] - d["minx"]))
        return candidatos[0]

    def _mejor_vert(x_target):
        candidatos = [
            d for d in vert
            if abs(((d["minx"] + d["maxx"]) / 2.0) - x_target) <= tol
        ]
        if not candidatos:
            return None
        candidatos.sort(key=lambda d: -(d["maxy"] - d["miny"]))
        return candidatos[0]

    # Cada intento es (nombre_pared, dir_cota, curva_ext, curva_int, punto_texto)
    intentos = []

    off = max(0.5, thk_sheet * 3)

    # Pared izquierda (cota HORIZONTAL de minx_g a minx_i)
    v_ext_izq = _mejor_vert(minx_g)
    v_int_izq = _mejor_vert(minx_i)
    if v_ext_izq is not None and v_int_izq is not None and v_ext_izq is not v_int_izq:
        pt = _clampear_punto_hoja(
            hoja, tg,
            (minx_g + minx_i) / 2.0,
            miny_g - off,
        )
        intentos.append(("izquierda", kHorizontalDimensionType, v_ext_izq, v_int_izq, pt))

    # Pared inferior (cota VERTICAL de miny_g a miny_i)
    h_ext_inf = _mejor_horiz(miny_g)
    h_int_inf = _mejor_horiz(miny_i)
    if h_ext_inf is not None and h_int_inf is not None and h_ext_inf is not h_int_inf:
        pt = _clampear_punto_hoja(
            hoja, tg,
            minx_g - off,
            (miny_g + miny_i) / 2.0,
        )
        intentos.append(("inferior", kVerticalDimensionType, h_ext_inf, h_int_inf, pt))

    # Pared derecha (cota HORIZONTAL de maxx_i a maxx_g)
    v_ext_der = _mejor_vert(maxx_g)
    v_int_der = _mejor_vert(maxx_i)
    if v_ext_der is not None and v_int_der is not None and v_ext_der is not v_int_der:
        pt = _clampear_punto_hoja(
            hoja, tg,
            (maxx_g + maxx_i) / 2.0,
            miny_g - off,
        )
        intentos.append(("derecha", kHorizontalDimensionType, v_ext_der, v_int_der, pt))

    # Pared superior (cota VERTICAL)
    h_ext_sup = _mejor_horiz(maxy_g)
    h_int_sup = _mejor_horiz(maxy_i)
    if h_ext_sup is not None and h_int_sup is not None and h_ext_sup is not h_int_sup:
        pt = _clampear_punto_hoja(
            hoja, tg,
            minx_g - off,
            (maxy_g + maxy_i) / 2.0,
        )
        intentos.append(("superior", kVerticalDimensionType, h_ext_sup, h_int_sup, pt))

    if not intentos:
        print(
            f"⚠️ {nombre_hoja}: HSS detectado pero ninguna pared exterior/interior "
            f"quedó bien matcheada dentro de tol={tol:.3f}cm."
        )
        return False, None

    ultimo_error = None
    for nombre_pared, dir_dim, curva_ext, curva_int, pt_texto in intentos:
        try:
            int_ext = hoja.CreateGeometryIntent(curva_ext["curve"])
            int_int = hoja.CreateGeometryIntent(curva_int["curve"])
            dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                pt_texto, int_ext, int_int, dir_dim
            )
            try:
                dim.Text.Text = _formato_valor(snap["valor_in"])
            except Exception:
                pass
            aplicar_estilo_cota(dim, hoja=hoja)
            origen = "catálogo" if snap.get("desde_catalogo") else "medido"
            print(
                f"✅ {nombre_hoja}: THK HSS = {snap['valor_in']:.4f} in "
                f"(pared {nombre_pared}, {origen}, detectado {valor_cm / IN_TO_CM:.4f} in)"
            )
            return True, {
                "gap_sheet": thk_sheet,
                "valor_cm": valor_cm,
                "valor_in": snap["valor_in"],
                "snap": snap,
            }
        except Exception as e:
            ultimo_error = e
            continue

    print(
        f"⚠️ {nombre_hoja}: HSS detectado pero Inventor rechazó AddLinear en las "
        f"4 paredes intentadas. Último error: {ultimo_error}"
    )
    return False, None


def _clasificar_lado(datos):
    """
    Devuelve:
    - ("circular_solid", outer, None)
    - ("circular_hollow", outer, inner)
    - ("rect_hollow", None, None)  (tubo rectangular hueco, HSS)
    - ("prismatic", None, None)
    """
    if not datos:
        return ("prismatic", None, None)

    minx, maxx, miny, maxy = _bbox_global(datos)
    global_w = maxx - minx
    global_h = maxy - miny

    circulos = _buscar_circulos(datos)
    if not circulos:
        # Sin círculos no es circular; puede ser HSS (tubo rectangular hueco).
        if _es_tubo_rectangular_hueco(datos):
            return ("rect_hollow", None, None)
        return ("prismatic", None, None)

    outer = max(circulos, key=lambda d: d["dx"])

    # Para considerar que el contorno principal es circular,
    # el círculo mayor debe parecer coincidir con la envolvente global.
    tol_w = max(0.05, global_w * 0.10)
    tol_h = max(0.05, global_h * 0.10)

    bbox_match = (
        abs(outer["dx"] - global_w) <= tol_w and
        abs(outer["dy"] - global_h) <= tol_h
    )

    if not bbox_match:
        if _es_tubo_rectangular_hueco(datos):
            return ("rect_hollow", None, None)
        return ("prismatic", None, None)

    # Buscar círculos interiores concéntricos con el exterior
    inner_candidates = []
    center_tol = max(0.03, outer["dx"] * 0.03)

    for c in circulos:
        if c is outer:
            continue

        if abs(c["cx"] - outer["cx"]) <= center_tol and abs(c["cy"] - outer["cy"]) <= center_tol:
            if c["dx"] < outer["dx"]:
                inner_candidates.append(c)

    if inner_candidates:
        # Para THK circular queremos el interior MÁS GRANDE (el más cercano a la pared)
        inner = max(inner_candidates, key=lambda d: d["dx"])
        return ("circular_hollow", outer, inner)

    # Aún si el contorno "principal" parece circular, si además hay una
    # silueta clara de rectángulo hueco, dejamos que el resolver de HSS lo
    # atrape (esquinas redondeadas de tubos rectangulares).
    if _es_tubo_rectangular_hueco(datos):
        return ("rect_hollow", None, None)

    return ("circular_solid", outer, None)


def _buscar_candidatos_lineales(datos):
    """
    Devuelve candidatos de espesor lineal en piezas prismáticas.
    Busca separaciones entre pares de líneas casi paralelas con buen traslape.
    """
    candidatos = []

    verticales = []
    horizontales = []

    for d in datos:
        # Línea predominantemente vertical
        if d["dy"] >= max(0.05, d["dx"] * 5.0):
            verticales.append(d)
        # Línea predominantemente horizontal
        elif d["dx"] >= max(0.05, d["dy"] * 5.0):
            horizontales.append(d)

    # Espesor horizontal entre líneas verticales
    for i in range(len(verticales)):
        for j in range(i + 1, len(verticales)):
            a = verticales[i]
            b = verticales[j]

            overlap_y = min(a["maxy"], b["maxy"]) - max(a["miny"], b["miny"])
            if overlap_y <= 0:
                continue

            min_len = min(a["dy"], b["dy"])
            if overlap_y < min_len * 0.40:
                continue

            xa = (a["minx"] + a["maxx"]) / 2.0
            xb = (b["minx"] + b["maxx"]) / 2.0
            gap = abs(xa - xb)

            if gap <= EPS:
                continue

            # Caras deben ser más largas que el "espesor": evita medir la
            # pata recta de un U de barra (HV Parking) entre tope corto y
            # tangente de radio.
            cara = min(a["dy"], b["dy"])
            # Cara de espesor debe ser claramente más larga que el gap.
            # 1.25x no bastó (pata U HV Parking ~cara≈gap).
            if cara < max(0.4, gap * 3.0):
                continue

            candidatos.append({
                "tipo": "horizontal",
                "gap_sheet": gap,
                "a": a,
                "b": b,
                "overlap": overlap_y
            })

    # Espesor vertical entre líneas horizontales
    for i in range(len(horizontales)):
        for j in range(i + 1, len(horizontales)):
            a = horizontales[i]
            b = horizontales[j]

            overlap_x = min(a["maxx"], b["maxx"]) - max(a["minx"], b["minx"])
            if overlap_x <= 0:
                continue

            min_len = min(a["dx"], b["dx"])
            if overlap_x < min_len * 0.40:
                continue

            ya = (a["miny"] + a["maxy"]) / 2.0
            yb = (b["miny"] + b["maxy"]) / 2.0
            gap = abs(ya - yb)

            if gap <= EPS:
                continue

            cara = min(a["dx"], b["dx"])
            if cara < max(0.4, gap * 3.0):
                continue

            candidatos.append({
                "tipo": "vertical",
                "gap_sheet": gap,
                "a": a,
                "b": b,
                "overlap": overlap_x
            })

    return candidatos


def _candidato_bbox_menor(datos):
    """
    Rescate para vistas de canto (placa larga): el espesor ≈ el lado menor
    del bbox global, usando curvas en los extremos.
    """
    if not datos:
        return None
    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if min(w, h) <= EPS:
        return None

    if h <= w:
        a = max(datos, key=lambda d: d["maxy"])
        b = min(datos, key=lambda d: d["miny"])
        return {
            "tipo": "vertical",
            "gap_sheet": h,
            "a": a,
            "b": b,
            "overlap": w,
        }

    a = max(datos, key=lambda d: d["maxx"])
    b = min(datos, key=lambda d: d["minx"])
    return {
        "tipo": "horizontal",
        "gap_sheet": w,
        "a": a,
        "b": b,
        "overlap": h,
    }


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2, evitar_bbox=None):
    """Point2d dentro del sheet; nunca dentro de la silueta de la pieza."""
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
        if punto_dentro_bbox(x, y, evitar_bbox, holgura=0.15):
            x, y = empujar_punto_fuera_bbox(x, y, evitar_bbox, lado, clr)
    return tg.CreatePoint2d(x, y)


def _dibujar_cota_prismatica(hoja, tg, mejor):
    a = mejor["a"]
    b = mejor["b"]
    tipo = mejor["tipo"]

    if tipo == "horizontal":
        int_a = _intent_en_extremo(hoja, tg, a, "der" if a["cx"] >= b["cx"] else "izq")
        int_b = _intent_en_extremo(hoja, tg, b, "izq" if a["cx"] >= b["cx"] else "der")
        if int_a is None or int_b is None:
            int_a = hoja.CreateGeometryIntent(a["curve"])
            int_b = hoja.CreateGeometryIntent(b["curve"])
        x_texto = ((a["cx"] + b["cx"]) / 2.0)
        y_texto = max(a["maxy"], b["maxy"]) + OFFSET_COTA
        pieza_bb = (
            min(a["minx"], b["minx"]),
            max(a["maxx"], b["maxx"]),
            min(a["miny"], b["miny"]),
            max(a["maxy"], b["maxy"]),
        )
        pt_texto = _clampear_punto_hoja(
            hoja, tg, x_texto, y_texto, evitar_bbox=pieza_bb
        )
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_a, int_b, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)
        asegurar_cota_fuera_pieza(dim, tg, pieza_bb, "H")
        return dim
    else:
        int_a = _intent_en_extremo(hoja, tg, a, "sup" if a["cy"] >= b["cy"] else "inf")
        int_b = _intent_en_extremo(hoja, tg, b, "inf" if a["cy"] >= b["cy"] else "sup")
        if int_a is None or int_b is None:
            int_a = hoja.CreateGeometryIntent(a["curve"])
            int_b = hoja.CreateGeometryIntent(b["curve"])
        x_texto = min(a["minx"], b["minx"]) - OFFSET_COTA
        y_texto = ((a["cy"] + b["cy"]) / 2.0)
        pieza_bb = (
            min(a["minx"], b["minx"]),
            max(a["maxx"], b["maxx"]),
            min(a["miny"], b["miny"]),
            max(a["maxy"], b["maxy"]),
        )
        pt_texto = _clampear_punto_hoja(
            hoja, tg, x_texto, y_texto, evitar_bbox=pieza_bb
        )
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_a, int_b, kVerticalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)
        asegurar_cota_fuera_pieza(dim, tg, pieza_bb, "V")
        return dim


def _envolvente_espesor(datos):
    """
    Para vistas de canto (brida, placa con resaltes): el espesor real ≈ el
    lado menor del bbox global (la pieza suele ser ancha y poco alta).
    """
    if not datos:
        return None
    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if min(w, h) <= EPS:
        return None
    if h <= w:
        return {
            "gap_sheet": h,
            "tipo": "vertical",
            "a": max(datos, key=lambda d: d["maxy"]),
            "b": min(datos, key=lambda d: d["miny"]),
            "aspect": (w / h) if h > EPS else 999.0,
        }
    return {
        "gap_sheet": w,
        "tipo": "horizontal",
        "a": max(datos, key=lambda d: d["maxx"]),
        "b": min(datos, key=lambda d: d["minx"]),
        "aspect": (h / w) if w > EPS else 999.0,
    }


def _es_altura_hasta_radio(gap_sheet, overall_sheet):
    """
    True si el gap es casi la envolvente pero corto (pata U hasta el fillet).

    Caso HV Parking: medía 0.73 in en vez del alto exterior 0.875 in.
    """
    if gap_sheet is None or overall_sheet is None:
        return False
    if float(overall_sheet) <= EPS:
        return False
    ratio = float(gap_sheet) / float(overall_sheet)
    return 0.72 <= ratio < 0.97


def _nombre_parece_parking_u(nombre_hoja):
    u = str(nombre_hoja or "").upper()
    return "PARKING" in u


def _nombre_parece_brida(nombre_hoja):
    u = str(nombre_hoja or "").upper()
    return any(x in u for x in ("FLANGE", "FLANE", "BRIDA", "WELDNECK"))


def _nombre_parece_tierra(nombre_hoja):
    u = str(nombre_hoja or "").upper()
    return "TIERRA" in u or "GROUND" in u


def _nombre_parece_nipple(nombre_hoja):
    u = str(nombre_hoja or "").upper()
    return "NIPPLE" in u


def _necesita_alto_axial(nombre_hoja):
    """
    Cilindros / stubs / studs / bridas cuyo FRENTE es Ø y falta el largo
    axial (``_ALTO`` / HEIGHT).

    Incluye nipples, STUD, SP-751, PIPE/TUBE y bridas (PIPE FLANGE / FLANE /
    BRIDA). Antes se excluía ``FLANGE`` y solo el typo ``FLANE`` generaba
    HEIGHT (Vantran 0.375 sí, 0.250/1 no).
    """
    u = str(nombre_hoja or "").upper()
    if _nombre_parece_nipple(u):
        return True
    if "STUD" in u:
        return True
    if "SP-751" in u:
        return True
    # Bridas: HEIGHT = cuerpo (hub+disco), distinto del THK del disco.
    if _nombre_parece_brida(u):
        return True
    if "PIPE" in u:
        return True
    if re.search(r"\bTUBE\b", u) or u.endswith("TUBE") or "_TUBE" in u:
        return True
    return False


def _nombre_parece_jacking_escuadra(nombre_hoja):
    """Pieza L (jacking pads.ipt): THK debe ser la escuadra, no la cara plana."""
    u = str(nombre_hoja or "").upper()
    if "SOLERA" in u:
        return False
    return "JACKING" in u and "PAD" in u


def _nombre_parece_solera_jacking(nombre_hoja):
    """Solera Jacking Pad: necesita LARGO (eje mayor) además del THK de canto."""
    u = str(nombre_hoja or "").upper()
    return "SOLERA" in u and "JACKING" in u


def _resolver_parking_stands(hoja, vista, tg, datos, nombre_hoja):
    """
    HV Parking (U de barra/chapa):

    En la hoja LADO se acota patas → base (alto exterior de la silueta)
    para verificar bend deduction — NO el espesor de pared ni la pata
    hasta el radio (0.73).

    El THK real (Thickness de sheet metal / canto de barra) se resuelve
    aparte en una hoja extra.
    """
    if not datos:
        return False, None
    minx, maxx, miny, maxy = _bbox_global(datos)
    w_env = maxx - minx
    h_env = maxy - miny
    if max(w_env, h_env) <= EPS:
        return False, None
    # Canal ancho: el alto (patas→base) es el lado menor del bbox.
    orient = "V" if h_env <= w_env else "H"
    span = h_env if orient == "V" else w_env
    print(
        f"↩️ {nombre_hoja}: Parking — patas→base "
        f"({orient}, ≈ {_esperado_modelo(vista, span) / IN_TO_CM:.3f} in)"
    )
    ok = _acotar_lado_bbox(
        hoja, vista, tg, datos, nombre_hoja, orient, "patas-base"
    )
    if not ok:
        return False, None
    valor_cm = _esperado_modelo(vista, span)
    snap = _snap_o_medido(valor_cm)
    print(
        f"✅ {nombre_hoja}: patas→base = {snap['valor_in']:.4f} in "
        f"(bend deduction; no es THK de chapa)"
    )
    thk_chapa_cm = _espesor_chapa_desde_vista(vista)
    return True, {
        "gap_sheet": span,
        "valor_cm": valor_cm,
        "valor_in": snap["valor_in"],
        "es_patas_base": True,
        "thk_chapa_cm": thk_chapa_cm,
    }


def _resolver_prismatico(hoja, vista, tg, datos, nombre_hoja):
    """
    Returns (ok: bool, meta: dict|None).
    meta incluye gap_sheet / valor_cm para decidir si conviene hoja ALTO.
    """
    if _es_vista_cara_plana(datos):
        print(
            f"⚠️ {nombre_hoja}: vista parece cara plana (con agujero); "
            f"se intentará THK igualmente priorizando espesor de chapa."
        )

    # Parking SIEMPRE: aunque la vista parezca "cara plana" (U de frente).
    # Antes se saltaba y caía a nota-only → JPG omitido (sin ALTO/THK).
    if _nombre_parece_parking_u(nombre_hoja):
        return _resolver_parking_stands(hoja, vista, tg, datos, nombre_hoja)

    envolvente = _envolvente_espesor(datos)
    overall_sheet = envolvente["gap_sheet"] if envolvente else None
    overall_cm = (
        _esperado_modelo(vista, overall_sheet) if overall_sheet else None
    )
    # Solo en siluetas "aplanadas" (brida de canto: mucho más ancha que alta)
    # el espesor real es la envolvente. En L altas la franja 0.38 SÍ es THK.
    minx, maxx, miny, maxy = _bbox_global(datos)
    w_env = maxx - minx
    h_env = maxy - miny
    es_canto_aplanado = (
        overall_cm is not None
        and h_env > EPS
        and w_env >= h_env * 1.5
    )
    if overall_cm and es_canto_aplanado:
        print(
            f"  {nombre_hoja}: envolvente de espesor ≈ "
            f"{overall_cm / IN_TO_CM:.4f} in (canto aplanado)"
        )

    candidatos = _buscar_candidatos_lineales(datos)
    _dbg(
        f"{nombre_hoja}: prismatico bbox w={w_env / IN_TO_CM:.3f}in "
        f"h={h_env / IN_TO_CM:.3f}in canto_aplanado={es_canto_aplanado} "
        f"curvas={len(datos)} candidatos_lineales={len(candidatos) if candidatos else 0}"
    )
    if not candidatos:
        # Barra redonda / alambre en U: sin pares de "caras" largas; intentar Ø.
        # TIERRA/GROUND: Ø de la cara NO es THK (bug 3.14 flotante).
        circs = _buscar_circulos(datos)
        if circs and not _nombre_parece_tierra(nombre_hoja):
            outer = max(circs, key=lambda c: c.get("radius") or c["dx"] * 0.5)
            diam_sheet = (
                float(outer["radius"]) * 2.0
                if outer.get("radius")
                else float(outer["dx"])
            )
            if 0.05 < _esperado_modelo(vista, diam_sheet) / IN_TO_CM < 3.0:
                print(
                    f"↩️ {nombre_hoja}: sin pares de cara; rescate Ø barra "
                    f"≈ {_esperado_modelo(vista, diam_sheet) / IN_TO_CM:.3f} in"
                )
                try:
                    ok_c, meta_c = _resolver_circular_solid(
                        hoja, vista, tg, outer, nombre_hoja
                    )
                    if ok_c:
                        return ok_c, meta_c
                except Exception as exc_c:
                    _dbg(f"{nombre_hoja}: rescate Ø falló ({exc_c})")

        # Envolvente solo si parece canto aplanado (espesor << largo).
        if (
            envolvente
            and not _es_vista_cara_plana(datos)
            and (
                es_canto_aplanado
                or (
                    overall_cm is not None
                    and overall_cm <= 1.25 * IN_TO_CM
                )
            )
        ):
            print(f"↩️ {nombre_hoja}: sin pares lineales, rescate por envolvente.")
            candidatos = [envolvente]
        else:
            print(f"⚠️ {nombre_hoja}: no se encontraron candidatos prismáticos.")
            _dbg(
                f"{nombre_hoja}: descartado sin candidatos. "
                f"cara_plana={_es_vista_cara_plana(datos)} "
                f"envolvente={'sí' if envolvente else 'no'}"
            )
            return False, None
    elif envolvente and es_canto_aplanado:
        # Siempre considerar el extremo global (ancho real), no solo resaltes.
        candidatos = list(candidatos) + [envolvente]
    if _THK_LOG:
        for i, c in enumerate(candidatos):
            gap = c.get("gap_sheet")
            val_in = _esperado_modelo(vista, gap) / IN_TO_CM if gap else None
            _dbg(
                f"  cand#{i}: gap_sheet={gap:.4f} val≈{val_in:.4f}in tipo={c.get('tipo', '?')}"
                if val_in is not None else f"  cand#{i}: gap_sheet=None"
            )

    thk_chapa_cm = _espesor_chapa_desde_vista(vista)
    if thk_chapa_cm:
        print(
            f"  {nombre_hoja}: espesor de chapa del modelo = "
            f"{thk_chapa_cm / IN_TO_CM:.4f} in"
        )

    ranqueados = []
    for c in candidatos:
        valor_cm = _esperado_modelo(vista, c["gap_sheet"])
        # Descartar "espesores" absurdos (> 12 in): suelen ser largo/ancho.
        if valor_cm > 12.0 * IN_TO_CM:
            _dbg(f"  descarta cand val={valor_cm / IN_TO_CM:.3f}in >12in (largo)")
            continue
        # En cara plana, ignorar gaps enormes (son anchos de placa).
        if _es_vista_cara_plana(datos) and valor_cm > 1.5 * IN_TO_CM:
            _dbg(f"  descarta cand val={valor_cm / IN_TO_CM:.3f}in >1.5in (cara plana)")
            continue
        # En bridas de canto: descartar resaltes (0.06) << cuerpo real.
        if es_canto_aplanado and overall_cm and valor_cm < overall_cm * 0.35:
            _dbg(
                f"  descarta cand val={valor_cm / IN_TO_CM:.3f}in <35% de envolvente "
                f"{overall_cm / IN_TO_CM:.3f}in (resalte)"
            )
            continue
        # Pata U hasta el radio (0.73 vs alto 0.875): no es THK.
        overall_eje = h_env if c.get("tipo") == "vertical" else w_env
        if _es_altura_hasta_radio(c.get("gap_sheet"), overall_eje):
            _dbg(
                f"  descarta cand val={valor_cm / IN_TO_CM:.3f}in "
                f"(altura hasta radio; silueta≈"
                f"{_esperado_modelo(vista, overall_eje) / IN_TO_CM:.3f} in)"
            )
            continue
        c = dict(c)
        c["valor_cm"] = valor_cm
        c["snap"] = _snap_o_medido(valor_cm)
        ranqueados.append(c)

    if not ranqueados:
        # Si todo quedó filtrado por "superficial", forzar la envolvente.
        if envolvente and es_canto_aplanado and not _es_vista_cara_plana(datos):
            valor_cm = _esperado_modelo(vista, envolvente["gap_sheet"])
            c = dict(envolvente)
            c["valor_cm"] = valor_cm
            c["snap"] = _snap_o_medido(valor_cm)
            ranqueados = [c]
            print(f"↩️ {nombre_hoja}: se fuerza envolvente real de espesor.")
        elif thk_chapa_cm and candidatos:
            for c in candidatos:
                valor_cm = _esperado_modelo(vista, c["gap_sheet"])
                c = dict(c)
                c["valor_cm"] = valor_cm
                c["snap"] = _snap_o_medido(valor_cm)
                ranqueados.append(c)
        if not ranqueados:
            print(f"⚠️ {nombre_hoja}: candidatos prismáticos no usables.")
            return False, None

    # ¿El Thickness de chapa representa el cuerpo o solo un resalte?
    chapa_es_cuerpo = False
    if thk_chapa_cm is not None and es_canto_aplanado and overall_cm:
        chapa_es_cuerpo = thk_chapa_cm >= overall_cm * 0.50
    elif thk_chapa_cm is not None and not es_canto_aplanado:
        chapa_es_cuerpo = True

    def _clave(c):
        if es_canto_aplanado and overall_cm:
            cerca_env = abs(c["valor_cm"] - overall_cm) / max(overall_cm, EPS)
        else:
            cerca_env = 0.0
        match_chapa = 1
        if chapa_es_cuerpo and thk_chapa_cm is not None:
            tol = max(TOL_CM * 3, abs(thk_chapa_cm) * 0.20)
            if abs(c["valor_cm"] - thk_chapa_cm) <= tol:
                match_chapa = 0
        # En L/U (no aplanado): menor espesor de catálogo sigue siendo THK.
        prefer_menor = c["valor_cm"] if not es_canto_aplanado else -c["valor_cm"]
        return (
            cerca_env if es_canto_aplanado else 0.0,
            match_chapa,
            0 if (not es_canto_aplanado and c["snap"].get("desde_catalogo")) else 1,
            prefer_menor,
        )

    ranqueados.sort(key=_clave)
    mejor = ranqueados[0]

    if es_canto_aplanado and overall_cm:
        tol_env = max(TOL_CM * 4, overall_cm * 0.12)
        cercanos_env = [
            c for c in ranqueados
            if abs(c["valor_cm"] - overall_cm) <= tol_env
        ]
        if cercanos_env:
            mejor = min(
                cercanos_env, key=lambda x: abs(x["valor_cm"] - overall_cm)
            )
        elif chapa_es_cuerpo and thk_chapa_cm is not None:
            tol = max(TOL_CM * 3, abs(thk_chapa_cm) * 0.20)
            cercanos = [
                c for c in ranqueados
                if abs(c["valor_cm"] - thk_chapa_cm) <= tol
            ]
            if cercanos:
                mejor = min(
                    cercanos, key=lambda x: abs(x["valor_cm"] - thk_chapa_cm)
                )
    elif chapa_es_cuerpo and thk_chapa_cm is not None:
        tol = max(TOL_CM * 3, abs(thk_chapa_cm) * 0.20)
        cercanos = [
            c for c in ranqueados
            if abs(c["valor_cm"] - thk_chapa_cm) <= tol
        ]
        if cercanos:
            mejor = min(cercanos, key=lambda x: abs(x["valor_cm"] - thk_chapa_cm))
        elif _es_vista_cara_plana(datos):
            print(
                f"⚠️ {nombre_hoja}: cara plana sin gap cercano a "
                f"{thk_chapa_cm / IN_TO_CM:.4f} in — no se fuerza cota errónea."
            )
            return False, None

    # GUARDARRAÍL "ancho de placa disfrazado de THK":
    # Corrida 08:50 dejó cotas absurdas en P11=8.25 in, P32=6.5 in, P30=7.0 in,
    # 62176-1247-P04=15.62 in, SP-800=9.5 in, etc — todas piezas de chapa
    # 0.12 in donde `_snap_o_medido` capturó el ancho del bbox como "espesor".
    # Si el valor final es MEDIDO (sin catálogo), es mucho mayor que el
    # thk_chapa del modelo Y también mayor que la mitad del bbox menor,
    # es imposible que sea espesor. Preferir un candidato cercano al thk_chapa;
    # si no hay, abortar sin dibujar cota falsa.
    if (
        thk_chapa_cm is not None
        and thk_chapa_cm > EPS
        and not mejor["snap"].get("desde_catalogo", False)
    ):
        ratio_chapa = mejor["valor_cm"] / thk_chapa_cm
        menor_bbox = min(w_env, h_env)
        # Si mide más de 5x el thk_chapa Y más de 40% del bbox menor,
        # con casi total certeza es un ancho/largo, no un espesor.
        if ratio_chapa > 5.0 and (
            menor_bbox <= EPS or mejor["valor_cm"] > menor_bbox * 0.40
        ):
            print(
                f"⚠️ {nombre_hoja}: THK medido "
                f"{mejor['valor_cm'] / IN_TO_CM:.4f} in es "
                f"{ratio_chapa:.0f}x el thk_chapa "
                f"({thk_chapa_cm / IN_TO_CM:.4f} in): probable ancho de placa."
            )
            cercanos_chapa = [
                c for c in ranqueados
                if abs(c["valor_cm"] - thk_chapa_cm)
                <= max(TOL_CM * 3, abs(thk_chapa_cm) * 0.20)
            ]
            if cercanos_chapa:
                mejor = min(
                    cercanos_chapa,
                    key=lambda x: abs(x["valor_cm"] - thk_chapa_cm),
                )
                print(
                    f"  {nombre_hoja}: usando candidato cercano al thk_chapa "
                    f"= {mejor['valor_cm'] / IN_TO_CM:.4f} in."
                )
            else:
                print(
                    f"⚠️ {nombre_hoja}: sin candidato cercano al thk_chapa; "
                    f"no se dibuja cota falsa (pendiente para revisión manual)."
                )
                return False, None

    try:
        # Bridas / canto aplanado: silueta exterior con validación de span.
        # Evita el bug de filete (PIPE FLANE 0.375: intent circular → 0.75
        # flotante) frente a PIPE FLANGE 0.250 (tope plano, OK).
        overall_eje = h_env if mejor.get("tipo") == "vertical" else w_env
        forzar_silueta = (
            _es_altura_hasta_radio(mejor.get("gap_sheet"), overall_eje)
            or _nombre_parece_brida(nombre_hoja)
            or (
                es_canto_aplanado
                and overall_cm
                and abs(mejor["valor_cm"] - overall_cm)
                <= max(TOL_CM * 4, overall_cm * 0.12)
            )
        )
        if forzar_silueta:
            orient = "V" if mejor.get("tipo") == "vertical" else "H"
            if _es_altura_hasta_radio(mejor.get("gap_sheet"), overall_eje):
                print(
                    f"↩️ {nombre_hoja}: candidato hasta radio "
                    f"({mejor['valor_cm'] / IN_TO_CM:.3f} in); "
                    f"silueta exterior"
                )
            elif _nombre_parece_brida(nombre_hoja):
                print(
                    f"↩️ {nombre_hoja}: brida — THK por silueta exterior "
                    f"(evita filete→cuadrante)"
                )
            if _acotar_lado_bbox(
                hoja, vista, tg, datos, nombre_hoja, orient, "THK"
            ):
                valor_cm = _esperado_modelo(vista, overall_eje)
                snap = _snap_o_medido(valor_cm)
                print(
                    f"✅ {nombre_hoja}: THK prismático = {snap['valor_in']:.4f} in "
                    f"(silueta exterior)"
                )
                return True, {
                    "gap_sheet": overall_eje,
                    "valor_cm": valor_cm,
                    "valor_in": snap["valor_in"],
                }

        dim = _dibujar_cota_prismatica(hoja, tg, mejor)
        # Guardarraíl: si ModelValue no coincide con el gap elegido (filete
        # mal anclado), borrar y forzar silueta.
        if dim is not None and not _validar_span_cota(
            dim, mejor["gap_sheet"], vista, nombre_hoja, "THK"
        ):
            orient = "V" if mejor.get("tipo") == "vertical" else "H"
            if _acotar_lado_bbox(
                hoja, vista, tg, datos, nombre_hoja, orient, "THK"
            ):
                valor_cm = _esperado_modelo(vista, mejor["gap_sheet"])
                snap = _snap_o_medido(valor_cm)
                print(
                    f"✅ {nombre_hoja}: THK prismático = {snap['valor_in']:.4f} in "
                    f"(rescate silueta tras span inválido)"
                )
                return True, {
                    "gap_sheet": mejor["gap_sheet"],
                    "valor_cm": valor_cm,
                    "valor_in": snap["valor_in"],
                }
            return False, None

        origen = "envolvente" if (
            es_canto_aplanado
            and overall_cm
            and abs(mejor["valor_cm"] - overall_cm)
            <= max(TOL_CM * 4, overall_cm * 0.12)
        ) else ("catálogo" if mejor["snap"].get("desde_catalogo") else "medido")
        if chapa_es_cuerpo and thk_chapa_cm and abs(
            mejor["valor_cm"] - thk_chapa_cm
        ) <= max(TOL_CM * 3, abs(thk_chapa_cm) * 0.20):
            origen = "chapa+" + origen
        print(
            f"✅ {nombre_hoja}: THK prismático = {mejor['snap']['valor_in']:.4f} in "
            f"({origen}, detectado {mejor['valor_cm']/IN_TO_CM:.4f} in)"
        )
        return True, {
            "gap_sheet": mejor["gap_sheet"],
            "valor_cm": mejor["valor_cm"],
            "valor_in": mejor["snap"]["valor_in"],
        }
    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó la cota prismática -> {e}")
        return False, None


def _es_perfil_l_por_franja(datos, thk_sheet):
    """
    Escuadra L vista de frente a una pata (caso P47):
    gran rectángulo + franja delgada en un borde = espesor de la otra pata.
    """
    if not datos or thk_sheet is None or thk_sheet <= EPS:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    mayor = max(w, h)
    menor = min(w, h)
    if menor <= EPS:
        return False
    if mayor < thk_sheet * 6.0:
        return False
    # Tira solo-espesor muy larga → no es L de frente
    if mayor / menor >= 14.0:
        return False

    tol = max(0.025, thk_sheet * 0.40)
    borde_tol = max(tol * 3.0, thk_sheet * 1.5)

    for c in _buscar_candidatos_lineales(datos):
        if abs(c["gap_sheet"] - thk_sheet) > tol:
            continue

        if c["tipo"] == "vertical":
            ya = (c["a"]["miny"] + c["a"]["maxy"]) / 2.0
            yb = (c["b"]["miny"] + c["b"]["maxy"]) / 2.0
            y_lo, y_hi = min(ya, yb), max(ya, yb)
            cerca = (
                abs(y_lo - miny) <= borde_tol
                or abs(y_hi - maxy) <= borde_tol
            )
        else:
            xa = (c["a"]["minx"] + c["a"]["maxx"]) / 2.0
            xb = (c["b"]["minx"] + c["b"]["maxx"]) / 2.0
            x_lo, x_hi = min(xa, xb), max(xa, xb)
            cerca = (
                abs(x_lo - minx) <= borde_tol
                or abs(x_hi - maxx) <= borde_tol
            )

        if cerca and (mayor - thk_sheet) > thk_sheet * 3.0:
            return True
    return False


def _es_perfil_u_o_l(datos, thk_sheet):
    """
    Detecta perfiles doblados (canal U / escuadra L) vs placa de canto.

    - Tira muy alargada (aspecto >= 6) → placa, no ALTO (salvo franja L).
    - Perfil compacto con alto >> espesor y >=2 paredes del mismo THK → U/L.
    - Franja de espesor en el borde de una cara grande → L de frente (P47).
    """
    if not datos or thk_sheet is None or thk_sheet <= EPS:
        return False

    if _es_perfil_l_por_franja(datos, thk_sheet):
        return True

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    mayor = max(w, h)
    menor = min(w, h)
    if menor <= EPS:
        return False

    aspect = mayor / menor
    if aspect >= 6.0:
        return False
    if mayor < thk_sheet * 4.0:
        return False

    cands = _buscar_candidatos_lineales(datos)
    tol = max(0.02, thk_sheet * 0.35)
    similares = [
        c for c in cands
        if abs(c["gap_sheet"] - thk_sheet) <= tol
    ]
    if len(similares) >= 2:
        return True

    # Antes: ``len(datos)>=8 and aspect<4.5`` marcaba placas (P12_2 aspect
    # ~1.8) como perfil y generaba ALTO/PATA vacíos. Exigir al menos un
    # candidato THK + aspect de sección (no placa alargada ni cara plana).
    if len(similares) >= 1 and 1.15 <= aspect < 4.0 and mayor >= thk_sheet * 6.0:
        return True
    return False


def _span_bordes_seleccionados(a, b, orientacion):
    """Distancia en hoja entre los extremos de los dos bordes elegidos."""
    if orientacion == "V":
        return abs(float(a["maxy"]) - float(b["miny"]))
    return abs(float(a["maxx"]) - float(b["minx"]))


def _pares_borde_silueta(datos, orientacion, tol):
    """
    Candidatos (superior/inferior o der/izq) que SÍ tocan el extremo global.

    Evita anclar en tangencia: el borde inferior debe tener miny≈miny_global
    y el superior maxy≈maxy_global (idem en X).
    """
    minx, maxx, miny, maxy = _bbox_global(datos)
    if orientacion == "V":
        a = _elegir_curva_extrema_thk(datos, "sup", tol)
        b = _elegir_curva_extrema_thk(datos, "inf", tol)
        if a is None:
            a = max(datos, key=lambda d: d["maxy"])
        if b is None:
            b = min(datos, key=lambda d: d["miny"])
        # Exigir que los bordes toquen la silueta exterior.
        if a is not None and a["maxy"] < maxy - tol:
            tops = [d for d in datos if d["maxy"] >= maxy - tol]
            if tops:
                a = max(tops, key=lambda d: (d["dx"], d["dy"]))
        if b is not None and b["miny"] > miny + tol:
            bots = [d for d in datos if d["miny"] <= miny + tol]
            if bots:
                b = max(bots, key=lambda d: (d["dx"], d["dy"]))
        return a, b, "sup", "inf", (maxy - miny)
    a = _elegir_curva_extrema_thk(datos, "der", tol)
    b = _elegir_curva_extrema_thk(datos, "izq", tol)
    if a is None:
        a = max(datos, key=lambda d: d["maxx"])
    if b is None:
        b = min(datos, key=lambda d: d["minx"])
    if a is not None and a["maxx"] < maxx - tol:
        ders = [d for d in datos if d["maxx"] >= maxx - tol]
        if ders:
            a = max(ders, key=lambda d: (d["dy"], d["dx"]))
    if b is not None and b["minx"] > minx + tol:
        izqs = [d for d in datos if d["minx"] <= minx + tol]
        if izqs:
            b = max(izqs, key=lambda d: (d["dy"], d["dx"]))
    return a, b, "der", "izq", (maxx - minx)


def _acotar_lado_bbox(
    hoja, vista, tg, datos, nombre_hoja, orientacion, etiqueta, ratio_min=0.97
):
    """
    Dibuja UNA cota lineal (vertical u horizontal) sobre la silueta exterior.

    No usa ``min/max`` crudo de una sola curva (eso ancla en la tangencia
    del doblez). Elige bordes reales (recto horizontal arriba/abajo o
    vertical izq/der) y fija GeometryIntent en el punto extremo.
    Valida ModelValue ≈ span de silueta; si queda corto, reintenta.

    ``ratio_min``: cobertura mínima bordes/silueta. Bajar a ~0.80 en U
    con filetes (Parking) donde la tangencia recorta el span.
    """
    if not datos:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if max(w, h) <= EPS:
        return False

    inv_app = None
    try:
        inv_app = conectar_inventor()
    except Exception:
        inv_app = None

    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
    except Exception:
        sheet_w = None
        sheet_h = None

    span = h if orientacion == "V" else w
    tol = max(0.03, max(w, h) * 0.02)
    ratio_ok = float(ratio_min) if ratio_min else 0.97

    def _pt_para_orient(a, b, dx_off, dy_off):
        clr = clearance_texto_cota_cm()
        if orientacion == "V":
            x = min(a["minx"], b["minx"]) - clr - dx_off
            y = (miny + maxy) / 2.0 + dy_off
        else:
            x = (minx + maxx) / 2.0 + dx_off
            y = max(a["maxy"], b["maxy"]) + clr + dy_off
        return _clampear_punto_hoja(
            hoja, tg, x, y, evitar_bbox=(minx, maxx, miny, maxy)
        )

    def _cota_dentro_de_sheet(dim_obj):
        if sheet_w is None or sheet_h is None:
            return True
        try:
            rb = dim_obj.RangeBox
            dminx = float(rb.MinPoint.X)
            dmaxx = float(rb.MaxPoint.X)
            dminy = float(rb.MinPoint.Y)
            dmaxy = float(rb.MaxPoint.Y)
        except Exception:
            return True
        margen = 0.3
        return (
            dminx >= margen
            and dmaxx <= sheet_w - margen
            and dminy >= margen
            and dmaxy <= sheet_h - margen
        )

    def _cota_limpia(dim_obj):
        """Fuera de la pieza Y (preferible) dentro del sheet."""
        pieza_bb = (minx, maxx, miny, maxy)
        if dim_solapa_pieza(dim_obj, pieza_bb):
            asegurar_cota_fuera_pieza(
                dim_obj, tg, pieza_bb, orientacion
            )
        if dim_solapa_pieza(dim_obj, pieza_bb):
            return False
        return _cota_dentro_de_sheet(dim_obj)

    ultimo_error = None
    datos_iter = list(datos)
    for intento in range(3):
        try:
            a, b, lado_a, lado_b, span = _pares_borde_silueta(
                datos_iter, orientacion, tol
            )
            if a is None or b is None or a.get("curve") is b.get("curve"):
                raise RuntimeError("sin pares de borde para silueta exterior")

            span_bordes = _span_bordes_seleccionados(a, b, orientacion)
            if span > EPS and span_bordes / span < ratio_ok:
                raise RuntimeError(
                    f"bordes no cubren silueta "
                    f"({span_bordes:.3f}/{span:.3f})"
                )

            int_a = _intent_en_extremo(hoja, tg, a, lado_a)
            int_b = _intent_en_extremo(hoja, tg, b, lado_b)
            if int_a is None or int_b is None:
                raise RuntimeError("GeometryIntent de borde falló")

            pieza_bb = (minx, maxx, miny, maxy)
            clr = clearance_texto_cota_cm()
            # Primero candidatos geométricos fuera de silueta; luego offsets.
            offsets_a_probar = []
            for x, y, _lado in candidatos_texto_fuera_pieza(
                pieza_bb, orientacion, clearance=clr
            ):
                offsets_a_probar.append(("xy", x, y))
            for dx_off, dy_off in (
                (0.0, 0.0),
                (1.5, 0.0),
                (3.0, 0.0),
                (4.5, 0.0),
                (0.0, 1.5),
                (0.0, 3.0),
            ):
                offsets_a_probar.append(("d", dx_off, dy_off))

            dim_creado = None
            dim_fallback = None
            for modo, v1, v2 in offsets_a_probar:
                if modo == "xy":
                    pt = _clampear_punto_hoja(
                        hoja, tg, v1, v2, evitar_bbox=pieza_bb
                    )
                else:
                    pt = _pt_para_orient(a, b, v1, v2)
                try:
                    if orientacion == "V":
                        dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                            pt, int_a, int_b, kVerticalDimensionType
                        )
                    else:
                        dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                            pt, int_a, int_b, kHorizontalDimensionType
                        )
                except Exception:
                    continue

                if not _validar_span_cota(
                    dim_test, span, vista, nombre_hoja, etiqueta
                ):
                    try:
                        dim_test.Delete()
                    except Exception:
                        pass
                    continue

                if _cota_limpia(dim_test):
                    dim_creado = dim_test
                    break

                # Preferir sin solape aunque quede cerca del borde del sheet.
                if not dim_solapa_pieza(dim_test, pieza_bb):
                    if dim_fallback is None:
                        dim_fallback = dim_test
                    else:
                        try:
                            dim_test.Delete()
                        except Exception:
                            pass
                else:
                    try:
                        dim_test.Delete()
                    except Exception:
                        pass

            if dim_creado is None and dim_fallback is not None:
                if _validar_span_cota(
                    dim_fallback, span, vista, nombre_hoja, etiqueta
                ) and not dim_solapa_pieza(dim_fallback, pieza_bb):
                    dim_creado = dim_fallback
                    _dbg(
                        f"{nombre_hoja}: cota aceptada cerca del borde "
                        f"(sin solape con pieza)."
                    )
                else:
                    try:
                        dim_fallback.Delete()
                    except Exception:
                        pass
                    dim_fallback = None

            if dim_creado is None:
                raise RuntimeError(
                    "no se pudo crear cota fuera de la pieza"
                )

            aplicar_estilo_cota(dim_creado, hoja=hoja)
            asegurar_cota_fuera_pieza(
                dim_creado, tg, pieza_bb, orientacion
            )

            valor_in = _esperado_modelo(vista, span) / IN_TO_CM
            print(f"✅ {nombre_hoja}: {etiqueta} = {valor_in:.4f} in (silueta)")
            return True
        except Exception as e:
            ultimo_error = e
            try:
                hoja.Activate()
            except Exception:
                pass
            try:
                if inv_app is not None:
                    inv_app.ActiveView.Update()
            except Exception:
                pass
            time.sleep(0.5 * (intento + 1))
            try:
                fresh = _obtener_curvas_validas(vista)
                if fresh:
                    datos_iter = fresh
                    minx, maxx, miny, maxy = _bbox_global(datos_iter)
                    w = maxx - minx
                    h = maxy - miny
                    span = h if orientacion == "V" else w
                    tol = max(0.03, max(w, h) * 0.02)
            except Exception:
                pass

    print(f"⚠️ {nombre_hoja}: no se pudo acotar {etiqueta} -> {ultimo_error}")
    return False


def _acotar_alto_perfil(hoja, vista, tg, datos, nombre_hoja, thk_sheet=None):
    """
    Cota las dimensiones globales del perfil visto de canto.

    - En perfiles U/C compactos (canal, escuadra doblada) se dibujan DOS
      cotas: el LADO MAYOR del bbox (alto del cuerpo) y el LADO MENOR del
      bbox (largo de la pata / ala). Esto permite verificar bend deduction
      contra la pieza física.
    - En L-por-franja (una pata muy larga con franja del otro espesor en el
      borde) el "menor" del bbox ES el espesor → sólo se cota el mayor
      para no duplicar el THK.
    - En semicírculo (media caña) se dibujan también ambas cotas.

    Devuelve ``True`` si logró colocar al menos la cota del lado mayor.
    """
    if not datos:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if max(w, h) <= EPS:
        return False

    # Decidir orientación de la cota mayor.
    orient_mayor = 'V' if h >= w else 'H'
    orient_menor = 'H' if h >= w else 'V'
    menor = min(w, h)

    # Cota mayor (siempre): "ALTO perfil" (mantiene el label anterior en log).
    ok_mayor = _acotar_lado_bbox(
        hoja, vista, tg, datos, nombre_hoja, orient_mayor, "ALTO perfil"
    )
    if not ok_mayor:
        return False

    # Decidir si además cotamos el menor del bbox (largo de la pata).
    # Sólo lo hacemos si el menor es realmente distinto al espesor de chapa
    # (≥ 2× thk_sheet). Si thk_sheet es desconocido, aceptamos por defecto
    # cuando el menor sea >= 0.6 cm (evita cotar franjas puramente de espesor).
    cotar_menor = False
    if thk_sheet is not None and thk_sheet > EPS:
        if menor >= thk_sheet * 2.2:
            cotar_menor = True
    else:
        if menor >= 0.6:
            cotar_menor = True

    if cotar_menor:
        # Refrescar curvas antes de la segunda cota: la primera cambia
        # DrawingCurves en algunas versiones y sin refresco AddLinear
        # puede tirar HRESULT en la 2a llamada.
        try:
            fresh = _obtener_curvas_validas(vista)
            if fresh:
                datos_menor = fresh
            else:
                datos_menor = datos
        except Exception:
            datos_menor = datos

        _acotar_lado_bbox(
            hoja, vista, tg, datos_menor, nombre_hoja, orient_menor,
            "LARGO pata"
        )

    return True


def _borrar_cotas_hoja(hoja):
    try:
        dims = hoja.DrawingDimensions.GeneralDimensions
        for i in range(dims.Count, 0, -1):
            try:
                dims.Item(i).Delete()
            except Exception:
                pass
    except Exception:
        pass


def _nombre_hoja_alto(nombre_lado):
    return _nombre_hoja_variante(nombre_lado, "_ALTO")


def _nombre_hoja_largo_pata(nombre_lado):
    return _nombre_hoja_variante(nombre_lado, "_LARGO_PATA")


def _acotar_thk_asociativa_forzada(hoja, vista, tg, datos, nombre_hoja, thk_cm):
    """
    Cota asociativa (AddLinear) cuyo valor ≈ thk_cm.

    Orden:
    1) Par lineal cercano al Thickness de chapa/barra.
    2) Eje de silueta (H/V) cuyo span coincide con thk_cm.
    3) Eje menor de la silueta (canto típico).
    4) Envolvente menor (Parking U con filetes: ratio relajado).

    Prefiere GeneralDimension; si nada cuadra, el caller puede poner nota.
    """
    if not datos or thk_cm is None or thk_cm <= EPS:
        return False

    # Parking U: filetes recortan el span de bordes rectos → ratio más laxo.
    ratio_sil = 0.80 if _nombre_parece_parking_u(nombre_hoja) else 0.97

    if _acotar_espesor_cercano_chapa(
        hoja, vista, tg, datos, nombre_hoja, thk_cm
    ):
        return True

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if max(w, h) <= EPS:
        return False

    tol = max(TOL_CM * 4, abs(thk_cm) * 0.25)
    ejes = []
    if h > EPS:
        ejes.append(("V", h, abs(_esperado_modelo(vista, h) - thk_cm)))
    if w > EPS:
        ejes.append(("H", w, abs(_esperado_modelo(vista, w) - thk_cm)))
    ejes.sort(key=lambda t: t[2])

    for orient, span, delta in ejes:
        if delta > tol:
            continue
        if _acotar_lado_bbox(
            hoja,
            vista,
            tg,
            datos,
            nombre_hoja,
            orient,
            "THK",
            ratio_min=ratio_sil,
        ):
            print(
                f"↩️ {nombre_hoja}: THK asociativa por silueta {orient} "
                f"≈ {_esperado_modelo(vista, span) / IN_TO_CM:.4f} in"
            )
            return True

    # Último intento silueta: lado menor del bbox (canto).
    orient = "V" if h <= w else "H"
    if _acotar_lado_bbox(
        hoja,
        vista,
        tg,
        datos,
        nombre_hoja,
        orient,
        "THK",
        ratio_min=ratio_sil,
    ):
        span = h if orient == "V" else w
        print(
            f"↩️ {nombre_hoja}: THK asociativa eje menor {orient} "
            f"≈ {_esperado_modelo(vista, span) / IN_TO_CM:.4f} in"
        )
        return True

    # Envolvente menor (min/max de curvas): acepta si ModelValue ≈ thk.
    envolvente = _envolvente_espesor(datos)
    if envolvente is not None:
        gap_cm = _esperado_modelo(vista, envolvente["gap_sheet"])
        if abs(gap_cm - thk_cm) <= max(TOL_CM * 4, abs(thk_cm) * 0.30):
            try:
                dim = _dibujar_cota_prismatica(hoja, tg, envolvente)
                if dim is not None and _validar_span_cota(
                    dim,
                    envolvente["gap_sheet"],
                    vista,
                    nombre_hoja,
                    "THK envolvente",
                ):
                    print(
                        f"↩️ {nombre_hoja}: THK por envolvente "
                        f"≈ {gap_cm / IN_TO_CM:.4f} in"
                    )
                    return True
            except Exception:
                pass
    return False


def _finalizar_parking_patas_base(plano, hoja_lado, tg, vista, meta, nombre_lado):
    """
    Tras acotar patas→base en LADO (Parking):
    1) Renombra esa hoja a ``_ALTO`` (bend deduction).
    2) Crea ``_THK`` solo si el espesor está validado (Sheet Metal o pared 2D).
       Si no, omite THK (no inventa 0.875 desde bbox).
    """
    creadas = set()
    try:
        nombre_alto = _nombre_hoja_alto(nombre_lado)
        if str(hoja_lado.Name).upper() != nombre_alto.upper():
            hoja_lado.Name = nombre_alto
        creadas.add(str(hoja_lado.Name))
        print(f"  {nombre_lado}: hoja patas→base renombrada a {hoja_lado.Name}")
    except Exception as exc:
        print(f"⚠️ {nombre_lado}: no se pudo renombrar LADO→ALTO ({exc})")

    alto_cm = (meta or {}).get("valor_cm")
    datos_lado = None
    try:
        datos_lado = _obtener_curvas_validas(vista)
    except Exception:
        datos_lado = None

    thk_cm = (meta or {}).get("thk_chapa_cm")
    origen = "meta_chapa"
    if thk_cm is None or thk_cm <= EPS or (
        alto_cm and _thk_duplica_referencia(thk_cm, alto_cm)
    ):
        thk_cm, origen = _espesor_thk_validado(
            vista,
            alto_cm=alto_cm,
            datos=datos_lado,
            nombre_hoja=nombre_lado,
        )
    if thk_cm is None or thk_cm <= EPS:
        print(
            f"⚠️ {nombre_lado}: ALTO OK; THK omitido (sin espesor validado). "
            f"Fijar Sheet Metal Thickness o mejorar vista de canto."
        )
        # Marcar pendiente para no renombrar/capturar un THK falso.
        try:
            LAST_PENDIENTES_THK.append(str(nombre_lado))
        except Exception:
            pass
        return creadas

    nombre_thk = _nombre_hoja_variante(nombre_lado, "_THK")
    try:
        for i in range(1, plano.Sheets.Count + 1):
            if str(plano.Sheets.Item(i).Name).upper() == nombre_thk.upper():
                print(f"  {nombre_lado}: ya existe {nombre_thk}, no se duplica")
                creadas.add(nombre_thk)
                return creadas
        hoja_thk = hoja_lado.CopyTo(plano)
        hoja_thk.Name = nombre_thk
        try:
            dims = hoja_thk.DrawingDimensions.GeneralDimensions
            for i in range(dims.Count, 0, -1):
                dims.Item(i).Delete()
        except Exception:
            pass
        try:
            vista_thk = hoja_thk.DrawingViews.Item(1)
        except Exception:
            vista_thk = vista

        ok_geom = False
        datos_thk = None
        try:
            datos_thk = _obtener_curvas_validas(vista_thk)
        except Exception:
            datos_thk = None

        if datos_thk:
            ok_geom = _acotar_espesor_cercano_chapa(
                hoja_thk, vista_thk, tg, datos_thk, nombre_thk, thk_cm
            )
            if not ok_geom:
                ok_geom = _acotar_thk_asociativa_forzada(
                    hoja_thk, vista_thk, tg, datos_thk, nombre_thk, thk_cm
                )

        if ok_geom:
            creadas.add(nombre_thk)
            print(
                f"✅ {nombre_thk}: THK = "
                f"{thk_cm / IN_TO_CM:.4f} in (cota asociativa, {origen})"
            )
        else:
            ok_nota, _ = _forzar_cota_thk_desde_modelo(
                hoja_thk,
                tg,
                vista_thk,
                nombre_thk,
                alto_cm=alto_cm,
                datos=datos_thk or datos_lado,
            )
            if ok_nota:
                creadas.add(nombre_thk)
                print(
                    f"⚠️ {nombre_thk}: sin AddLinear; nota THK validada "
                    f"= {thk_cm / IN_TO_CM:.4f} in ({origen})"
                )
            else:
                try:
                    hoja_thk.Delete()
                except Exception:
                    pass
                print(
                    f"⚠️ {nombre_lado}: no se crea hoja THK "
                    f"(espesor no dibujable ni validado como nota)."
                )
                try:
                    LAST_PENDIENTES_THK.append(str(nombre_lado))
                except Exception:
                    pass
    except Exception as exc:
        print(f"⚠️ {nombre_lado}: no se pudo crear hoja THK Parking ({exc})")
        try:
            LAST_PENDIENTES_THK.append(str(nombre_lado))
        except Exception:
            pass
    return creadas


def _acotar_espesor_cercano_chapa(hoja, vista, tg, datos, nombre_hoja, thk_cm):
    """Dibuja cota lineal del candidato más cercano al Thickness de chapa."""
    if not datos or thk_cm is None or thk_cm <= EPS:
        return False
    candidatos = _buscar_candidatos_lineales(datos)
    if not candidatos:
        return False
    tol = max(TOL_CM * 3, abs(thk_cm) * 0.35)
    cercanos = []
    for c in candidatos:
        valor_cm = _esperado_modelo(vista, c["gap_sheet"])
        if abs(valor_cm - thk_cm) <= tol:
            c2 = dict(c)
            c2["valor_cm"] = valor_cm
            cercanos.append(c2)
    if not cercanos:
        return False
    mejor = min(cercanos, key=lambda x: abs(x["valor_cm"] - thk_cm))
    try:
        dim = _dibujar_cota_prismatica(hoja, tg, mejor)
        if dim is None:
            return False
        if not _validar_span_cota(
            dim, mejor["gap_sheet"], vista, nombre_hoja, "THK pared"
        ):
            return False
        return True
    except Exception:
        return False


def _finalizar_largo_desde_vista_thk(
    plano, hoja_lado, tg, vista, datos, thk_sheet, nombre_lado
):
    """
    Solera Jacking Pad: hoja extra con el LARGO real (= arista más larga 3D).

    La cámara del THK suele mostrar el canto: ahí el eje "mayor" de la
    silueta es el ancho de cara (~2.50 in), NO el largo de base (~3.536 in).
    Por eso se ancla al mayor del bbox 3D y se prueban cámaras hasta que
    la silueta proyecte ese valor.
    """
    creadas = set()
    if not _nombre_parece_solera_jacking(nombre_lado):
        return creadas
    if not datos:
        return creadas

    # Objetivo 3D: arista más larga (excluye espesor y el ancho medio).
    target_cm = None
    dims3d = _dimensiones_bbox_3d(vista)
    if dims3d:
        orden = sorted(d for d in dims3d if d is not None and d > EPS)
        if orden:
            target_cm = orden[-1]
            print(
                f"  {nombre_lado}: LARGO objetivo 3D = "
                f"{target_cm / IN_TO_CM:.4f} in "
                f"(bbox {[round(d / IN_TO_CM, 3) for d in orden]} in)"
            )

    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if max(w, h) <= EPS and target_cm is None:
        return creadas

    nombre_largo = _nombre_hoja_largo_pata(nombre_lado)
    try:
        for i in range(1, plano.Sheets.Count + 1):
            if str(plano.Sheets.Item(i).Name).upper() == nombre_largo.upper():
                creadas.add(nombre_largo)
                return creadas
    except Exception:
        pass

    inv_app = None
    try:
        inv_app = conectar_inventor()
    except Exception:
        inv_app = None

    camaras_alt = []
    if vista is not None and inv_app is not None:
        try:
            part_doc = vista.ReferencedDocumentDescriptor.ReferencedDocument
            to = inv_app.TransientObjects
            camaras_alt = list(_camaras_transversales(part_doc, tg, to) or [])
        except Exception:
            camaras_alt = []

    # Tolerancia: silueta mayor ≈ target 3D (acepta ~8% o 0.15 in).
    def _coincide_largo(datos_n, vista_n):
        if not datos_n:
            return False, None, None, None
        aminx, amaxx, aminy, amaxy = _bbox_global(datos_n)
        aw = amaxx - aminx
        ah = amaxy - aminy
        if max(aw, ah) <= EPS:
            return False, None, None, None
        orient = "H" if aw >= ah else "V"
        span = max(aw, ah)
        menor = min(aw, ah)
        mayor_cm = _esperado_modelo(vista_n, span)
        # Descartar franja de canto (mayor ≈ THK).
        if thk_sheet is not None and thk_sheet > EPS:
            if mayor_cm < thk_sheet * 2.2:
                return False, None, None, None
        if target_cm is not None and target_cm > EPS:
            tol = max(0.15 * IN_TO_CM, abs(target_cm) * 0.08)
            if abs(mayor_cm - target_cm) > tol:
                # También aceptar si el MENOR proyecta el largo (raro).
                menor_cm = _esperado_modelo(vista_n, menor)
                if abs(menor_cm - target_cm) <= tol:
                    orient = "V" if orient == "H" else "H"
                    span = menor
                    mayor_cm = menor_cm
                else:
                    return False, orient, span, mayor_cm
        elif span < menor * 1.5:
            return False, None, None, None
        return True, orient, span, mayor_cm

    # Vista actual primero; luego transversales (cara triangular / base).
    intentos = [None] + camaras_alt
    mejor_fallback = None  # (delta, cam, hoja, vista, datos, orient, span)

    for cam in intentos:
        hoja_n, vista_n = _clonar_hoja_lado_para_cota(
            plano, hoja_lado, tg, nombre_largo, inv_app, camara_alt=cam
        )
        if hoja_n is None:
            continue
        try:
            datos_n = _obtener_curvas_validas(vista_n)
        except Exception:
            datos_n = None
        if not datos_n:
            try:
                hoja_n.Delete()
            except Exception:
                pass
            continue

        ok_sil, orient, span, mayor_cm = _coincide_largo(datos_n, vista_n)

        # Guardar el candidato más cercano al target por si ninguno pasa tol.
        if target_cm is not None:
            aminx, amaxx, aminy, amaxy = _bbox_global(datos_n)
            aw = amaxx - aminx
            ah = amaxy - aminy
            deltas_eje = []
            if aw > EPS:
                deltas_eje.append(abs(_esperado_modelo(vista_n, aw) - target_cm))
            if ah > EPS:
                deltas_eje.append(abs(_esperado_modelo(vista_n, ah) - target_cm))
            delta = min(deltas_eje) if deltas_eje else None
            if delta is not None and (
                mejor_fallback is None or delta < mejor_fallback[0]
            ):
                if mejor_fallback is not None:
                    try:
                        mejor_fallback[2].Delete()
                    except Exception:
                        pass
                mejor_fallback = (
                    delta, cam, hoja_n, vista_n, datos_n, orient, span
                )
                if not ok_sil:
                    continue
            elif not ok_sil:
                try:
                    hoja_n.Delete()
                except Exception:
                    pass
                continue
        elif not ok_sil:
            try:
                hoja_n.Delete()
            except Exception:
                pass
            continue

        # Coincide: acotar y listo.
        try:
            dims = hoja_n.DrawingDimensions.GeneralDimensions
            for i in range(dims.Count, 0, -1):
                dims.Item(i).Delete()
        except Exception:
            pass

        if _acotar_lado_bbox(
            hoja_n, vista_n, tg, datos_n, nombre_largo, orient, "LARGO solera"
        ):
            # Si había un fallback distinto, borrarlo.
            if mejor_fallback is not None and mejor_fallback[2] is not hoja_n:
                try:
                    mejor_fallback[2].Delete()
                except Exception:
                    pass
            creadas.add(nombre_largo)
            print(
                f"📐 {nombre_lado}: creada {nombre_largo} "
                f"(largo≈{mayor_cm / IN_TO_CM:.4f} in, {orient})"
            )
            return creadas

        try:
            hoja_n.Delete()
        except Exception:
            pass
        if mejor_fallback is not None and mejor_fallback[2] is hoja_n:
            mejor_fallback = None

    # Fallback: mejor aproximación al largo 3D (aunque fuera de tol estricta).
    if mejor_fallback is not None:
        delta, _cam, hoja_n, vista_n, datos_n, orient, span = mejor_fallback
        # No re-acotar el ancho de cara (~2.50) si sigue lejos del largo 3D.
        if target_cm is not None and target_cm > EPS:
            tol_fb = max(0.25 * IN_TO_CM, abs(target_cm) * 0.15)
            if delta > tol_fb:
                print(
                    f"⚠️ {nombre_lado}: ninguna vista proyecta el largo 3D "
                    f"(mejor delta={delta / IN_TO_CM:.3f} in); "
                    f"no se crea LARGO"
                )
                try:
                    hoja_n.Delete()
                except Exception:
                    pass
                return creadas
        if orient is None:
            aminx, amaxx, aminy, amaxy = _bbox_global(datos_n)
            aw = amaxx - aminx
            ah = amaxy - aminy
            orient = "H" if aw >= ah else "V"
        try:
            dims = hoja_n.DrawingDimensions.GeneralDimensions
            for i in range(dims.Count, 0, -1):
                dims.Item(i).Delete()
        except Exception:
            pass
        # Preferir eje cuyo span modelo esté más cerca del target.
        if target_cm is not None and datos_n:
            aminx, amaxx, aminy, amaxy = _bbox_global(datos_n)
            aw = amaxx - aminx
            ah = amaxy - aminy
            cand = []
            if aw > EPS:
                cand.append(("H", aw, abs(_esperado_modelo(vista_n, aw) - target_cm)))
            if ah > EPS:
                cand.append(("V", ah, abs(_esperado_modelo(vista_n, ah) - target_cm)))
            if cand:
                cand.sort(key=lambda t: t[2])
                orient = cand[0][0]
        if _acotar_lado_bbox(
            hoja_n, vista_n, tg, datos_n, nombre_largo, orient, "LARGO solera"
        ):
            creadas.add(nombre_largo)
            print(
                f"📐 {nombre_lado}: creada {nombre_largo} "
                f"(mejor cámara vs largo 3D, {orient})"
            )
            return creadas
        try:
            hoja_n.Delete()
        except Exception:
            pass

    print(f"⚠️ {nombre_lado}: no se pudo crear LARGO solera (~3D max)")
    return creadas


def _finalizar_nipple_alto(plano, hoja_lado, tg, vista, datos, meta, nombre_lado):
    """
    Cilindros / nipples / STUD / tubos: FRENTE suele ser Ø; falta longitud axial
    como ``_ALTO``.

    Si LADO ya muestra el perfil (rectángulo alargado), acota el eje mayor.
    Si LADO es cara circular, prueba cámaras transversales.
    """
    creadas = set()
    if not _necesita_alto_axial(nombre_lado):
        return creadas

    nombre_alto = _nombre_hoja_alto(nombre_lado)
    try:
        for i in range(1, plano.Sheets.Count + 1):
            if str(plano.Sheets.Item(i).Name).upper() == nombre_alto.upper():
                creadas.add(nombre_alto)
                return creadas
    except Exception:
        pass

    inv_app = None
    try:
        inv_app = conectar_inventor()
    except Exception:
        inv_app = None

    thk_sheet = (meta or {}).get("gap_sheet")
    camaras_alt = []
    if vista is not None and inv_app is not None:
        try:
            part_doc = vista.ReferencedDocumentDescriptor.ReferencedDocument
            to = inv_app.TransientObjects
            camaras_alt = _camaras_transversales(part_doc, tg, to)
        except Exception:
            camaras_alt = []

    def _silueta_alargada(datos_n):
        if not datos_n:
            return False, None, None
        minx, maxx, miny, maxy = _bbox_global(datos_n)
        w = maxx - minx
        h = maxy - miny
        if max(w, h) <= EPS:
            return False, None, None
        aspect = max(w, h) / max(min(w, h), EPS)
        # Bridas compactas: el cuerpo puede ser solo un poco más alto/ancho
        # que el Ø (antes 1.25 dejaba fuera algunas PIPE FLANGE).
        umbral = 1.12 if _nombre_parece_brida(nombre_lado) else 1.25
        if aspect < umbral:
            return False, None, None
        orient = "V" if h >= w else "H"
        return True, orient, max(w, h)

    # Primero la vista LADO actual; luego cámaras alt.
    intentos = [None] + list(camaras_alt)
    for cam in intentos:
        hoja_n, vista_n = _clonar_hoja_lado_para_cota(
            plano, hoja_lado, tg, nombre_alto, inv_app, camara_alt=cam
        )
        if hoja_n is None:
            continue
        try:
            datos_n = _obtener_curvas_validas(vista_n)
        except Exception:
            datos_n = None
        ok_sil, orient, _span = _silueta_alargada(datos_n)
        if not ok_sil:
            try:
                hoja_n.Delete()
            except Exception:
                pass
            continue
        # Quitar cotas copiadas (Ø/THK) antes de poner ALTO.
        try:
            dims = hoja_n.DrawingDimensions.GeneralDimensions
            for i in range(dims.Count, 0, -1):
                dims.Item(i).Delete()
        except Exception:
            pass
        if _acotar_lado_bbox(
            hoja_n, vista_n, tg, datos_n, nombre_alto, orient, "ALTO axial"
        ):
            creadas.add(nombre_alto)
            print(f"📐 {nombre_lado}: creada hoja {nombre_alto} (largo axial)")
            return creadas
        try:
            hoja_n.Delete()
        except Exception:
            pass

    print(f"⚠️ {nombre_lado}: no se pudo crear ALTO de longitud axial")
    return creadas


def _nombre_hoja_variante(nombre_lado, sufijo_final):
    """Reemplaza ``_LADO``/``_THK`` por ``sufijo_final`` conservando el resto."""
    base = str(nombre_lado)
    if ":" in base:
        base = base.rsplit(":", 1)[0]
    up = base.upper()
    if "_LADO" in up:
        idx = up.rfind("_LADO")
        return base[:idx] + sufijo_final
    if "_THK" in up:
        idx = up.rfind("_THK")
        return base[:idx] + sufijo_final
    return base + sufijo_final


def _es_perfil_por_modelo_3d(vista, thk_sheet=None):
    """
    Detecta perfil L/U/C/canal desde el bbox 3D del modelo, sin depender
    de que la vista 2D los muestre correctamente.

    Criterios:
    - Aspect ratio mayor/medio ≥ 3 (pieza tipo viga/canal larga).
    - Menor > 1.5× espesor de chapa (para descartar placas planas).
    - Menor > 0.5 cm en absoluto (para descartar barras muy finas).
    """
    dims = _dimensiones_bbox_3d(vista)
    if not dims or len(dims) < 3:
        return False
    dims_sorted = sorted(dims)
    menor, medio, mayor = dims_sorted[0], dims_sorted[1], dims_sorted[2]
    if medio <= EPS or mayor <= EPS:
        return False
    if mayor / medio < 3.0:
        return False
    if thk_sheet is not None and thk_sheet > EPS:
        if menor <= thk_sheet * 1.5:
            return False
    if menor < 0.5:
        return False
    return True


def _debe_generar_alto(datos, thk_sheet, vista=None):
    """U/L doblados o media caña semicircular, o perfil detectado por modelo 3D."""
    if os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return False
    if _es_perfil_u_o_l(datos, thk_sheet):
        return True
    # Fallback: detectar por modelo 3D. Cuando la vista LADO cae en cara
    # plana (P05_Default_As Machined, P47), _es_perfil_u_o_l retorna False
    # aunque el modelo SÍ sea un perfil L/U/C. Esta rama nos permite
    # generar _ALTO / _LARGO_PATA con una cámara alternativa.
    if vista is not None and _es_perfil_por_modelo_3d(vista, thk_sheet):
        return True
    if not thk_sheet or not _es_perfil_semicircular(datos):
        return False
    minx, maxx, miny, maxy = _bbox_global(datos)
    mayor = max(maxx - minx, maxy - miny)
    return mayor >= thk_sheet * 4.0


def _parece_silueta_franja_canto(datos):
    """True si la vista es una tira larga (cara de espesor), no escuadra L."""
    if not datos:
        return True
    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    menor = min(w, h)
    mayor = max(w, h)
    if menor <= EPS:
        return True
    return (mayor / menor) >= 5.0


def _reorientar_lado_a_escuadra(plano, hoja_lado, tg, inv_app, nombre_lado):
    """
    Jacking pads: si LADO quedó de cara (rectángulo), recrea la vista con
    cámara de escuadra L (eje de doblez / transversales).

    Sustituye la hoja LADO in-place (misma nombre) y devuelve
    ``(hoja, vista, datos)`` o ``(None, None, None)`` si no mejora.

    Bloqueado en DESPLIEGUE / SOLO_FLAT_CORTE: esa cámara es del sólido
    doblado y produce el perfil L con radio de doblez.
    """
    nombre_u = str(nombre_lado or "").upper()
    if "_DESPLIEGUE_" in nombre_u:
        return None, None, None
    if os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return None, None, None
    if inv_app is None:
        return None, None, None
    try:
        vista_orig = hoja_lado.DrawingViews.Item(1)
        part_doc = vista_orig.ReferencedDocumentDescriptor.ReferencedDocument
        to = inv_app.TransientObjects
    except Exception:
        return None, None, None

    camaras = []
    if _creador_vistas is not None:
        try:
            ori = _creador_vistas._orientacion_lado_doblado(
                part_doc, tg, to, None
            )
            if ori is not None:
                cam = _creador_vistas.crear_camara(
                    part_doc, tg, to,
                    ori["cx"], ori["cy"], ori["cz"],
                    ori["v_lado"], ori["v_up"],
                )
                if cam is not None:
                    camaras.append(cam)
        except Exception:
            pass
    try:
        camaras.extend(_camaras_transversales(part_doc, tg, to) or [])
    except Exception:
        pass

    if not camaras:
        return None, None, None

    nombre_tmp = str(hoja_lado.Name) + "_ESC_TMP"
    for cam in camaras:
        hoja_n, vista_n = _clonar_hoja_lado_para_cota(
            plano, hoja_lado, tg, nombre_tmp, inv_app, camara_alt=cam
        )
        if hoja_n is None:
            continue
        try:
            datos_n = _obtener_curvas_validas(vista_n)
        except Exception:
            datos_n = None
        if not datos_n or _parece_silueta_franja_canto(datos_n):
            try:
                hoja_n.Delete()
            except Exception:
                pass
            continue
        # Silueta más compacta = perfil L visible.
        try:
            nombre_final = str(hoja_lado.Name)
            hoja_lado.Name = nombre_final + "_OLD_FACE"
            hoja_n.Name = nombre_final
            try:
                hoja_lado.Delete()
            except Exception:
                pass
            print(
                f"↩️ {nombre_lado}: LADO reorientado a escuadra L "
                f"({len(datos_n)} curvas)"
            )
            return hoja_n, vista_n, datos_n
        except Exception as exc:
            print(f"⚠️ {nombre_lado}: no se pudo sustituir LADO ({exc})")
            try:
                hoja_n.Delete()
            except Exception:
                pass
            return None, None, None

    return None, None, None


def _calcular_camara_transversal(part_doc, tg, to):
    """
    Primera cámara transversal (compat). Preferir ``_camaras_transversales``.
    """
    cams = _camaras_transversales(part_doc, tg, to)
    return cams[0] if cams else None


def _camaras_transversales(part_doc, tg, to):
    """
    Lista de cámaras candidatas para sección transversal de perfil L/U/C.

    Orden: orientación doblada (si existe) → mirar eje mayor → eje medio.
    Así P16 / perfiles con LADO plano pueden reintentar otra dirección.
    """
    if part_doc is None or tg is None or to is None:
        return []

    cams = []
    if _creador_vistas is not None:
        try:
            ori = _creador_vistas._orientacion_lado_doblado(
                part_doc, tg, to, None
            )
            if ori is not None:
                cam = _creador_vistas.crear_camara(
                    part_doc, tg, to,
                    ori["cx"], ori["cy"], ori["cz"],
                    ori["v_lado"], ori["v_up"],
                )
                if cam is not None:
                    cams.append(cam)
        except Exception:
            pass

    if _creador_vistas is None:
        return cams

    try:
        rb = part_doc.ComponentDefinition.RangeBox
        cx = (float(rb.MaxPoint.X) + float(rb.MinPoint.X)) / 2.0
        cy = (float(rb.MaxPoint.Y) + float(rb.MinPoint.Y)) / 2.0
        cz = (float(rb.MaxPoint.Z) + float(rb.MinPoint.Z)) / 2.0
        dx = abs(float(rb.MaxPoint.X) - float(rb.MinPoint.X))
        dy = abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y))
        dz = abs(float(rb.MaxPoint.Z) - float(rb.MinPoint.Z))
        ejes = sorted(
            [("X", dx), ("Y", dy), ("Z", dz)],
            key=lambda t: t[1],
            reverse=True,
        )
        vec = {"X": (1, 0, 0), "Y": (0, 1, 0), "Z": (0, 0, 1)}
        # Mirar a lo largo del mayor (sección) y, si falla, del medio.
        for i_eye, i_up in ((0, 1), (1, 0), (0, 2)):
            if i_eye >= len(ejes) or i_up >= len(ejes):
                continue
            eye_dir = tg.CreateVector(*vec[ejes[i_eye][0]])
            up_hint = tg.CreateVector(*vec[ejes[i_up][0]])
            try:
                cam = _creador_vistas.crear_camara(
                    part_doc, tg, to, cx, cy, cz, eye_dir, up_hint
                )
                if cam is not None:
                    cams.append(cam)
            except Exception:
                continue
    except Exception:
        pass
    return cams


def _vista_parece_seccion_perfil(datos, thk_sheet, vista=None):
    """
    True si el bbox 2D (y opcionalmente 3D) luce como sección U/L, no placa.
    """
    if not datos:
        return False
    if _es_perfil_u_o_l(datos, thk_sheet):
        return True
    minx, maxx, miny, maxy = _bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    mayor = max(w, h)
    menor = min(w, h)
    if menor <= EPS or thk_sheet is None or thk_sheet <= EPS:
        return False
    # Sección útil: alto del canal >> espesor; no una tira de canto fina.
    if mayor < thk_sheet * 4.0:
        return False
    if menor < thk_sheet * 1.8:
        return False
    if vista is not None:
        dims = _dimensiones_bbox_3d(vista)
        if dims and len(dims) >= 3:
            d_ord = sorted(dims)
            # El menor 2D no debe ser el largo de la viga (mayor 3D).
            if mayor > d_ord[2] * 0.85:
                return False
            # Debe parecerse al medio o menor 3D (sección).
            if menor < d_ord[0] * 0.5:
                return False
    return True


def _es_com_transitorio_thk(exc):
    """
    Reconoce COM errores transitorios que suelen resolverse tras un pump
    de mensajes + delay. Se usa en los reintentos de CopyTo/AddLinear.
    """
    try:
        args = getattr(exc, "args", ())
        for a in args:
            if isinstance(a, int) and a in (-2147352567, -2147418111, -2147417846):
                return True
        s = str(exc).lower()
        if "-2147352567" in s or "-2147418111" in s or "-2147417846" in s:
            return True
        if "call was rejected" in s or "rpc" in s:
            return True
    except Exception:
        pass
    return False


def _clonar_hoja_lado_para_cota(
    plano, hoja_lado, tg, nombre_nueva, inv_app, camara_alt=None
):
    """
    Clona la hoja LADO usando ``hoja.CopyTo(plano)`` (mismo mecanismo que
    ``creador_vistas._crear_hoja_vista``, que sí funciona), borra la vista
    original de la copia, y crea una BaseView fresca del mismo modelo.

    ``camara_alt`` (opcional): cámara a usar en lugar de la del LADO
    original. Se usa cuando el LADO original cae en una cara plana y el
    modelo 3D indica que es un perfil (P05, P47) — así podemos forzar
    una vista transversal correcta.

    Devuelve ``(nueva_sheet, vista_nueva)`` o ``(None, None)`` si falla.

    Este approach reemplaza el uso previo de ``plano.Sheets.Add()``, que
    tiraba COM error ``-2147352567`` en el 100% de las corridas por la
    forma en que pywin32 maneja los parámetros opcionales de esa firma.
    """
    # Extraer modelo y cámara ANTES de clonar (así aunque la vista original
    # se dañe al clonar, tenemos los datos).
    try:
        vista_orig = hoja_lado.DrawingViews.Item(1)
    except Exception as e:
        print(f"⚠️ {nombre_nueva}: no se pudo leer vista LADO original -> {e}")
        return None, None

    try:
        part_doc = vista_orig.ReferencedDocumentDescriptor.ReferencedDocument
    except Exception as e:
        print(f"⚠️ {nombre_nueva}: no se pudo leer modelo referenciado -> {e}")
        return None, None

    if camara_alt is not None:
        cam_orig = camara_alt
    else:
        try:
            cam_orig = vista_orig.Camera
        except Exception:
            cam_orig = None

    # Borrar hoja previa con el mismo nombre (o cualquier sufijo :N) para
    # que ``.Name = nombre_nueva`` no falle por duplicado.
    try:
        for i in range(plano.Sheets.Count, 0, -1):
            try:
                h = plano.Sheets.Item(i)
                if str(h.Name).upper().startswith(nombre_nueva.upper()):
                    h.Delete()
            except Exception:
                continue
    except Exception:
        pass

    # CopyTo con reintentos ante COM transitorios.
    nueva = None
    ultimo_err = None
    for intento in range(3):
        try:
            nueva = hoja_lado.CopyTo(plano)
            break
        except Exception as e:
            ultimo_err = e
            nueva = None
            if not _es_com_transitorio_thk(e) and intento == 2:
                break
            if inv_app is not None:
                try:
                    inv_app.UserInterfaceManager.DoEvents()
                except Exception:
                    pass
            time.sleep(0.6 * (intento + 1))

    if nueva is None:
        print(f"⚠️ {nombre_nueva}: CopyTo falló -> {ultimo_err}")
        return None, None

    # Renombrar. Si Inventor pone auto-sufijo :N no importa (se maneja al
    # renombrar final del flujo).
    try:
        nueva.Name = nombre_nueva
    except Exception:
        pass

    try:
        nueva.Activate()
    except Exception:
        pass

    # Borrar TODAS las vistas heredadas del CopyTo (los proxies COM de esas
    # vistas son los que causaban los errores en la versión anterior).
    try:
        for v in range(nueva.DrawingViews.Count, 0, -1):
            try:
                nueva.DrawingViews.Item(v).Delete()
            except Exception:
                continue
    except Exception:
        pass

    # CRÍTICO: limpiar border, title block, sketches, notas, símbolos y
    # tablas heredados del machote. Sin este pase, el JPG exportado sale
    # dominado por líneas de plantilla y la vista real queda diminuta
    # en una esquina (bug visto en P14_LARGO_PATA_61.jpg).
    if _creador_vistas is not None:
        try:
            _creador_vistas._limpiar_border_y_titleblock(nueva)
        except Exception as exc_clean:
            print(f"AVISO {nombre_nueva}: limpieza border/titleblock falló ({exc_clean}); "
                  "el JPG puede salir con residuos del machote.")

    if inv_app is not None:
        try:
            inv_app.ActiveDocument.Update()
        except Exception:
            pass
        try:
            inv_app.UserInterfaceManager.DoEvents()
        except Exception:
            pass

    # Añadir BaseView fresca con la misma cámara.
    px = float(nueva.Width) / 2.0
    py = float(nueva.Height) / 2.0
    pt_centro = tg.CreatePoint2d(px, py)

    vista_nueva = None
    if cam_orig is not None:
        try:
            vista_nueva = nueva.DrawingViews.AddBaseView(
                part_doc,
                pt_centro,
                1.0,
                kArbitraryViewOrientation,
                kHiddenLineRemovedDrawingViewStyle,
                "",
                cam_orig,
            )
        except Exception:
            vista_nueva = None

    if vista_nueva is None:
        try:
            vista_nueva = nueva.DrawingViews.AddBaseView(
                part_doc,
                pt_centro,
                1.0,
                kDefaultViewOrientation,
                kHiddenLineRemovedDrawingViewStyle,
            )
        except Exception as e:
            print(f"⚠️ {nombre_nueva}: no se pudo crear BaseView -> {e}")
            try:
                nueva.Delete()
            except Exception:
                pass
            return None, None

    # Reescalar la vista de forma DEFENSIVA *después* de materializar curvas.
    # Medir RangeBox antes de que existan DrawingCurves dejaba escalas
    # minúsculas → JPG HEIGHT/LEG borrosos / puntito en blanco (P14_2, P15_1).
    if inv_app is not None:
        try:
            inv_app.ActiveDocument.Update()
        except Exception:
            pass
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        try:
            inv_app.UserInterfaceManager.DoEvents()
        except Exception:
            pass
    time.sleep(0.45)

    try:
        sheet_w = float(nueva.Width)
        sheet_h = float(nueva.Height)
    except Exception:
        sheet_w = 0.0
        sheet_h = 0.0

    w_v = 0.0
    h_v = 0.0
    # Preferir bbox de curvas HLR (geometría real) sobre RangeBox de vista
    # (puede incluir residuos o quedar vacío al inicio).
    try:
        datos_esc = _obtener_curvas_validas(vista_nueva)
    except Exception:
        datos_esc = None
    if datos_esc:
        try:
            minx, maxx, miny, maxy = _bbox_global(datos_esc)
            w_v = max(0.0, maxx - minx)
            h_v = max(0.0, maxy - miny)
        except Exception:
            w_v = 0.0
            h_v = 0.0
    if w_v <= EPS or h_v <= EPS:
        try:
            rb_view = vista_nueva.RangeBox
            w_v = float(rb_view.MaxPoint.X) - float(rb_view.MinPoint.X)
            h_v = float(rb_view.MaxPoint.Y) - float(rb_view.MinPoint.Y)
        except Exception:
            w_v = 0.0
            h_v = 0.0

    if sheet_w > 0 and sheet_h > 0 and w_v > EPS and h_v > EPS:
        try:
            scale_actual = float(vista_nueva.Scale)
        except Exception:
            scale_actual = 1.0
        if scale_actual <= EPS:
            scale_actual = 1.0
        # Dimensiones reales de la pieza (independientes de escala actual).
        real_w = w_v / scale_actual
        real_h = h_v / scale_actual

        # Objetivo: ~50 % del sheet (margen para cotas).
        max_w_pieza = sheet_w * 0.50
        max_h_pieza = sheet_h * 0.50
        escala_ideal = min(
            max_w_pieza / max(real_w, EPS),
            max_h_pieza / max(real_h, EPS),
        )

        escalas_discretas = [
            5.0, 4.0, 3.0, 2.0, 1.5, 1.0, 0.75, 0.5,
            0.4, 0.3, 0.25, 0.2, 0.15, 0.1, 0.08, 0.05,
            0.04, 0.03, 0.02, 0.015, 0.01,
        ]
        escala_elegida = 0.01
        for e in escalas_discretas:
            if e <= escala_ideal:
                escala_elegida = e
                break

        # Si aún quedaría demasiado pequeña (<12 % del sheet), subir un paso.
        cob_w = (real_w * escala_elegida) / sheet_w
        cob_h = (real_h * escala_elegida) / sheet_h
        if max(cob_w, cob_h) < 0.12:
            for e in reversed(escalas_discretas):
                if e > escala_elegida:
                    prueba_w = (real_w * e) / sheet_w
                    prueba_h = (real_h * e) / sheet_h
                    if max(prueba_w, prueba_h) <= 0.72:
                        escala_elegida = e
                        if max(prueba_w, prueba_h) >= 0.18:
                            break

        try:
            vista_nueva.Scale = escala_elegida
        except Exception:
            pass

        try:
            vista_nueva.Position = tg.CreatePoint2d(sheet_w / 2.0, sheet_h / 2.0)
        except Exception:
            pass

        if inv_app is not None:
            try:
                inv_app.ActiveDocument.Update()
            except Exception:
                pass

    if inv_app is not None:
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        try:
            inv_app.UserInterfaceManager.DoEvents()
        except Exception:
            pass
    time.sleep(0.35)

    return nueva, vista_nueva


def _crear_hoja_alto(plano, hoja_lado, tg, datos, thk_sheet, nombre_lado):
    """
    Crea hasta DOS hojas nuevas para dimensionar el perfil U/L/semicírculo:

    - ``_ALTO``: cota del lado MAYOR del bbox 2D (alto del cuerpo del canal).
    - ``_LARGO_PATA``: cota del lado MENOR del bbox 2D (largo del ala). Sólo
      se genera si el menor del bbox es ≥ 2.2× el espesor (para no duplicar
      el THK cuando la silueta es una franja L de canto).

    Cada hoja lleva UNA sola cota, para que la exportación JPG dé una foto
    limpia por dimensión (requerido para verificar bend deduction contra
    la pieza física).

    Cuando la vista _LADO original NO muestra la sección transversal (cae
    en cara plana: P05, P47), se detecta por modelo 3D y se recalcula una
    cámara alternativa que mira A LO LARGO del eje mayor del bbox 3D, para
    forzar una vista transversal correcta en las hojas extra.

    Devuelve un ``set`` con los nombres de hojas creadas (vacío si no se
    generó ninguna).
    """
    creadas = set()
    nombre_u = str(nombre_lado or "").upper()
    if "_DESPLIEGUE_" in nombre_u:
        return creadas
    if os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    ):
        return creadas

    inv_app = None
    try:
        inv_app = conectar_inventor()
    except Exception:
        inv_app = None

    # Necesitamos la vista original para pasar a _debe_generar_alto (para
    # que pueda inspeccionar el bbox 3D del modelo si la vista 2D no basta).
    try:
        vista_orig = hoja_lado.DrawingViews.Item(1)
    except Exception:
        vista_orig = None

    if not _debe_generar_alto(datos, thk_sheet, vista=vista_orig):
        return creadas

    # ¿La vista LADO 2D muestra el perfil correctamente?
    perfil_ok_en_lado = _es_perfil_u_o_l(datos, thk_sheet)

    # Cámaras transversales candidatas (reintentos si LADO es cara plana).
    camaras_alt = []
    if not perfil_ok_en_lado and vista_orig is not None and inv_app is not None:
        try:
            part_doc = vista_orig.ReferencedDocumentDescriptor.ReferencedDocument
            to = inv_app.TransientObjects
            camaras_alt = _camaras_transversales(part_doc, tg, to)
            if camaras_alt:
                _dbg(
                    f"{nombre_lado}: LADO no muestra perfil; "
                    f"{len(camaras_alt)} cámara(s) transversal(es) a probar."
                )
        except Exception as exc_cam:
            _dbg(f"{nombre_lado}: no se pudo calcular cámara alterna ({exc_cam})")
            camaras_alt = []

    def _clonar_con_reintentos(nombre_nueva):
        """Prueba LADO (si ya es sección) y luego cámaras transversales."""
        if perfil_ok_en_lado:
            intentos = [None] + list(camaras_alt)
        else:
            intentos = list(camaras_alt) if camaras_alt else [None]
        for idx, cam in enumerate(intentos):
            hoja_n, vista_n = _clonar_hoja_lado_para_cota(
                plano, hoja_lado, tg, nombre_nueva, inv_app, camara_alt=cam
            )
            if hoja_n is None:
                continue
            datos_n = None
            for intento in range(3):
                try:
                    datos_n = _obtener_curvas_validas(vista_n)
                except Exception:
                    datos_n = None
                if datos_n:
                    break
                if inv_app is not None:
                    try:
                        inv_app.UserInterfaceManager.DoEvents()
                    except Exception:
                        pass
                time.sleep(0.35 * (intento + 1))
            if not datos_n:
                try:
                    hoja_n.Delete()
                except Exception:
                    pass
                continue
            if not _vista_parece_seccion_perfil(datos_n, thk_sheet, vista_n):
                _dbg(
                    f"{nombre_nueva}: intento {idx} descartado "
                    "(vista no parece sección de perfil)."
                )
                try:
                    hoja_n.Delete()
                except Exception:
                    pass
                continue
            return hoja_n, vista_n, datos_n
        return None, None, None

    # ============================================================
    # Hoja 1: _ALTO  (solo cota asociativa; sin nota-only → se borra)
    # ============================================================
    nombre_alto = _nombre_hoja_alto(nombre_lado)
    hoja_alto, vista_alto, datos_alto = _clonar_con_reintentos(nombre_alto)
    if hoja_alto is None:
        print(f"⚠️ {nombre_lado}: no se pudo crear hoja usable {nombre_alto}")
    else:
        aminx, amaxx, aminy, amaxy = _bbox_global(datos_alto)
        aw = amaxx - aminx
        ah = amaxy - aminy
        orient_mayor_a = 'V' if ah >= aw else 'H'
        ok_alto = _acotar_lado_bbox(
            hoja_alto, vista_alto, tg, datos_alto, nombre_alto,
            orient_mayor_a, "ALTO perfil"
        )
        if ok_alto:
            creadas.add(nombre_alto)
            print(f"📐 {nombre_lado}: creada hoja extra {nombre_alto}")
        else:
            print(
                f"🗑️ {nombre_alto}: sin cota asociativa; se elimina "
                "(no se exporta JPG solo-nota)."
            )
            try:
                hoja_alto.Delete()
            except Exception:
                pass

    # ============================================================
    # Hoja 2: _LARGO_PATA (sólo perfiles con ambas dimensiones útiles)
    # ============================================================
    nombre_pata = _nombre_hoja_largo_pata(nombre_lado)
    hoja_pata, vista_pata, datos_pata = _clonar_con_reintentos(nombre_pata)
    if hoja_pata is None:
        print(f"⚠️ {nombre_lado}: no se pudo crear hoja usable {nombre_pata}")
    else:
        pminx, pmaxx, pminy, pmaxy = _bbox_global(datos_pata)
        pw = pmaxx - pminx
        ph = pmaxy - pminy
        menor_p = min(pw, ph)

        cotar_menor = False
        if thk_sheet is not None and thk_sheet > EPS:
            if menor_p >= thk_sheet * 2.2:
                cotar_menor = True
        else:
            if menor_p >= 0.6:
                cotar_menor = True

        if not cotar_menor:
            _dbg(
                f"{nombre_pata}: se omite (menor_bbox={menor_p:.3f}cm "
                f"≈ thk_sheet, no aporta cota nueva)."
            )
            try:
                hoja_pata.Delete()
            except Exception:
                pass
        else:
            orient_menor_p = 'H' if ph >= pw else 'V'
            ok_pata = _acotar_lado_bbox(
                hoja_pata, vista_pata, tg, datos_pata, nombre_pata,
                orient_menor_p, "LARGO pata"
            )
            if ok_pata:
                creadas.add(nombre_pata)
                print(f"📐 {nombre_lado}: creada hoja extra {nombre_pata}")
            else:
                print(
                    f"🗑️ {nombre_pata}: sin cota asociativa; se elimina "
                    "(no se exporta JPG solo-nota)."
                )
                try:
                    hoja_pata.Delete()
                except Exception:
                    pass

    return creadas


def _sanear_thk_si_duplica_alto(hoja_lado, vista, tg, datos, meta, nombre_lado, extras):
    """
    Si tras crear ``_ALTO`` el THK de LADO ≈ alto del perfil (P04/P05),
    sustituye por espesor real de chapa / gap fino / nota de modelo.

    Evita exportar el mismo 8.00 como HEIGHT y como THK.
    """
    if not extras or not datos:
        return False
    extras_up = {str(x).upper() for x in extras}
    if not any("_ALTO" in x for x in extras_up):
        return False

    valor_thk = (meta or {}).get("valor_cm")
    if valor_thk is None or valor_thk <= EPS:
        return False

    minx, maxx, miny, maxy = _bbox_global(datos)
    mayor_sheet = max(maxx - minx, maxy - miny)
    if mayor_sheet <= EPS:
        return False
    mayor_model = _esperado_modelo(vista, mayor_sheet)
    if mayor_model <= EPS:
        return False
    if abs(float(valor_thk) - float(mayor_model)) / mayor_model > 0.12:
        return False

    # THK midió el alto del perfil → corregir solo con espesor VALIDADO.
    real, _origen = _espesor_thk_validado(
        vista, alto_cm=valor_thk, datos=datos, nombre_hoja=nombre_lado
    )
    if real is None or real <= EPS:
        # último recurso: lado menor del bbox 2D si es fino vs mayor
        menor_sheet = min(maxx - minx, maxy - miny)
        menor_model = _esperado_modelo(vista, menor_sheet)
        if (
            menor_model > EPS
            and menor_model < mayor_model * 0.35
            and abs(menor_model - valor_thk) / mayor_model > 0.12
        ):
            real = menor_model
    if real is None or real <= EPS:
        print(
            f"⚠️ {nombre_lado}: THK≈ALTO ({valor_thk / IN_TO_CM:.3f} in) "
            f"pero no hay espesor validado para sustituir; se deja pendiente."
        )
        try:
            LAST_PENDIENTES_THK.append(str(nombre_lado))
        except Exception:
            pass
        return False
    if abs(float(real) - float(valor_thk)) / max(float(valor_thk), EPS) < 0.05:
        return False

    try:
        dims = hoja_lado.DrawingDimensions.GeneralDimensions
        for i in range(dims.Count, 0, -1):
            try:
                dims.Item(i).Delete()
            except Exception:
                pass
    except Exception:
        pass

    ok = _acotar_thk_asociativa_forzada(
        hoja_lado, vista, tg, datos, nombre_lado, real
    )
    if not ok:
        ok, _ = _forzar_cota_thk_desde_modelo(
            hoja_lado,
            tg,
            vista,
            nombre_lado,
            alto_cm=valor_thk,
            datos=datos,
        )
    if ok:
        print(
            f"↩️ {nombre_lado}: THK era duplicado de ALTO "
            f"({valor_thk / IN_TO_CM:.3f}); corregido a "
            f"{real / IN_TO_CM:.4f} in"
        )
        if isinstance(meta, dict):
            meta["valor_cm"] = real
            meta["gap_sheet"] = real
        return True
    return False


def _resolver_circular_solid(hoja, vista, tg, outer, nombre_hoja):
    diam_cm = _esperado_modelo(vista, outer["dx"])
    snap = _snap_o_medido(diam_cm)

    try:
        intencion = hoja.CreateGeometryIntent(outer["curve"])
        # Silueta completa de la vista para no poner Ø encima de la pieza.
        try:
            datos_v = _obtener_curvas_validas(vista)
            pieza_bb = _bbox_global(datos_v) if datos_v else None
        except Exception:
            pieza_bb = (
                outer["minx"],
                outer["maxx"],
                outer["miny"],
                outer["maxy"],
            )
        clr = clearance_texto_cota_cm(10) + max(1.0, float(outer["dx"]) * 0.35)
        if pieza_bb is not None:
            from cota_estilo import empujar_punto_fuera_bbox

            try:
                sheet_w = float(hoja.Width)
                sheet_h = float(hoja.Height)
                minx, maxx, miny, maxy = pieza_bb
                aire = {
                    "der": sheet_w - maxx,
                    "izq": minx,
                    "sup": sheet_h - maxy,
                    "inf": miny,
                }
                lado = max(aire, key=aire.get)
            except Exception:
                lado = "der"
            tx, ty = empujar_punto_fuera_bbox(
                outer["cx"], outer["cy"], pieza_bb, lado, clr
            )
        else:
            tx = outer["cx"] + (outer["dx"] / 2.0) + 1.0
            ty = outer["cy"] + (outer["dx"] / 2.0) + 1.0
        punto_texto = _clampear_punto_hoja(
            hoja, tg, tx, ty, evitar_bbox=pieza_bb
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
        aplicar_estilo_cota(dim, hoja=hoja)
        if pieza_bb is not None:
            asegurar_cota_fuera_pieza_robusto(
                dim, tg, pieza_bb, holgura=0.6, n_chars=12
            )

        origen = "catálogo" if snap.get("desde_catalogo") else "medido"
        print(
            f"✅ {nombre_hoja}: Ø sólido = {snap['valor_in']:.4f} in "
            f"({origen}, detectado {diam_cm/IN_TO_CM:.4f} in)"
        )
        return True, {
            "gap_sheet": None,
            "valor_cm": diam_cm,
            "valor_in": snap["valor_in"],
        }

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó Ø sólido -> {e}")
        return False, None


def _resolver_circular_hollow(hoja, vista, tg, outer, inner, nombre_hoja):
    # espesor radial = (Dext - Dint) / 2
    thk_sheet = (outer["dx"] - inner["dx"]) / 2.0
    thk_cm = _esperado_modelo(vista, thk_sheet)
    snap = _snap_o_medido(thk_cm)

    try:
        outer_r = outer["dx"] / 2.0
        inner_r = inner["dx"] / 2.0

        int_outer = _crear_intent_punto2d(
            hoja, tg, outer["curve"],
            outer["cx"] + outer_r, outer["cy"]
        )
        int_inner = _crear_intent_punto2d(
            hoja, tg, inner["curve"],
            inner["cx"] + inner_r, inner["cy"]
        )

        if not int_outer or not int_inner:
            print(f"⚠️ {nombre_hoja}: no se pudieron crear intents radiales.")
            return False, None

        try:
            datos_v = _obtener_curvas_validas(vista)
            pieza_bb = _bbox_global(datos_v) if datos_v else None
        except Exception:
            pieza_bb = (
                outer["minx"],
                outer["maxx"],
                outer["miny"],
                outer["maxy"],
            )
        clr = clearance_texto_cota_cm(10) + OFFSET_COTA
        if pieza_bb is not None:
            from cota_estilo import empujar_punto_fuera_bbox

            try:
                sheet_w = float(hoja.Width)
                sheet_h = float(hoja.Height)
                minx, maxx, miny, maxy = pieza_bb
                aire = {
                    "der": sheet_w - maxx,
                    "izq": minx,
                    "sup": sheet_h - maxy,
                    "inf": miny,
                }
                lado = max(aire, key=aire.get)
            except Exception:
                lado = "der"
            tx, ty = empujar_punto_fuera_bbox(
                outer["cx"], outer["cy"], pieza_bb, lado, clr
            )
        else:
            tx = outer["cx"] + outer_r + OFFSET_COTA
            ty = outer["cy"] + OFFSET_COTA
        pt_texto = _clampear_punto_hoja(
            hoja, tg, tx, ty, evitar_bbox=pieza_bb
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inner, int_outer, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)
        if pieza_bb is not None:
            asegurar_cota_fuera_pieza_robusto(
                dim, tg, pieza_bb, holgura=0.6, n_chars=12
            )

        origen = "catálogo" if snap.get("desde_catalogo") else "medido"
        print(
            f"✅ {nombre_hoja}: THK circular = {snap['valor_in']:.4f} in "
            f"({origen}, detectado {thk_cm/IN_TO_CM:.4f} in)"
        )
        return True, {
            "gap_sheet": thk_sheet,
            "valor_cm": thk_cm,
            "valor_in": snap["valor_in"],
        }

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó THK circular -> {e}")
        return False, None


def acotar_thk(nombres_permitidos=None):
    """
    Aplica cotas THK sobre hojas _LADO.

    Parametros
    ----------
    nombres_permitidos : set[str] | None
        Si se provee, solo se procesan hojas cuyo nombre (upper) esté en el
        set. Útil para procesar por lotes (modo D).

    Returns
    -------
    set[str]
        Nombres de hojas extra creadas (p. ej. ``_ALTO`` para perfiles U/L).
    """
    print("📏 THK.py: Iniciando pruebas para hojas _LADO...")
    if _THK_LOG:
        print("[THK_LOG] modo diagnóstico ACTIVO (THK_LOG=1)")

    solo_flat = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    if solo_flat:
        print("  SOLO_FLAT_CORTE: solo *_DESPLIEGUE_LADO* (no LADO doblado)")
        print(
            "  SOLO_FLAT_CORTE: bloqueado reorientar a escuadra L/U "
            "y hojas ALTO/LARGO_PATA"
        )

    hojas_extra = set()
    permitidos_up = None
    if nombres_permitidos is not None:
        permitidos_up = {str(x).upper() for x in nombres_permitidos}
        print(f"  Modo lote: {len(permitidos_up)} hojas permitidas")

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return hojas_extra

    tg = inv_app.TransientGeometry

    procesadas = 0
    pendientes = []

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_completo = str(hoja.Name)
        nombre_hoja = nombre_completo.upper()
        base_up = _base_hoja(nombre_completo).upper()

        if permitidos_up is not None and base_up not in permitidos_up:
            continue

        if solo_flat and "_DESPLIEGUE_LADO" not in nombre_hoja:
            continue

        if "_LADO" not in nombre_hoja:
            continue

        if hoja.DrawingViews.Count == 0:
            print(f"⏭️ {nombre_hoja}: sin vistas.")
            continue

        vista = hoja.DrawingViews.Item(1)
        datos = _obtener_curvas_validas(vista)

        if not datos:
            print(f"⚠️ {nombre_hoja}: sin curvas válidas.")
            pendientes.append(nombre_hoja)
            continue

        # Jacking pads / escuadras L: LADO debe ser la escuadra, no la cara.
        # SOLO_FLAT / DESPLIEGUE: NUNCA reorientar a perfil doblado (L/U).
        # Eso reemplazaba el canto flat por la vista de pieza ya doblada.
        thk_chk = _espesor_chapa_desde_vista(vista) or _espesor_desde_bbox_3d(
            vista
        )
        necesita_escuadra = False
        es_despliegue = "_DESPLIEGUE_" in nombre_hoja
        if solo_flat or es_despliegue:
            necesita_escuadra = False
        elif _nombre_parece_jacking_escuadra(nombre_hoja):
            necesita_escuadra = (
                _parece_silueta_franja_canto(datos)
                or _es_vista_cara_plana(datos)
                or not _es_perfil_u_o_l(datos, thk_chk)
            )
        elif (
            _parece_silueta_franja_canto(datos)
            and vista is not None
            and _es_perfil_por_modelo_3d(vista, thk_chk)
        ):
            # Escuadras genéricas (P91): bbox 3D de L pero LADO de cara.
            necesita_escuadra = True
        if necesita_escuadra:
            hoja2, vista2, datos2 = _reorientar_lado_a_escuadra(
                plano, hoja, tg, inv_app, str(hoja.Name)
            )
            if datos2 is not None:
                hoja, vista, datos = hoja2, vista2, datos2

        tipo, outer, inner = _clasificar_lado(datos)
        print(f"🔎 {nombre_hoja}: clasificado como {tipo}")

        ok = False
        meta = None

        if tipo == "circular_solid":
            # TIERRA/GROUND: la cara redonda NO es el THK (Ø≈3.14 flotante).
            # Forzar camino prismático / modelo.
            if _nombre_parece_tierra(nombre_hoja):
                print(
                    f"↩️ {nombre_hoja}: TIERRA/GROUND — Ø de cara no es THK; "
                    f"se resuelve espesor"
                )
                ok, meta = _resolver_prismatico(
                    hoja, vista, tg, datos, nombre_hoja
                )
            else:
                ok, meta = _resolver_circular_solid(
                    hoja, vista, tg, outer, nombre_hoja
                )
                if ok:
                    try:
                        LAST_OD_SOLID_LADO.append(str(nombre_hoja))
                    except Exception:
                        pass

        elif tipo == "circular_hollow":
            if _nombre_parece_tierra(nombre_hoja):
                print(
                    f"↩️ {nombre_hoja}: TIERRA/GROUND hollow — espesor, no Ø"
                )
                ok, meta = _resolver_prismatico(
                    hoja, vista, tg, datos, nombre_hoja
                )
            else:
                ok, meta = _resolver_circular_hollow(
                    hoja, vista, tg, outer, inner, nombre_hoja
                )

        elif tipo == "rect_hollow":
            ok, meta = _resolver_rectangular_hollow(
                hoja, vista, tg, datos, nombre_hoja
            )

        elif _es_perfil_semicircular(datos):
            ok, meta = _resolver_semicircular(hoja, vista, tg, datos, nombre_hoja)

        else:
            ok, meta = _resolver_prismatico(hoja, vista, tg, datos, nombre_hoja)

        if ok:
            procesadas += 1
            thk_sheet = (meta or {}).get("gap_sheet")
            valor_meta = (meta or {}).get("valor_cm")
            try:
                if _alinear_cota_thk_a_chapa_board(
                    hoja, vista, nombre_hoja, meta
                ):
                    thk_sheet = (meta or {}).get("gap_sheet") or thk_sheet
                    valor_meta = (meta or {}).get("valor_cm") or valor_meta
            except Exception as exc_al:
                _dbg(f"{nombre_hoja}: alinear THK chapa falló ({exc_al})")
            # Cara plana con "THK" enorme = midió el ancho de placa (TOP P01).
            if (
                valor_meta
                and _es_vista_cara_plana(datos)
                and tipo not in ("circular_solid", "circular_hollow")
            ):
                thk_real, _ = _espesor_thk_validado(
                    vista,
                    alto_cm=valor_meta,
                    datos=datos,
                    nombre_hoja=nombre_hoja,
                )
                if (
                    thk_real
                    and thk_real > EPS
                    and float(valor_meta) > float(thk_real) * 2.5
                ):
                    try:
                        dims = hoja.DrawingDimensions.GeneralDimensions
                        for di in range(dims.Count, 0, -1):
                            dims.Item(di).Delete()
                    except Exception:
                        pass
                    if _acotar_thk_asociativa_forzada(
                        hoja, vista, tg, datos, nombre_hoja, thk_real
                    ) or _forzar_cota_thk_desde_modelo(
                        hoja,
                        tg,
                        vista,
                        nombre_hoja,
                        alto_cm=valor_meta,
                        datos=datos,
                    )[0]:
                        print(
                            f"↩️ {nombre_hoja}: THK era cara plana "
                            f"({valor_meta / IN_TO_CM:.3f} in); "
                            f"forzado a {thk_real / IN_TO_CM:.4f} in"
                        )
                        if isinstance(meta, dict):
                            meta["valor_cm"] = thk_real
                            meta["gap_sheet"] = thk_real
                        thk_sheet = thk_real
                        valor_meta = thk_real
                    else:
                        print(
                            f"⚠️ {nombre_hoja}: cara plana midió "
                            f"{valor_meta / IN_TO_CM:.3f} in y no hay THK "
                            f"validado; se deja pendiente."
                        )
                        try:
                            dims = hoja.DrawingDimensions.GeneralDimensions
                            for di in range(dims.Count, 0, -1):
                                dims.Item(di).Delete()
                        except Exception:
                            pass
                        procesadas = max(0, procesadas - 1)
                        pendientes.append(nombre_hoja)
                        continue
            _dbg(
                f"{nombre_hoja}: OK "
                f"valor={valor_meta / IN_TO_CM:.4f}in "
                f"gap_sheet={thk_sheet:.4f}" if (valor_meta and thk_sheet)
                else f"{nombre_hoja}: OK (sin metadata numérica)"
            )
            # HSS ya expone su sección en _LADO y su largo en las vistas
            # frontales; no requiere hojas _ALTO/_LARGO_PATA adicionales.
            # Para perfiles U/C/L doblados _crear_hoja_alto crea DOS hojas
            # separadas (una por cota) para que cada JPG traiga UNA sola
            # dimensión — así el usuario puede verificar bend deduction
            # midiendo la pieza física contra cada foto.
            # SOLO_FLAT / DESPLIEGUE: solo THK de canto flat — nada de ALTO,
            # LARGO_PATA ni vistas de perfil doblado.
            if (
                thk_sheet
                and tipo != "rect_hollow"
                and not solo_flat
                and not es_despliegue
            ):
                if (meta or {}).get("es_patas_base"):
                    # Parking: LADO ya tiene patas→base. Renombrar a ALTO y
                    # crear hoja THK real desde Thickness de chapa/barra.
                    extras = _finalizar_parking_patas_base(
                        plano, hoja, tg, vista, meta, str(hoja.Name)
                    )
                else:
                    extras = _crear_hoja_alto(
                        plano, hoja, tg, datos, thk_sheet, str(hoja.Name)
                    )
                if extras:
                    hojas_extra.update(extras)
                    _dbg(f"{nombre_hoja}: hojas extra creadas -> {sorted(extras)}")
                    try:
                        _sanear_thk_si_duplica_alto(
                            hoja, vista, tg, datos, meta, nombre_hoja, extras
                        )
                    except Exception as exc_san:
                        _dbg(f"{nombre_hoja}: saneo THK/ALTO falló ({exc_san})")

            # Solera Jacking Pad: LARGO (eje mayor) desde la misma vista THK.
            if (
                not solo_flat
                and not es_despliegue
                and _nombre_parece_solera_jacking(nombre_hoja)
            ):
                extras_s = _finalizar_largo_desde_vista_thk(
                    plano,
                    hoja,
                    tg,
                    vista,
                    datos,
                    thk_sheet or (meta or {}).get("gap_sheet"),
                    str(hoja.Name),
                )
                if extras_s:
                    hojas_extra.update(extras_s)

            # Nipple / STUD / tubo: longitud axial aparte (FRENTE suele ser solo Ø).
            if (
                not solo_flat
                and not es_despliegue
                and _necesita_alto_axial(nombre_hoja)
            ):
                extras_n = _finalizar_nipple_alto(
                    plano, hoja, tg, vista, datos, meta, str(hoja.Name)
                )
                if extras_n:
                    hojas_extra.update(extras_n)
        else:
            # Preferir cota asociativa antes que leyenda tipográfica.
            # Nunca forzar THK desde bbox crudo si no está validado.
            minx_f, maxx_f, miny_f, maxy_f = _bbox_global(datos) if datos else (0, 0, 0, 0)
            alto_sil = None
            if datos:
                alto_sil = _esperado_modelo(
                    vista, min(maxx_f - minx_f, maxy_f - miny_f)
                )
            thk_fallback, origen_fb = _espesor_thk_validado(
                vista,
                alto_cm=alto_sil,
                datos=datos,
                nombre_hoja=nombre_hoja,
            )
            forzado_ok = False
            _valor_cm = None
            if thk_fallback and datos:
                forzado_ok = _acotar_thk_asociativa_forzada(
                    hoja, vista, tg, datos, nombre_hoja, thk_fallback
                )
                if forzado_ok:
                    _valor_cm = thk_fallback
                    print(
                        f"↩️ {nombre_hoja}: THK asociativa forzada "
                        f"= {thk_fallback / IN_TO_CM:.4f} in ({origen_fb})"
                    )
            if not forzado_ok:
                forzado_ok, _valor_cm = _forzar_cota_thk_desde_modelo(
                    hoja,
                    tg,
                    vista,
                    nombre_hoja,
                    alto_cm=alto_sil,
                    datos=datos,
                )
            if forzado_ok:
                procesadas += 1
                _dbg(f"{nombre_hoja}: THK resuelto por fallback validado.")
                # Aun con nota, intentar LARGO desde la vista de canto.
                if _nombre_parece_solera_jacking(nombre_hoja):
                    extras_s = _finalizar_largo_desde_vista_thk(
                        plano, hoja, tg, vista, datos, _valor_cm, str(hoja.Name)
                    )
                    if extras_s:
                        hojas_extra.update(extras_s)
                # Nipple / STUD / tubo aún puede necesitar ALTO de longitud.
                if _necesita_alto_axial(nombre_hoja):
                    extras_n = _finalizar_nipple_alto(
                        plano, hoja, tg, vista, datos, None, str(hoja.Name)
                    )
                    if extras_n:
                        hojas_extra.update(extras_n)
            else:
                pendientes.append(nombre_hoja)
                _dbg(f"{nombre_hoja}: NO resuelto (tipo={tipo})")

    print(f"\n✅ THK.py finalizado. Hojas cotadas: {procesadas}")
    if hojas_extra:
        print(f"📐 Hojas ALTO creadas: {len(hojas_extra)}")

    if pendientes:
        print("⚠️ Hojas _LADO no resueltas:")
        for h in pendientes:
            print(f"   - {h}")
        # Acumular en el registro global para el reporte de fin de flujo.
        for h in pendientes:
            LAST_PENDIENTES_THK.append(str(h))

    return hojas_extra


if __name__ == "__main__":
    acotar_thk()
