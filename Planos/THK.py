import math
import os
import time
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota

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


def reset_pendientes_thk():
    """Vacía el registro global de pendientes THK. Llamar al inicio del flujo."""
    LAST_PENDIENTES_THK.clear()


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
OFFSET_COTA = 1.5

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
                    pt = _clampear_punto_hoja(
                        hoja, tg,
                        cx + ro + OFFSET_COTA,
                        cy + OFFSET_COTA,
                    )
                    dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                        pt, int_i, int_o, kHorizontalDimensionType
                    )
                    aplicar_estilo_cota(dim, hoja=hoja)
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
    """
    dims = _dimensiones_bbox_3d(vista)
    if not dims:
        return None
    dims.sort()
    menor = dims[0]
    if menor <= EPS:
        return None
    return menor


def _forzar_cota_thk_desde_modelo(hoja, tg, vista, nombre_hoja):
    """
    Fallback final cuando ningún resolver geométrico pudo dibujar el THK.

    Estrategia: lee el espesor real del modelo (Thickness de sheet metal, o
    lado más pequeño del bbox 3D si es una parte normal) y coloca una
    ``GeneralNote`` de texto ``THK = X.XXX in`` sobre la hoja. Esto permite
    que el JPG final tenga el dato aunque la vista sea cara plana o de
    orientación dudosa. La hoja se renombra a ``_THK`` porque SÍ tiene el
    espesor informado (aunque venga del modelo, no de la geometría 2D).

    Retorna (True, valor_cm) si logró agregar la nota, (False, None) si no.
    """
    valor_cm = _espesor_chapa_desde_vista(vista)
    origen = "sheet_metal_thickness"
    if valor_cm is None or valor_cm <= EPS:
        valor_cm = _espesor_desde_bbox_3d(vista)
        origen = "bbox_3d_menor"

    if valor_cm is None or valor_cm <= EPS:
        return False, None

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
    except Exception as exc:
        print(
            f"⚠️ {nombre_hoja}: no se pudo crear nota THK forzada ({exc})."
        )
        return False, None

    print(
        f"↩️ {nombre_hoja}: THK forzado desde modelo = {valor_in:.3f} in "
        f"(origen={origen})"
    )
    return True, valor_cm


def _forzar_nota_dimension_individual(
    hoja, tg, vista, nombre_hoja, etiqueta, valor_cm
):
    """
    Coloca UNA sola ``GeneralNote`` centrada bajo la vista, con el formato
    ``ETIQUETA = X.XXX in``. Se usa como fallback cuando la cota geométrica
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
    valor_in = valor_cm / IN_TO_CM
    texto = f"{etiqueta} = {valor_in:.3f} in"

    try:
        hoja.DrawingNotes.GeneralNotes.AddFitted(pt, texto)
    except Exception as exc:
        print(f"⚠️ {nombre_hoja}: no se pudo crear nota {etiqueta} ({exc}).")
        return False

    print(f"↩️ {nombre_hoja}: {etiqueta} forzado desde bbox 3D = {valor_in:.3f} in")
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


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2):
    """Devuelve un Point2d garantizado dentro de la hoja física con margen.

    Sin esto, cotas dibujadas cerca del borde quedan fuera del rectángulo que
    la cámara exporta como JPG (Inventor las acepta pero el bitmap no las
    incluye), y el resultado es la clásica "captura sin cota". El margen
    (1.2 cm) da espacio al número y la flecha para caber enteros.
    """
    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        x = max(margen, min(sheet_w - margen, x))
        y = max(margen, min(sheet_h - margen, y))
    except Exception:
        pass
    return tg.CreatePoint2d(x, y)


def _dibujar_cota_prismatica(hoja, tg, mejor):
    a = mejor["a"]
    b = mejor["b"]
    tipo = mejor["tipo"]

    if tipo == "horizontal":
        int_a = hoja.CreateGeometryIntent(a["curve"])
        int_b = hoja.CreateGeometryIntent(b["curve"])
        x_texto = ((a["cx"] + b["cx"]) / 2.0)
        y_texto = max(a["maxy"], b["maxy"]) + OFFSET_COTA
        pt_texto = _clampear_punto_hoja(hoja, tg, x_texto, y_texto)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_a, int_b, kHorizontalDimensionType
        )
    else:
        int_a = hoja.CreateGeometryIntent(a["curve"])
        int_b = hoja.CreateGeometryIntent(b["curve"])
        x_texto = min(a["minx"], b["minx"]) - OFFSET_COTA
        y_texto = ((a["cy"] + b["cy"]) / 2.0)
        pt_texto = _clampear_punto_hoja(hoja, tg, x_texto, y_texto)
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_a, int_b, kVerticalDimensionType
        )
    aplicar_estilo_cota(dim, hoja=hoja)
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
        if envolvente and not _es_vista_cara_plana(datos):
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
        _dibujar_cota_prismatica(hoja, tg, mejor)
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

    # Canales/escuadras con fillets suelen fragmentar curvas: aceptar
    # perfiles compactos con bastante geometría.
    if len(datos) >= 8 and aspect < 4.5:
        return True
    return False


def _acotar_lado_bbox(hoja, vista, tg, datos, nombre_hoja, orientacion, etiqueta):
    """
    Dibuja UNA cota lineal (vertical u horizontal) sobre el bbox global de
    ``datos``. ``orientacion`` es 'V' (mide alto del bbox) o 'H' (mide ancho).

    Reintenta hasta 3 veces si Inventor devuelve -2147352567 (curvas COM
    en estado inestable tras crear la vista).

    Validación post-dibujo: después de crear la cota, verificamos que su
    ``RangeBox`` quede dentro del sheet físico con al menos 0.5 cm de
    margen. Si NO cabe, borramos la cota, movemos el punto de anclaje
    hacia el centro y reintentamos. Sin esta validación, la cota se
    puede dibujar en un lugar donde el JPG exportado no la incluye
    (P05_LARGO_PATA sin cota visible).

    Devuelve ``True`` si logró colocar la cota Y la cota queda dentro del
    sheet, ``False`` en caso contrario.
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

    def _pt_para_orient(dx_off, dy_off):
        """Calcula el punto de anclaje de la cota, con offsets adicionales
        que empujan hacia el centro del sheet si la primera posición se sale."""
        if orientacion == 'V':
            x = min(a["minx"], b["minx"]) - OFFSET_COTA - dx_off
            y = (miny + maxy) / 2.0 + dy_off
        else:
            x = (minx + maxx) / 2.0 + dx_off
            y = max(a["maxy"], b["maxy"]) + OFFSET_COTA + dy_off
        return _clampear_punto_hoja(hoja, tg, x, y)

    def _cota_dentro_de_sheet(dim_obj):
        """Verifica que el RangeBox de la cota (línea + texto + flechas)
        quede DENTRO del sheet físico con margen mínimo de 0.3 cm."""
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

    ultimo_error = None
    datos_iter = datos
    for intento in range(3):
        try:
            if orientacion == 'V':
                a = max(datos_iter, key=lambda d: d["maxy"])
                b = min(datos_iter, key=lambda d: d["miny"])
            else:
                a = max(datos_iter, key=lambda d: d["maxx"])
                b = min(datos_iter, key=lambda d: d["minx"])

            int_a = hoja.CreateGeometryIntent(a["curve"])
            int_b = hoja.CreateGeometryIntent(b["curve"])

            # Intentamos varias posiciones para la cota, empujando cada vez
            # más hacia el centro del sheet si la anterior no cupo. Si
            # ninguna cabe perfecto, nos quedamos con la PRIMERA como
            # último recurso (mejor una cota parcialmente al borde que
            # ninguna cota).
            offsets_a_probar = [
                (0.0, 0.0),
                (1.5, 0.0),
                (3.0, 0.0),
                (0.0, 1.5),
                (0.0, 3.0),
            ]

            dim_creado = None
            dim_fallback = None
            for dx_off, dy_off in offsets_a_probar:
                pt = _pt_para_orient(dx_off, dy_off)
                try:
                    if orientacion == 'V':
                        dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                            pt, int_a, int_b, kVerticalDimensionType
                        )
                    else:
                        dim_test = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                            pt, int_a, int_b, kHorizontalDimensionType
                        )
                except Exception:
                    continue

                if _cota_dentro_de_sheet(dim_test):
                    dim_creado = dim_test
                    break

                # Guardar como fallback la primera cota que sí se pudo
                # crear (aunque no quepa perfectamente).
                if dim_fallback is None:
                    dim_fallback = dim_test
                else:
                    try:
                        dim_test.Delete()
                    except Exception:
                        pass

            if dim_creado is None and dim_fallback is not None:
                dim_creado = dim_fallback
                _dbg(
                    f"{nombre_hoja}: cota aceptada como fallback (cerca del "
                    f"borde), mejor eso que hoja sin dimensión."
                )

            if dim_creado is None:
                raise RuntimeError(
                    "no se pudo crear cota en ningún offset probado"
                )

            aplicar_estilo_cota(dim_creado, hoja=hoja)

            span = h if orientacion == 'V' else w
            valor_in = _esperado_modelo(vista, span) / IN_TO_CM
            print(f"✅ {nombre_hoja}: {etiqueta} = {valor_in:.4f} in")
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


def _calcular_camara_transversal(part_doc, tg, to):
    """
    Calcula una cámara que mire A LO LARGO del eje MAYOR del bbox 3D del
    modelo, para mostrar la SECCIÓN TRANSVERSAL de un perfil L/U/C.

    Reutiliza ``creador_vistas._orientacion_lado_doblado`` cuando está
    disponible (detección específica para perfiles doblados con normales
    perpendiculares). Si no, calcula manualmente desde el bbox 3D.

    Devuelve un ``Camera`` de Inventor listo para ``AddBaseView``, o
    ``None`` si no se pudo determinar.
    """
    if part_doc is None or tg is None or to is None:
        return None

    # Vía preferida: reutilizar la lógica probada del flujo normal.
    if _creador_vistas is not None:
        try:
            ori = _creador_vistas._orientacion_lado_doblado(
                part_doc, tg, to, None
            )
            if ori is not None:
                eye_dir = ori["v_lado"]
                up_hint = ori["v_up"]
                cx, cy, cz = ori["cx"], ori["cy"], ori["cz"]
                return _creador_vistas.crear_camara(
                    part_doc, tg, to, cx, cy, cz, eye_dir, up_hint
                )
        except Exception:
            pass

    # Fallback manual: eye = eje mayor del bbox 3D.
    try:
        rb = part_doc.ComponentDefinition.RangeBox
        cx = (float(rb.MaxPoint.X) + float(rb.MinPoint.X)) / 2.0
        cy = (float(rb.MaxPoint.Y) + float(rb.MinPoint.Y)) / 2.0
        cz = (float(rb.MaxPoint.Z) + float(rb.MinPoint.Z)) / 2.0
        dx = float(rb.MaxPoint.X) - float(rb.MinPoint.X)
        dy = float(rb.MaxPoint.Y) - float(rb.MinPoint.Y)
        dz = float(rb.MaxPoint.Z) - float(rb.MinPoint.Z)
        ejes = sorted(
            [("X", abs(dx)), ("Y", abs(dy)), ("Z", abs(dz))],
            key=lambda t: t[1],
            reverse=True,
        )
        eje_mayor = ejes[0][0]
        eje_medio = ejes[1][0]
        vec = {"X": (1, 0, 0), "Y": (0, 1, 0), "Z": (0, 0, 1)}
        eye_dir = tg.CreateVector(*vec[eje_mayor])
        up_hint = tg.CreateVector(*vec[eje_medio])
        if _creador_vistas is not None:
            return _creador_vistas.crear_camara(
                part_doc, tg, to, cx, cy, cz, eye_dir, up_hint
            )
    except Exception:
        pass
    return None


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

    # Reescalar la vista de forma DEFENSIVA: la meta es que la pieza ocupe
    # ~55 % del sheet dejando ~22 % de margen por lado para dibujar cotas.
    # NO usamos ``escalar_vista`` de creador_vistas aquí porque su loop de
    # verificación puede dejar la vista con escala mínima (0.001) cuando
    # la geometría del sheet clonado es asimétrica, dando lugar a vistas
    # inservibles (fue lo que rompió P14 en la corrida anterior).
    try:
        sheet_w = float(nueva.Width)
        sheet_h = float(nueva.Height)
    except Exception:
        sheet_w = 0.0
        sheet_h = 0.0

    try:
        rb_view = vista_nueva.RangeBox
        w_v = float(rb_view.MaxPoint.X) - float(rb_view.MinPoint.X)
        h_v = float(rb_view.MaxPoint.Y) - float(rb_view.MinPoint.Y)
    except Exception:
        w_v = 0.0
        h_v = 0.0

    if sheet_w > 0 and sheet_h > 0 and w_v > 0 and h_v > 0:
        try:
            scale_actual = float(vista_nueva.Scale)
        except Exception:
            scale_actual = 1.0
        # Dimensiones reales de la pieza (independientes de escala actual).
        real_w = w_v / max(scale_actual, EPS)
        real_h = h_v / max(scale_actual, EPS)

        # Objetivo: 55 % del sheet en cada dirección (deja 22 % por lado
        # para cotas). Elegimos la escala más pequeña de las dos.
        max_w_pieza = sheet_w * 0.55
        max_h_pieza = sheet_h * 0.55
        escala_ideal = min(
            max_w_pieza / max(real_w, EPS),
            max_h_pieza / max(real_h, EPS),
        )

        escalas_discretas = [
            5.0, 4.0, 3.0, 2.0, 1.5, 1.0, 0.75, 0.5,
            0.4, 0.3, 0.25, 0.2, 0.15, 0.1, 0.08, 0.05,
            0.04, 0.03, 0.02, 0.015, 0.01,
        ]
        # Escala más grande que sea ≤ ideal.
        escala_elegida = 0.01
        for e in escalas_discretas:
            if e <= escala_ideal:
                escala_elegida = e
                break

        try:
            vista_nueva.Scale = escala_elegida
        except Exception:
            pass

        # Centrar la vista en el sheet.
        try:
            vista_nueva.Position = tg.CreatePoint2d(sheet_w / 2.0, sheet_h / 2.0)
        except Exception:
            pass

        # Update para que Inventor materialice el cambio de escala.
        if inv_app is not None:
            try:
                inv_app.ActiveDocument.Update()
            except Exception:
                pass

    # Esperar a que Inventor materialice las DrawingCurves de la vista nueva.
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
    time.sleep(0.5)

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

    # Si NO lo muestra pero SÍ es perfil por modelo 3D, calcular cámara
    # transversal alternativa. Esto atiende P05_Default_As Machined y
    # similares donde la vista LADO cayó en cara plana.
    camara_alt = None
    if not perfil_ok_en_lado and vista_orig is not None and inv_app is not None:
        try:
            part_doc = vista_orig.ReferencedDocumentDescriptor.ReferencedDocument
            to = inv_app.TransientObjects
            camara_alt = _calcular_camara_transversal(part_doc, tg, to)
            if camara_alt is not None:
                _dbg(
                    f"{nombre_lado}: LADO no muestra perfil; usando cámara "
                    f"transversal alternativa para _ALTO/_LARGO_PATA."
                )
        except Exception as exc_cam:
            _dbg(f"{nombre_lado}: no se pudo calcular cámara alterna ({exc_cam})")
            camara_alt = None

    # ============================================================
    # Hoja 1: _ALTO
    # ============================================================
    nombre_alto = _nombre_hoja_alto(nombre_lado)
    hoja_alto, vista_alto = _clonar_hoja_lado_para_cota(
        plano, hoja_lado, tg, nombre_alto, inv_app, camara_alt=camara_alt
    )
    if hoja_alto is None:
        print(f"⚠️ {nombre_lado}: no se pudo crear hoja {nombre_alto}")
    else:
        datos_alto = None
        for intento in range(3):
            try:
                datos_alto = _obtener_curvas_validas(vista_alto)
            except Exception:
                datos_alto = None
            if datos_alto:
                break
            if inv_app is not None:
                try:
                    inv_app.UserInterfaceManager.DoEvents()
                except Exception:
                    pass
            time.sleep(0.4 * (intento + 1))

        if not datos_alto:
            print(f"⚠️ {nombre_alto}: sin curvas 2D en la vista clonada.")
            try:
                hoja_alto.Delete()
            except Exception:
                pass
        else:
            # Bbox 2D calculado sobre la vista NUEVA (así funciona igual
            # si venimos de LADO original que si venimos de cámara alterna).
            aminx, amaxx, aminy, amaxy = _bbox_global(datos_alto)
            aw = amaxx - aminx
            ah = amaxy - aminy
            orient_mayor_a = 'V' if ah >= aw else 'H'
            ok_alto = _acotar_lado_bbox(
                hoja_alto, vista_alto, tg, datos_alto, nombre_alto,
                orient_mayor_a, "ALTO perfil"
            )
            if not ok_alto:
                # Fallback: nota con el valor del bbox 3D (dimensión media).
                dims_3d = _dimensiones_bbox_3d(vista_alto)
                valor_alto_cm = None
                if dims_3d and len(dims_3d) >= 3:
                    dims_ord = sorted(dims_3d)
                    valor_alto_cm = dims_ord[1]  # medio = ALTO del perfil
                ok_alto = _forzar_nota_dimension_individual(
                    hoja_alto, tg, vista_alto, nombre_alto,
                    "ALTO", valor_alto_cm
                )
            if ok_alto:
                creadas.add(nombre_alto)
                print(f"📐 {nombre_lado}: creada hoja extra {nombre_alto}")
            else:
                try:
                    hoja_alto.Delete()
                except Exception:
                    pass

    # ============================================================
    # Hoja 2: _LARGO_PATA (sólo perfiles con ambas dimensiones útiles)
    # ============================================================
    nombre_pata = _nombre_hoja_largo_pata(nombre_lado)
    hoja_pata, vista_pata = _clonar_hoja_lado_para_cota(
        plano, hoja_lado, tg, nombre_pata, inv_app, camara_alt=camara_alt
    )
    if hoja_pata is None:
        print(f"⚠️ {nombre_lado}: no se pudo crear hoja {nombre_pata}")
    else:
        datos_pata = None
        for intento in range(3):
            try:
                datos_pata = _obtener_curvas_validas(vista_pata)
            except Exception:
                datos_pata = None
            if datos_pata:
                break
            if inv_app is not None:
                try:
                    inv_app.UserInterfaceManager.DoEvents()
                except Exception:
                    pass
            time.sleep(0.4 * (intento + 1))

        if not datos_pata:
            print(f"⚠️ {nombre_pata}: sin curvas 2D en la vista clonada.")
            try:
                hoja_pata.Delete()
            except Exception:
                pass
        else:
            pminx, pmaxx, pminy, pmaxy = _bbox_global(datos_pata)
            pw = pmaxx - pminx
            ph = pmaxy - pminy
            menor_p = min(pw, ph)

            # Sólo cotar la pata cuando el menor del bbox NO es prácticamente
            # el espesor (evita duplicar THK en franjas L de canto).
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
                if not ok_pata:
                    # Fallback: nota con el valor del bbox 3D (menor
                    # descartando el espesor de chapa).
                    dims_3d = _dimensiones_bbox_3d(vista_pata)
                    valor_pata_cm = None
                    if dims_3d and len(dims_3d) >= 3:
                        dims_ord = sorted(dims_3d)
                        # menor > thk (ya validado por cotar_menor).
                        valor_pata_cm = dims_ord[0]
                    ok_pata = _forzar_nota_dimension_individual(
                        hoja_pata, tg, vista_pata, nombre_pata,
                        "LARGO_PATA", valor_pata_cm
                    )
                if ok_pata:
                    creadas.add(nombre_pata)
                    print(f"📐 {nombre_lado}: creada hoja extra {nombre_pata}")
                else:
                    try:
                        hoja_pata.Delete()
                    except Exception:
                        pass

    return creadas


def _resolver_circular_solid(hoja, vista, tg, outer, nombre_hoja):
    diam_cm = _esperado_modelo(vista, outer["dx"])
    snap = _snap_o_medido(diam_cm)

    try:
        intencion = hoja.CreateGeometryIntent(outer["curve"])
        offset = (outer["dx"] / 2.0) + 1.0
        punto_texto = _clampear_punto_hoja(
            hoja, tg,
            outer["cx"] + offset,
            outer["cy"] + offset,
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
        aplicar_estilo_cota(dim, hoja=hoja)

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

        pt_texto = _clampear_punto_hoja(
            hoja, tg,
            outer["cx"] + outer_r + OFFSET_COTA,
            outer["cy"] + OFFSET_COTA,
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inner, int_outer, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

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

        tipo, outer, inner = _clasificar_lado(datos)
        print(f"🔎 {nombre_hoja}: clasificado como {tipo}")

        ok = False
        meta = None

        if tipo == "circular_solid":
            ok, meta = _resolver_circular_solid(hoja, vista, tg, outer, nombre_hoja)

        elif tipo == "circular_hollow":
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
            if thk_sheet and tipo != "rect_hollow":
                extras = _crear_hoja_alto(
                    plano, hoja, tg, datos, thk_sheet, str(hoja.Name)
                )
                if extras:
                    hojas_extra.update(extras)
                    _dbg(f"{nombre_hoja}: hojas extra creadas -> {sorted(extras)}")
        else:
            # Fallback final: si el resolver geométrico no encontró el THK
            # (típicamente por vista de cara plana o pieza sin par lineal
            # claro), tomamos el espesor directamente del modelo (Thickness
            # de sheet metal, o menor lado del bbox 3D). Así ninguna hoja
            # LADO queda sin dato de espesor.
            forzado_ok, _valor_cm = _forzar_cota_thk_desde_modelo(
                hoja, tg, vista, nombre_hoja
            )
            if forzado_ok:
                procesadas += 1
                _dbg(f"{nombre_hoja}: THK resuelto por fallback de modelo.")
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
