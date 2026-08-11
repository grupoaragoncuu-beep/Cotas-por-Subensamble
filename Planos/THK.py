import math
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota

kHorizontalDimensionType = 60162
kVerticalDimensionType = 60163

EPS = 0.0001
OFFSET_COTA = 1.5

# Inventor internamente trabaja en cm
IN_TO_CM = 2.54
TOL_IN = 0.005
TOL_CM = TOL_IN * IN_TO_CM

ALLOWED_IN = [
    0.06,
    0.07,
    0.105,
    0.119,
    0.187,
    0.25,
    0.3125,
    0.375,
    0.5,
    0.625,
    0.75,
    0.875,
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
    2.0
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

        return {
            "curve": curva,
            "minx": minx,
            "maxx": maxx,
            "miny": miny,
            "maxy": maxy,
            "dx": dx,
            "dy": dy,
            "cx": (minx + maxx) / 2.0,
            "cy": (miny + maxy) / 2.0
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
            "diff_cm": mejor_diff
        }

    return None


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


def _clasificar_lado(datos):
    """
    Devuelve:
    - ("circular_solid", outer, None)
    - ("circular_hollow", outer, inner)
    - ("prismatic", None, None)
    """
    if not datos:
        return ("prismatic", None, None)

    minx, maxx, miny, maxy = _bbox_global(datos)
    global_w = maxx - minx
    global_h = maxy - miny

    circulos = _buscar_circulos(datos)
    if not circulos:
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
            if overlap_y < min_len * 0.60:
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
            if overlap_x < min_len * 0.60:
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


def _resolver_prismatico(hoja, vista, tg, datos, nombre_hoja):
    candidatos = _buscar_candidatos_lineales(datos)
    if not candidatos:
        print(f"⚠️ {nombre_hoja}: no se encontraron candidatos prismáticos.")
        return False

    candidatos_validos = []

    for c in candidatos:
        valor_cm = _esperado_modelo(vista, c["gap_sheet"])
        snap = _snap_a_catalogo(valor_cm)
        if snap:
            c["valor_cm"] = valor_cm
            c["snap"] = snap
            candidatos_validos.append(c)

    if not candidatos_validos:
        print(f"⚠️ {nombre_hoja}: candidatos prismáticos fuera de catálogo.")
        return False

    # Elegimos el menor espesor válido
    mejor = min(candidatos_validos, key=lambda x: x["valor_cm"])

    a = mejor["a"]
    b = mejor["b"]
    tipo = mejor["tipo"]

    try:
        if tipo == "horizontal":
            int_a = hoja.CreateGeometryIntent(a["curve"])
            int_b = hoja.CreateGeometryIntent(b["curve"])

            x_texto = ((a["cx"] + b["cx"]) / 2.0)
            y_texto = max(a["maxy"], b["maxy"]) + OFFSET_COTA
            pt_texto = tg.CreatePoint2d(x_texto, y_texto)

            dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                pt_texto, int_a, int_b, kHorizontalDimensionType
            )
            aplicar_estilo_cota(dim, hoja=hoja)

        else:  # vertical
            int_a = hoja.CreateGeometryIntent(a["curve"])
            int_b = hoja.CreateGeometryIntent(b["curve"])

            x_texto = min(a["minx"], b["minx"]) - OFFSET_COTA
            y_texto = ((a["cy"] + b["cy"]) / 2.0)
            pt_texto = tg.CreatePoint2d(x_texto, y_texto)

            dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                pt_texto, int_a, int_b, kVerticalDimensionType
            )
            aplicar_estilo_cota(dim, hoja=hoja)

        print(
            f"✅ {nombre_hoja}: THK prismático = {mejor['snap']['valor_in']:.4f} in "
            f"(detectado {mejor['valor_cm']/IN_TO_CM:.4f} in)"
        )
        return True

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó la cota prismática -> {e}")
        return False


def _resolver_circular_solid(hoja, vista, tg, outer, nombre_hoja):
    diam_cm = _esperado_modelo(vista, outer["dx"])
    snap = _snap_a_catalogo(diam_cm)

    if not snap:
        print(
            f"⚠️ {nombre_hoja}: Ø sólido fuera de catálogo "
            f"({diam_cm/IN_TO_CM:.4f} in)"
        )
        return False

    try:
        intencion = hoja.CreateGeometryIntent(outer["curve"])
        offset = (outer["dx"] / 2.0) + 1.0
        punto_texto = tg.CreatePoint2d(
            outer["cx"] + offset,
            outer["cy"] + offset
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
        aplicar_estilo_cota(dim, hoja=hoja)

        print(
            f"✅ {nombre_hoja}: Ø sólido = {snap['valor_in']:.4f} in "
            f"(detectado {diam_cm/IN_TO_CM:.4f} in)"
        )
        return True

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó Ø sólido -> {e}")
        return False


def _resolver_circular_hollow(hoja, vista, tg, outer, inner, nombre_hoja):
    # espesor radial = (Dext - Dint) / 2
    thk_sheet = (outer["dx"] - inner["dx"]) / 2.0
    thk_cm = _esperado_modelo(vista, thk_sheet)
    snap = _snap_a_catalogo(thk_cm)

    if not snap:
        print(
            f"⚠️ {nombre_hoja}: THK radial fuera de catálogo "
            f"({thk_cm/IN_TO_CM:.4f} in)"
        )
        return False

    try:
        # Medimos radialmente por el lado derecho
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
            return False

        pt_texto = tg.CreatePoint2d(
            outer["cx"] + outer_r + OFFSET_COTA,
            outer["cy"] + OFFSET_COTA
        )

        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            pt_texto, int_inner, int_outer, kHorizontalDimensionType
        )
        aplicar_estilo_cota(dim, hoja=hoja)

        print(
            f"✅ {nombre_hoja}: THK circular = {snap['valor_in']:.4f} in "
            f"(detectado {thk_cm/IN_TO_CM:.4f} in)"
        )
        return True

    except Exception as e:
        print(f"⚠️ {nombre_hoja}: Inventor rechazó THK circular -> {e}")
        return False


def acotar_thk():
    print("📏 THK.py: Iniciando pruebas para hojas _LADO...")

    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except:
        print("❌ No hay un DrawingDocument activo.")
        return

    tg = inv_app.TransientGeometry

    procesadas = 0
    pendientes = []

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja = str(hoja.Name).upper()

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

        if tipo == "circular_solid":
            ok = _resolver_circular_solid(hoja, vista, tg, outer, nombre_hoja)

        elif tipo == "circular_hollow":
            ok = _resolver_circular_hollow(hoja, vista, tg, outer, inner, nombre_hoja)

        else:
            ok = _resolver_prismatico(hoja, vista, tg, datos, nombre_hoja)

        if ok:
            procesadas += 1
        else:
            pendientes.append(nombre_hoja)

    print(f"\n✅ THK.py finalizado. Hojas cotadas: {procesadas}")

    if pendientes:
        print("⚠️ Hojas _LADO no resueltas:")
        for h in pendientes:
            print(f"   - {h}")


if __name__ == "__main__":
    acotar_thk()