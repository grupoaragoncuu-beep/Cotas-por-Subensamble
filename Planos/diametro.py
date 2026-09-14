import os
import re
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import (
    aplicar_estilo_cota,
    asegurar_cota_fuera_pieza_robusto,
    clearance_texto_cota_cm,
    empujar_punto_fuera_bbox,
    punto_dentro_bbox,
)
from rutas_runtime import ruta_piezas_solidas


# Archivo donde publicamos los nombres BASE de piezas identificadas como
# "cilindros sólidos sin interior concéntrico". El flujo de exportación
# (``exportar_hojas_jpg``) lee este archivo para borrar cualquier JPG
# ``_DIAMETRO_INTERIOR_*.jpg`` residual que haya quedado de corridas viejas
# antes de que se agregara la eliminación de hojas huérfanas.
# Portable: Planos/.runtime/ (antes C:\Temp\...)
RUTA_PIEZAS_SOLIDAS = ruta_piezas_solidas()


def _publicar_piezas_solidas(nombres_hojas_solidas):
    """Escribe los nombres BASE de piezas sólidas para consumo del exportador."""
    if not nombres_hojas_solidas:
        return
    try:
        os.makedirs(os.path.dirname(RUTA_PIEZAS_SOLIDAS), exist_ok=True)
    except Exception:
        pass
    piezas = set()
    for hoja_nombre in nombres_hojas_solidas:
        base = str(hoja_nombre).split(":", 1)[0]
        base = re.sub(r"_FRENTE_[12]$", "", base, flags=re.IGNORECASE)
        base = re.sub(r"_DIAMETRO_(INTERIOR|EXTERIOR)$", "", base, flags=re.IGNORECASE)
        piezas.add(base.strip().upper())

    if not piezas:
        return

    existentes = set()
    try:
        if os.path.exists(RUTA_PIEZAS_SOLIDAS):
            with open(RUTA_PIEZAS_SOLIDAS, "r", encoding="utf-8") as f:
                for linea in f:
                    linea = linea.strip()
                    if linea:
                        existentes.add(linea.upper())
    except Exception:
        existentes = set()

    piezas.update(existentes)
    try:
        with open(RUTA_PIEZAS_SOLIDAS, "w", encoding="utf-8") as f:
            for p in sorted(piezas):
                f.write(p + "\n")
    except Exception:
        pass


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2, evitar_bbox=None):
    """
    Point2d dentro del sheet y, si se pasa ``evitar_bbox``, fuera de la
    silueta de la pieza (texto legible para visión/OCR).
    """
    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        x = max(margen, min(sheet_w - margen, float(x)))
        y = max(margen, min(sheet_h - margen, float(y)))
    except Exception:
        x, y = float(x), float(y)
    if evitar_bbox is not None and punto_dentro_bbox(
        x, y, evitar_bbox, holgura=0.35
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
        clr = clearance_texto_cota_cm(10)
        x, y = empujar_punto_fuera_bbox(x, y, evitar_bbox, lado, clr)
        try:
            sheet_w = float(hoja.Width)
            sheet_h = float(hoja.Height)
            x = max(margen, min(sheet_w - margen, x))
            y = max(margen, min(sheet_h - margen, y))
        except Exception:
            pass
    return tg.CreatePoint2d(x, y)


def _bbox_pieza_vista(vista):
    """(minx, maxx, miny, maxy) silueta 2D de la vista o None."""
    sil = _silueta_vista(vista)
    if not sil:
        return None
    return sil[0], sil[1], sil[2], sil[3]


def _punto_texto_barreno(hoja, tg, vista, cx, cy, tam):
    """
    Origen de texto para Ø/barreno: siempre fuera de la silueta de la pieza,
    con holgura generosa (visión).
    """
    pieza_bb = _bbox_pieza_vista(vista)
    clr = clearance_texto_cota_cm(10) + max(1.0, float(tam) * 0.45)
    if pieza_bb is None:
        return _clampear_punto_hoja(
            hoja, tg, float(cx) + clr, float(cy) + clr, margen=1.2
        )
    # Preferir el lado del sheet con más aire respecto al centro del barreno.
    minx, maxx, miny, maxy = pieza_bb
    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        aire = {
            "der": sheet_w - maxx,
            "izq": minx,
            "sup": sheet_h - maxy,
            "inf": miny,
        }
        lado = max(aire, key=aire.get)
    except Exception:
        lado = "der"
    x, y = empujar_punto_fuera_bbox(cx, cy, pieza_bb, lado, clr)
    return _clampear_punto_hoja(
        hoja, tg, x, y, margen=1.2, evitar_bbox=pieza_bb
    )


def _aplicar_estilo_y_fuera_pieza(dim, hoja, tg, vista):
    """Estilo navy + forzar texto fuera de la silueta."""
    aplicar_estilo_cota(dim, hoja=hoja)
    pieza_bb = _bbox_pieza_vista(vista)
    if pieza_bb is not None:
        ok = asegurar_cota_fuera_pieza_robusto(
            dim, tg, pieza_bb, holgura=0.6, n_chars=12
        )
        if not ok:
            print(
                "  AVISO: cota Ø/arco aún cerca de la pieza tras reposicionar."
            )
    return dim


# Curve2dTypeEnum (DrawingCurve.CurveType)
kCircularArcCurve2d = 5121
kCircleCurve2d = 5122
# En HLR de chapa GIGA, Inventor suele proyectar barrenos así:
kEllipseFullCurve2d = 5124
kEllipticalArcCurve2d = 5125

_TIPOS_CIRCULO = (kCircleCurve2d, kEllipseFullCurve2d)
_TIPOS_ARCO = (kCircularArcCurve2d, kEllipticalArcCurve2d)


def _silueta_vista(vista, solo_lineas=False):
    """
    BBox de curvas de la vista → (minx, maxx, miny, maxy, span).

    ``solo_lineas=True``: excluye círculos/arcos/elipses (barrenos). Así el
    origen IL es la placa, no el cluster de agujeros (evita cotas "paso"
    entre barrenos tipo 38.506 en vez de borde→centro).
    """
    minx = miny = None
    maxx = maxy = None
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return None
    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            if solo_lineas:
                try:
                    ct = int(curva.CurveType)
                except Exception:
                    ct = None
                if ct in _TIPOS_CIRCULO or ct in _TIPOS_ARCO:
                    continue
            caja = curva.Evaluator2D.RangeBox
            x0, x1 = float(caja.MinPoint.X), float(caja.MaxPoint.X)
            y0, y1 = float(caja.MinPoint.Y), float(caja.MaxPoint.Y)
            minx = x0 if minx is None else min(minx, x0)
            maxx = x1 if maxx is None else max(maxx, x1)
            miny = y0 if miny is None else min(miny, y0)
            maxy = y1 if maxy is None else max(maxy, y1)
        except Exception:
            continue
    if minx is None:
        return None
    span = max(maxx - minx, maxy - miny)
    return minx, maxx, miny, maxy, span


def _silueta_placa_vista(vista):
    """Silueta de PLACA (sin barrenos); fallback a bbox total si no hay líneas."""
    sil = _silueta_vista(vista, solo_lineas=True)
    if sil is not None:
        return sil
    return _silueta_vista(vista, solo_lineas=False)


def _arco_casi_cerrado(curva, tam_caja):
    """
    True si el arco/elipse es un barreno cerrado (no radio de doblez ~90°).
    """
    try:
        ct = int(curva.CurveType)
    except Exception:
        return True
    if ct in _TIPOS_CIRCULO:
        return True
    if ct not in _TIPOS_ARCO:
        return False
    try:
        sp = curva.StartPoint
        ep = curva.EndPoint
        dx = float(sp.X) - float(ep.X)
        dy = float(sp.Y) - float(ep.Y)
        gap = (dx * dx + dy * dy) ** 0.5
    except Exception:
        # Sin extremos: full circle/ellipse = cerrado; arco elíptico = abierto.
        return ct in _TIPOS_CIRCULO
    if tam_caja <= 1e-9:
        return False
    # Umbral RELATIVO al tamaño: un semicírculo de slot pequeño (~0.10 cm
    # hoja a escala 0.1) tiene gap≈D y antes caía bajo max(0.12,...) → se
    # descartaba como "cerrado" y se perdían TODOS los óvalos.
    return gap <= max(0.015, tam_caja * 0.18)


def _min_tam_hoja(vista, piso=0.05):
    """Umbral mínimo en cm de hoja (~3 mm de modelo × escala)."""
    try:
        sc = abs(float(vista.Scale))
        if sc > 1e-9:
            # A escala baja (0.1) un slot Ø~0.5" mide ~0.05 cm: no subir el piso.
            return max(min(float(piso), 0.04), 0.25 * sc)
    except Exception:
        pass
    return float(piso)


def _centro_interior_silueta(cx, cy, minx, maxx, miny, maxy, margen_frac=0.10):
    """Rechaza centros cerca del borde (típico de radio de doblez en esquina)."""
    span = max(maxx - minx, maxy - miny)
    if span <= 1e-9:
        return False
    m = span * float(margen_frac)
    return (minx + m) <= cx <= (maxx - m) and (miny + m) <= cy <= (maxy - m)


def _anillos_en_vista(vista, min_tam=None, max_frac=0.55):
    """
    Círculos/elipses cerradas (barrenos redondos).

    Incluye CurveType 5122 (circle) y 5124 (ellipse full) — típicos en HLR GIGA.
    Excluye radios de doblez (arcos abiertos / centros en esquina).
    """
    anillos = []
    sil = _silueta_vista(vista)
    if not sil:
        return anillos
    minx, maxx, miny, maxy, span = sil
    tope = span * max_frac if span > 0 else 1e9
    if min_tam is None:
        min_tam = _min_tam_hoja(vista)

    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return anillos

    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            try:
                ct = int(curva.CurveType)
            except Exception:
                ct = None
            # Preferir tipos circulares/elípticos; si no hay tipo, caer a bbox.
            if ct is not None and ct not in _TIPOS_CIRCULO and ct not in _TIPOS_ARCO:
                continue
            caja = curva.Evaluator2D.RangeBox
            ancho = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
            alto = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
            if ancho < min_tam or alto < min_tam:
                continue
            if abs(ancho - alto) > max(ancho, alto) * 0.22:
                continue
            tam = (ancho + alto) * 0.5
            if tam > tope:
                continue
            cx = (float(caja.MaxPoint.X) + float(caja.MinPoint.X)) / 2.0
            cy = (float(caja.MaxPoint.Y) + float(caja.MinPoint.Y)) / 2.0
            if ct in _TIPOS_ARCO and not _arco_casi_cerrado(curva, tam):
                continue
            if ct in _TIPOS_CIRCULO:
                pass  # elipse/círculo full: OK
            elif ct is None and not _arco_casi_cerrado(curva, tam):
                continue
            if not _centro_interior_silueta(
                cx, cy, minx, maxx, miny, maxy, margen_frac=0.03
            ):
                continue
            anillos.append(
                {
                    "curva": curva,
                    "tamaño": tam,
                    "cx": cx,
                    "cy": cy,
                    "tipo": "circulo",
                }
            )
        except Exception:
            continue
    return anillos


def _arcos_extremos_ranura(vista, min_tam=None, max_frac=0.45):
    """
    Arcos ~semicirculares / elípticos (extremos de barreno ovalado).

    Tipos 5121/5125. Bbox de semicírculo ≈ D × D/2 (aspecto ~2).
    """
    sil = _silueta_vista(vista)
    if not sil:
        return []
    minx, maxx, miny, maxy, span = sil
    tope = span * max_frac if span > 0 else 1e9
    if min_tam is None:
        min_tam = _min_tam_hoja(vista, piso=0.04)
    extremos = []
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return extremos

    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            try:
                ct = int(curva.CurveType)
            except Exception:
                continue
            if ct not in _TIPOS_ARCO:
                continue
            caja = curva.Evaluator2D.RangeBox
            ancho = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
            alto = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
            major = max(ancho, alto)
            minor = min(ancho, alto)
            if major < min_tam or minor < min_tam * 0.30:
                continue
            if major > tope:
                continue
            if minor <= 1e-9 or major / minor < 1.25 or major / minor > 4.5:
                continue
            diam = major
            cx = (float(caja.MaxPoint.X) + float(caja.MinPoint.X)) / 2.0
            cy = (float(caja.MaxPoint.Y) + float(caja.MinPoint.Y)) / 2.0
            if not _centro_interior_silueta(
                cx, cy, minx, maxx, miny, maxy, margen_frac=0.02
            ):
                continue
            if _arco_casi_cerrado(curva, major):
                continue
            extremos.append(
                {
                    "curva": curva,
                    "diam": diam,
                    "cx": cx,
                    "cy": cy,
                    "eje": "H" if ancho >= alto else "V",
                }
            )
        except Exception:
            continue
    return extremos


def _emparejar_ranuras(extremos, tol_diam_frac=0.12):
    """
    Empareja dos extremos semicirculares del mismo Ø → una ranura ovalada.

    ``tamaño`` = diámetro menor del slot (ancho libre del óvalo).
    """
    ranuras = []
    usados = set()
    n = len(extremos)
    for i in range(n):
        if i in usados:
            continue
        a = extremos[i]
        mejor_j = None
        mejor_sep = None
        for j in range(i + 1, n):
            if j in usados:
                continue
            b = extremos[j]
            if abs(a["diam"] - b["diam"]) > max(0.06, a["diam"] * tol_diam_frac):
                continue
            dx = abs(float(a["cx"]) - float(b["cx"]))
            dy = abs(float(a["cy"]) - float(b["cy"]))
            diam = (a["diam"] + b["diam"]) * 0.5
            # Slot horizontal: centros alineados en Y, separados en X.
            if dy <= diam * 0.40 and dx >= diam * 0.45:
                sep = dx
            # Slot vertical: alineados en X, separados en Y.
            elif dx <= diam * 0.40 and dy >= diam * 0.45:
                sep = dy
            else:
                continue
            # Evitar emparejar extremos lejanos (filas distintas):
            # largo típico de ranura < ~4·Ø.
            if sep > max(diam * 4.0, 0.35):
                continue
            if mejor_j is None or sep < mejor_sep:
                mejor_j = j
                mejor_sep = sep
        if mejor_j is None:
            continue
        b = extremos[mejor_j]
        usados.add(i)
        usados.add(mejor_j)
        diam = (a["diam"] + b["diam"]) * 0.5
        cx_a, cy_a = float(a["cx"]), float(a["cy"])
        cx_b, cy_b = float(b["cx"]), float(b["cy"])
        ranuras.append(
            {
                "curva": a["curva"],  # AddDiameter sobre un extremo
                "curva_b": b["curva"],
                "tamaño": diam,
                "cx": (cx_a + cx_b) * 0.5,
                "cy": (cy_a + cy_b) * 0.5,
                "cx_a": cx_a,
                "cy_a": cy_a,
                "cx_b": cx_b,
                "cy_b": cy_b,
                "tipo": "oval",
                "largo": float(mejor_sep) + diam,
            }
        )
    return ranuras


def _ranuras_en_vista(vista, min_tam=None, max_frac=0.45):
    """Barrenos ovalados (slots) por par de extremos semicirculares."""
    return _emparejar_ranuras(
        _arcos_extremos_ranura(vista, min_tam=min_tam, max_frac=max_frac)
    )


def _barrenos_en_vista(vista):
    """Círculos + óvalos/ranuras unificados para HOLE##."""
    return list(_anillos_en_vista(vista)) + list(_ranuras_en_vista(vista))


def _agrupar_diametros_unicos(anillos, tol_frac=0.06, max_grupos=8):
    grupos = _agrupar_diametros_grupos(anillos, tol_frac=tol_frac, max_grupos=max_grupos)
    return [max(g, key=lambda x: x["tamaño"]) for g in grupos]


def _agrupar_diametros_grupos(anillos, tol_frac=0.03, max_grupos=12):
    """Grupos de barrenos por Ø (círculo) o ancho de oval — sin mezclar tipos."""
    if not anillos:
        return []
    ordenados = sorted(anillos, key=lambda a: a["tamaño"], reverse=True)
    grupos = []
    for a in ordenados:
        tipo_a = str(a.get("tipo") or "circulo")
        colocado = False
        for g in grupos:
            ref = g[0]
            tipo_g = str(ref.get("tipo") or "circulo")
            if tipo_a != tipo_g:
                continue
            rt = float(ref["tamaño"])
            if abs(float(a["tamaño"]) - rt) <= max(0.004, rt * tol_frac):
                g.append(a)
                colocado = True
                break
        if not colocado:
            grupos.append([a])
        if len(grupos) >= max_grupos:
            break
    return grupos


def _marcar_barrenos_azules(hoja, vista, anillos_grupo, tg, inv_app):
    """
    Círculos azules (estilo TYP) sobre cada barreno del mismo Ø en la hoja HOLE.
    """
    if not anillos_grupo:
        return
    try:
        sketch = hoja.Sketches.Add()
    except Exception:
        try:
            sketch = hoja.DrawingSketches.Add()
        except Exception:
            return
    try:
        color = inv_app.TransientObjects.CreateColor(0, 0, 220)
    except Exception:
        color = None
    offsets = (-0.03, 0.0, 0.03)
    for a in anillos_grupo:
        try:
            cx = float(a["cx"])
            cy = float(a["cy"])
            radio_base = max(0.12, float(a["tamaño"]) * 0.55)
        except Exception:
            continue
        for delta in offsets:
            r = radio_base + delta
            if r <= 0.05:
                continue
            try:
                circ = sketch.SketchCircles.AddByCenterRadius(
                    tg.CreatePoint2d(cx, cy), r
                )
                if color is not None:
                    try:
                        circ.OverrideColor = color
                    except Exception:
                        pass
                try:
                    circ.LineWeight = 0.05
                except Exception:
                    pass
            except Exception:
                continue
    try:
        sketch.Visible = True
    except Exception:
        pass


def acotar_barrenos_placas(nombres_frente_ok=None):
    """
    Crea hojas ``*_DIAMETRO_Hnn`` — UNA por cada tamaño distinto de barreno
    (círculo Ø o ancho menor de óvalo/ranura).

    Ej.: todos Ø11 → 1 hoja; Ø11 + oval 13 → 2 hojas (H01 y H02).
    Si el grupo tiene ≥2 ocurrencias, marca azules (TYP visual).

    Acepta FRENTE doblado y DESPLIEGUE_FRENTE_* (flat Corte).
    """
    print("diametro.py: barrenos por TIPO de tamaño (circulos + ovalos)...")
    inv_app = conectar_inventor()
    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except Exception:
        print("Error: No hay un plano de Inventor abierto.")
        return []

    tg = inv_app.TransientGeometry
    objetivo = None
    if nombres_frente_ok is not None:
        objetivo = {str(h).upper().rsplit(":", 1)[0] for h in nombres_frente_ok}

    creadas = 0
    creadas_nombres = []
    for i in range(1, plano.Sheets.Count + 1):
        try:
            hoja = plano.Sheets.Item(i)
        except Exception:
            continue
        nombre = str(hoja.Name)
        nombre_up = nombre.upper()
        es_flat_frente = (
            "_DESPLIEGUE_" in nombre_up
            and ("_FRENTE_1" in nombre_up or "_FRENTE_2" in nombre_up)
        )
        es_frente_doblado = (
            "_DESPLIEGUE_" not in nombre_up
            and ("_FRENTE_1" in nombre_up or "_FRENTE_2" in nombre_up)
        )
        if not (es_flat_frente or es_frente_doblado):
            continue
        # No re-procesar hojas ya de diametro/xy
        if "_DIAMETRO_" in nombre_up or "_XCENTRO" in nombre_up or "_YCENTRO" in nombre_up:
            continue
        base_cmp = nombre_up.rsplit(":", 1)[0]
        if objetivo is not None and base_cmp not in objetivo:
            continue
        if hoja.DrawingViews.Count < 1:
            continue
        vista = hoja.DrawingViews.Item(1)
        grupos = _agrupar_diametros_grupos(_barrenos_en_vista(vista))
        if not grupos:
            continue

        base_nombre = nombre.rsplit(":", 1)[0]
        base_nombre = re.sub(
            r"_FRENTE_[12]$", "", base_nombre, flags=re.IGNORECASE
        )
        marcar_typ = True  # azules si el grupo tiene ≥2 (abajo)

        print(
            f"  {base_cmp}: {len(grupos)} tipo(s) de tamaño "
            f"({sum(len(g) for g in grupos)} barrenos)"
        )

        for idx, grupo in enumerate(grupos, start=1):
            anillo = max(grupo, key=lambda x: x["tamaño"])
            nombre_nueva = f"{base_nombre}_DIAMETRO_H{idx:02d}"
            try:
                for j in range(plano.Sheets.Count, 0, -1):
                    try:
                        h = plano.Sheets.Item(j)
                        if str(h.Name).upper().startswith(nombre_nueva.upper()):
                            h.Delete()
                    except Exception:
                        continue
            except Exception:
                pass

            try:
                nueva = hoja.CopyTo(plano)
            except Exception as e:
                print(f"  {nombre_nueva}: CopyTo fallo ({e})")
                continue
            try:
                nueva.Name = nombre_nueva
            except Exception:
                pass
            try:
                nueva.Activate()
            except Exception:
                pass
            try:
                dims = nueva.DrawingDimensions.GeneralDimensions
                for k in range(dims.Count, 0, -1):
                    try:
                        dims.Item(k).Delete()
                    except Exception:
                        pass
            except Exception:
                pass

            if nueva.DrawingViews.Count < 1:
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            vista_n = nueva.DrawingViews.Item(1)
            anillos_n = _barrenos_en_vista(vista_n)
            if not anillos_n:
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            # Reagrupar en la hoja clonada: mismo TIPO + mismo Ø (tol estricta).
            # Antes tol 0.08 mezclaba circulo 0.110 con oval 0.103 → dos hojas
            # con el mismo Ø 11.000.
            tipo_objetivo = str(anillo.get("tipo") or "circulo")
            grupos_n = _agrupar_diametros_grupos(anillos_n)
            grupo_n = None
            for g in grupos_n:
                ref = max(g, key=lambda x: x["tamaño"])
                if str(ref.get("tipo") or "circulo") != tipo_objetivo:
                    continue
                if abs(float(ref["tamaño"]) - float(anillo["tamaño"])) <= max(
                    0.004, float(anillo["tamaño"]) * 0.03
                ):
                    grupo_n = g
                    break
            if grupo_n is None:
                # Fallback: filtrar por tipo en la lista plana
                mismos = [
                    a
                    for a in anillos_n
                    if str(a.get("tipo") or "circulo") == tipo_objetivo
                ]
                if mismos:
                    grupo_n = mismos
                else:
                    try:
                        nueva.Delete()
                    except Exception:
                        pass
                    continue
            # Preferir miembro del mismo tipo; para oval el tamaño = ancho menor
            objetivo_a = max(grupo_n, key=lambda x: x["tamaño"])
            if tipo_objetivo == "oval":
                # Asegurar cota del ancho del slot (no un circulo colado)
                ovals = [
                    a
                    for a in grupo_n
                    if str(a.get("tipo") or "") == "oval"
                ]
                if ovals:
                    objetivo_a = max(ovals, key=lambda x: x["tamaño"])
            dim = None
            try:
                intent = nueva.CreateGeometryIntent(objetivo_a["curva"])
                pt = _punto_texto_barreno(
                    nueva,
                    tg,
                    vista_n,
                    objetivo_a["cx"],
                    objetivo_a["cy"],
                    objetivo_a["tamaño"],
                )
                dim = nueva.DrawingDimensions.GeneralDimensions.AddDiameter(
                    pt, intent
                )
            except Exception as e_diam:
                # Elipses HLR (5124/5125) a veces rechazan AddDiameter →
                # cota lineal del diámetro/ancho del barreno.
                try:
                    curva = objetivo_a["curva"]
                    caja = curva.Evaluator2D.RangeBox
                    x0, x1 = float(caja.MinPoint.X), float(caja.MaxPoint.X)
                    y0, y1 = float(caja.MinPoint.Y), float(caja.MaxPoint.Y)
                    if (x1 - x0) >= (y1 - y0):
                        p1 = tg.CreatePoint2d(x0, (y0 + y1) * 0.5)
                        p2 = tg.CreatePoint2d(x1, (y0 + y1) * 0.5)
                        pt = _punto_texto_barreno(
                            nueva,
                            tg,
                            vista_n,
                            (x0 + x1) * 0.5,
                            (y0 + y1) * 0.5,
                            objetivo_a["tamaño"],
                        )
                        dim = (
                            nueva.DrawingDimensions.GeneralDimensions.AddLinear(
                                pt, p1, p2
                            )
                        )
                    else:
                        p1 = tg.CreatePoint2d((x0 + x1) * 0.5, y0)
                        p2 = tg.CreatePoint2d((x0 + x1) * 0.5, y1)
                        pt = _punto_texto_barreno(
                            nueva,
                            tg,
                            vista_n,
                            (x0 + x1) * 0.5,
                            (y0 + y1) * 0.5,
                            objetivo_a["tamaño"],
                        )
                        dim = (
                            nueva.DrawingDimensions.GeneralDimensions.AddLinear(
                                pt, p1, p2
                            )
                        )
                    print(
                        f"  ↩️ {nombre_nueva}: AddDiameter no aplicó "
                        f"({e_diam}); cota lineal del Ø/ancho."
                    )
                except Exception as e_lin:
                    print(
                        f"⚠️ {nombre_nueva}: no se pudo cotar barreno "
                        f"(Ø={e_diam}; lineal={e_lin})"
                    )
                    try:
                        nueva.Delete()
                    except Exception:
                        pass
                    continue
            try:
                _aplicar_estilo_y_fuera_pieza(dim, nueva, tg, vista_n)
                # Forzar texto con conversion GIGA (pulg 3dec → mm), no ModelValue*10 crudo
                try:
                    from cota_estilo import (
                        texto_cota_dibujo,
                        COTA_FONT_SIZE_CM,
                        COTA_BOLD,
                    )

                    mv = float(dim.ModelValue)
                    esc = abs(float(vista_n.Scale)) or 1.0
                    tam_mod = float(objetivo_a.get("tamaño") or 0) / esc
                    if str(objetivo_a.get("tipo") or "") == "oval" and tam_mod > 1e-9:
                        # Ancho menor del slot (modelo cm)
                        if abs(mv - tam_mod) > max(0.05, tam_mod * 0.2):
                            mv = tam_mod
                    txt = texto_cota_dibujo(mv, nueva)
                    if txt:
                        dim.HideValue = True
                        bold = "True" if COTA_BOLD else "False"
                        dim.Text.FormattedText = (
                            f"<StyleOverride FontSize='{COTA_FONT_SIZE_CM}' "
                            f"Bold='{bold}'>{txt}</StyleOverride>"
                        )
                except Exception:
                    pass
                if marcar_typ and len(grupo_n) >= 2:
                    try:
                        _marcar_barrenos_azules(
                            nueva, vista_n, grupo_n, tg, inv_app
                        )
                        print(
                            f"  🔵 {nombre_nueva}: {len(grupo_n)} barrenos "
                            f"marcados (multi-Ø)"
                        )
                    except Exception as e_m:
                        print(f"  AVISO marca barrenos: {e_m}")
                try:
                    from cota_estilo import texto_cota_dibujo

                    txt_d = texto_cota_dibujo(objetivo_a["tamaño"])
                except Exception:
                    txt_d = f"{objetivo_a['tamaño'] / 2.54:.3f} in"
                tipo = str(objetivo_a.get("tipo") or "circulo")
                tag = "oval" if tipo == "oval" else "Ø"
                print(f"✅ {nombre_nueva}: barreno {tag} {txt_d}")
                creadas += 1
                creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])
            except Exception as e:
                print(f"⚠️ {nombre_nueva}: estilo/export barreno falló ({e})")
                try:
                    nueva.Delete()
                except Exception:
                    pass

    print(f"✅ diametro.py barrenos: {creadas} hojas DIAMETRO_H* creadas")
    return creadas_nombres


def acotar_diametros(hojas_pendientes=None):
    print("⭕ diametro.py: Iniciando escáner de límites (Ext/Int)...")
    inv_app = conectar_inventor()

    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, 'DrawingDocument')
    except:
        print("❌ Error: No hay un plano de Inventor abierto.")
        return list(hojas_pendientes) if hojas_pendientes is not None else []

    tg = inv_app.TransientGeometry
    procesadas = 0
    procesadas_nombres = set()

    hojas_objetivo = None
    objetivo_set = None

    if hojas_pendientes is not None:
        hojas_objetivo = [str(h).upper() for h in hojas_pendientes]
        objetivo_set = set(hojas_objetivo)

    # Nombres de hojas _FRENTE_2 a ELIMINAR al final del loop porque la pieza
    # cilíndrica no tiene borde interior concéntrico (es una barra/pin/stud
    # sólido) y no tiene sentido exportar un JPG con solo el contorno
    # exterior duplicado.
    hojas_a_eliminar = []

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja_original = str(hoja.Name)
        nombre_hoja = nombre_hoja_original.upper()

        if objetivo_set is not None and nombre_hoja not in objetivo_set:
            continue

        if "_FRENTE" not in nombre_hoja:
            continue

        if hoja.DrawingViews.Count == 0:
            continue

        vista = hoja.DrawingViews.Item(1)
        anillos = _anillos_en_vista(vista, min_tam=0.1, max_frac=0.98)

        if not anillos:
            continue

        try:
            anillo_objetivo = None
            etiqueta = ""

            anillos_validos = [a for a in anillos if a['tamaño'] > 0.3]

            if "_FRENTE_1" in nombre_hoja:
                if anillos_validos:
                    anillo_objetivo = max(anillos_validos, key=lambda x: x['tamaño'])
                    etiqueta = "EXTERIOR"

            elif "_FRENTE_2" in nombre_hoja:
                if anillos_validos:
                    anillo_exterior = max(anillos_validos, key=lambda x: x['tamaño'])

                    tolerancia_centro = max(0.15, anillo_exterior['tamaño'] * 0.05)

                    interiores_concentricos = []
                    for a in anillos_validos:
                        if a is anillo_exterior:
                            continue

                        dx_c = abs(a['cx'] - anillo_exterior['cx'])
                        dy_c = abs(a['cy'] - anillo_exterior['cy'])

                        if dx_c <= tolerancia_centro and dy_c <= tolerancia_centro and a['tamaño'] < anillo_exterior['tamaño']:
                            interiores_concentricos.append(a)

                    if interiores_concentricos:
                        # Para _FRENTE_2 queremos el círculo MÁS PEQUEÑO
                        # del mismo centro, o sea el límite interior real.
                        anillo_objetivo = min(interiores_concentricos, key=lambda x: x['tamaño'])
                        etiqueta = "INTERIOR"
                    else:
                        # Sin interior concéntrico → pieza cilíndrica SÓLIDA
                        # (barra, pin, stud). No tiene diámetro interior,
                        # así que eliminamos la hoja para que no aparezca un
                        # JPG mudo con solo el contorno exterior.
                        print(
                            f"🗑️ {nombre_hoja}: pieza cilíndrica sólida sin "
                            f"interior concéntrico; se elimina la hoja "
                            f"_FRENTE_2 (no aplica DIAMETRO_INTERIOR)."
                        )
                        hojas_a_eliminar.append(nombre_hoja_original)
                        procesadas_nombres.add(nombre_hoja)
                        continue

            if anillo_objetivo:
                intencion = hoja.CreateGeometryIntent(anillo_objetivo['curva'])
                punto_texto = _punto_texto_barreno(
                    hoja,
                    tg,
                    vista,
                    anillo_objetivo['cx'],
                    anillo_objetivo['cy'],
                    anillo_objetivo['tamaño'],
                )

                dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
                _aplicar_estilo_y_fuera_pieza(dim, hoja, tg, vista)
                print(f"✅ {nombre_hoja}: Límite {etiqueta} (Ø {anillo_objetivo['tamaño']:.2f})")
                procesadas += 1
                procesadas_nombres.add(nombre_hoja)

        except Exception:
            print(f"⚠️ {nombre_hoja}: Inventor rechazó el diámetro.")

    # Eliminar hojas _FRENTE_2 de piezas cilíndricas sólidas. Iteramos por
    # nombre y de atrás hacia adelante en el arreglo de sheets para no
    # invalidar índices.
    for nombre_borrar in hojas_a_eliminar:
        try:
            for j in range(plano.Sheets.Count, 0, -1):
                try:
                    if str(plano.Sheets.Item(j).Name) == nombre_borrar:
                        plano.Sheets.Item(j).Delete()
                        break
                except Exception:
                    continue
        except Exception as e:
            print(f"⚠️ No se pudo eliminar la hoja '{nombre_borrar}': {e}")

    print(f"✅ diametro.py finalizado. Piezas acotadas: {procesadas}")
    if hojas_a_eliminar:
        print(f"🗑️ Hojas _FRENTE_2 eliminadas por ser sólidas: {len(hojas_a_eliminar)}")
        # Publicar la lista de piezas sólidas para que el flujo de exportación
        # borre cualquier JPG _DIAMETRO_INTERIOR residual de corridas anteriores.
        try:
            _publicar_piezas_solidas(hojas_a_eliminar)
        except Exception as pub_err:
            print(f"AVISO: no se pudo publicar piezas sólidas: {pub_err}")

    if hojas_objetivo is not None:
        no_resueltas = [h for h in hojas_objetivo if h not in procesadas_nombres]
        return no_resueltas

    return []


if __name__ == "__main__":
    acotar_diametros()