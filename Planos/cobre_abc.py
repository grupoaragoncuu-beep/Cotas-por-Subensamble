# -*- coding: utf-8 -*-
"""
Hojas combinadas de cobre GIGA.

Con dobleces: ISO (letras de zona), DOBLADO y CORTE.
Sin dobleces: solo CORTE. No publica a la base de datos.
"""
from __future__ import annotations

import math
import os
import shutil
import sys
import time

_PLANOS = r"C:\Proyectos\COTAS\Planos"
if _PLANOS not in sys.path:
    sys.path.insert(0, _PLANOS)

import pythoncom
import win32com.client

MACHOTE = r"C:\Proyectos\COTAS\Planos\MACHOTE PLANOS.dwg"
TMP = r"C:\Proyectos\COTAS\.runtime\_cobre_abc"
K_ARB = 10763
K_HLR = 32258
K_HOR = 60162
K_VER = 60163
K_CIRC = {"izq": 57862, "der": 57863, "sup": 57864, "inf": 57865}

_ULTIMO_LINEAL = ""
_PLANO = None
_CAPA_SOLIDA = None
_ESCALA = 1.0
_TRAZOS = []
K_CONTINUOUS = 37633


def _clave(nombre):
    n = str(nombre or "").upper()
    if ":" in n:
        n = n.split(":", 1)[0]
    if n.endswith(".IPT"):
        n = n[:-4]
    return n.strip()


def _ensamble_de(raw):
    try:
        if int(raw.DocumentType) != 12291:
            return None
        if "9919-BOARD 1" not in str(raw.DisplayName).upper():
            return None
    except Exception:
        return None
    refs = raw.AllReferencedDocuments
    for j in range(1, int(refs.Count) + 1):
        doc = refs.Item(j)
        try:
            if int(doc.DocumentType) != 12290:
                continue
        except Exception:
            continue
        if _clave(doc.DisplayName) == STEM:
            return win32com.client.CastTo(doc, "PartDocument")
    return None


def _part(inv):
    """La pieza sale del ensamble abierto. No se toca ningún plano ya cargado."""
    try:
        pieza = _ensamble_de(inv.ActiveDocument)
    except Exception:
        pieza = None
    if pieza is not None:
        return pieza
    for i in range(1, int(inv.Documents.Count) + 1):
        pieza = _ensamble_de(inv.Documents.Item(i))
        if pieza is not None:
            return pieza
    for i in range(1, int(inv.Documents.Count) + 1):
        raw = inv.Documents.Item(i)
        try:
            if int(raw.DocumentType) != 12290:
                continue
        except Exception:
            continue
        if _clave(raw.DisplayName) == STEM:
            return win32com.client.CastTo(raw, "PartDocument")
    return None


def _huecos_cara(face):
    n = 0
    try:
        loops = int(face.EdgeLoops.Count)
    except Exception:
        return 0
    for li in range(1, loops + 1):
        try:
            loop = face.EdgeLoops.Item(li)
            if bool(getattr(loop, "IsOuterEdgeLoop", True)):
                continue
            n += 1
        except Exception:
            continue
    return n


def _grupos(part):
    sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
    thk = abs(float(sm.Thickness.Value)) * 10.0
    body = sm.SurfaceBodies.Item(1)
    caras = []
    for i in range(1, int(body.Faces.Count) + 1):
        face = body.Faces.Item(i)
        try:
            if int(face.SurfaceType) != 5890:
                continue
            area = float(face.Evaluator.Area)
        except Exception:
            continue
        if area < 8.0:
            continue
        box = face.Evaluator.RangeBox
        ext = sorted(
            abs(float(getattr(box.MaxPoint, eje)) - float(getattr(box.MinPoint, eje))) * 10.0
            for eje in ("X", "Y", "Z")
        )
        # El canto del espesor no es una zona de doblez.
        if ext[1] <= max(thk * 1.4, thk + 1.5):
            continue
        nx = ny = nz = 0.0
        try:
            normal = face.Geometry.Normal
            nx, ny, nz = float(normal.X), float(normal.Y), float(normal.Z)
        except Exception:
            pass
        caras.append(
            {
                "cx": (float(box.MinPoint.X) + float(box.MaxPoint.X)) * 5.0,
                "cy": (float(box.MinPoint.Y) + float(box.MaxPoint.Y)) * 5.0,
                "cz": (float(box.MinPoint.Z) + float(box.MaxPoint.Z)) * 5.0,
                "area": area,
                "huecos": _huecos_cara(face),
                "face": face,
                "nx": nx,
                "ny": ny,
                "nz": nz,
            }
        )
    grupos = []
    for c in caras:
        puesto = False
        for g in grupos:
            if abs(c["cy"] - g["cy"]) >= 12 or abs(c["cx"] - g["cx"]) >= 12:
                continue
            g["huecos"] = max(g["huecos"], c["huecos"])
            cam_c = c["cx"] * 0.9 + c["cy"] * 0.55 + c["cz"] * 0.85
            cam_g = g["cx"] * 0.9 + g["cy"] * 0.55 + g["cz"] * 0.85
            if c["area"] > g["area"] + 1.0 or (c["area"] >= g["area"] - 1.0 and cam_c > cam_g):
                g["area"] = max(g["area"], c["area"])
                g["cx"], g["cy"], g["cz"] = c["cx"], c["cy"], c["cz"]
                g["face"] = c.get("face")
                g["nx"], g["ny"], g["nz"] = c.get("nx", 0), c.get("ny", 0), c.get("nz", 0)
            puesto = True
            break
        if not puesto:
            grupos.append(dict(c))
    grupos.sort(key=lambda g: g["area"], reverse=True)
    return grupos[:3]


def _letras(grupos):
    """A = ala con más barrenos, B = alma, C = la otra punta. Con menos caras, menos letras."""
    if len(grupos) >= 3:
        orden = sorted(grupos, key=lambda g: g["cy"])
        medio = orden[len(orden) // 2]
        puntas = [g for g in orden if g is not medio]
        puntas = sorted(puntas, key=lambda g: -g["huecos"])
        puntas[0]["letra"] = "A"
        medio["letra"] = "B"
        puntas[1]["letra"] = "C"
        return [puntas[0], medio, puntas[1]]
    if len(grupos) == 2:
        orden = sorted(grupos, key=lambda g: (-g["huecos"], g["cy"]))
        orden[0]["letra"] = "A"
        orden[1]["letra"] = "B"
        return orden
    if len(grupos) == 1:
        grupos[0]["letra"] = "A"
        return grupos
    return []


def _mm_hoja(delta):
    return abs(float(delta)) / max(float(_ESCALA), 1e-6) * 10.0


def _letra_seq(i):
    n = int(i)
    s = ""
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _letra_cerca(grupos, vista, tg, x, y):
    mejor = None
    for g in grupos or []:
        try:
            p = vista.ModelToSheetSpace(
                tg.CreatePoint(g["cx"] / 10.0, g["cy"] / 10.0, g["cz"] / 10.0)
            )
            d = math.hypot(float(p.X) - x, float(p.Y) - y)
        except Exception:
            continue
        if mejor is None or d < mejor[0]:
            mejor = (d, g.get("letra"))
    if mejor is None:
        return None
    return mejor[1]


def _espesor_mm(part):
    try:
        sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
        return abs(float(sm.Thickness.Value)) * 10.0
    except Exception:
        return None


def _vista_doblez(part):
    """Cámara a lo largo del eje de doblez. Con eje Z se conserva la vista de 711."""
    eje = (0.0, 0.0, 1.0)
    try:
        sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
        if int(sm.Bends.Count) > 0:
            b = sm.Bends.Item(1)
            ff = b.FrontFaces.Item(1).Geometry
            bf = b.BackFaces.Item(1).Geometry
            inner = ff if float(ff.Radius) <= float(bf.Radius) else bf
            eje = (
                float(inner.AxisVector.X),
                float(inner.AxisVector.Y),
                float(inner.AxisVector.Z),
            )
    except Exception as exc:
        print("  eje", exc)
    ax, ay, az = abs(eje[0]), abs(eje[1]), abs(eje[2])
    if az >= ax and az >= ay:
        return (0.0, 0.0, 1.0), (0.0, 1.0, 0.0)
    if ay >= ax:
        return (0.0, 1.0, 0.0), (0.0, 0.0, 1.0)
    return (1.0, 0.0, 0.0), (0.0, 1.0, 0.0)


def _tiene_dobleces(part):
    try:
        sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
        return int(sm.Bends.Count) > 0
    except Exception:
        return False


def _respirar(segundos=1.2):
    time.sleep(segundos)


def _abrir_copia(inv):
    """Copia limpia del machote. Nunca se reabre el dibujo que ya quedó en memoria."""
    copia = os.path.join(TMP, f"_machote_{time.time_ns()}.dwg")
    shutil.copy2(MACHOTE, copia)
    _respirar(0.8)
    plano = win32com.client.CastTo(inv.Documents.Open(copia, True), "DrawingDocument")
    print("  copia", os.path.basename(copia))
    return plano


def _cerrar_copias(inv, conservar=None):
    """Cierra las copias temporales del machote. No toca el ensamble."""
    conservar_fn = ""
    if conservar is not None:
        try:
            conservar_fn = str(conservar.FullFileName)
        except Exception:
            conservar_fn = ""
    pendientes = []
    for i in range(1, int(inv.Documents.Count) + 1):
        doc = inv.Documents.Item(i)
        try:
            if int(doc.DocumentType) != 12292:
                continue
            fn = str(doc.FullFileName)
        except Exception:
            continue
        if not os.path.basename(fn).lower().startswith("_machote_"):
            continue
        if conservar_fn and os.path.normcase(fn) == os.path.normcase(conservar_fn):
            continue
        pendientes.append(doc)
    for doc in pendientes:
        try:
            print("  cierro", os.path.basename(str(doc.FullFileName)))
            doc.Close(False)
        except Exception as exc:
            print("  close", exc)
        _respirar(0.5)


def _hoja(plano, base, nombre):
    _respirar(0.8)
    nueva = base.CopyTo(plano)
    nueva.Name = nombre
    from creador_vistas import _limpiar_border_y_titleblock

    _limpiar_border_y_titleblock(nueva)
    return nueva


def _vista(hoja, part, tg, to, cam, flat, ancho_cm, alto_cm, margen_x=7.0, margen_y=6.0):
    """Escala la pieza al área útil de la hoja, dejando margen para las cotas."""
    global _ESCALA
    opciones = to.CreateNameValueMap()
    try:
        opciones.Add("SheetMetalFoldedModel", not flat)
    except Exception:
        pass
    px = float(hoja.Width) * 0.50
    py = float(hoja.Height) * 0.52
    usable_w = float(hoja.Width) * 0.82 - margen_x
    usable_h = float(hoja.Height) * 0.72 - margen_y
    escala = min(
        usable_w / max(float(ancho_cm), 0.1),
        usable_h / max(float(alto_cm), 0.1),
    )
    # La 711 quedó cerca de 0.60. Subir más la escala agranda la pieza
    # y la letra se ve chica respecto a esas tres hojas.
    escala = max(0.05, min(escala, 0.72))
    _ESCALA = escala
    print(
        "  escala", round(escala, 3),
        "hoja", round(float(hoja.Width), 1), round(float(hoja.Height), 1),
        "pieza", round(float(ancho_cm), 1), round(float(alto_cm), 1),
    )
    return hoja.DrawingViews.AddBaseView(
        part,
        tg.CreatePoint2d(px, py),
        escala,
        K_ARB,
        K_HLR,
        "",
        cam,
        opciones,
    )


def _font():
    """Mitad del tamaño de cota de cobre, para estas tres hojas."""
    from cota_estilo import get_cota_font_size_cm

    return get_cota_font_size_cm() * 0.5


# Prueba 711: la X del centro va en #043D56 y la cota con sus líneas en #444444.
_COLOR_LINEA = None
_COLOR_MARCA = None
_ESTILO_COLOR = None


def _azul(inv):
    """Mismo azul de las letras del isométrico, no el marino de las cotas."""
    return inv.TransientObjects.CreateColor(0, 0, 250)


def _color_linea(inv):
    if _COLOR_LINEA is not None:
        return inv.TransientObjects.CreateColor(*_COLOR_LINEA)
    return _azul(inv)


def _color_marca(inv):
    if _COLOR_MARCA is not None:
        return inv.TransientObjects.CreateColor(*_COLOR_MARCA)
    return _azul(inv)


def _capa_solida(inv):
    """Línea de cota continua. Las extensiones se dibujan aparte, punteadas."""
    global _CAPA_SOLIDA
    if _CAPA_SOLIDA is not None or _PLANO is None:
        return _CAPA_SOLIDA
    try:
        layers = _PLANO.StylesManager.Layers
        try:
            raw = layers.Item("COTA_SOLIDA")
        except Exception:
            raw = layers.Item(1).Copy("COTA_SOLIDA")
        capa = win32com.client.CastTo(raw, "Layer")
        capa.LineType = K_CONTINUOUS
        try:
            capa.Color = _color_linea(inv)
            print(
                "  capa color",
                int(capa.Color.Red), int(capa.Color.Green), int(capa.Color.Blue),
            )
        except Exception as exc:
            print("  capa color", exc)
        _CAPA_SOLIDA = capa
    except Exception as exc:
        print("  capa", exc)
    return _CAPA_SOLIDA


def _linea_solida(dim, inv):
    capa = _capa_solida(inv)
    if capa is None or dim is None:
        return
    try:
        dim.Layer = capa
    except Exception as exc:
        print("  solida", exc)
    global _ESTILO_COLOR
    if _COLOR_LINEA is None or _PLANO is None:
        return
    try:
        if _ESTILO_COLOR is None:
            nombre = "COTA_COLOR_711"
            estilos = _PLANO.StylesManager.DimensionStyles
            try:
                copia = estilos.Item(nombre)
            except Exception:
                copia = dim.Style.Copy(nombre)
            copia = win32com.client.CastTo(copia, "DimensionStyle")
            copia.Color = inv.TransientObjects.CreateColor(*_COLOR_LINEA)
            _ESTILO_COLOR = copia
        dim.Style = _ESTILO_COLOR
    except Exception as exc:
        print("  estilo", exc)


def _capturar_extensiones(dim):
    """Guarda las extensiones (origen → cota) y las oculta. La línea del texto se queda."""
    try:
        lin = win32com.client.CastTo(dim, "LinearGeneralDimension")
    except Exception:
        return
    for attr, vis in (
        ("ExtensionLineOne", "ExtensionLineOneVisible"),
        ("ExtensionLineTwo", "ExtensionLineTwoVisible"),
    ):
        try:
            ext = getattr(lin, attr)
            p1, p2 = ext.StartPoint, ext.EndPoint
            _TRAZOS.append((float(p1.X), float(p1.Y), float(p2.X), float(p2.Y)))
            setattr(lin, vis, False)
        except Exception as exc:
            print("  ext", attr, exc)


def _volcar_trazos(hoja, tg, inv):
    """Extensiones punteadas, en un solo sketch por hoja."""
    global _TRAZOS
    if not _TRAZOS:
        return
    try:
        sketch = hoja.Sketches.Add()
        sketch.Edit()
        azul = _color_linea(inv)
        for x1, y1, x2, y2 in _TRAZOS:
            dx, dy = x2 - x1, y2 - y1
            largo = math.hypot(dx, dy)
            if largo < 0.05:
                continue
            ux, uy = dx / largo, dy / largo
            t = 0.0
            on = True
            while t < largo - 1e-6:
                paso = 0.28 if on else 0.16
                t2 = min(largo, t + paso)
                if on and t2 - t > 0.04:
                    linea = sketch.SketchLines.AddByTwoPoints(
                        tg.CreatePoint2d(x1 + ux * t, y1 + uy * t),
                        tg.CreatePoint2d(x1 + ux * t2, y1 + uy * t2),
                    )
                    try:
                        linea.OverrideColor = azul
                    except Exception:
                        pass
                t = t2
                on = not on
        sketch.Visible = True
        sketch.ExitEdit()
    except Exception as exc:
        print("  trazos", exc)
    _TRAZOS = []


def _centrar_en(caja, tg, cx, cy):
    """Deja el texto en el centro del globo."""
    try:
        caja.HorizontalJustification = 19970
        caja.VerticalJustification = 19973
        caja.Origin = tg.CreatePoint2d(cx, cy)
        return
    except Exception:
        pass
    try:
        w = float(caja.Width)
        h = float(caja.Height)
        caja.Origin = tg.CreatePoint2d(cx - w / 2.0, cy - h / 2.0)
    except Exception:
        pass


def _etiqueta_girada(hoja, tg, inv, x, y, texto):
    """Texto de cota vertical, leído de abajo hacia arriba, como las capturas."""
    try:
        sketch = hoja.Sketches.Add()
        sketch.Edit()
        caja = sketch.TextBoxes.AddFitted(tg.CreatePoint2d(x, y), texto)
        from cota_estilo import aplicar_estilo_texto_cota, armar_formatted_texto_cota

        aplicar_estilo_texto_cota(caja, texto, inv, vertical=False)
        font = _font()
        formado = armar_formatted_texto_cota(
            texto, font_cm=font, bold=True, vertical=False
        )
        if _COLOR_LINEA is not None:
            r, g, b = _COLOR_LINEA
            formado = formado.replace(
                "<StyleOverride ",
                f"<StyleOverride Color='{r},{g},{b}' ",
            )
        caja.FormattedText = formado
        caja.Rotation = math.pi / 2.0
        try:
            caja.Color = _color_linea(inv)
        except Exception:
            pass
        sketch.Visible = True
        sketch.ExitEdit()
    except Exception as exc:
        print("  etiq", exc)


_DEFS_GLOBO = {}


def _def_globo(inv, tg, texto):
    """Símbolo de dibujo: círculo con la letra centrada. El globo de pieza no existe aquí."""
    if texto in _DEFS_GLOBO:
        return _DEFS_GLOBO[texto]
    nombre = "GLOBO_" + "".join(ch if ch.isalnum() else "_" for ch in texto)
    defs = _PLANO.SketchedSymbolDefinitions
    try:
        defin = defs.Item(nombre)
    except Exception:
        defin = defs.Add(nombre)
        sketch = defin.Edit()
        radio = 0.40 if len(texto) <= 2 else max(0.72, 0.15 * len(texto))
        azul = _azul(inv)
        circ = sketch.SketchCircles.AddByCenterRadius(tg.CreatePoint2d(0.0, 0.0), radio)
        try:
            circ.OverrideColor = azul
        except Exception:
            pass
        from cota_estilo import armar_formatted_texto_cota

        font = 0.20 if len(texto) <= 2 else 0.16
        caja = sketch.TextBoxes.AddFitted(tg.CreatePoint2d(0.0, 0.0), texto)
        caja.FormattedText = armar_formatted_texto_cota(texto, font_cm=font, bold=True)
        caja.HorizontalJustification = 19969
        caja.VerticalJustification = 25601
        caja.Origin = tg.CreatePoint2d(0.0, 0.0)
        try:
            caja.Color = azul
        except Exception:
            pass
        defin.ExitEdit(True)
    _DEFS_GLOBO[texto] = defin
    return defin


def _ancla_cara_mm(g, grupos):
    """Punto de la cara, corrido hacia la punta libre para no caer en el doblez."""
    cx, cy, cz = g["cx"], g["cy"], g["cz"]
    otros = [o for o in grupos if o is not g]
    if not otros:
        return cx, cy, cz
    ox = sum(o["cx"] for o in otros) / len(otros)
    oy = sum(o["cy"] for o in otros) / len(otros)
    oz = sum(o["cz"] for o in otros) / len(otros)
    vx, vy, vz = cx - ox, cy - oy, cz - oz
    nx, ny, nz = g.get("nx") or 0.0, g.get("ny") or 0.0, g.get("nz") or 0.0
    dot = vx * nx + vy * ny + vz * nz
    vx, vy, vz = vx - dot * nx, vy - dot * ny, vz - dot * nz
    largo = math.sqrt(vx * vx + vy * vy + vz * vz)
    if largo < 8.0:
        return cx, cy, cz
    paso = min(55.0, largo * 0.45)
    return cx + vx / largo * paso, cy + vy / largo * paso, cz + vz / largo * paso


def _proyectar(vista, tg, xmm, ymm, zmm):
    p = vista.ModelToSheetSpace(tg.CreatePoint(xmm / 10.0, ymm / 10.0, zmm / 10.0))
    return float(p.X), float(p.Y)


def _aristas_cara(vista, face):
    try:
        dcs = vista.DrawingCurves(face)
        n = int(dcs.Count)
    except Exception:
        return []
    aristas = []
    for i in range(1, n + 1):
        c = dcs.Item(i)
        try:
            x1, y1 = float(c.StartPoint.X), float(c.StartPoint.Y)
            x2, y2 = float(c.EndPoint.X), float(c.EndPoint.Y)
            aristas.append(((x1 + x2) / 2.0, (y1 + y2) / 2.0, math.hypot(x2 - x1, y2 - y1)))
        except Exception:
            continue
    return aristas


def _flecha_en_cara(vista, tg, g, grupos):
    """Punta en la esquina del ala más lejos de las otras zonas."""
    try:
        cx, cy = _proyectar(vista, tg, g["cx"], g["cy"], g["cz"])
    except Exception:
        return None, None
    face = g.get("face")
    otros = [o for o in grupos if o is not g]
    if face is None or not otros:
        return cx, cy
    ox = sum(o["cx"] for o in otros) / len(otros)
    oy = sum(o["cy"] for o in otros) / len(otros)
    oz = sum(o["cz"] for o in otros) / len(otros)
    try:
        box = face.Evaluator.RangeBox
        xs = (float(box.MinPoint.X), float(box.MaxPoint.X))
        ys = (float(box.MinPoint.Y), float(box.MaxPoint.Y))
        zs = (float(box.MinPoint.Z), float(box.MaxPoint.Z))
    except Exception:
        return cx, cy
    esquinas = []
    for x in xs:
        for y in ys:
            for z in zs:
                px, py, pz = x * 10.0, y * 10.0, z * 10.0
                d3 = math.sqrt((px - ox) ** 2 + (py - oy) ** 2 + (pz - oz) ** 2)
                try:
                    sx, sy = _proyectar(vista, tg, px, py, pz)
                except Exception:
                    continue
                esquinas.append((d3, sx, sy, px, py, pz))
    if not esquinas:
        return cx, cy
    tope = max(e[0] for e in esquinas) * 0.90
    libres = [e for e in esquinas if e[0] >= tope] or esquinas
    otros_xy = []
    for o in otros:
        try:
            otros_xy.append(_proyectar(vista, tg, o["cx"], o["cy"], o["cz"]))
        except Exception:
            continue
    if otros_xy:
        def lejos(e):
            return min(math.hypot(e[1] - qx, e[2] - qy) for qx, qy in otros_xy)
        libre = max(libres, key=lejos)
    else:
        libre = max(libres, key=lambda e: e[0])
    px = g["cx"] + (libre[3] - g["cx"]) * 0.42
    py = g["cy"] + (libre[4] - g["cy"]) * 0.42
    pz = g["cz"] + (libre[5] - g["cz"]) * 0.42
    try:
        return _proyectar(vista, tg, px, py, pz)
    except Exception:
        return cx, cy


def _globos_iso(hoja, vista, tg, inv, grupos):
    """Globo junto a su cara. El líder sale de la arista de esa zona."""
    vcx = float(vista.Center.X)
    vcy = float(vista.Center.Y)
    hw = float(vista.Width) * 0.5
    hh = float(vista.Height) * 0.5
    puestos = []
    for g in grupos:
        fx, fy = _flecha_en_cara(vista, tg, g, grupos)
        if fx is None:
            continue
        letra = g.get("letra") or "A"
        dx, dy = fx - vcx, fy - vcy
        dist = math.hypot(dx, dy)
        if dist < 0.9:
            if letra == "C":
                dx, dy = -0.5, -1.0
            elif letra == "B":
                dx, dy = -1.0, 0.15
            else:
                dx, dy = 0.85, 0.7
            dist = math.hypot(dx, dy)
        ux, uy = dx / dist, dy / dist
        gx, gy = fx, fy
        for _ in range(14):
            if abs(gx - vcx) > hw + 1.15 or abs(gy - vcy) > hh + 1.15:
                break
            gx += ux * 0.4
            gy += uy * 0.4
        gx += ux * 0.9
        gy += uy * 0.9
        for px, py in puestos:
            if math.hypot(gx - px, gy - py) < 1.35:
                gx += -uy * 1.3
                gy += ux * 1.3
        puestos.append((gx, gy))
        print(f"  globo {letra} flecha {fx:.2f},{fy:.2f} texto {gx:.2f},{gy:.2f}")
        _globo(hoja, tg, inv, fx, fy, gx, gy, letra)


def _globo(hoja, tg, inv, x_flecha, y_flecha, x_globo, y_globo, texto, curva=None):
    """Símbolo con líder. La letra va centrada en el círculo del símbolo."""
    texto = str(texto)
    try:
        defin = _def_globo(inv, tg, texto)
        pts = inv.TransientObjects.CreateObjectCollection()
        # El primer punto es el centro del símbolo; el último, la punta.
        pts.Add(tg.CreatePoint2d(x_globo, y_globo))
        pts.Add(tg.CreatePoint2d(x_flecha, y_flecha))
        vacio = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, [])
        sim = hoja.SketchedSymbols.AddWithLeader(defin, pts, 0.0, 1.0, vacio, True, False)
        try:
            sim.Leader.ArrowheadType = 71941
        except Exception:
            pass
        print("  globo simbolo", texto)
    except Exception as exc:
        print("  globo simbolo", texto, exc)


def _texto(dim, inv, texto, vertical=False, diametro=False):
    from cota_estilo import armar_formatted_texto_cota, get_cota_font_size_cm

    font = _font()
    cuerpo = texto
    prefijo = ""
    if "=" in texto and diametro:
        prefijo, cuerpo = texto.split("=", 1)
        prefijo = prefijo.strip() + "="
        cuerpo = cuerpo.replace("Ø", "").strip()
    dim.HideValue = True
    formado = armar_formatted_texto_cota(
        cuerpo, font_cm=font, bold=True, vertical=vertical, simbolo_diametro=diametro
    )
    if prefijo:
        formado = (
            f"<StyleOverride FontSize='{font}' Bold='True'>{prefijo}</StyleOverride>"
            + formado
        )
    if _COLOR_LINEA is not None:
        r, g, b = _COLOR_LINEA
        formado = formado.replace(
            "<StyleOverride ",
            f"<StyleOverride Color='{r},{g},{b}' ",
        )
    dim.Text.FormattedText = formado
    try:
        dim.Text.Color = _color_linea(inv)
    except Exception:
        pass


def _mm(dim):
    try:
        return abs(float(dim.ModelValue)) * 10.0
    except Exception:
        return None


def _curvas(vista):
    out = []
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return out
    for i in range(1, n + 1):
        c = vista.DrawingCurves.Item(i)
        try:
            s, e = c.StartPoint, c.EndPoint
            x1, y1 = float(s.X), float(s.Y)
            x2, y2 = float(e.X), float(e.Y)
        except Exception:
            try:
                cen = c.CenterPoint
                x1 = x2 = float(cen.X)
                y1 = y2 = float(cen.Y)
            except Exception:
                continue
            out.append(
                {
                    "c": c,
                    "x1": x1,
                    "y1": y1,
                    "x2": x2,
                    "y2": y2,
                    "mx": x1,
                    "my": y1,
                    "arco": True,
                    "dx": 0.0,
                    "dy": 0.0,
                }
            )
            continue
        arco = False
        try:
            arco = c.CenterPoint is not None
        except Exception:
            arco = False
        out.append(
            {
                "c": c,
                "x1": x1,
                "y1": y1,
                "x2": x2,
                "y2": y2,
                "mx": (x1 + x2) * 0.5,
                "my": (y1 + y2) * 0.5,
                "arco": arco,
                "dx": abs(x2 - x1),
                "dy": abs(y2 - y1),
            }
        )
    return out


def _intent_linea(hoja, curva, lado, cerca_x=None):
    global _ULTIMO_LINEAL
    c = curva["c"]
    if not curva["arco"] and cerca_x is not None:
        try:
            if abs(curva["x1"] - cerca_x) <= abs(curva["x2"] - cerca_x):
                return hoja.CreateGeometryIntent(c, c.StartPoint)
            return hoja.CreateGeometryIntent(c, c.EndPoint)
        except Exception:
            pass
    if curva["arco"] and lado == "centro":
        for codigo in (57861,):
            try:
                return hoja.CreateGeometryIntent(c, codigo)
            except Exception as exc:
                _ULTIMO_LINEAL = f"intent centro: {exc}"
        try:
            return hoja.CreateGeometryIntent(c, c.CenterPoint)
        except Exception:
            return None
    if curva["arco"]:
        codigos = [K_CIRC[lado]]
        try:
            nombre = {
                "izq": "kCircularLeftPointIntent",
                "der": "kCircularRightPointIntent",
                "sup": "kCircularTopPointIntent",
                "inf": "kCircularBottomPointIntent",
            }[lado]
            codigos.insert(0, getattr(win32com.client.constants, nombre))
        except Exception:
            pass
        for codigo in codigos:
            try:
                return hoja.CreateGeometryIntent(c, codigo)
            except Exception as exc:
                _ULTIMO_LINEAL = f"intent {lado}: {exc}"
        return None
    if lado == "izq":
        p = c.StartPoint if curva["x1"] <= curva["x2"] else c.EndPoint
    elif lado == "der":
        p = c.StartPoint if curva["x1"] >= curva["x2"] else c.EndPoint
    elif lado == "inf":
        p = c.StartPoint if curva["y1"] <= curva["y2"] else c.EndPoint
    else:
        p = c.StartPoint if curva["y1"] >= curva["y2"] else c.EndPoint
    try:
        return hoja.CreateGeometryIntent(c, p)
    except Exception:
        return None


def _intent_punto_hoja(hoja, tg, sx, sy):
    """Punto de la hoja, en el centro proyectado del barreno."""
    sketch = hoja.Sketches.Add()
    sketch.Edit()
    punto = sketch.SketchPoints.Add(tg.CreatePoint2d(sx, sy))
    sketch.ExitEdit()
    return hoja.CreateGeometryIntent(punto)


def _lineal(hoja, tg, a, b, tx, ty, orient, lado_a, lado_b, cerca_x=None):
    i1 = _intent_linea(hoja, a, lado_a, cerca_x)
    i2 = _intent_linea(hoja, b, lado_b, cerca_x)
    if i1 is None or i2 is None:
        return None
    try:
        return hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            tg.CreatePoint2d(tx, ty), i1, i2, orient
        )
    except Exception as exc:
        global _ULTIMO_LINEAL
        _ULTIMO_LINEAL = str(exc)
        return None


def _buscar_par(hoja, tg, inv, pool_a, pool_b, target, orient, texto, tx, ty, lado_a, lado_b, tol=1.2, vertical=False, separacion=0.0, cerca_x=None, plano=False):
    mejor = None
    peor = None
    esperado = float(target) * _ESCALA / 10.0
    margen = max(0.35, (float(tol) + 9.0) * _ESCALA / 10.0)
    for a in pool_a:
        for b in pool_b:
            if a["c"] is b["c"]:
                continue
            proy = abs(a["my"] - b["my"]) if orient == K_VER else abs(a["mx"] - b["mx"])
            if abs(proy - esperado) > margen:
                continue
            dim = _lineal(hoja, tg, a, b, tx, ty, orient, lado_a, lado_b, cerca_x)
            if dim is None:
                continue
            mm = _mm(dim)
            if mm is not None and (peor is None or abs(mm - target) < peor[0]):
                peor = (abs(mm - target), mm)
            if mm is None or abs(mm - target) > tol:
                try:
                    dim.Delete()
                except Exception:
                    pass
                continue
            if mejor is None or abs(mm - target) < mejor[0]:
                if mejor is not None:
                    try:
                        mejor[2].Delete()
                    except Exception:
                        pass
                mejor = (abs(mm - target), mm, dim)
            else:
                try:
                    dim.Delete()
                except Exception:
                    pass
    if mejor is None:
        print(
            "  SIN", texto, "target", target,
            "cerca", None if peor is None else round(peor[1], 2),
            "err", _ULTIMO_LINEAL[:140],
        )
        return None
    _linea_solida(mejor[2], inv)
    if plano:
        try:
            mejor[2].HideValue = True
            mejor[2].Text.FormattedText = " "
        except Exception:
            pass
        # Mismo texto que la ordenada: perpendicular a la línea, en gris.
        if separacion < 0:
            _nota_color(
                hoja, tg, tx + separacion, ty, texto, (68, 68, 68), ancla="derecha",
            )
        else:
            _nota_color(
                hoja, tg, tx + separacion, ty, texto, (68, 68, 68), ancla="izquierda",
            )
        print("  COTA", texto, round(mejor[1], 2))
        return mejor[2]
    _capturar_extensiones(mejor[2])
    if vertical:
        mejor[2].HideValue = True
        try:
            mejor[2].Text.FormattedText = " "
        except Exception:
            pass
        # El texto sale de la línea: separacion < 0 a la izquierda, > 0 a la derecha.
        _etiqueta_girada(hoja, tg, inv, tx + separacion, ty, texto)
    else:
        _texto(mejor[2], inv, texto, vertical=False)
    print("  COTA", texto, round(mejor[1], 2))
    return mejor[2]


def _anillos_referencia(hoja, tg, inv, vista, miembros, eje, tam=0.16):
    """X en el centro que la cota señala."""
    for x, y, _d in miembros:
        sx, sy = _sheet(vista, tg, x, y)
        _marca_x(hoja, tg, inv, sx, sy, tam=tam)


def _mover_nota(nota, tg, x, y, ancla="centro"):
    """Lleva el RangeBox de la nota al punto. ancla: centro, abajo o derecha."""
    for _ in range(2):
        box = nota.RangeBox
        minx, miny = float(box.MinPoint.X), float(box.MinPoint.Y)
        maxx, maxy = float(box.MaxPoint.X), float(box.MaxPoint.Y)
        if ancla == "abajo":
            cx, cy = (minx + maxx) / 2.0, miny
        elif ancla == "arriba":
            cx, cy = (minx + maxx) / 2.0, maxy
        elif ancla == "eje":
            # El trazo del texto girado queda en el lado izquierdo de la caja.
            cx, cy = minx + _font() * 0.45, miny
        elif ancla == "derecha":
            cx, cy = maxx, (miny + maxy) / 2.0
        elif ancla == "izquierda":
            cx, cy = minx, (miny + maxy) / 2.0
        else:
            cx, cy = (minx + maxx) / 2.0, (miny + maxy) / 2.0
        pos = nota.Position
        nota.Position = tg.CreatePoint2d(float(pos.X) + (x - cx), float(pos.Y) + (y - cy))


def _nota_color(hoja, tg, x, y, texto, color, rot=0.0, ancla="centro", fuente=None):
    """Nota con color que sí sale en el JPG. El texto de sketch se queda negro."""
    r, g, b = color
    font = _font() if fuente is None else fuente
    nota = hoja.DrawingNotes.GeneralNotes.AddFitted(tg.CreatePoint2d(x, y), texto)
    nota.FormattedText = (
        f"<StyleOverride FontSize='{font}' Bold='True' "
        f"Font='Arial' Color='{r},{g},{b}'>{texto}</StyleOverride>"
    )
    if rot:
        nota.Rotation = rot
    nota.Color = _PLANO.Parent.TransientObjects.CreateColor(r, g, b)
    _mover_nota(nota, tg, x, y, ancla)
    return nota


def _linea_sketch(sketch, tg, inv, x1, y1, x2, y2, color, peso=0.04):
    linea = sketch.SketchLines.AddByTwoPoints(
        tg.CreatePoint2d(x1, y1), tg.CreatePoint2d(x2, y2)
    )
    try:
        linea.OverrideColor = inv.TransientObjects.CreateColor(*color)
    except Exception:
        pass
    try:
        linea.LineWeight = peso
    except Exception:
        pass
    return linea


def _marca_x(hoja, tg, inv, x, y, tam=0.17):
    """Equis de dos líneas. El color se aplica al salir de la edición: ahí sí llega al JPG."""
    color = _COLOR_MARCA if _COLOR_MARCA is not None else (4, 61, 86)
    try:
        sketch = hoja.Sketches.Add()
        sketch.Edit()
        lineas = [
            sketch.SketchLines.AddByTwoPoints(
                tg.CreatePoint2d(x - tam, y - tam), tg.CreatePoint2d(x + tam, y + tam)
            ),
            sketch.SketchLines.AddByTwoPoints(
                tg.CreatePoint2d(x - tam, y + tam), tg.CreatePoint2d(x + tam, y - tam)
            ),
        ]
        for linea in lineas:
            try:
                linea.LineWeight = 0.05
            except Exception:
                pass
        sketch.Visible = True
        sketch.ExitEdit()
        tinta = inv.TransientObjects.CreateColor(*color)
        try:
            capas = _PLANO.StylesManager.Layers
            try:
                raw = capas.Item("COTA_MARCA")
            except Exception:
                raw = capas.Item(1).Copy("COTA_MARCA")
            capa = win32com.client.CastTo(raw, "Layer")
            capa.Color = tinta
        except Exception as exc:
            capa = None
            print("  capa marca", exc)
        for linea in lineas:
            if capa is not None:
                try:
                    linea.Layer = capa
                except Exception:
                    pass
            try:
                linea.OverrideColor = tinta
            except Exception:
                pass
    except Exception as exc:
        print("  marca x", exc)


def _borde_ranura(curvas, sx, sy):
    """Lado recto superior de la ranura, en coordenadas de hoja."""
    mejor = None
    for c in curvas:
        if c["arco"] or c["dy"] > 0.12 or c["dx"] < 0.15:
            continue
        if abs(c["mx"] - sx) > 0.35 or abs(c["my"] - sy) > 0.9:
            continue
        if mejor is None or c["my"] > mejor:
            mejor = c["my"]
    return mejor


def _lider_sobre_centro(hoja, tg, inv, x, y_borde, y_texto):
    """Líder vertical del diámetro, en la misma X que la ordenada."""
    color = _COLOR_LINEA if _COLOR_LINEA is not None else (68, 68, 68)
    try:
        sketch = hoja.Sketches.Add()
        sketch.Edit()
        _linea_sketch(sketch, tg, inv, x, y_borde, x, y_texto, color, 0.03)
        _linea_sketch(sketch, tg, inv, x, y_borde, x - 0.09, y_borde + 0.16, color, 0.03)
        _linea_sketch(sketch, tg, inv, x, y_borde, x + 0.09, y_borde + 0.16, color, 0.03)
        sketch.Visible = True
        sketch.ExitEdit()
    except Exception as exc:
        print("  lider", exc)


def _anillo_sobre_cota(hoja, tg, inv, x, y):
    _marca_x(hoja, tg, inv, x, y, tam=0.22)


def _nota(hoja, vista, tg, x_mm, y_mm, z_mm, texto, dx=0.0, dy=0.0):
    p3 = tg.CreatePoint(x_mm / 10.0, y_mm / 10.0, z_mm / 10.0)
    p2 = vista.ModelToSheetSpace(p3)
    pt = tg.CreatePoint2d(float(p2.X) + dx, float(p2.Y) + dy)
    nota = hoja.DrawingNotes.GeneralNotes.AddFitted(pt, texto)
    if _COLOR_LINEA is not None and _PLANO is not None:
        try:
            nota.Color = _PLANO.Parent.TransientObjects.CreateColor(*_COLOR_LINEA)
        except Exception as exc:
            print("  color nota", exc)


def _teñir_equis(ruta):
    """El JPG no guarda el color del sketch. La equis sale negra y aquí pasa a #043D56."""
    if _COLOR_MARCA is None:
        return
    try:
        from PIL import Image
    except Exception as exc:
        print("  tinta", exc)
        return
    im = Image.open(ruta).convert("RGB")
    pix = im.load()
    ancho, alto = im.size

    def gris(p):
        r, g, b = p
        return 45 < r < 105 and abs(r - g) < 14 and abs(r - b) < 20 and abs(g - b) < 20

    def negro(p):
        return max(p) < 30

    def grupos(vals, salto=3):
        if not vals:
            return []
        bloques = [[vals[0]]]
        for v in vals[1:]:
            if v - bloques[-1][-1] <= salto:
                bloques[-1].append(v)
            else:
                bloques.append([v])
        return [b[len(b) // 2] for b in bloques]

    cols, filas = [], []
    for x in range(ancho):
        corrido = mejor = 0
        for y in range(alto):
            if gris(pix[x, y]):
                corrido += 1
                mejor = max(mejor, corrido)
            else:
                corrido = 0
        if mejor > 35:
            cols.append(x)
    for y in range(alto):
        corrido = mejor = 0
        for x in range(ancho):
            if gris(pix[x, y]):
                corrido += 1
                mejor = max(mejor, corrido)
            else:
                corrido = 0
        if mejor > 25:
            filas.append(y)
    tinta = 0
    for x in grupos(cols):
        for y in grupos(filas):
            negros = []
            for dy in range(-14, 15):
                for dx in range(-14, 15):
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < ancho and 0 <= yy < alto and negro(pix[xx, yy]):
                        negros.append((xx, yy))
            diag = [
                p for p in negros
                if abs(abs(p[0] - x) - abs(p[1] - y)) <= 2
            ]
            if len(diag) < 40:
                continue
            for xx, yy in negros:
                pix[xx, yy] = _COLOR_MARCA
                tinta += 1
    if tinta:
        im.save(ruta, quality=95)
    print("  equis", tinta)


def _exportar(inv, to, hoja, ruta):
    from generador_vistas import _recortar_exportacion_jpg

    hoja.Activate()
    try:
        inv.ActiveView.Update()
    except Exception:
        pass
    time.sleep(1.2)
    tmp = ruta + ".tmp.jpg"
    inv.ActiveView.Camera.SaveAsBitmap(
        tmp, 2400, 1600, to.CreateColor(255, 255, 255)
    )
    _recortar_exportacion_jpg(hoja, tmp, ruta)
    _teñir_equis(ruta)
    print("jpg", ruta, os.path.getsize(ruta))


def _huecos_flat(fbody):
    vistos = []
    for fi in range(1, int(fbody.Faces.Count) + 1):
        face = fbody.Faces.Item(fi)
        try:
            nloops = int(face.EdgeLoops.Count)
        except Exception:
            continue
        for li in range(1, nloops + 1):
            loop = face.EdgeLoops.Item(li)
            try:
                if bool(getattr(loop, "IsOuterEdgeLoop", True)):
                    continue
            except Exception:
                continue
            xs, ys = [], []
            for ei in range(1, int(loop.Edges.Count) + 1):
                try:
                    box = loop.Edges.Item(ei).Evaluator.RangeBox
                    xs += [float(box.MinPoint.X) * 10.0, float(box.MaxPoint.X) * 10.0]
                    ys += [float(box.MinPoint.Y) * 10.0, float(box.MaxPoint.Y) * 10.0]
                except Exception:
                    continue
            if not xs:
                continue
            xmin, xmax = min(xs), max(xs)
            ymin, ymax = min(ys), max(ys)
            x = (xmin + xmax) * 0.5
            y = (ymin + ymax) * 0.5
            # El ancho es el eje corto: círculo o lado recto del slot.
            d = min(xmax - xmin, ymax - ymin)
            if d < 0.4:
                continue
            if any(abs(x - a) < 2.5 and abs(y - b) < 2.5 for a, b, *_ in vistos):
                continue
            vistos.append((x, y, d, xmin, xmax, ymin, ymax))
    return vistos


def _sheet(vista, tg, x_mm, y_mm, z_mm=0.0):
    p = vista.ModelToSheetSpace(tg.CreatePoint(x_mm / 10.0, y_mm / 10.0, z_mm / 10.0))
    return float(p.X), float(p.Y)


def _dist_seg(c, sx, sy):
    dx, dy = c["x2"] - c["x1"], c["y2"] - c["y1"]
    largo = dx * dx + dy * dy
    if largo < 1e-12:
        return math.hypot(sx - c["x1"], sy - c["y1"])
    t = max(0.0, min(1.0, ((sx - c["x1"]) * dx + (sy - c["y1"]) * dy) / largo))
    return math.hypot(sx - (c["x1"] + t * dx), sy - (c["y1"] + t * dy))


def _tope_orilla(d_mm):
    """El centro del arco está a un radio de la orilla que se acota."""
    return max(0.8, (float(d_mm) / 2.0) * float(_ESCALA) / 10.0 + 0.35)


def _cerca(curvas, sx, sy, arco=None, vertical=None, tope=1.6):
    mejor = None
    for c in curvas:
        if arco is True and not c["arco"]:
            continue
        if arco is False and c["arco"]:
            continue
        if vertical is True and c["dx"] > 0.15:
            continue
        if vertical is False and c["dy"] > 0.15:
            continue
        if c["arco"]:
            try:
                cen = c["c"].CenterPoint
                d = math.hypot(float(cen.X) - sx, float(cen.Y) - sy)
                # La orilla del barreno está a un radio del centro. En escalas
                # grandes ese radio supera el tope y la cota se pierde.
                try:
                    radio = float(c["c"].ModelGeometry.Radius) * float(_ESCALA)
                    if radio > 1e-4:
                        d = min(d, abs(d - radio))
                except Exception:
                    pass
            except Exception:
                d = _dist_seg(c, sx, sy)
        else:
            d = _dist_seg(c, sx, sy)
        if mejor is None or d < mejor[0]:
            mejor = (d, c)
    if mejor is None or mejor[0] > tope:
        return None
    return mejor[1]


def _alas_letra(hoja, vista, tg, inv, part):
    """Punta del ala a la cara de afuera del alma. La de arriba es A."""
    from wing_cobre import medir_alas_mm

    info = medir_alas_mm(part)
    if not isinstance(info, dict):
        print("  SIN alas", info)
        return
    puestos = []
    for par, valor in zip(info["puntos"], info["alas"]):
        try:
            s = vista.ModelToSheetSpace(tg.CreatePoint(float(par[0][0]), float(par[0][1]), float(par[0][2])))
            sy = float(s.Y)
        except Exception as exc:
            print("  ala proj", exc)
            continue
        puestos.append((sy, float(valor), par))
    puestos.sort(key=lambda t: -t[0])
    for (sy, valor, par), letra in zip(puestos, ("A", "C")):
        s1 = vista.ModelToSheetSpace(tg.CreatePoint(*par[0]))
        s2 = vista.ModelToSheetSpace(tg.CreatePoint(*par[1]))
        curvas = _curvas(vista)
        c1 = _cerca(curvas, float(s1.X), float(s1.Y), arco=False)
        c2 = _cerca(curvas, float(s2.X), float(s2.Y), arco=False)
        if c1 is None or c2 is None:
            print("  SIN curva ala", letra)
            continue
        tx = max(float(s1.X), float(s2.X)) + (1.7 if letra == "A" else -1.7)
        ty = (float(s1.Y) + float(s2.Y)) / 2.0
        dim = _lineal(hoja, tg, c1, c2, tx, ty, K_VER, "sup", "inf")
        if dim is None:
            print("  SIN dim ala", letra)
            continue
        mm = _mm(dim)
        if mm is None or abs(mm - valor) > 3.0:
            print("  ala", letra, "midio", mm, "esperado", valor)
            try:
                dim.Delete()
            except Exception:
                pass
            continue
        _texto(dim, inv, f"{letra}={valor:.2f} mm")
        print("  COTA", letra, round(mm, 2))


def _acotar_doblado(hoja, vista, tg, inv, part, grupos):
    curvas = _curvas(vista)
    horiz = [c for c in curvas if c["dy"] < 0.08 and c["dx"] > 0.04 and not c["arco"]]
    vert = [c for c in curvas if c["dx"] < 0.08 and c["dy"] > 0.04 and not c["arco"]]
    print("perfil lineas H", len(horiz), "V", len(vert))
    for c in horiz + vert:
        print(
            "  ln",
            "H" if c["dy"] < 0.08 else "V",
            round(c["x1"], 2), round(c["y1"], 2),
            round(c["x2"], 2), round(c["y2"], 2),
        )
    if not horiz or not vert:
        return
    ys = [c["my"] for c in horiz]
    xs = [c["mx"] for c in vert]
    y_mid = (min(ys) + max(ys)) / 2.0
    x_min, x_max = min(xs), max(xs)
    height = _mm_hoja(max(ys) - min(ys))
    _buscar_par(
        hoja, tg, inv, horiz, horiz, height, K_VER, f"HEIGHT={height:.2f} mm",
        x_min - 2.85, (y_mid + max(ys)) / 2.0, "sup", "inf", tol=2.0,
        vertical=True, separacion=-0.8,
    )
    leg = _mm_hoja(x_max - x_min)
    thk = _espesor_mm(part) or 0.0
    if leg > max(thk * 3.0, 15.0):
        letra_alma = "B" if any(g.get("letra") == "B" for g in grupos) else "LEG"
        _buscar_par(
            hoja, tg, inv, vert, vert, leg, K_HOR, f"{letra_alma}={leg:.2f} mm",
            (x_min + x_max) / 2.0, max(ys) + 1.6, "izq", "der", tol=2.0,
        )
    largas = [c for c in horiz if c["dx"] > 1.5]
    cortas = [c for c in horiz if c["dx"] <= 1.5]
    if largas and cortas:
        punta_sup = max(cortas, key=lambda c: c["my"])
        punta_inf = min(cortas, key=lambda c: c["my"])
        web_lejos_sup = min(largas, key=lambda c: c["my"])
        web_lejos_inf = max(largas, key=lambda c: c["my"])
        ala_sup = _mm_hoja(punta_sup["my"] - web_lejos_sup["my"])
        ala_inf = _mm_hoja(punta_inf["my"] - web_lejos_inf["my"])
        letra_sup = _letra_cerca(grupos, vista, tg, punta_sup["mx"], punta_sup["my"]) or "A"
        letra_inf = _letra_cerca(grupos, vista, tg, punta_inf["mx"], punta_inf["my"]) or "C"
        if letra_inf == letra_sup:
            letra_inf = "C" if letra_sup != "C" else "A"
        if abs(ala_sup - height) > 3.0:
            _buscar_par(
                hoja, tg, inv, [punta_sup], [web_lejos_sup], ala_sup, K_VER,
                f"{letra_sup}={ala_sup:.2f} mm",
                x_max + 1.15, (punta_sup["my"] + web_lejos_sup["my"]) / 2.0,
                "sup", "inf", tol=2.0, vertical=True, separacion=0.75,
            )
        if abs(punta_sup["my"] - punta_inf["my"]) > 0.4 and abs(ala_inf - height) > 3.0:
            _buscar_par(
                hoja, tg, inv, [punta_inf], [web_lejos_inf], ala_inf, K_VER,
                f"{letra_inf}={ala_inf:.2f} mm",
                x_min - 1.15, (punta_inf["my"] + web_lejos_inf["my"]) / 2.0,
                "sup", "inf", tol=2.0, vertical=True, separacion=-0.75,
            )
    if thk > 0.2:
        _buscar_par(
            hoja, tg, inv, horiz + vert, horiz + vert, thk, K_HOR, f"THK={thk:.2f} mm",
            x_max + 3.2, min(ys) + 1.4, "izq", "der", tol=0.6,
        )
    _angulos(hoja, vista, tg, inv, curvas, horiz, vert, y_mid, grupos)
    _volcar_trazos(hoja, tg, inv)


def _centro_radio(curvas, hit):
    """Centro del filete de doblez más cercano al codo. Ese lado es el bolsillo."""
    mejor = None
    for c in curvas:
        if not c.get("arco"):
            continue
        d = math.hypot(c["mx"] - hit[0], c["my"] - hit[1])
        if d > 1.4:
            continue
        try:
            cen = c["c"].CenterPoint
            cx, cy = float(cen.X), float(cen.Y)
        except Exception:
            continue
        if mejor is None or d < mejor[0]:
            mejor = (d, cx, cy)
    if mejor is None:
        return None
    return mejor[1], mejor[2]


def _colocar_angulo(hoja, tg, inv, dims, h, v, ix, iy, ox, oy, etiqueta):
    """Arco de 90 en (ix, iy). (ox, oy) apunta hacia fuera de la pieza."""
    i1 = _intent_linea(hoja, h, "izq")
    i2 = _intent_linea(hoja, v, "sup")
    if i1 is None or i2 is None:
        print("  SIN angulo", etiqueta)
        return
    nrm = math.hypot(ox, oy) or 1.0
    ux, uy = ox / nrm, oy / nrm
    elegido = None
    for escala in (0.9, -0.9):
        tx = ix + ux * escala
        ty = iy + uy * escala
        try:
            dim = dims.AddAngular(tg.CreatePoint2d(tx, ty), i1, i2)
        except Exception:
            continue
        deg = abs(float(dim.ModelValue)) * 180.0 / math.pi
        if deg > 180.0:
            deg = 360.0 - deg
        if abs(deg - 90.0) > 8.0:
            try:
                dim.Delete()
            except Exception:
                pass
            continue
        if escala > 0 or elegido is None:
            if elegido is not None:
                try:
                    elegido[0].Delete()
                except Exception:
                    pass
            elegido = (dim, deg, tx, ty, escala)
        else:
            try:
                dim.Delete()
            except Exception:
                pass
        if escala > 0:
            break
    if elegido is None:
        print("  SIN angulo", etiqueta)
        return
    dim, deg, tx, ty, escala = elegido
    dim.HideValue = True
    try:
        dim.Text.FormattedText = " "
        dim.Text.Color = _azul(inv)
    except Exception:
        pass
    _linea_solida(dim, inv)
    _globo(
        hoja, tg, inv, ix, iy,
        ix + ux * 2.6 * (1.0 if escala > 0 else -1.0),
        iy + uy * 2.6 * (1.0 if escala > 0 else -1.0),
        f"{etiqueta}∠{deg:.0f}°",
    )
    print(
        "  ANG", etiqueta, round(deg, 1),
        "esquina", round(ix, 2), round(iy, 2),
        "texto", round(tx, 2), round(ty, 2),
    )


def _extremo_hacia(curva, x, y):
    d1 = math.hypot(curva["x1"] - x, curva["y1"] - y)
    d2 = math.hypot(curva["x2"] - x, curva["y2"] - y)
    if d1 <= d2:
        return curva["x1"], curva["y1"], True
    return curva["x2"], curva["y2"], False


def _intent_punto(hoja, curva, x, y):
    """Intent en el extremo de la línea que llega al doblez, no en la punta lejana."""
    c = curva["c"]
    try:
        s = c.StartPoint
        e = c.EndPoint
        ds = math.hypot(float(s.X) - x, float(s.Y) - y)
        de = math.hypot(float(e.X) - x, float(e.Y) - y)
        return hoja.CreateGeometryIntent(c, s if ds <= de else e)
    except Exception:
        return None


def _intent_extremo(hoja, curva, inicio):
    c = curva["c"]
    try:
        p = c.StartPoint if inicio else c.EndPoint
        return hoja.CreateGeometryIntent(c, p)
    except Exception:
        return None


def _angulo_en_extremos(hoja, tg, inv, dims, h, v, ox, oy, etiqueta):
    """Ángulo entre los extremos de h y v que se acercan al doblez."""
    hx, hy, hi = _extremo_hacia(h, v["mx"], v["my"])
    vx, vy, vi = _extremo_hacia(v, h["mx"], h["my"])
    i1 = _intent_extremo(hoja, h, hi)
    i2 = _intent_extremo(hoja, v, vi)
    print(
        "  caras", etiqueta,
        round(hx, 2), round(hy, 2), "->", round(vx, 2), round(vy, 2),
    )
    if i1 is None or i2 is None:
        print("  SIN angulo", etiqueta)
        return
    ix, iy = (hx + vx) / 2.0, (hy + vy) / 2.0
    nrm = math.hypot(ox, oy) or 1.0
    ux, uy = ox / nrm, oy / nrm
    elegido = None
    for paso in (0.35, 0.9):
        tx, ty = ix + ux * paso, iy + uy * paso
        try:
            dim = dims.AddAngular(tg.CreatePoint2d(tx, ty), i1, i2)
        except Exception:
            continue
        deg = abs(float(dim.ModelValue)) * 180.0 / math.pi
        if deg > 180.0:
            deg = 360.0 - deg
        if abs(deg - 90.0) > 8.0:
            try:
                dim.Delete()
            except Exception:
                pass
            continue
        elegido = (dim, deg, tx, ty)
        break
    if elegido is None:
        print("  SIN angulo", etiqueta)
        return
    dim, deg, tx, ty = elegido
    dim.HideValue = True
    try:
        dim.Text.FormattedText = " "
        dim.Text.Color = _azul(inv)
    except Exception:
        pass
    _linea_solida(dim, inv)
    _globo(
        hoja, tg, inv, ix, iy,
        ix + ux * 1.8, iy + uy * 1.8,
        f"{etiqueta}∠{deg:.0f}°",
    )
    print(
        "  ANG", etiqueta, round(deg, 1),
        "codo", round(ix, 2), round(iy, 2),
        "texto", round(tx, 2), round(ty, 2),
    )


def _angulos(hoja, vista, tg, inv, curvas, horiz, vert, y_mid, grupos):
    """Ángulo en cada codo, sobre las caras que se juntan en el radio."""
    dims = hoja.DrawingDimensions.GeneralDimensions
    alma = [h for h in horiz if h["dx"] > 1.5]
    if not alma or not vert:
        print("  SIN angulo: sin alma")
        return
    x_mid = sum(v["mx"] for v in vert) / len(vert)
    objetivos = (
        ("der", [v for v in vert if v["mx"] >= x_mid]),
        ("izq", [v for v in vert if v["mx"] < x_mid]),
    )
    usadas = set()
    for lado, alas in objetivos:
        if not alas:
            continue
        puesto = False
        mejor = None
        for v in alas:
            if _mm_hoja(v["dy"]) < 15.0:
                continue
            for h in alma:
                pares = []
                for ax, ay in ((h["x1"], h["y1"]), (h["x2"], h["y2"])):
                    for bx, by in ((v["x1"], v["y1"]), (v["x2"], v["y2"])):
                        pares.append((math.hypot(ax - bx, ay - by), ax, ay, bx, by))
                dist, ax, ay, bx, by = min(pares)
                if dist < 0.12 or dist > 0.9:
                    continue
                if mejor is None or dist < mejor[0]:
                    mejor = (dist, h, v, ax, ay, bx, by)
        if mejor is None:
            print("  SIN angulo", lado)
            continue
        dist, h, v, ax, ay, bx, by = mejor
        hit = ((ax + bx) / 2.0, (ay + by) / 2.0)
        etiqueta = _letra_cerca(grupos, vista, tg, hit[0], hit[1])
        if not etiqueta or etiqueta in usadas or etiqueta == "B":
            etiqueta = "A" if lado == "der" else "C"
            if etiqueta in usadas:
                etiqueta = "C" if etiqueta == "A" else "A"
        usadas.add(etiqueta)
        i1 = _intent_punto(hoja, h, ax, ay)
        i2 = _intent_punto(hoja, v, bx, by)
        print(
            "  caras", etiqueta,
            "H", round(h["my"], 2), "extremo", round(ax, 2), round(ay, 2),
            "V", round(v["mx"], 2), "extremo", round(bx, 2), round(by, 2),
            "sep", round(dist, 2),
        )
        if i1 is None or i2 is None:
            print("  SIN angulo", etiqueta)
            continue
        centro = _centro_radio(curvas, hit)
        if centro is None:
            vx = x_mid - hit[0]
            vy = y_mid - hit[1]
        else:
            vx = centro[0] - hit[0]
            vy = centro[1] - hit[1]
        nrm = math.hypot(vx, vy) or 1.0
        # El arco va en el codo, entre las dos caras del doblez.
        preferido = 1.0
        elegido = None
        for signo in (preferido, -preferido):
            tx = hit[0] + signo * vx / nrm * 0.9
            ty = hit[1] + signo * vy / nrm * 0.9
            try:
                dim = dims.AddAngular(tg.CreatePoint2d(tx, ty), i1, i2)
            except Exception:
                continue
            deg = abs(float(dim.ModelValue)) * 180.0 / math.pi
            if deg > 180.0:
                deg = 360.0 - deg
            if abs(deg - 90.0) > 8.0:
                try:
                    dim.Delete()
                except Exception:
                    pass
                continue
            if signo == preferido or elegido is None:
                if elegido is not None:
                    try:
                        elegido[0].Delete()
                    except Exception:
                        pass
                elegido = (dim, deg, tx, ty, signo)
            else:
                try:
                    dim.Delete()
                except Exception:
                    pass
            if signo == preferido:
                break
        if elegido is None:
            print("  SIN angulo", etiqueta)
            continue
        dim, deg, tx, ty, signo = elegido
        dim.HideValue = True
        try:
            dim.Text.FormattedText = " "
            dim.Text.Color = _azul(inv)
        except Exception:
            pass
        _linea_solida(dim, inv)
        _globo(
            hoja, tg, inv, hit[0], hit[1],
            hit[0] + signo * vx / nrm * 1.7,
            hit[1] + signo * vy / nrm * 1.7,
            f"{etiqueta}∠{deg:.0f}°",
        )
        print(
            "  ANG", etiqueta, round(deg, 1),
            "en", round(hit[0], 2), round(hit[1], 2),
            "texto", round(tx, 2), round(ty, 2), "signo", signo,
        )



def _agrupar(items, clave, tol=0.25):
    grupos = []
    for item in sorted(items, key=clave):
        v = clave(item)
        if grupos and abs(v - grupos[-1][0]) <= tol:
            grupos[-1][1].append(item)
        else:
            grupos.append([v, [item]])
    return grupos


def _segs_exterior(body):
    mejor = None
    for fi in range(1, int(body.Faces.Count) + 1):
        face = body.Faces.Item(fi)
        try:
            area = float(face.Evaluator.Area)
        except Exception:
            continue
        if mejor is None or area > mejor[0]:
            mejor = (area, face)
    if mejor is None:
        return []
    segs = []
    face = mejor[1]
    for li in range(1, int(face.EdgeLoops.Count) + 1):
        loop = face.EdgeLoops.Item(li)
        try:
            if not bool(loop.IsOuterEdgeLoop):
                continue
        except Exception:
            continue
        for ei in range(1, int(loop.Edges.Count) + 1):
            box = loop.Edges.Item(ei).Evaluator.RangeBox
            segs.append((
                float(box.MinPoint.X) * 10.0, float(box.MinPoint.Y) * 10.0,
                float(box.MaxPoint.X) * 10.0, float(box.MaxPoint.Y) * 10.0,
            ))
    return segs


def _ys_cubre(segs, x):
    ys = []
    for xa, ya, xb, yb in segs:
        if abs(yb - ya) > 2.5 or abs(xb - xa) < 8.0:
            continue
        if min(xa, xb) - 0.8 <= x <= max(xa, xb) + 0.8:
            ys.append((ya + yb) * 0.5)
    if len(ys) < 2:
        return None
    return min(ys), max(ys)


def _cortes_doblez(fp):
    cortes = []
    try:
        n = int(fp.FlatBendResults.Count)
    except Exception:
        return cortes
    for i in range(1, n + 1):
        try:
            box = fp.FlatBendResults.Item(i).Edge.Evaluator.RangeBox
            cortes.append((float(box.MinPoint.X) + float(box.MaxPoint.X)) * 5.0)
        except Exception:
            continue
    unicos = []
    for c in sorted(cortes):
        if not unicos or abs(c - unicos[-1]) > 3.0:
            unicos.append(c)
    return unicos


def _desfase_y(fp):
    """Las dos tiras miden menos que el alto total: el desarrollo va desfasado."""
    extremos = _extremos_ancho(fp)
    if len(extremos) < 2:
        return False
    rb = fp.Body.RangeBox
    env = abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y)) * 10.0
    return all(env - t["ancho"] > 3.0 for t in extremos)


def _extremos_ancho(fp):
    """Ancho real en cada extremo. El X chico queda a la derecha de la hoja."""
    body = fp.Body
    rb = body.RangeBox
    x0 = float(rb.MinPoint.X) * 10.0
    x1 = float(rb.MaxPoint.X) * 10.0
    segs = _segs_exterior(body)
    out = []
    for frac, lado in ((0.12, "der"), (0.88, "izq")):
        x = x0 + (x1 - x0) * frac
        ys = _ys_cubre(segs, x)
        if not ys:
            continue
        y0, y1 = ys
        out.append({"lado": lado, "x": x, "y0": y0, "y1": y1, "ancho": y1 - y0})
    return out


def _paneles_flat(fp, huecos):
    body = fp.Body
    rb = body.RangeBox
    x0 = float(rb.MinPoint.X) * 10.0
    x1 = float(rb.MaxPoint.X) * 10.0
    cortes = [c for c in _cortes_doblez(fp) if x0 + 2.0 < c < x1 - 2.0]
    limites = [x0] + cortes + [x1]
    segs = _segs_exterior(body)
    y_med = (float(rb.MinPoint.Y) + float(rb.MaxPoint.Y)) * 5.0
    paneles = []
    for a, b in zip(limites, limites[1:]):
        if b - a < 4.0:
            continue
        cx = (a + b) * 0.5
        ys = _ys_cubre(segs, cx)
        cy = (ys[0] + ys[1]) * 0.5 if ys else y_med
        n = sum(1 for h in huecos if a - 1.0 <= h[0] <= b + 1.0)
        paneles.append({"x0": a, "x1": b, "cx": cx, "cy": cy, "huecos": n})
    return paneles


def _letras_paneles(paneles):
    """La del medio es B. Las puntas: A la de más barrenos, C la otra."""
    if not paneles:
        return []
    if len(paneles) == 1:
        paneles[0]["letra"] = "A"
        return paneles
    if len(paneles) == 2:
        orden = sorted(paneles, key=lambda g: -g["huecos"])
        orden[0]["letra"] = "A"
        orden[1]["letra"] = "B"
        return paneles
    orden = sorted(paneles, key=lambda g: g["cx"])
    medio = orden[len(orden) // 2]
    puntas = [g for g in orden if g is not medio]
    puntas.sort(key=lambda g: -g["huecos"])
    puntas[0]["letra"] = "A"
    medio["letra"] = "B"
    puntas[1]["letra"] = "C"
    return paneles


def _cota_ancho_local(hoja, vista, tg, inv, horiz, tramo, texto, x_cota, sep, plano=False):
    sx, sy_a = _sheet(vista, tg, tramo["x"], tramo["y0"])
    _sx, sy_b = _sheet(vista, tg, tramo["x"], tramo["y1"])
    a = _cerca(horiz, sx, sy_a, arco=False, vertical=False, tope=1.2)
    b = _cerca(horiz, sx, sy_b, arco=False, vertical=False, tope=1.2)
    if a is None or b is None or a is b:
        print("  SIN ancho", round(tramo["ancho"], 2))
        return None
    return _buscar_par(
        hoja, tg, inv, [a], [b], tramo["ancho"], K_VER, texto,
        x_cota, (sy_a + sy_b) / 2.0, "inf", "sup",
        tol=2.5, vertical=True, separacion=sep, plano=plano,
    )


def _intent_centro_arco(hoja, curva):
    c = curva["c"]
    try:
        return hoja.CreateGeometryIntent(c, 57860)
    except Exception:
        return hoja.CreateGeometryIntent(c, c.CenterPoint)


def _centros_radios(cx, cy, d, xmin, xmax, ymin, ymax):
    """Círculo: un centro. Slot: el centro de cada radio, no el de la ranura."""
    largo_x = xmax - xmin
    largo_y = ymax - ymin
    if max(largo_x, largo_y) <= d + 0.8:
        return [(cx, cy)]
    radio = d / 2.0
    if largo_x >= largo_y:
        return [(xmin + radio, cy), (xmax - radio, cy)]
    return [(cx, ymin + radio), (cx, ymax - radio)]


def _arco_en_centro(curvas, sx, sy, tope=0.28):
    """El arco cuyo centro cae en el punto. Sin el ajuste de radio, que confunde los dos del slot."""
    mejor = None
    for c in curvas:
        if not c["arco"]:
            continue
        try:
            cen = c["c"].CenterPoint
            dist = math.hypot(float(cen.X) - sx, float(cen.Y) - sy)
        except Exception:
            continue
        if mejor is None or dist < mejor[0]:
            mejor = (dist, c)
    if mejor is None or mejor[0] > tope:
        return None
    return mejor[1]


def _intent_centro_ranura(hoja, curvas, sx, sy):
    """Punto medio del lado recto de la ranura. El arco mide un extremo, no el centro."""
    mejor = None
    for c in curvas:
        if c["arco"] or c["dy"] > 0.12 or c["dx"] < 0.15:
            continue
        if abs(c["mx"] - sx) > 0.35 or abs(c["my"] - sy) > 0.9:
            continue
        if mejor is None or c["my"] > mejor["my"]:
            mejor = c
    if mejor is None:
        return None
    try:
        return hoja.CreateGeometryIntent(mejor["c"], 57859)
    except Exception as exc:
        print("  ranura", exc)
        return None


def _ancho_texto(texto):
    return max(0.45, len(texto) * _font() * 0.80)


def _repartir_sin_montar(items, indice, holgura):
    """Filas cuya coordenada queda separada, para que el texto no se monte."""
    filas = []
    for item in items:
        pos = item[indice]
        destino = None
        for fila in filas:
            if all(abs(pos - otro[indice]) >= holgura for otro in fila):
                destino = fila
                break
        if destino is None:
            filas.append([item])
        else:
            destino.append(item)
    return filas


def _etiqueta_recta(hoja, tg, inv, x, y, texto):
    """Texto horizontal. x, y es la esquina inferior izquierda."""
    try:
        sketch = hoja.Sketches.Add()
        sketch.Edit()
        caja = sketch.TextBoxes.AddFitted(tg.CreatePoint2d(x, y), texto)
        from cota_estilo import aplicar_estilo_texto_cota, armar_formatted_texto_cota

        aplicar_estilo_texto_cota(caja, texto, inv, vertical=False)
        font = _font()
        formado = armar_formatted_texto_cota(
            texto, font_cm=font, bold=True, vertical=False
        )
        if _COLOR_LINEA is not None:
            r, g, b = _COLOR_LINEA
            formado = formado.replace(
                "<StyleOverride ",
                f"<StyleOverride Color='{r},{g},{b}' ",
            )
        caja.FormattedText = formado
        try:
            caja.Color = _color_linea(inv)
        except Exception:
            pass
        sketch.Visible = True
        sketch.ExitEdit()
    except Exception as exc:
        print("  etiq", exc)


def _juego_ordenado(hoja, tg, inv, items, orient, px, py, nota_origen=True):
    """Un juego ordenado. El primer intento es el origen .00. Una extensión por valor."""
    col = inv.TransientObjects.CreateObjectCollection()
    limpios = []
    for intent, texto, valor, ax, ay in items:
        if intent is None:
            print("  SIN ordenada", texto)
            continue
        col.Add(intent)
        limpios.append((texto, float(valor), float(ax), float(ay)))
    if int(col.Count) < 2:
        print("  SIN ordenada conjunto")
        return None
    try:
        conjunto = hoja.DrawingDimensions.OrdinateDimensionSets.Add(
            col, tg.CreatePoint2d(px, py), orient,
        )
    except Exception as exc:
        print("  SIN ordenada", exc)
        return None
    miembros = []
    try:
        n = int(conjunto.Members.Count)
        for i in range(1, n + 1):
            miembros.append(conjunto.Members.Item(i))
    except Exception as exc:
        print("  miembros ordenada", exc)
        return conjunto
    try:
        if miembros:
            conjunto.OriginMember = miembros[0]
    except Exception as exc:
        print("  origen ordenada", exc)
    def _con_medido(texto, mm):
        if mm is None or "=" not in texto:
            return texto
        izq, der = texto.split("=", 1)
        suf = " TYP" if "TYP" in der else ""
        return f"{izq}={mm:.2f}{suf} mm"

    puestos = []
    usados = set()
    sueltos = []
    for dim in miembros:
        mm = _mm(dim)
        mejor = None
        mejor_d = 1e9
        for j, (texto, valor, ax, ay) in enumerate(limpios):
            if j in usados:
                continue
            d = abs((0.0 if mm is None else mm) - valor)
            if d < mejor_d:
                mejor_d = d
                mejor = j
        if mejor is None or mejor_d > 2.0:
            sueltos.append((dim, mm))
            continue
        usados.add(mejor)
        texto, valor, ax, ay = limpios[mejor]
        texto = _con_medido(texto, mm)
        _linea_solida(dim, inv)
        try:
            dim.HideValue = True
            dim.Text.FormattedText = " "
            if orient == K_HOR:
                dim.Text.Origin = tg.CreatePoint2d(ax, py)
            else:
                dim.Text.Origin = tg.CreatePoint2d(px, ay)
        except Exception:
            pass
        puestos.append((texto, ax, ay))
        print("  COTA", texto, None if mm is None else round(mm, 2))
    libres = [t for j, t in enumerate(limpios) if j not in usados]
    sueltos.sort(key=lambda t: 0.0 if t[1] is None else t[1])
    libres.sort(key=lambda t: t[1])
    for (dim, mm), (texto, valor, ax, ay) in zip(sueltos, libres):
        texto = _con_medido(texto, mm)
        _linea_solida(dim, inv)
        try:
            dim.HideValue = True
            dim.Text.FormattedText = " "
            if orient == K_HOR:
                dim.Text.Origin = tg.CreatePoint2d(ax, py)
            else:
                dim.Text.Origin = tg.CreatePoint2d(px, ay)
        except Exception:
            pass
        puestos.append((texto, ax, ay))
        print("  COTA", texto, None if mm is None else round(mm, 2))
    gris = (68, 68, 68)
    if orient == K_HOR:
        for texto, ax, ay in puestos:
            if not nota_origen and texto.strip() == "0.00":
                continue
            h = _ancho_texto(texto)
            _nota_color(
                hoja, tg, ax, py - 0.12 - h, texto, gris,
                rot=math.pi / 2.0, ancla="eje",
            )
    else:
        for texto, ax, ay in puestos:
            if not nota_origen and texto.strip() == "0.00":
                continue
            _nota_color(hoja, tg, px - 0.06, ay, texto, gris, ancla="derecha")
    return conjunto


def _acotar_corte(hoja, vista, tg, inv, huecos, fr, grupos, thk, fp=None, origen_xmin=False, al_centro=False):
    curvas = _curvas(vista)
    vert = [c for c in curvas if c["dx"] < 0.1 and c["dy"] > 0.15]
    horiz = [c for c in curvas if c["dy"] < 0.1 and c["dx"] > 0.15]
    print("flat lineas H", len(horiz), "V", len(vert), "curvas", len(curvas))
    x_min = float(fr.MinPoint.X) * 10.0
    x_max = float(fr.MaxPoint.X) * 10.0
    oy = float(fr.MinPoint.Y) * 10.0
    y_sup = float(fr.MaxPoint.Y) * 10.0
    if origen_xmin:
        ox, x_far = x_min, x_max
    else:
        ox, x_far = x_max, x_min
    sx0, sy0 = _sheet(vista, tg, ox, oy)
    sx1, sy1 = _sheet(vista, tg, x_far, y_sup)
    if sx0 > sx1:
        ox, x_far = x_far, ox
        origen_xmin = not origen_xmin
        sx0, sy0 = _sheet(vista, tg, ox, oy)
        sx1, sy1 = _sheet(vista, tg, x_far, y_sup)
    print("origen", "xmin" if origen_xmin else "xmax", "sx", round(sx0, 2), round(sx1, 2))
    largo = abs(ox - x_far)
    ancho = abs(y_sup - oy)
    borde_x = _cerca(vert, sx0, (sy0 + sy1) / 2.0, arco=False, vertical=True)
    borde_y = _cerca(horiz, (sx0 + sx1) / 2.0, sy0, arco=False, vertical=False)
    borde_izq = _cerca(vert, sx1, (sy0 + sy1) / 2.0, arco=False, vertical=True)
    borde_sup = _cerca(horiz, (sx0 + sx1) / 2.0, sy1, arco=False, vertical=False)
    if borde_y is None and horiz:
        borde_y = min(horiz, key=lambda c: c["my"])
    if borde_sup is None and horiz:
        borde_sup = max(horiz, key=lambda c: c["my"])
    if borde_y and borde_sup:
        ancho = _mm_hoja(borde_sup["my"] - borde_y["my"])
    if borde_x and borde_izq and not al_centro:
        _buscar_par(
            hoja, tg, inv, [borde_izq], [borde_x], largo, K_HOR,
            f"LENGTH={largo:.2f} mm", (sx0 + sx1) / 2.0, sy1 + 1.5,
            "izq", "der", tol=2.0,
        )
    fila_x = sy0 - 1.35
    items_x = []
    if al_centro and borde_x is not None:
        items_x.append((_intent_linea(hoja, borde_x, "inf"), "0.00", 0.0, sx0, sy0))
    n_letra = 0
    medidas_x = []
    for cx, cy, d, xmin, xmax, ymin, ymax in huecos:
        if al_centro:
            for rx, ry in _centros_radios(cx, cy, d, xmin, xmax, ymin, ymax):
                valor = rx - ox if origen_xmin else ox - rx
                medidas_x.append((round(valor, 2), rx, ry, d, rx, ymin))
        else:
            # Orilla del barreno que mira al origen, que queda a la izquierda de la hoja.
            ancla_x = xmin if origen_xmin else xmax
            valor = ancla_x - ox if origen_xmin else ox - ancla_x
            medidas_x.append((round(valor, 2), cx, cy, d, ancla_x, ymin))
    cols = _agrupar(medidas_x, lambda t: t[0], 0.3)
    cols.sort(key=lambda g: g[0])
    piso = max(6.4, float(hoja.Height) * 0.22)
    n_cols = len(cols)
    libre = (sy0 - 0.65) - (piso + 0.95)
    paso_x = 0.58
    if n_cols > 1 and libre > 0:
        paso_x = min(0.58, max(0.30, libre / (n_cols - 1)))
    for i, (valor, miembros) in enumerate(cols):
        letra = _letra_seq(n_letra)
        n_letra += 1
        sx, sy = _sheet(vista, tg, miembros[0][4], miembros[0][2])
        fracs = (0.42, 0.62, 0.78)
        frac = fracs[i % len(fracs)]
        y_cota = sy0 - 0.7 - i * paso_x
        typ = " TYP" if len(miembros) >= 2 else ""
        texto_x = f"{letra}={valor:.2f}{typ} mm"
        if al_centro:
            ancla_m = min(miembros, key=lambda m: m[2])
            sx, sy = _sheet(vista, tg, ancla_m[4], ancla_m[2])
            objetivo = _arco_en_centro(curvas, sx, sy)
            intent = _intent_centro_arco(hoja, objetivo) if objetivo is not None else None
            if intent is None:
                print("  SIN X", valor)
            else:
                items_x.append((intent, texto_x, valor, sx, sy))
        else:
            tope_x = _tope_orilla(miembros[0][3])
            objetivo = _cerca(curvas, sx, sy, arco=True, tope=tope_x)
            if objetivo is None:
                objetivo = _cerca(vert, sx, sy, arco=False, vertical=True, tope=0.45)
            if borde_x is None or objetivo is None:
                print("  SIN X", valor)
                continue
            _buscar_par(
                hoja, tg, inv, [borde_x], [objetivo], valor, K_HOR,
                texto_x,
                sx0 + (sx - sx0) * frac, y_cota,
                "izq", "izq", tol=1.5,
            )
        if al_centro or len(miembros) >= 2:
            _anillos_referencia(
                hoja, tg, inv, vista,
                [(m[4], m[2], m[3]) for m in miembros], "X",
                tam=0.11,
            )
    if al_centro and borde_izq is not None:
        items_x.append((
            _intent_linea(hoja, borde_izq, "inf"),
            f"LENGTH={largo:.2f} mm",
            largo,
            sx1,
            sy0,
        ))
    if al_centro and items_x:
        origen = items_x[0]
        resto = items_x[1:]
        # Un juego por cota. Si dos centros cercanos comparten juego, Inventor
        # quiebra la extensión y el texto deja de caer sobre esa X.
        # 0.50 cm es el ancho de la letra vertical: más cerca, baja a otra fila.
        filas_x = _repartir_sin_montar(resto, 3, 0.50)
        y_fila = fila_x
        nota = True
        for fila in filas_x:
            for item in fila:
                _juego_ordenado(
                    hoja, tg, inv, [origen, item], K_HOR, sx0, y_fila,
                    nota_origen=nota,
                )
                nota = False
            alto = max(_ancho_texto(it[1]) for it in fila)
            y_fila = y_fila - alto - 0.40
    medidas_y = []
    for cx, cy, d, xmin, xmax, ymin, ymax in huecos:
        if al_centro:
            for rx, ry in _centros_radios(cx, cy, d, xmin, xmax, ymin, ymax):
                medidas_y.append((round(ry - oy, 2), rx, ry, d, xmax, ry))
        else:
            ancla_y = ymin
            valor = ancla_y - oy
            medidas_y.append((round(valor, 2), cx, cy, d, xmax, ancla_y))
    filas = _agrupar(medidas_y, lambda t: t[0], 0.3)
    filas.sort(key=lambda g: g[0])
    extremos = _extremos_ancho(fp) if fp is not None else []
    distintos = []
    for tramo in extremos:
        if not any(abs(tramo["ancho"] - u) < 3.0 for u in distintos):
            distintos.append(tramo["ancho"])
    local = bool(distintos) and (
        len(distintos) > 1 or abs(distintos[0] - ancho) > 3.0
    )
    if al_centro and borde_y is not None:
        col_y = sx0 - 1.7
        items_y = [(_intent_linea(hoja, borde_y, "izq"), "0.00", 0.0, col_y, sy0)]
        for valor, miembros in filas:
            letra = _letra_seq(n_letra)
            n_letra += 1
            # El de la derecha, para que la extensión cruce todos los radios.
            if origen_xmin:
                ancla = max(miembros, key=lambda m: m[1])
            else:
                ancla = min(miembros, key=lambda m: m[1])
            sx, sy = _sheet(vista, tg, ancla[1], ancla[5])
            objetivo = _arco_en_centro(curvas, sx, sy)
            if objetivo is None:
                print("  SIN Y", valor)
                continue
            typ = " TYP" if len(miembros) >= 2 else ""
            items_y.append((
                _intent_centro_arco(hoja, objetivo),
                f"{letra}={valor:.2f}{typ} mm",
                valor,
                col_y,
                sy,
            ))
        if borde_sup is not None and not local:
            items_y.append((
                _intent_linea(hoja, borde_sup, "izq"),
                f"WIDTH={ancho:.2f} mm",
                ancho,
                col_y,
                sy1,
            ))
        elif local:
            # El ancho que nace en el origen entra en la misma ordenada.
            # El desfasado se acota solo, más abajo, también en ordenada.
            for tramo in extremos:
                sx_t = _sheet(vista, tg, tramo["x"], (tramo["y0"] + tramo["y1"]) / 2.0)[0]
                tramo["lado"] = "izq" if abs(sx_t - sx0) <= abs(sx_t - sx1) else "der"
            orden_w = [t for lado in ("izq", "der") for t in extremos if t["lado"] == lado]
            for i, tramo in enumerate(orden_w, start=1):
                tramo["n"] = i
                y_top = tramo["y1"] if abs(tramo["y1"] - oy) >= abs(tramo["y0"] - oy) else tramo["y0"]
                tramo["en_origen"] = abs(min(tramo["y0"], tramo["y1"]) - oy) < 4.0
                if not tramo["en_origen"]:
                    continue
                sx_e, sy_e = _sheet(vista, tg, tramo["x"], y_top)
                curva = _cerca(horiz, sx_e, sy_e, arco=False, vertical=False, tope=1.4)
                if curva is None:
                    print("  SIN ancho", round(tramo["ancho"], 2))
                    continue
                items_y.append((
                    _intent_linea(hoja, curva, "izq"),
                    f"WIDTH {i}={tramo['ancho']:.2f} mm",
                    abs(y_top - oy),
                    col_y,
                    sy_e,
                ))
            if borde_sup is not None and all(abs(ancho - u) > 3.0 for u in distintos):
                items_y.append((
                    _intent_linea(hoja, borde_sup, "izq"),
                    f"WIDTH TOTAL={ancho:.2f} mm",
                    ancho,
                    col_y,
                    sy1,
                ))
        origen_y = items_y[0]
        # Igual que en X: un juego por cota, para que la línea no se quiebre.
        columnas = _repartir_sin_montar(items_y[1:], 4, 0.48)
        x_col = col_y
        nota = True
        for columna in columnas:
            for item in columna:
                intent, texto, valor, _ax, ay = item
                _juego_ordenado(
                    hoja, tg, inv,
                    [origen_y, (intent, texto, valor, x_col, ay)],
                    K_VER, x_col, sy0,
                    nota_origen=nota,
                )
                nota = False
            ancho_txt = max(_ancho_texto(it[1]) for it in columna)
            x_col = x_col - ancho_txt - 0.30
        for tramo in extremos if local else []:
            if not tramo.get("en_origen"):
                y_top = tramo["y1"] if abs(tramo["y1"] - oy) >= abs(tramo["y0"] - oy) else tramo["y0"]
                y_bot = tramo["y0"] if y_top == tramo["y1"] else tramo["y1"]
                sx_b, sy_b = _sheet(vista, tg, tramo["x"], y_bot)
                sx_t, sy_t = _sheet(vista, tg, tramo["x"], y_top)
                curva_b = _cerca(horiz, sx_b, sy_b, arco=False, vertical=False, tope=1.4)
                curva_t = _cerca(horiz, sx_t, sy_t, arco=False, vertical=False, tope=1.4)
                if curva_b is None or curva_t is None:
                    print("  SIN ancho", round(tramo["ancho"], 2))
                    continue
                texto_w = f"WIDTH {tramo['n']}={tramo['ancho']:.2f} mm"
                if tramo["lado"] == "der":
                    x_w = sx1 + _ancho_texto(texto_w) + 0.55
                else:
                    x_w = x_col - _ancho_texto(texto_w) - 0.4
                _juego_ordenado(
                    hoja, tg, inv,
                    [
                        (_intent_linea(hoja, curva_b, "izq"), "0.00", 0.0, x_w, sy_b),
                        (
                            _intent_linea(hoja, curva_t, "izq"),
                            texto_w,
                            tramo["ancho"],
                            x_w,
                            sy_t,
                        ),
                    ],
                    K_VER, x_w, sy_b, nota_origen=False,
                )
        filas = []
    izq, der = [], []
    anterior_izq = None
    for valor, miembros in filas:
        ancla = min(
            miembros,
            key=lambda m: min(
                abs(_sheet(vista, tg, m[1], m[2])[0] - sx0),
                abs(_sheet(vista, tg, m[1], m[2])[0] - sx1),
            ),
        )
        sx_a = _sheet(vista, tg, ancla[1], ancla[2])[0]
        letra = _letra_seq(n_letra)
        n_letra += 1
        # De abajo hacia arriba: un lado y el siguiente al otro, para que no se monten.
        if anterior_izq is None:
            a_izq = abs(sx_a - sx0) <= abs(sx_a - sx1)
        else:
            a_izq = not anterior_izq
        anterior_izq = a_izq
        lado = izq if a_izq else der
        lado.append((letra, valor, miembros, ancla))
    def _apilar(lado, hacia_izq):
        n = len(lado)
        if n <= 0:
            return
        paso = min(1.15, 2.3 / n)
        for i, (letra, valor, miembros, ancla) in enumerate(lado):
            sx, sy = _sheet(vista, tg, ancla[1], ancla[5])
            tope_y = 1.2 if al_centro else _tope_orilla(ancla[3])
            objetivo = _cerca(curvas, sx, sy, arco=True, tope=tope_y)
            if objetivo is None:
                objetivo = _cerca(horiz, sx, sy, arco=False, vertical=False, tope=0.45)
            if borde_y is None or objetivo is None:
                print("  SIN Y", valor)
                continue
            if hacia_izq:
                x_cota = sx0 - (1.15 + i * paso)
                sep = -0.65
            else:
                x_cota = sx1 + (1.05 + i * paso)
                sep = 0.45
            typ = " TYP" if len(miembros) >= 2 else ""
            _buscar_par(
                hoja, tg, inv, [borde_y], [objetivo], valor, K_VER,
                f"{letra}={valor:.2f}{typ} mm",
                x_cota, (sy0 + sy) / 2.0,
                "inf", "centro" if al_centro else "inf", tol=1.2, vertical=True, separacion=sep,
                cerca_x=x_cota,
            )
            if len(miembros) >= 2 and not al_centro:
                _anillos_referencia(
                    hoja, tg, inv, vista,
                    [(m[1], m[5], m[3]) for m in miembros], "Y",
                )
    _apilar(izq, True)
    _apilar(der, False)
    n_der = len(der)
    paso_der = min(1.15, 2.3 / n_der) if n_der else 0.0
    x_width = sx1 + 1.05 + n_der * paso_der + 1.15
    if x_width > sx1 + 3.35:
        x_width = sx1 + 3.35
    if local and not al_centro:
        for tramo in extremos:
            sx_t = _sheet(vista, tg, tramo["x"], (tramo["y0"] + tramo["y1"]) / 2.0)[0]
            tramo["lado"] = "izq" if abs(sx_t - sx0) <= abs(sx_t - sx1) else "der"
        orden = [t for lado in ("izq", "der") for t in extremos if t["lado"] == lado]
        paso_izq = min(1.15, 2.3 / len(izq)) if izq else 0.0
        fuera_izq = 1.15 + max(len(izq) - 1, 0) * paso_izq
        if al_centro:
            # La columna Y ya ocupa este tramo. Las WIDTH salen más afuera,
            # con el texto horizontal para que quepa junto a esa columna.
            fuera_izq = max(fuera_izq, 1.7 + 2.5)
        fuera_der = 1.05 + max(n_der - 1, 0) * paso_der
        if al_centro:
            x_izq = sx0 - fuera_izq - 0.35
            x_der = sx1 + fuera_der + 0.45
            for i, tramo in enumerate(orden, start=1):
                texto_w = f"WIDTH {i}={tramo['ancho']:.2f} mm"
                if tramo["lado"] == "izq":
                    _cota_ancho_local(
                        hoja, vista, tg, inv, horiz, tramo, texto_w,
                        x_izq, -0.08, plano=True,
                    )
                    x_izq -= _ancho_texto(texto_w) + 0.45
                else:
                    _cota_ancho_local(
                        hoja, vista, tg, inv, horiz, tramo, texto_w,
                        x_der, 0.08, plano=True,
                    )
                    x_der += _ancho_texto(texto_w) + 0.45
            if borde_y and borde_sup and all(abs(ancho - u) > 3.0 for u in distintos):
                texto_total = f"WIDTH TOTAL={ancho:.2f} mm"
                _buscar_par(
                    hoja, tg, inv, [borde_y], [borde_sup], ancho, K_VER,
                    texto_total, x_izq, (sy0 + sy1) / 2.0,
                    "inf", "sup", tol=2.0, separacion=-0.08, plano=True,
                )
        else:
            x_w1 = sx0 - (fuera_izq + 1.6)
            x_w2 = sx1 + (fuera_der + 1.6)
            x_total = x_w1 - 1.7
            for i, tramo in enumerate(orden, start=1):
                texto_w = f"WIDTH {i}={tramo['ancho']:.2f} mm"
                if tramo["lado"] == "izq":
                    _cota_ancho_local(
                        hoja, vista, tg, inv, horiz, tramo, texto_w,
                        x_w1, -0.55,
                    )
                else:
                    _cota_ancho_local(
                        hoja, vista, tg, inv, horiz, tramo, texto_w,
                        x_w2, 0.5,
                    )
            if borde_y and borde_sup and all(abs(ancho - u) > 3.0 for u in distintos):
                _buscar_par(
                    hoja, tg, inv, [borde_y], [borde_sup], ancho, K_VER,
                    f"WIDTH TOTAL={ancho:.2f} mm", x_total, (sy0 + sy1) / 2.0,
                    "inf", "sup", tol=2.0, vertical=True, separacion=-0.55,
                )
    elif borde_y and borde_sup and not al_centro:
        _buscar_par(
            hoja, tg, inv, [borde_y], [borde_sup], ancho, K_VER,
            f"WIDTH={ancho:.2f} mm", x_width, (sy0 + sy1) / 2.0,
            "inf", "sup", tol=2.0, vertical=True, separacion=0.5,
        )
    dims = hoja.DrawingDimensions.GeneralDimensions
    por_diam = _agrupar(
        [(round(d, 2), cx, cy, d) for cx, cy, d, *_resto in huecos],
        lambda t: t[0],
        0.05,
    )
    por_diam.sort(key=lambda g: (-len(g[1]), -g[0]))
    for i, (diam, miembros) in enumerate(por_diam):
        letra = _letra_seq(n_letra)
        n_letra += 1
        if al_centro:
            h = min(
                miembros,
                key=lambda m: (
                    min(
                        abs(_sheet(vista, tg, m[1], m[2])[0] - sx0),
                        abs(_sheet(vista, tg, m[1], m[2])[0] - sx1),
                    ),
                    -_sheet(vista, tg, m[1], m[2])[1],
                ),
            )
        else:
            h = miembros[0]
        hueco = next(
            (t for t in huecos if abs(t[0] - h[1]) < 0.4 and abs(t[1] - h[2]) < 0.4),
            None,
        )
        if al_centro and hueco is not None:
            radios = _centros_radios(*hueco)
            elegido = min(
                radios,
                key=lambda p: min(
                    abs(_sheet(vista, tg, p[0], p[1])[0] - sx0),
                    abs(_sheet(vista, tg, p[0], p[1])[0] - sx1),
                ),
            )
            sx, sy = _sheet(vista, tg, elegido[0], elegido[1])
        else:
            sx, sy = _sheet(vista, tg, h[1], h[2])
        arco = _cerca(curvas, sx, sy, arco=True)
        if arco is None:
            print("  SIN diam", diam)
            continue
        intent = _intent_linea(hoja, arco, "sup")
        if intent is None:
            continue
        typ = " TYP" if len(miembros) >= 2 else ""
        texto = f"{letra}=Ø{diam:.2f}{typ} mm"
        tope = _borde_ranura(curvas, sx, sy) if al_centro else None
        if al_centro and tope is None:
            try:
                radio = float(arco["c"].ModelGeometry.Radius) * float(_ESCALA)
                tope = float(arco["c"].CenterPoint.Y) + radio
            except Exception:
                tope = sy
        try:
            if al_centro:
                ptx, pty = sx, sy1 + 0.55
            else:
                ptx = sx0 + (i + 1) * (sx1 - sx0) / (len(por_diam) + 1)
                pty = sy1 + 1.8 + (i % 2) * 0.7
            dim = dims.AddDiameter(tg.CreatePoint2d(ptx, pty), intent)
        except Exception as exc:
            print("  diam", texto, exc)
            continue
        medido = _mm(dim)
        if al_centro:
            try:
                dim.Delete()
            except Exception as exc:
                print("  diam borrar", exc)
            y_nota = sy1 + 0.72
            _lider_sobre_centro(hoja, tg, inv, sx, tope, y_nota - 0.28)
            _nota_color(hoja, tg, sx, y_nota, texto, (68, 68, 68), ancla="centro")
            print("  COTA", texto, round(medido or 0, 2))
            continue
        _texto(dim, inv, texto, diametro=True)
        _linea_solida(dim, inv)
        print("  COTA", texto, round(medido or 0, 2))
    _zonas_flat(hoja, vista, tg, huecos, fp)
    if thk and thk > 0.2:
        try:
            y_ultima = sy0 - 0.7 - max(len(cols) - 1, 0) * paso_x
            y_thk = y_ultima - 0.62
            if al_centro:
                y_thk = min(y_thk, sy0 - 4.2)
            if y_thk < piso:
                y_thk = piso
            if al_centro:
                _nota_color(
                    hoja, tg, (sx0 + sx1) / 2.0, y_thk,
                    f"THK={thk:.2f} mm", (68, 68, 68), ancla="centro",
                )
            else:
                sketch = hoja.Sketches.Add()
                sketch.Edit()
                caja = sketch.TextBoxes.AddFitted(
                    tg.CreatePoint2d((sx0 + sx1) / 2.0 - 0.7, y_thk),
                    f"THK={thk:.2f} mm",
                )
                from cota_estilo import aplicar_estilo_texto_cota, armar_formatted_texto_cota

                aplicar_estilo_texto_cota(caja, f"THK={thk:.2f} mm", inv, vertical=False)
                formado = armar_formatted_texto_cota(
                    f"THK={thk:.2f} mm", font_cm=_font(), bold=True
                )
                caja.FormattedText = formado
                sketch.Visible = True
                sketch.ExitEdit()
        except Exception as exc:
            print("  thk", exc)
    _volcar_trazos(hoja, tg, inv)


def _zonas_flat(hoja, vista, tg, huecos, fp):
    """Letra de cada panel del desarrollo, el mismo criterio que la isométrica."""
    if fp is None:
        return
    paneles = _letras_paneles(_paneles_flat(fp, huecos))
    for panel in paneles:
        x, y = panel["cx"], panel["cy"]
        for h in huecos:
            if abs(h[0] - x) < 16.0 and abs(h[1] - y) < 16.0:
                x = min(panel["x1"] - 10.0, x + 22.0)
                break
        print("  zona", panel["letra"], "huecos", panel["huecos"], "x", round(x, 1))
        _nota(hoja, vista, tg, x, y, 0, panel["letra"])


def _flat(part):
    sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
    try:
        if not bool(sm.HasFlatPattern):
            try:
                sm.Unfold()
            except Exception as exc:
                print("  unfold", exc)
        if not bool(sm.HasFlatPattern):
            cara = None
            mejor = 0.0
            body = sm.SurfaceBodies.Item(1)
            for i in range(1, int(body.Faces.Count) + 1):
                face = body.Faces.Item(i)
                try:
                    area = float(face.Evaluator.Area)
                except Exception:
                    continue
                if area > mejor:
                    mejor = area
                    cara = face
            if cara is not None:
                try:
                    sm.Unfold2(cara)
                except Exception as exc:
                    print("  unfold2", exc)
    except Exception as exc:
        print("  flat", exc)
    try:
        if not bool(sm.HasFlatPattern):
            return None
        return sm.FlatPattern
    except Exception:
        return None


def _reset_dibujo(plano):
    global _PLANO, _CAPA_SOLIDA, _TRAZOS, _DEFS_GLOBO
    _PLANO = plano
    _CAPA_SOLIDA = None
    _TRAZOS = []
    _DEFS_GLOBO = {}


def _hoja_base(plano):
    base = plano.Sheets.Item(1)
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        if "MODELO" in str(h.Name).upper():
            return h
    return base


def acotar_pieza(inv, plano, base, part, stem, destino):
    """Iso en todas. Doblado solo si hay dobleces. Corte siempre."""
    from creador_vistas import crear_camara

    global _TRAZOS
    _TRAZOS = []
    os.makedirs(destino, exist_ok=True)
    tg = inv.TransientGeometry
    to = inv.TransientObjects
    dobleces = _tiene_dobleces(part)
    grupos = _letras(_grupos(part))
    if not dobleces and grupos:
        principal = max(grupos, key=lambda g: g.get("area", 0))
        principal["letra"] = "A"
        grupos = [principal]
    for g in grupos:
        print(
            stem, g.get("letra"), "huecos", g.get("huecos"),
            "c", tuple(round(g[k], 1) for k in ("cx", "cy", "cz")),
        )
    sm = win32com.client.CastTo(part.ComponentDefinition, "SheetMetalComponentDefinition")
    body = sm.SurfaceBodies.Item(1)
    rb = body.RangeBox
    cx = (float(rb.MinPoint.X) + float(rb.MaxPoint.X)) / 2.0
    cy = (float(rb.MinPoint.Y) + float(rb.MaxPoint.Y)) / 2.0
    cz = (float(rb.MinPoint.Z) + float(rb.MaxPoint.Z)) / 2.0
    dx = abs(float(rb.MaxPoint.X) - float(rb.MinPoint.X))
    dy = abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y))
    dz = abs(float(rb.MaxPoint.Z) - float(rb.MinPoint.Z))
    thk = _espesor_mm(part)
    hechas = []
    solo_corte = os.environ.get("COTAS_COBRE_SOLO_CORTE") == "1"
    if not solo_corte:
        _reset_dibujo(plano)
        h_iso = _hoja(plano, base, f"{stem}_ISO_ABC")
        v_iso = _vista(
            h_iso, part, tg, to,
            crear_camara(part, tg, to, cx, cy, cz, tg.CreateVector(0.9, 0.55, 0.85), tg.CreateVector(0, 1, 0)),
            False, max(dx, dz), max(dy, dz), margen_x=4.5, margen_y=4.5,
        )
        _globos_iso(h_iso, v_iso, tg, inv, grupos)
        ruta = os.path.join(destino, f"{stem}__ISO_ABC.jpg")
        _exportar(inv, to, h_iso, ruta)
        hechas.append(ruta)
        _respirar(2.0)
        if os.environ.get("COTAS_COBRE_SOLO_ISO") == "1":
            return hechas
        if dobleces:
            h_dob = _hoja(plano, base, f"{stem}_DOBLADO_ABC")
            ojo, arriba = _vista_doblez(part)
            if abs(ojo[1]) >= 0.5:
                ancho_p, alto_p = dx, dz
            elif abs(ojo[0]) >= 0.5:
                ancho_p, alto_p = dz, dy
            else:
                ancho_p, alto_p = dx, dy
            v_dob = _vista(
                h_dob, part, tg, to,
                crear_camara(
                    part, tg, to, cx, cy, cz,
                    tg.CreateVector(*ojo), tg.CreateVector(*arriba),
                ),
                False, ancho_p, alto_p, margen_x=8.0, margen_y=4.0,
            )
            _acotar_doblado(h_dob, v_dob, tg, inv, part, grupos)
            ruta = os.path.join(destino, f"{stem}__DOBLADO_ABC.jpg")
            _exportar(inv, to, h_dob, ruta)
            hechas.append(ruta)
            _respirar(2.0)
    fp = _flat(part)
    if fp is None:
        print(stem, "SIN FLAT")
        return hechas
    fr = fp.Body.RangeBox
    fcx = (float(fr.MinPoint.X) + float(fr.MaxPoint.X)) / 2.0
    fcy = (float(fr.MinPoint.Y) + float(fr.MaxPoint.Y)) / 2.0
    fcz = (float(fr.MinPoint.Z) + float(fr.MaxPoint.Z)) / 2.0
    fw = abs(float(fr.MaxPoint.X) - float(fr.MinPoint.X))
    fh = abs(float(fr.MaxPoint.Y) - float(fr.MinPoint.Y))
    desfase = _desfase_y(fp)
    extremos = _extremos_ancho(fp)
    bajo = min(extremos, key=lambda t: t["y0"]) if extremos else None
    # "der" es el X chico. Esa punta va a la izquierda solo si es la que apoya abajo.
    poner_xmin_izq = bool(desfase and bajo is not None and bajo["lado"] == "der")
    ojo_z = 1.0 if poner_xmin_izq else -1.0
    margen_x = 14.0 if desfase else 9.0
    h_cor = _hoja(plano, base, f"{stem}_CORTE_ABC")
    v_cor = _vista(
        h_cor, part, tg, to,
        crear_camara(
            part, tg, to, fcx, fcy, fcz,
            tg.CreateVector(0, 0, ojo_z), tg.CreateVector(0, 1, 0),
        ),
        True, fw, fh, margen_x=margen_x, margen_y=7.0,
    )
    huecos = _huecos_flat(fp.Body)
    print("huecos", [(round(a, 1), round(b, 1), round(c, 2)) for a, b, c, *_ in huecos])
    # Corte de todo el cobre: ordenada al centro de cada radio, X en #043D56 y cotas en #444444.
    global _COLOR_LINEA, _COLOR_MARCA, _ESTILO_COLOR
    _ESTILO_COLOR = None
    _COLOR_LINEA = (68, 68, 68)
    _COLOR_MARCA = (4, 61, 86)
    _acotar_corte(
        h_cor, v_cor, tg, inv, huecos, fr, grupos, thk, fp,
        origen_xmin=poner_xmin_izq, al_centro=True,
    )
    ruta = os.path.join(destino, f"{stem}__CORTE_ABC.jpg")
    _exportar(inv, to, h_cor, ruta)
    _COLOR_LINEA = None
    _COLOR_MARCA = None
    _ESTILO_COLOR = None
    hechas.append(ruta)
    return hechas


def acotar_cobre_ensamble(inv, ensamble=None, destino=None, solo=None):
    """
    Acota el cobre del ensamble abierto. Escribe JPG locales y no registra en la DB.
    ``solo`` es una lista de stems; vacío procesa todo el cobre del ensamble.
    """
    from creador_vistas import configurar_producto_flujo, set_nombre_pieza_completo
    from cota_estilo import set_unidad_cota
    from piezas_cobre import es_pieza_cobre
    from wing_cobre import _asm, _piezas

    configurar_producto_flujo("BOARD")
    set_nombre_pieza_completo(True)
    set_unidad_cota("mm")
    destino = destino or TMP
    os.makedirs(destino, exist_ok=True)
    asm = _asm(inv, ensamble)
    if asm is None:
        print("COBRE ABC: no hay ensamble")
        return []
    filtro = {s.strip().upper() for s in (solo or []) if str(s).strip()}
    env_solo = os.environ.get("COTAS_COBRE_ABC_SOLO", "").strip()
    if env_solo:
        filtro.update(t.strip().upper() for t in env_solo.replace(";", ",").split(",") if t.strip())
    piezas = []
    for stem, part in _piezas(asm):
        if not es_pieza_cobre(stem):
            continue
        if filtro and stem.upper() not in filtro:
            continue
        piezas.append((stem, win32com.client.CastTo(part, "PartDocument")))
    print("COBRE ABC", len(piezas), "piezas")
    if not piezas:
        return []
    previo = bool(getattr(inv, "SilentOperation", False))
    inv.SilentOperation = True
    _cerrar_copias(inv)
    _respirar(1.5)
    plano = _abrir_copia(inv)
    _reset_dibujo(plano)
    _respirar(1.5)
    base = _hoja_base(plano)
    informe = []
    try:
        for i, (stem, part) in enumerate(piezas, start=1):
            print(f"--- {i}/{len(piezas)} {stem}")
            try:
                rutas = acotar_pieza(inv, plano, base, part, stem, destino)
                informe.append((stem, "OK", len(rutas)))
            except Exception as exc:
                texto = str(exc)
                print("  FALLO", stem, texto)
                informe.append((stem, "FALLO", texto))
                if "RPC" in texto.upper() or "2147023174" in texto or "2147023170" in texto:
                    print("COBRE ABC: Inventor no responde, se detiene el lote")
                    break
            _respirar(2.0)
    finally:
        _cerrar_copias(inv)
        try:
            inv.SilentOperation = previo
        except Exception:
            pass
    ruta_inf = os.path.join(destino, "_informe.txt")
    with open(ruta_inf, "w", encoding="utf-8") as f:
        for fila in informe:
            f.write("\t".join(str(x) for x in fila) + "\n")
    print("informe", ruta_inf)
    return informe


def main():
    pythoncom.CoInitialize()
    raw = pythoncom.GetActiveObject("Inventor.Application")
    disp = raw.QueryInterface(pythoncom.IID_IDispatch)
    inv = win32com.client.Dispatch(disp)
    acotar_cobre_ensamble(inv)


if __name__ == "__main__":
    main()
