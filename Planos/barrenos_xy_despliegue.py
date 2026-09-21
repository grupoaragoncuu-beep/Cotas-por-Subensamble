# -*- coding: utf-8 -*-
"""
Barrenos flat (DESPLIEGUE): X/Y + TYP al CENTRO.

- Circulo: centro → X y Y
- Ovalo/slot: centro geometrico (medio entre extremos) → X y Y
- TYP si ≥2 refs comparten la misma distancia (tol geometrica estricta)
"""

from __future__ import annotations

import os
import re
import time

import win32com.client

from cota_estilo import (
    texto_cota_limpio,
    asegurar_unidad_cota,
    set_typ_letras_habilitadas,
    typ_letras_habilitadas,
)
from diametro import (
    _barrenos_en_vista,
    _marcar_barrenos_azules,
    _origen_il_pieza,
    _punto_cerca_contorno,
    _silueta_placa_vista,
    _silueta_vista,
)
from inventor_com import conectar_inventor
from nomenclatura_capturas import formatear_valor_en_nombre

_EPS = 1e-6


def _pieza_desde_hoja_despliegue(nombre_hoja: str) -> str:
    """``PART_DESPLIEGUE_FRENTE_1`` → ``PART``."""
    base = str(nombre_hoja or "").rsplit(":", 1)[0]
    base = re.sub(r"_DESPLIEGUE_FRENTE_[12]$", "", base, flags=re.IGNORECASE)
    if base.upper().endswith("_DESPLIEGUE"):
        base = base[: -len("_DESPLIEGUE")]
    return base.strip() or ""


def _es_cobre_hoja(nombre_hoja: str) -> bool:
    try:
        from piezas_cobre import es_pieza_cobre

        return bool(es_pieza_cobre(_pieza_desde_hoja_despliegue(nombre_hoja)))
    except Exception:
        return False
_MIN_DIST_HOJA = 0.05
_PREFIJO_SKETCH = "COTAS_XY_"


def _escala_vista(vista) -> float:
    try:
        esc = abs(float(vista.Scale))
        if esc > _EPS:
            return esc
    except Exception:
        pass
    return 1.0


def _valor_desde_hoja(vista, hoja, a: float, b: float) -> str:
    try:
        txt = texto_cota_limpio(abs(float(b) - float(a)) / _escala_vista(vista), hoja)
        if txt:
            return formatear_valor_en_nombre(txt)
    except Exception:
        pass
    return formatear_valor_en_nombre(
        abs(float(b) - float(a)) / _escala_vista(vista) / 2.54
    )


def _doc_vista(vista):
    for getter in (
        lambda: vista.ReferencedDocumentDescriptor.ReferencedDocument,
        lambda: vista.ReferencedFile.DocumentDescriptor.ReferencedDocument,
    ):
        try:
            doc = getter()
            if doc is not None:
                return doc
        except Exception:
            continue
    return None


def _cuerpos_para_barrenos(doc) -> list:
    """Preferir FlatPattern (DESPLIEGUE). En SOLO_FLAT no cae al sólido doblado."""
    cuerpos = []
    if doc is None:
        return cuerpos
    solo_flat = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    try:
        cdef = doc.ComponentDefinition
    except Exception:
        return cuerpos
    # HasFlatPattern vive en SheetMetalComponentDefinition (no Part*).
    sm = None
    try:
        sm = win32com.client.CastTo(cdef, "SheetMetalComponentDefinition")
    except Exception:
        sm = None
    try:
        src = sm if sm is not None else cdef
        if bool(getattr(src, "HasFlatPattern", False)):
            fp = src.FlatPattern
            try:
                cuerpos.append(fp.Body)
            except Exception:
                pass
            try:
                for i in range(1, int(fp.SurfaceBodies.Count) + 1):
                    cuerpos.append(fp.SurfaceBodies.Item(i))
            except Exception:
                pass
    except Exception:
        pass
    if not cuerpos and not solo_flat:
        try:
            for i in range(1, int(cdef.SurfaceBodies.Count) + 1):
                cuerpos.append(cdef.SurfaceBodies.Item(i))
        except Exception:
            pass
    # Flat vacío (pieza sin FlatPattern usable): caer a cuerpos plegados
    # para no dejar refs XY en cero (audit fallaba todo).
    if not cuerpos and solo_flat:
        try:
            for i in range(1, int(cdef.SurfaceBodies.Count) + 1):
                cuerpos.append(cdef.SurfaceBodies.Item(i))
        except Exception:
            pass
    # únicos por id
    out, vistos = [], set()
    for c in cuerpos:
        try:
            kid = id(c)
        except Exception:
            kid = None
        if kid in vistos:
            continue
        vistos.add(kid)
        out.append(c)
    return out


def _cortes_internos_inicio(vista, tg, sil=None) -> list[dict]:
    """
    Cortes internos NO circulares (cuadrado/rectángulo/poligonal).

    Como subensamble: se acotan por INICIO → Xmin (izq) e Ymin (inf)
    del bbox del hueco, no por centro.
    """
    if sil is None:
        sil = _silueta_vista(vista) or _silueta_placa_vista(vista)
    if not sil:
        return []
    minx, maxx, miny, maxy, span = sil
    if span <= _EPS:
        return []
    doc = _doc_vista(vista)
    if doc is None or tg is None:
        return []

    segs = []
    for body in _cuerpos_para_barrenos(doc):
        try:
            n_edges = int(body.Edges.Count)
        except Exception:
            continue
        for j in range(1, n_edges + 1):
            try:
                geom = body.Edges.Item(j).Geometry
                if geom is None:
                    continue
                tipo_g = str(type(geom)).upper()
                if "LINE" not in tipo_g:
                    continue
                p1 = geom.StartPoint
                p2 = geom.EndPoint
                s1 = vista.ModelToSheetSpace(
                    tg.CreatePoint(float(p1.X), float(p1.Y), float(p1.Z))
                )
                s2 = vista.ModelToSheetSpace(
                    tg.CreatePoint(float(p2.X), float(p2.Y), float(p2.Z))
                )
                segs.append(
                    (float(s1.X), float(s1.Y), float(s2.X), float(s2.Y))
                )
            except Exception:
                continue
    if len(segs) < 4:
        return []

    # Solo segmentos cuyo punto medio cae claramente DENTRO de la placa
    # (huecos internos; no contorno exterior).
    margen = max(0.04, 0.025 * span)
    interior = []
    for x1, y1, x2, y2 in segs:
        cx = 0.5 * (x1 + x2)
        cy = 0.5 * (y1 + y2)
        if (minx + margen) < cx < (maxx - margen) and (
            miny + margen
        ) < cy < (maxy - margen):
            interior.append((x1, y1, x2, y2))
    if len(interior) < 4:
        return []

    # Componentes conexas por proximidad de extremos
    tol = max(0.04, 0.008 * span)
    used = [False] * len(interior)
    clusters: list[tuple[float, float, float, float, int]] = []
    for a in range(len(interior)):
        if used[a]:
            continue
        stack = [a]
        used[a] = True
        comp = [interior[a]]
        while stack:
            i = stack.pop()
            x1, y1, x2, y2 = interior[i]
            ends = ((x1, y1), (x2, y2))
            for b in range(len(interior)):
                if used[b]:
                    continue
                bx1, by1, bx2, by2 = interior[b]
                bends = ((bx1, by1), (bx2, by2))
                hit = False
                for p in ends:
                    for q in bends:
                        if abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol:
                            hit = True
                            break
                    if hit:
                        break
                if hit:
                    used[b] = True
                    stack.append(b)
                    comp.append(interior[b])
        if len(comp) < 4:
            continue
        xs: list[float] = []
        ys: list[float] = []
        for x1, y1, x2, y2 in comp:
            xs.extend((x1, x2))
            ys.extend((y1, y2))
        clusters.append((min(xs), max(xs), min(ys), max(ys), len(comp)))

    # Tamaño mínimo de hueco en hoja (flat: aceptar ~8–9 mm modelo)
    esc = _escala_vista(vista)
    solo_flat = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "si",
        "on",
    )
    min_lado = max(0.03, 0.5 * esc) if solo_flat else max(0.08, 1.0 * esc)
    plate_dx = maxx - minx
    plate_dy = maxy - miny
    out: list[dict] = []
    vistos = set()
    for x0, x1, y0, y1, nseg in clusters:
        dx = x1 - x0
        dy = y1 - y0
        if dx < min_lado or dy < min_lado:
            continue
        # Contorno mal clasificado como "interior"
        if dx > 0.80 * plate_dx and dy > 0.80 * plate_dy:
            continue
        clave = (round(x0, 2), round(y0, 2), round(dx, 2), round(dy, 2))
        if clave in vistos:
            continue
        vistos.add(clave)
        # Punto de acotado = esquina inicio (izq-inf) del hueco
        out.append(
            {
                "cx": float(x0),  # Xmin del corte
                "cy": float(y0),  # Ymin del corte
                "xmin": float(x0),
                "ymin": float(y0),
                "xmax": float(x1),
                "ymax": float(y1),
                "dx": float(dx),  # ancho del hueco (hoja)
                "dy": float(dy),  # largo del hueco (hoja)
                "tamaño": float(min(dx, dy)),
                "fuente": "corte_interno",
                "tipo": "corte",
                "medida": "inicio",
                "ejes": ("X", "Y"),
                "nseg": int(nseg),
            }
        )
    return out


def _edges_loops_interiores(body) -> set:
    """
    Edges que pertenecen a bucles INTERIORES (cortes/barrenos reales).

    Las muescas de canto / contorno exterior viven en IsOuterEdgeLoop y
    no deben contarse como barrenos aunque su geometría sea Circle/Arc.
    """
    interiores: set = set()
    try:
        n_faces = int(body.Faces.Count)
    except Exception:
        return interiores
    for fi in range(1, n_faces + 1):
        try:
            face = body.Faces.Item(fi)
            n_loops = int(face.EdgeLoops.Count)
        except Exception:
            continue
        for li in range(1, n_loops + 1):
            try:
                loop = face.EdgeLoops.Item(li)
                # Outer = contorno / muescas abiertas al borde
                if bool(getattr(loop, "IsOuterEdgeLoop", True)):
                    continue
                for ei in range(1, int(loop.Edges.Count) + 1):
                    try:
                        interiores.add(int(loop.Edges.Item(ei).TransientKey))
                    except Exception:
                        try:
                            interiores.add(id(loop.Edges.Item(ei)))
                        except Exception:
                            pass
            except Exception:
                continue
    return interiores


def _centros_barrenos_modelo(vista, tg, sil=None) -> list[dict]:
    """
    Centros REALES de circunferencia/elipse de CORTES INTERIORES.

    Solo edges de bucles internos del FlatPattern (no muescas de borde).
    """
    if sil is None:
        sil = _silueta_vista(vista)
    if not sil:
        return []
    minx, maxx, miny, maxy, span = sil
    lim = 0.48 * max(span, _EPS)
    escala = _escala_vista(vista)
    doc = _doc_vista(vista)
    centros = []
    vistos = set()

    for body in _cuerpos_para_barrenos(doc):
        edges_ok = _edges_loops_interiores(body)
        # Si no hay loops interiores detectables, no inventar barrenos
        # desde el contorno (evita muescas A/B como Ø/XCENTRO).
        if not edges_ok:
            continue
        try:
            n_edges = int(body.Edges.Count)
        except Exception:
            continue
        for j in range(1, n_edges + 1):
            try:
                edge = body.Edges.Item(j)
                try:
                    key = int(edge.TransientKey)
                except Exception:
                    key = id(edge)
                if key not in edges_ok:
                    continue
                geom = edge.Geometry
                if geom is None:
                    continue
                tipo_g = str(type(geom)).upper()
                es_circ = "CIRCLE" in tipo_g
                es_ell = "ELLIPSE" in tipo_g
                if not es_circ and not es_ell:
                    continue
                radio = float(getattr(geom, "Radius", 0) or 0)
                if radio <= 0 and es_ell:
                    # Elipse: radio menor ≈ tamaño de barreno
                    try:
                        rmaj = float(getattr(geom, "MajorRadius", 0) or 0)
                        rmin = float(getattr(geom, "MinorRadius", 0) or 0)
                        radio = min(rmaj, rmin) if rmin > 0 else rmaj
                    except Exception:
                        radio = 0
                if radio <= 0:
                    continue
                radio_hoja = radio * escala
                # Flat: aceptar Ø chicos (~1 mm modelo)
                solo_flat = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
                    "1",
                    "true",
                    "yes",
                    "si",
                    "on",
                )
                if solo_flat:
                    # Ø1.27 mm @ esc=0.1 → radio hoja ≈ 0.00635 cm
                    min_rh = max(0.002, 0.004 * escala)
                else:
                    min_rh = max(0.008, 0.03 * escala)
                if radio_hoja > lim or radio_hoja < min_rh:
                    continue
                c = geom.Center
                p2 = vista.ModelToSheetSpace(
                    tg.CreatePoint(float(c.X), float(c.Y), float(c.Z))
                )
                cx, cy = float(p2.X), float(p2.Y)
                if not (
                    minx - 0.5 <= cx <= maxx + 0.5
                    and miny - 0.5 <= cy <= maxy + 0.5
                ):
                    continue
                # Contención geométrica adicional (círculo dentro de placa).
                from diametro import _es_barreno_circular_interior

                if not _es_barreno_circular_interior(cx, cy, radio_hoja, sil):
                    continue
                clave = (round(cx, 3), round(cy, 3))
                if clave in vistos:
                    continue
                vistos.add(clave)
                centros.append(
                    {
                        "cx": cx,
                        "cy": cy,
                        "tamaño": 2.0 * radio_hoja,
                        "radio_hoja": radio_hoja,
                        "fuente": "modelo",
                        "tipo": "oval" if es_ell else "circulo",
                    }
                )
            except Exception:
                continue
    return centros


def _elipses_alargadas_hlr(vista) -> list[dict]:
    """
    Elipses HLR alargadas (5124) que _anillos_en_vista descarta por aspecto.
    Centro = bbox; son barrenos ovalados proyectados como elipse completa.
    """
    from diametro import (
        _TIPOS_CIRCULO,
        _min_tam_hoja,
        _centro_interior_silueta,
        kEllipseFullCurve2d,
    )

    sil = _silueta_vista(vista)
    if not sil:
        return []
    minx, maxx, miny, maxy, span = sil
    tope = span * 0.55 if span > 0 else 1e9
    min_tam = _min_tam_hoja(vista)
    out = []
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return out
    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            ct = int(curva.CurveType)
            if ct != int(kEllipseFullCurve2d) and ct not in _TIPOS_CIRCULO:
                continue
            caja = curva.Evaluator2D.RangeBox
            ancho = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
            alto = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
            if ancho < min_tam or alto < min_tam:
                continue
            major, minor = max(ancho, alto), min(ancho, alto)
            if major > tope or minor <= 1e-9:
                continue
            aspect = major / minor
            # Círculos (~1) ya van por _anillos; aquí óvalos 1.2..5
            if aspect < 1.20 or aspect > 5.0:
                continue
            cx = (float(caja.MaxPoint.X) + float(caja.MinPoint.X)) / 2.0
            cy = (float(caja.MaxPoint.Y) + float(caja.MinPoint.Y)) / 2.0
            margen = 0.02
            try:
                if os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
                    "1",
                    "true",
                    "yes",
                    "si",
                    "on",
                ):
                    margen = 0.02
            except Exception:
                pass
            if not _centro_interior_silueta(
                cx, cy, minx, maxx, miny, maxy, margen_frac=margen
            ):
                continue
            # Oval/slot también debe estar contenido (no muesca de canto).
            from diametro import _es_barreno_circular_interior

            if not _es_barreno_circular_interior(cx, cy, float(minor) * 0.5, sil):
                continue
            out.append(
                {
                    "cx": cx,
                    "cy": cy,
                    "tamaño": float(minor),
                    "curva": curva,
                    "fuente": "hlr_oval",
                    "tipo": "oval",
                }
            )
        except Exception:
            continue
    return out


def _centros_barrenos_hlr(vista) -> list[dict]:
    """Círculos HLR (sin óvalos: esos van por extremos/centro)."""
    out = []
    for b in _barrenos_en_vista(vista):
        try:
            if str(b.get("tipo") or "") == "oval":
                continue  # óvalos: ver _referencias_oval_extremos
            out.append(
                {
                    "cx": float(b["cx"]),
                    "cy": float(b["cy"]),
                    "tamaño": float(b.get("tamaño") or 0.2),
                    "curva": b.get("curva"),
                    "fuente": "hlr",
                    "tipo": "circulo",
                    "ejes": ("X", "Y"),
                }
            )
        except Exception:
            continue
    return out


def _referencias_oval_xy(vista) -> list[dict]:
    """
    Ovalos/ranuras -> UN solo punto: el CENTRO geometrico del slot
    (punto medio entre extremos), cotado en X y en Y.

    No se usan inicio/fin de semicirculos: el cliente exige centro
    forzosamente en ambos ejes (igual que un circulo).
    """
    from diametro import _arcos_extremos_ranura, _emparejar_ranuras

    out: list[dict] = []
    extremos = _arcos_extremos_ranura(vista)
    ranuras = _emparejar_ranuras(extremos)
    for r in ranuras:
        tam = float(r.get("tamaño") or 0.2)
        mid_x = float(r["cx"])
        mid_y = float(r["cy"])
        dx = abs(float(r["cx_a"]) - float(r["cx_b"]))
        dy = abs(float(r["cy_a"]) - float(r["cy_b"]))
        out.append(
            {
                "cx": mid_x,
                "cy": mid_y,
                "tamaño": tam,
                "fuente": "hlr_oval",
                "tipo": "oval",
                "ejes": ("X", "Y"),
                "orient": "H" if dx >= dy else "V",
            }
        )
    return out



def _fusionar_centros(*listas, tol=0.12) -> list[dict]:
    """Une fuentes sin duplicar el mismo punto (mismo cx/cy y mismos ejes)."""
    out = []
    for lista in listas:
        for b in lista or []:
            try:
                cx, cy = float(b["cx"]), float(b["cy"])
            except Exception:
                continue
            ejes = tuple(b.get("ejes") or ("X", "Y"))
            dup = False
            for e in out:
                ejes_e = tuple(e.get("ejes") or ("X", "Y"))
                if ejes != ejes_e:
                    continue
                if abs(float(e["cx"]) - cx) <= tol and abs(float(e["cy"]) - cy) <= tol:
                    if e.get("fuente") == "modelo":
                        dup = True
                        break
                    if b.get("fuente") == "modelo":
                        e.clear()
                        e.update(b)
                        e["ejes"] = ejes
                        dup = True
                        break
                    dup = True
                    break
            if not dup:
                nb = dict(b)
                nb["ejes"] = ejes
                out.append(nb)
    return out


def _centros_barrenos(vista, tg) -> list[dict]:
    """
    Referencias para cotas X/Y — siempre al CENTRO:

    - Círculo: centro → X y Y
    - Óvalo: centro del slot (medio entre extremos) → X y Y

    En SOLO_FLAT_CORTE: si hay centros del FlatPattern (modelo), NO se
    mezclan HLR (evita barrenos fantasma de radios/líneas de doblez).
    Cortes internos no circulares → Xmin/Ymin (ver ``medida=inicio``).
    """
    sil = _silueta_placa_vista(vista) or _silueta_vista(vista)
    modelo = _centros_barrenos_modelo(vista, tg, sil)
    for m in modelo:
        m.setdefault("ejes", ("X", "Y"))

    solo_flat = os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    hlr_circ = _centros_barrenos_hlr(vista)
    ovals = _referencias_oval_xy(vista)
    if solo_flat and modelo:
        # Preferir modelo, pero si HLR ve más círculos reales → fusionar
        # (modelo incompleto dejaba esperados vacíos y audit fallaba todo).
        if len(hlr_circ) <= len(
            [b for b in modelo if str(b.get("tipo") or "") == "circulo"]
        ) + 1:
            n_oval = sum(1 for b in modelo if str(b.get("tipo") or "") == "oval")
            n_circ = sum(1 for b in modelo if str(b.get("tipo") or "") == "circulo")
            print(
                f"    refs XY (modelo flat): total={len(modelo)} "
                f"(circulos={n_circ} oval={n_oval}) [sin HLR]"
            )
            fused = list(modelo)
        else:
            fused = _fusionar_centros(modelo, hlr_circ, ovals)
    else:
        fused = _fusionar_centros(modelo, hlr_circ, ovals)

    # Cortes internos no circulares (solo flat / cuando no hay círculos
    # suficientes, o siempre como complemento).
    cortes = []
    try:
        cortes = _cortes_internos_inicio(vista, tg, sil)
    except Exception:
        cortes = []
    if cortes:
        # No duplicar un corte cuyo bbox cubre un círculo ya detectado
        circ_pts = [
            (float(b["cx"]), float(b["cy"]))
            for b in fused
            if str(b.get("tipo") or "") in ("circulo", "oval", "")
        ]
        for c in cortes:
            x0, y0 = float(c["xmin"]), float(c["ymin"])
            x1, y1 = float(c["xmax"]), float(c["ymax"])
            tapa_circ = False
            for cx, cy in circ_pts:
                if x0 - 0.05 <= cx <= x1 + 0.05 and y0 - 0.05 <= cy <= y1 + 0.05:
                    tapa_circ = True
                    break
            if not tapa_circ:
                fused.append(c)
        if any(str(b.get("tipo") or "") == "corte" for b in fused):
            n_c = sum(1 for b in fused if str(b.get("tipo") or "") == "corte")
            print(f"    refs XY cortes internos (inicio Xmin/Ymin): {n_c}")

    # Descartar centros que no caen cerca del contorno/interior de placa
    limpios = []
    for b in fused:
        try:
            cx, cy = float(b["cx"]), float(b["cy"])
        except Exception:
            continue
        if sil:
            minx, maxx, miny, maxy, span = sil
            m = max(0.02, 0.01 * span)
            if not (minx - m <= cx <= maxx + m and miny - m <= cy <= maxy + m):
                continue
        limpios.append(b)
    fused = limpios
    n_oval = sum(1 for b in fused if str(b.get("tipo") or "") == "oval")
    n_circ = sum(1 for b in fused if str(b.get("tipo") or "") == "circulo")
    n_corte = sum(1 for b in fused if str(b.get("tipo") or "") == "corte")
    if fused:
        if os.environ.get("COTAS_DEBUG_REFS", "").strip() in ("1", "true", "yes"):
            print(
                f"    refs XY: total={len(fused)} "
                f"(circulos={n_circ} oval_centro={n_oval} corte={n_corte})"
            )
    return fused


def _tol_coincidencia_hoja(vista, miembros=None) -> float:
    """
    Tolerancia ESTRICTA en cm de hoja para misma coincidencia X o Y.

    ~0.25 mm de modelo × escala. El piso también escala: a Scale=0.1 un
    piso fijo 0.012 cm (=1.2 mm modelo) mezclaba círculo 1.500" con extremo
    de óvalo 1.52".
    """
    esc = _escala_vista(vista)
    tol = 0.025 * esc  # 0.25 mm modelo → hoja
    if miembros:
        try:
            r_min = 0.5 * min(float(m.get("tamaño") or 0.2) for m in miembros)
            tol = min(tol, max(0.010 * esc, r_min * 0.15))
        except Exception:
            pass
    piso = max(0.0015, 0.010 * esc)
    techo = max(0.04, 0.08 * min(esc, 1.0))
    return max(piso, min(techo, tol))


def _agrupar_coincidencias(miembros: list[dict], vista) -> list[dict]:
    """
    Una coincidencia = misma distancia geométrica en ese eje (no por texto).

    TYP solo si ≥2 barrenos tienen la MISMA coordenada (tol ~0.25 mm modelo).
    Sin encadenado: el span (max-min) del bucket debe quedar ≤ tol.
    """
    if not miembros:
        return []
    tol = _tol_coincidencia_hoja(vista, miembros)
    ordenados = sorted(miembros, key=lambda m: float(m["dist_hoja"]))
    buckets: list[list[dict]] = []
    for m in ordenados:
        d = float(m["dist_hoja"])
        if not buckets:
            buckets.append([m])
            continue
        d_min = min(float(x["dist_hoja"]) for x in buckets[-1])
        d_max = max(float(x["dist_hoja"]) for x in buckets[-1])
        # Solo misma coincidencia si TODO el grupo (incl. m) cabe en tol
        if (max(d_max, d) - min(d_min, d)) <= tol:
            buckets[-1].append(m)
        else:
            buckets.append([m])

    out = []
    for bucket in buckets:
        mems = sorted(
            bucket,
            key=lambda m: (float(m["dist_hoja"]), float(m["cy"]), float(m["cx"])),
        )
        dist_mean = sum(float(m["dist_hoja"]) for m in mems) / len(mems)
        # Span residual: si por redondeo quedó > tol, forzar 1 cota por barreno
        span = float(mems[-1]["dist_hoja"]) - float(mems[0]["dist_hoja"])
        if span > tol and len(mems) > 1:
            for solo in mems:
                out.append(
                    {
                        "clave": str(solo.get("clave") or ""),
                        "dist_hoja": float(solo["dist_hoja"]),
                        "cx": float(solo["cx"]),
                        "cy": float(solo["cy"]),
                        "tamaño": float(solo.get("tamaño") or 0.2),
                        "miembros": [solo],
                        "typ": False,
                        "tol": float(tol),
                    }
                )
            continue
        rep = min(mems, key=lambda m: abs(float(m["dist_hoja"]) - dist_mean))
        out.append(
            {
                "clave": str(rep.get("clave") or ""),
                "dist_hoja": float(dist_mean),
                "cx": float(rep["cx"]),
                "cy": float(rep["cy"]),
                "tamaño": float(rep.get("tamaño") or 0.2),
                "miembros": mems,
                "typ": len(mems) >= 2,
                "tol": float(tol),
            }
        )
    return out


def _borrar_hojas_prefijo(plano, prefijo: str) -> None:
    pref = prefijo.upper()
    try:
        for j in range(plano.Sheets.Count, 0, -1):
            try:
                h = plano.Sheets.Item(j)
                base = str(h.Name).upper().rsplit(":", 1)[0]
                if base == pref or re.match(re.escape(pref) + r"_\d{2}$", base):
                    h.Delete()
            except Exception:
                continue
    except Exception:
        pass


def _limpiar_dims_y_sketches(hoja) -> None:
    for attr in ("GeneralDimensions", "OrdinateDimensions"):
        try:
            col = getattr(hoja.DrawingDimensions, attr)
            for k in range(col.Count, 0, -1):
                try:
                    col.Item(k).Delete()
                except Exception:
                    pass
        except Exception:
            pass
    try:
        sketches = hoja.Sketches
    except Exception:
        try:
            sketches = hoja.DrawingSketches
        except Exception:
            return
    for k in range(sketches.Count, 0, -1):
        try:
            sk = sketches.Item(k)
            nom = str(getattr(sk, "Name", "") or "").upper()
            if nom.startswith(_PREFIJO_SKETCH) or "COTA" in nom:
                sk.Delete()
        except Exception:
            continue


def _pt(sketch, tg, x, y):
    p = tg.CreatePoint2d(float(x), float(y))
    try:
        return sketch.SheetToSketchSpace(p)
    except Exception:
        return p


def _linea(sketch, tg, x1, y1, x2, y2, color=None):
    try:
        ln = sketch.SketchLines.AddByTwoPoints(
            _pt(sketch, tg, x1, y1), _pt(sketch, tg, x2, y2)
        )
        if color is not None:
            for attr in ("OverrideColor", "Color"):
                try:
                    setattr(ln, attr, color)
                    break
                except Exception:
                    continue
        return ln
    except Exception:
        return None


def _texto(sketch, tg, x, y, txt, inv_app, vertical=False):
    try:
        caja = sketch.TextBoxes.AddFitted(_pt(sketch, tg, x, y), str(txt))
        try:
            caja.FormattedText = str(txt)
        except Exception:
            pass
        if vertical:
            try:
                import math

                caja.Rotation = math.pi / 2
            except Exception:
                pass
        try:
            from cota_estilo import aplicar_estilo_texto_cota

            aplicar_estilo_texto_cota(caja, str(txt), inv_app, vertical=False)
        except Exception:
            pass
        return caja
    except Exception:
        return None


def _dibujar_cota_centro_sketch(
    hoja,
    vista,
    tg,
    inv_app,
    eje: str,
    origen_x: float,
    origen_y: float,
    cx: float,
    cy: float,
    texto: str,
    miembros_typ: list | None = None,
):
    """Cota al CENTRO con motor de subensamble; fallback local con Edit()."""
    from generador_caras_tanque import _dibujar_cotas_hv_desde_origen

    dato = {
        "cx": float(cx),
        "cy": float(cy),
        "minx": float(cx),
        "maxx": float(cx),
        "miny": float(cy),
        "maxy": float(cy),
        "barreno": True,
    }
    # Cota = SOLO origen IL → centro representativo (nunca paso entre barrenos).
    # Miembros TYP extras: solo anillos A/B; todos comparten el mismo ``valor``.
    nivel = float(cx) if eje == "X" else float(cy)
    fuentes = miembros_typ if (miembros_typ and len(miembros_typ) >= 1) else [
        {"cx": cx, "cy": cy, "tamaño": 0.2}
    ]
    miembros = []
    for m in fuentes:
        try:
            mcx, mcy = float(m["cx"]), float(m["cy"])
        except Exception:
            continue
        miembros.append(
            {
                "valor": nivel,
                "dato": {
                    "cx": mcx,
                    "cy": mcy,
                    "minx": mcx,
                    "maxx": mcx,
                    "miny": mcy,
                    "maxy": mcy,
                    "barreno": True,
                    "tamaño": float(m.get("tamaño") or 0.2),
                },
                "lado": "centro",
                "pieza_id": f"BAR_{round(mcx, 3)}_{round(mcy, 3)}",
            }
        )
    if not miembros:
        miembros = [
            {
                "valor": nivel,
                "dato": dato,
                "lado": "centro",
                "pieza_id": "BAR",
            }
        ]
    # Representante primero: la línea de cota ancla a (cx,cy) del nivel.
    rep_id = f"BAR_{round(float(cx), 3)}_{round(float(cy), 3)}"
    miembros.sort(key=lambda m: 0 if m.get("pieza_id") == rep_id else 1)

    pos = {
        "valor": nivel,
        "typ": bool(miembros_typ and len(miembros_typ) >= 2),
        "miembros": miembros,
        "dato": dato,
        "lado": "centro",
        "pieza_id": miembros[0]["pieza_id"],
    }

    try:
        if eje == "X":
            creadas, fallos, _vals = _dibujar_cotas_hv_desde_origen(
                hoja, vista, tg, inv_app, [pos], [], float(origen_x), float(origen_y),
                typ_un_anillo=True,
            )
        else:
            creadas, fallos, _vals = _dibujar_cotas_hv_desde_origen(
                hoja, vista, tg, inv_app, [], [pos], float(origen_x), float(origen_y),
                typ_un_anillo=True,
            )
        if int(creadas or 0) >= 1:
            return True
        print(f"    motor subensamble: creadas={creadas} fallos={fallos}")
    except Exception as exc:
        print(f"    motor subensamble fallo: {exc}")

    if _dibujar_cota_centro_sketch_local(
        hoja, vista, tg, inv_app, eje, origen_x, origen_y, cx, cy, texto, miembros_typ
    ):
        return True

    # Último recurso: nota visible con el valor (la cota NO se pierde).
    try:
        if eje == "X":
            pt = tg.CreatePoint2d((origen_x + cx) * 0.5, min(origen_y, cy) - 1.2)
        else:
            pt = tg.CreatePoint2d(min(origen_x, cx) - 1.2, (origen_y + cy) * 0.5)
        nota = hoja.DrawingNotes.GeneralNotes.AddFitted(pt, str(texto))
        try:
            nota.Visible = True
        except Exception:
            pass
        print(f"    cota por GeneralNote (fallback): {texto}")
        return True
    except Exception as exc_n:
        print(f"    GeneralNote tambien fallo: {exc_n}")
        return False


def _dibujar_cota_centro_sketch_local(
    hoja,
    vista,
    tg,
    inv_app,
    eje: str,
    origen_x: float,
    origen_y: float,
    cx: float,
    cy: float,
    texto: str,
    miembros_typ: list | None = None,
):
    """Fallback: sketch.Edit() obligatorio."""
    try:
        sketch = hoja.Sketches.Add()
        sketch.Name = f"{_PREFIJO_SKETCH}{eje}"
        sketch.Edit()
    except Exception as exc1:
        try:
            sketch = hoja.DrawingSketches.Add()
            try:
                sketch.Name = f"{_PREFIJO_SKETCH}{eje}"
            except Exception:
                pass
            sketch.Edit()
        except Exception as exc2:
            print(f"    sketch Add/Edit fallo: {exc1} | {exc2}")
            return False

    color = None
    try:
        color = inv_app.TransientObjects.CreateColor(0, 0, 128)
    except Exception:
        pass

    try:
        left = float(vista.Left)
        right = left + float(vista.Width)
        top = float(vista.Top)
        bot = top - float(vista.Height)
    except Exception:
        left, right, top, bot = origen_x - 1, cx + 1, cy + 1, origen_y - 1

    try:
        sheet_h = float(hoja.Height)
        sheet_w = float(hoja.Width)
    except Exception:
        sheet_h, sheet_w = 40.0, 30.0

    lineas_ok = 0
    txtbox = None
    offset = 2.6
    margen = 0.6
    try:
        if eje == "X":
            y_dim = bot - offset
            if y_dim < margen:
                y_dim = min(sheet_h - margen, top + offset)
            for pts in (
                (origen_x, y_dim, cx, y_dim),
                (origen_x, origen_y, origen_x, y_dim),
                (cx, cy, cx, y_dim),
            ):
                if _linea(sketch, tg, *pts, color):
                    lineas_ok += 1
            txtbox = _texto(
                sketch, tg, (origen_x + cx) * 0.5, y_dim - 0.35, texto, inv_app
            )
        else:
            x_dim = left - offset
            if x_dim < margen:
                x_dim = min(sheet_w - margen, right + offset)
            for pts in (
                (x_dim, origen_y, x_dim, cy),
                (origen_x, origen_y, x_dim, origen_y),
                (cx, cy, x_dim, cy),
            ):
                if _linea(sketch, tg, *pts, color):
                    lineas_ok += 1
            txtbox = _texto(
                sketch, tg, x_dim - 0.35, (origen_y + cy) * 0.5, texto, inv_app,
                vertical=True,
            )
        if miembros_typ and len(miembros_typ) >= 2:
            try:
                color_typ = inv_app.TransientObjects.CreateColor(0, 0, 220)
            except Exception:
                color_typ = color
            for m in miembros_typ:
                try:
                    mcx, mcy = float(m["cx"]), float(m["cy"])
                    r = max(0.12, float(m.get("tamaño") or 0.2) * 0.55)
                    circ = sketch.SketchCircles.AddByCenterRadius(
                        _pt(sketch, tg, mcx, mcy), r
                    )
                    if color_typ is not None:
                        try:
                            circ.Color = color_typ
                        except Exception:
                            pass
                except Exception:
                    continue
    finally:
        try:
            sketch.Visible = True
        except Exception:
            pass
        try:
            sketch.ExitEdit()
        except Exception:
            pass

    ok = (txtbox is not None) or lineas_ok >= 1
    if not ok:
        print(f"    sketch local vacio (lineas={lineas_ok})")
    return bool(ok)



def _forzar_vista_lista(inv_app, hoja, vista) -> None:
    for intento in range(8):
        try:
            hoja.Activate()
        except Exception:
            pass
        if vista is not None:
            try:
                vista.Update()
            except Exception:
                pass
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        time.sleep(0.12 + 0.04 * intento)
        try:
            if vista is not None and int(vista.DrawingCurves.Count) >= 1:
                return
        except Exception:
            pass


def _dibujar_cota_tramo_sketch(
    hoja,
    vista,
    tg,
    inv_app,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    texto: str,
    eje: str,
):
    """
    Cota de tramo (largo/ancho de hueco) con sketch: línea entre (x1,y1)-(x2,y2)
    y texto con el valor.
    """
    try:
        sketch = hoja.Sketches.Add()
        sketch.Name = f"{_PREFIJO_SKETCH}CUT_{eje}"
        sketch.Edit()
    except Exception:
        try:
            sketch = hoja.DrawingSketches.Add()
            sketch.Edit()
        except Exception:
            return False

    color = None
    try:
        color = inv_app.TransientObjects.CreateColor(0, 90, 0)
    except Exception:
        pass

    try:
        left = float(vista.Left)
        right = left + float(vista.Width)
        top = float(vista.Top)
        bot = top - float(vista.Height)
    except Exception:
        left, right, top, bot = min(x1, x2) - 1, max(x1, x2) + 1, max(y1, y2) + 1, min(y1, y2) - 1

    offset = 1.8
    ok = False
    try:
        if eje.upper() == "X":  # ancho horizontal
            y_dim = min(y1, y2) - offset
            if y_dim < bot + 0.3:
                y_dim = max(y1, y2) + offset
            ok = bool(_linea(sketch, tg, x1, y_dim, x2, y_dim, color))
            ok = _linea(sketch, tg, x1, y1, x1, y_dim, color) or ok
            ok = _linea(sketch, tg, x2, y2, x2, y_dim, color) or ok
            _texto(sketch, tg, (x1 + x2) * 0.5, y_dim - 0.3, texto, inv_app)
        else:  # largo vertical
            x_dim = min(x1, x2) - offset
            if x_dim < left + 0.3:
                x_dim = max(x1, x2) + offset
            ok = bool(_linea(sketch, tg, x_dim, y1, x_dim, y2, color))
            ok = _linea(sketch, tg, x1, y1, x_dim, y1, color) or ok
            ok = _linea(sketch, tg, x2, y2, x_dim, y2, color) or ok
            _texto(
                sketch, tg, x_dim - 0.3, (y1 + y2) * 0.5, texto, inv_app, vertical=True
            )
    except Exception as exc:
        print(f"    cut-tramo sketch fallo: {exc}")
        ok = False
    finally:
        try:
            sketch.ExitEdit()
        except Exception:
            pass
    return ok


def _emitir_tamanos_corte(
    plano,
    hoja,
    vista,
    tg,
    inv_app,
    barrenos: list,
    base_nombre: str,
    creadas_nombres: list,
):
    """
    Por cada hueco rectangular: hojas CUT_WIDTH (dx) y CUT_LENGTH (dy).
    """
    cortes = [b for b in barrenos if str(b.get("tipo") or "") == "corte"]
    if not cortes:
        return
    vistos = set()
    for idx, c in enumerate(cortes, start=1):
        try:
            x0, y0 = float(c["xmin"]), float(c["ymin"])
            x1, y1 = float(c["xmax"]), float(c["ymax"])
        except Exception:
            continue
        dx = abs(x1 - x0)
        dy = abs(y1 - y0)
        if dx < _MIN_DIST_HOJA or dy < _MIN_DIST_HOJA:
            continue
        clave = (round(x0, 2), round(y0, 2), round(dx, 2), round(dy, 2))
        if clave in vistos:
            continue
        vistos.add(clave)
        suf = "" if len(cortes) <= 1 else f"_{idx:02d}"
        for etiqueta, eje, a0, b0, a1, b1 in (
            (f"CUT_WIDTH{suf}", "X", x0, y0, x1, y0),
            (f"CUT_LENGTH{suf}", "Y", x0, y0, x0, y1),
        ):
            nombre_nueva = f"{base_nombre}_{etiqueta}"
            _borrar_hojas_prefijo(plano, nombre_nueva)
            try:
                nueva = hoja.CopyTo(plano)
            except Exception as exc:
                print(f"⚠️ {nombre_nueva}: CopyTo falló ({exc})")
                continue
            try:
                nueva.Name = nombre_nueva
            except Exception:
                pass
            if nueva.DrawingViews.Count < 1:
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            vista_n = nueva.DrawingViews.Item(1)
            _forzar_vista_lista(inv_app, nueva, vista_n)
            _limpiar_dims_y_sketches(nueva)
            # Reproyectar bbox del mismo corte
            try:
                cortes_n = [
                    b
                    for b in _centros_barrenos(vista_n, tg)
                    if str(b.get("tipo") or "") == "corte"
                ]
            except Exception:
                cortes_n = []
            mejor = None
            mejor_d = 1e9
            for bn in cortes_n:
                try:
                    bx0, by0 = float(bn["xmin"]), float(bn["ymin"])
                except Exception:
                    continue
                d = (bx0 - x0) ** 2 + (by0 - y0) ** 2
                if d < mejor_d:
                    mejor_d = d
                    mejor = bn
            if mejor is not None:
                try:
                    x0 = float(mejor["xmin"])
                    y0 = float(mejor["ymin"])
                    x1 = float(mejor["xmax"])
                    y1 = float(mejor["ymax"])
                except Exception:
                    pass
            if eje == "X":
                val_txt = _valor_desde_hoja(vista_n, nueva, x0, x1)
                ok = _dibujar_cota_tramo_sketch(
                    nueva, vista_n, tg, inv_app, x0, y0, x1, y0, asegurar_unidad_cota(val_txt), "X"
                )
            else:
                val_txt = _valor_desde_hoja(vista_n, nueva, y0, y1)
                ok = _dibujar_cota_tramo_sketch(
                    nueva, vista_n, tg, inv_app, x0, y0, x0, y1, asegurar_unidad_cota(val_txt), "Y"
                )
            try:
                nota = nueva.DrawingNotes.GeneralNotes.AddFitted(
                    tg.CreatePoint2d(0.4, 0.4),
                    f"CUT={val_txt}",
                )
                try:
                    nota.Visible = False
                except Exception:
                    pass
            except Exception:
                pass
            if ok:
                kind = "WIDTH" if "WIDTH" in etiqueta.upper() else "LENGTH"
                print(f"✅ {nombre_nueva}: {kind}={val_txt} [cut]")
            else:
                print(f"⚠️ {nombre_nueva}: sketch cut falló (hoja conservada)")
            creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])


def _base_despliegue_desde_frente(nombre_hoja: str) -> str:
    base = str(nombre_hoja).rsplit(":", 1)[0]
    return re.sub(r"_FRENTE_[12]$", "", base, flags=re.IGNORECASE)


def acotar_barrenos_xy_despliegue(nombres_frente_ok=None):
    """
    Hojas ``*_DESPLIEGUE_XCENTRO[_TYP]`` / ``*_YCENTRO[_TYP]``
    midiendo al centro de cada circunferencia desde esquina IL.
    """
    # Flat GIGA: mm siempre (cota_estilo default histórico = in de tanques)
    if os.environ.get("SOLO_FLAT_CORTE", "").strip().lower() in (
        "1",
        "true",
        "yes",
        "si",
        "on",
    ):
        from cota_estilo import set_unidad_cota

        set_unidad_cota("mm")
    print(
        "barrenos_xy_despliegue: centros de circulo -> X/Y + TYP "
        "(metodo subensamble / sketch)..."
    )
    inv_app = conectar_inventor()
    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except Exception:
        print("❌ barrenos_xy: no hay plano de Inventor abierto.")
        return []

    tg = inv_app.TransientGeometry
    objetivo = None
    if nombres_frente_ok is not None:
        objetivo = {str(h).upper().rsplit(":", 1)[0] for h in nombres_frente_ok}

    creadas_nombres: list[str] = []
    for i in range(1, plano.Sheets.Count + 1):
        try:
            hoja = plano.Sheets.Item(i)
        except Exception:
            continue
        nombre = str(hoja.Name)
        nombre_up = nombre.upper()
        if "_DESPLIEGUE_FRENTE_1" not in nombre_up:
            continue
        # Nunca acotar barrenos en LADO/THK (canto).
        if "_LADO" in nombre_up or "_THK" in nombre_up:
            continue
        base_cmp = nombre_up.rsplit(":", 1)[0]
        if objetivo is not None and base_cmp not in objetivo:
            continue
        if hoja.DrawingViews.Count < 1:
            continue

        try:
            hoja.Activate()
            inv_app.ActiveView.Update()
            time.sleep(0.1)
        except Exception:
            pass

        vista = hoja.DrawingViews.Item(1)
        # Guardrail: cara principal no puede verse como canto fino.
        # Excepción: si hay barrenos detectables, igual se acota XY
        # (piezas largas/estrechas tipo OP/GS).
        try:
            vw, vh = float(vista.Width), float(vista.Height)
            if min(vw, vh) > 1e-6:
                ratio = max(vw, vh) / min(vw, vh)
                parece_canto = ratio >= 12.0 or min(vw, vh) < 0.35
                if parece_canto:
                    # Probar barrenos antes de omitir
                    try:
                        probe = _centros_barrenos(vista, tg)
                    except Exception:
                        probe = []
                    if not probe:
                        print(
                            f"  ⚠️ {nombre}: vista parece CANTO/THK "
                            f"(W={vw:.2f} H={vh:.2f}) — omitida para XY"
                        )
                        continue
                    print(
                        f"  ⚠️ {nombre}: vista estrecha "
                        f"(W={vw:.2f} H={vh:.2f}) pero {len(probe)} barrenos "
                        f"— se acota XY"
                    )
        except Exception:
            pass

        sil = _silueta_placa_vista(vista)
        if not sil:
            print(f"  ⚠️ {nombre}: sin silueta")
            continue
        from diametro import _origen_flota_en_vacio

        origen = _origen_il_pieza(vista)
        if origen is None or _origen_flota_en_vacio(vista, origen[0], origen[1]):
            print(
                f"  ⚠️ {nombre}: origen IL no ancla a geometria real "
                f"(NO se usa AABB flotante) — omitida"
            )
            continue
        origen_x, origen_y = origen
        _maxx, _maxy, _span = sil[1], sil[3], sil[4]
        aabb_il = (float(sil[0]), float(sil[2]))
        if (
            abs(aabb_il[0] - origen_x) > 0.05
            or abs(aabb_il[1] - origen_y) > 0.05
        ):
            print(
                f"    origen IL sobre geometria "
                f"({origen_x:.3f},{origen_y:.3f}) "
                f"≠ AABB ({aabb_il[0]:.3f},{aabb_il[1]:.3f})"
            )

        barrenos = _centros_barrenos(vista, tg)
        if not barrenos:
            print(f"  ⚠️ {nombre}: sin barrenos (modelo ni HLR)")
            continue

        # TANQUE + iProp Corte: NUNCA flat/barrenos/cortes internos
        # (DEBER_SER_COTAS_FLUJOS §3.1).
        try:
            from creador_vistas import producto_flujo_actual, _es_pieza_corte

            pieza = _pieza_desde_hoja_despliegue(nombre)
            if producto_flujo_actual() == "TANQUE" and _es_pieza_corte(pieza):
                print(
                    f"  {nombre.rsplit(':',1)[0]}: TANQUE/Corte → omitido "
                    f"XY/CUT/HOLE (sin flat ni cortes internos)"
                )
                continue
        except Exception:
            pass

        fuente = barrenos[0].get("fuente", "?")
        print(
            f"  {nombre.rsplit(':',1)[0]}: {len(barrenos)} centros "
            f"({fuente}), origen=({origen_x:.2f},{origen_y:.2f})"
        )

        mems_x, mems_y = [], []
        for b in barrenos:
            cx, cy = float(b["cx"]), float(b["cy"])
            tam = float(b.get("tamaño") or 0.2)
            ejes = tuple(b.get("ejes") or ("X", "Y"))
            dx, dy = cx - origen_x, cy - origen_y
            # Jog/S: origen en pad inferior; barrenos del tramo superior
            # pueden quedar a la IZQUIERDA (dx<0). Igual se acotan (distancia).
            if "X" in ejes and abs(dx) >= _MIN_DIST_HOJA:
                mems_x.append(
                    {
                        "cx": cx,
                        "cy": cy,
                        "tamaño": tam,
                        "dist_hoja": dx,  # con signo: no mezclar izq/der en TYP
                        "clave": _valor_desde_hoja(vista, hoja, origen_x, cx),
                        "tipo": b.get("tipo"),
                        "medida": b.get("medida") or "centro",
                    }
                )
            if "Y" in ejes and abs(dy) >= _MIN_DIST_HOJA:
                mems_y.append(
                    {
                        "cx": cx,
                        "cy": cy,
                        "tamaño": tam,
                        "dist_hoja": dy,
                        "clave": _valor_desde_hoja(vista, hoja, origen_y, cy),
                        "tipo": b.get("tipo"),
                        "medida": b.get("medida") or "centro",
                    }
                )

        grupos_x = _agrupar_coincidencias(mems_x, vista)
        grupos_y = _agrupar_coincidencias(mems_y, vista)
        base_nombre = _base_despliegue_desde_frente(nombre)

        # Largo×ancho del hueco aunque no haya X/Y (p.ej. corte en origen).
        _emitir_tamanos_corte(
            plano,
            hoja,
            vista,
            tg,
            inv_app,
            barrenos,
            base_nombre,
            creadas_nombres,
        )

        if not grupos_x and not grupos_y:
            continue
        print(
            f"    coincidencias estrictas: X={len(grupos_x)} Y={len(grupos_y)} "
            f"(tol≈{_tol_coincidencia_hoja(vista, barrenos):.4f} cm hoja)"
        )

        es_cobre = _es_cobre_hoja(nombre)
        letras_prev = typ_letras_habilitadas()
        if es_cobre:
            # Cobre/busbar: TYP sin nomenclatura A/B/C.
            set_typ_letras_habilitadas(False)

        def _emitir(eje: str, grupos: list[dict]) -> None:
            nonlocal creadas_nombres
            for idx, grupo in enumerate(grupos, start=1):
                # Cortes no circulares → XMIN/YMIN; barrenos → XCENTRO/YCENTRO
                es_inicio = any(
                    str(m.get("medida") or "") == "inicio"
                    or str(m.get("tipo") or "") == "corte"
                    for m in (grupo.get("miembros") or [])
                ) or str(grupo.get("medida") or "") == "inicio"
                # Propagar desde el miembro representativo
                if not es_inicio:
                    for b in barrenos:
                        try:
                            if abs(float(b["cx"]) - float(grupo["cx"])) < 1e-6 and abs(
                                float(b["cy"]) - float(grupo["cy"])
                            ) < 1e-6:
                                if str(b.get("tipo") or "") == "corte" or str(
                                    b.get("medida") or ""
                                ) == "inicio":
                                    es_inicio = True
                                break
                        except Exception:
                            pass
                etiqueta = f"{eje}MIN" if es_inicio else f"{eje}CENTRO"
                suf = f"{etiqueta}_TYP" if grupo["typ"] else etiqueta
                nombre_nueva = (
                    f"{base_nombre}_{suf}"
                    if len(grupos) <= 1
                    else f"{base_nombre}_{suf}_{idx:02d}"
                )
                # Calidad: primeras 2 posiciones CENTRO (no MIN) en cobre.
                if (
                    es_cobre
                    and not es_inicio
                    and idx <= 2
                    and "CENTRO" in etiqueta
                ):
                    print(
                        f"    Seleccionadas(calidad): {nombre_nueva} "
                        f"({eje}#{idx} desde 0,0)"
                    )
                _borrar_hojas_prefijo(plano, nombre_nueva)
                try:
                    nueva = hoja.CopyTo(plano)
                except Exception as exc:
                    print(f"⚠️ {nombre_nueva}: CopyTo falló ({exc})")
                    continue
                try:
                    nueva.Name = nombre_nueva
                except Exception:
                    pass
                if nueva.DrawingViews.Count < 1:
                    print(f"⚠️ {nombre_nueva}: CopyTo sin vistas")
                    try:
                        nueva.Delete()
                    except Exception:
                        pass
                    continue

                vista_n = nueva.DrawingViews.Item(1)
                _forzar_vista_lista(inv_app, nueva, vista_n)
                _limpiar_dims_y_sketches(nueva)

                sil_n = _silueta_placa_vista(vista_n)
                if not sil_n:
                    print(
                        f"⚠️ {nombre_nueva}: sin silueta tras CopyTo "
                        f"(se CONSERVA la hoja)"
                    )
                    creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])
                    continue
                origen_n = _origen_il_pieza(vista_n)
                from diametro import _origen_flota_en_vacio

                if (
                    origen_n is None
                    or _origen_flota_en_vacio(vista_n, origen_n[0], origen_n[1])
                ):
                    print(
                        f"⚠️ {nombre_nueva}: origen IL flotante tras CopyTo "
                        f"(se CONSERVA sin acotar)"
                    )
                    creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])
                    continue
                ox, oy = origen_n
                _mx, _my = sil_n[1], sil_n[3]

                # Reproyectar centros en la hoja copia
                barrenos_n = _centros_barrenos(vista_n, tg)
                cx, cy = float(grupo["cx"]), float(grupo["cy"])
                mejor = None
                mejor_d = 1e9
                for bn in barrenos_n:
                    try:
                        bcx, bcy = float(bn["cx"]), float(bn["cy"])
                    except Exception:
                        continue
                    d = (bcx - cx) ** 2 + (bcy - cy) ** 2
                    if d < mejor_d:
                        mejor_d = d
                        mejor = bn
                if mejor is None:
                    print(
                        f"⚠️ {nombre_nueva}: centro no reproyectado "
                        f"(se CONSERVA la hoja)"
                    )
                    creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])
                    continue
                cx, cy = float(mejor["cx"]), float(mejor["cy"])

                # Miembros TYP: misma distancia en ESTA hoja (ox/oy nuevos).
                # Referencia = centro reproyectado; no usar dist_hoja de la hoja origen.
                miembros_n = [mejor]
                if grupo["typ"]:
                    tol = float(
                        grupo.get("tol")
                        or _tol_coincidencia_hoja(vista_n, barrenos_n)
                    )
                    dist_ref = (
                        (cx - ox) if eje == "X" else (cy - oy)
                    )
                    miembros_n = []
                    for bn in barrenos_n:
                        dist = (
                            float(bn["cx"]) - ox
                            if eje == "X"
                            else float(bn["cy"]) - oy
                        )
                        if abs(dist - dist_ref) <= tol:
                            miembros_n.append(bn)
                    if len(miembros_n) < 2:
                        # No forzar TYP si al reproyectar no hay 2 al mismo nivel
                        miembros_n = [mejor]
                        grupo = dict(grupo)
                        grupo["typ"] = False

                val_txt = grupo["clave"]
                try:
                    if eje == "X":
                        val_txt = _valor_desde_hoja(vista_n, nueva, ox, cx)
                    else:
                        val_txt = _valor_desde_hoja(vista_n, nueva, oy, cy)
                except Exception:
                    pass
                es_typ = bool(grupo.get("typ")) and len(miembros_n) >= 2
                # Si al reproyectar ya no hay TYP, renombrar hoja sin _TYP
                if not es_typ and "_TYP" in nombre_nueva.upper():
                    try:
                        nombre_sin = re.sub(
                            r"_TYP", "", nombre_nueva, flags=re.IGNORECASE
                        )
                        nueva.Name = nombre_sin
                        nombre_nueva = nombre_sin
                    except Exception:
                        pass

                texto_vista = str(val_txt)
                if es_typ:
                    texto_vista = asegurar_unidad_cota(f"{val_txt} TYP")
                else:
                    texto_vista = asegurar_unidad_cota(str(val_txt))

                ok = _dibujar_cota_centro_sketch(
                    nueva,
                    vista_n,
                    tg,
                    inv_app,
                    eje,
                    ox,
                    oy,
                    cx,
                    cy,
                    texto_vista,
                    miembros_typ=miembros_n if es_typ else None,
                )
                # NUNCA borrar la hoja si falla el dibujo: se conserva para revisión.
                if not ok:
                    print(
                        f"⚠️ {nombre_nueva}: sketch falló "
                        f"(hoja CONSERVADA, no se borra)"
                    )
                else:
                    tag = "TYP" if es_typ else "1"
                    kind = "min" if es_inicio else "centro"
                    print(
                        f"✅ {nombre_nueva}: {eje}{kind}={val_txt} "
                        f"({tag}, n={len(miembros_n)}) [sketch]"
                    )

                # Nota con el valor: el exportador lee números de notas si no hay dim
                try:
                    nota = nueva.DrawingNotes.GeneralNotes.AddFitted(
                        tg.CreatePoint2d(0.4, 0.4),
                        f"XY={val_txt}",
                    )
                    try:
                        nota.Visible = False
                    except Exception:
                        pass
                except Exception:
                    pass

                # Marca azul extra (capa diametro) si TYP
                if es_typ and len(miembros_n) >= 2:
                    try:
                        _marcar_barrenos_azules(
                            nueva, vista_n, miembros_n, tg, inv_app
                        )
                    except Exception:
                        pass

                creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])

        try:
            _emitir("X", grupos_x)
            _emitir("Y", grupos_y)
        finally:
            if es_cobre:
                set_typ_letras_habilitadas(letras_prev)

    print(
        f"✅ barrenos_xy_despliegue: {len(creadas_nombres)} hojas X/Y creadas"
    )
    return creadas_nombres
