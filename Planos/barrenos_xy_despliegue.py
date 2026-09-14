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

from cota_estilo import texto_cota_limpio, asegurar_unidad_cota
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
    try:
        if bool(cdef.HasFlatPattern):
            fp = cdef.FlatPattern
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


def _centros_barrenos_modelo(vista, tg, sil=None) -> list[dict]:
    """
    Centros REALES de circunferencia/elipse (Circle/Ellipse del sólido → hoja).
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
                # Rechazar centros en el borde (radio de doblez / fantasma)
                from diametro import _centro_interior_silueta

                if not _centro_interior_silueta(
                    cx, cy, minx, maxx, miny, maxy, margen_frac=0.03
                ):
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
            if not _centro_interior_silueta(
                cx, cy, minx, maxx, miny, maxy, margen_frac=0.02
            ):
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
    if solo_flat and modelo:
        # Solo geometría real del flat: sin fantasmas HLR.
        n_oval = sum(1 for b in modelo if str(b.get("tipo") or "") == "oval")
        n_circ = sum(1 for b in modelo if str(b.get("tipo") or "") == "circulo")
        print(
            f"    refs XY (modelo flat): total={len(modelo)} "
            f"(circulos={n_circ} oval={n_oval}) [sin HLR]"
        )
        return modelo

    hlr_circ = _centros_barrenos_hlr(vista)
    ovals = _referencias_oval_xy(vista)
    fused = _fusionar_centros(modelo, hlr_circ, ovals)
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
    if fused:
        print(
            f"    refs XY: total={len(fused)} "
            f"(circulos={n_circ} oval_centro={n_oval})"
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


def _base_despliegue_desde_frente(nombre_hoja: str) -> str:
    base = str(nombre_hoja).rsplit(":", 1)[0]
    return re.sub(r"_FRENTE_[12]$", "", base, flags=re.IGNORECASE)


def acotar_barrenos_xy_despliegue(nombres_frente_ok=None):
    """
    Hojas ``*_DESPLIEGUE_XCENTRO[_TYP]`` / ``*_YCENTRO[_TYP]``
    midiendo al centro de cada circunferencia desde esquina IL.
    """
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
        try:
            vw, vh = float(vista.Width), float(vista.Height)
            if min(vw, vh) > 1e-6:
                ratio = max(vw, vh) / min(vw, vh)
                if ratio >= 12.0 or min(vw, vh) < 0.35:
                    print(
                        f"  ⚠️ {nombre}: vista parece CANTO/THK "
                        f"(W={vw:.2f} H={vh:.2f}) — omitida para XY"
                    )
                    continue
        except Exception:
            pass

        sil = _silueta_placa_vista(vista)
        if not sil:
            print(f"  ⚠️ {nombre}: sin silueta")
            continue
        origen = _origen_il_pieza(vista)
        if origen is None:
            origen_x, _maxx, origen_y, _maxy, _span = sil
        else:
            origen_x, origen_y = origen
            _maxx, _maxy, _span = sil[1], sil[3], sil[4]
            # Guardrail: si el AABB (minx,miny) no toca geometría y nuestro
            # IL sí, loguear — típico jog/S.
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
            if "X" in ejes and dx >= _MIN_DIST_HOJA:
                mems_x.append(
                    {
                        "cx": cx,
                        "cy": cy,
                        "tamaño": tam,
                        "dist_hoja": dx,
                        "clave": _valor_desde_hoja(vista, hoja, origen_x, cx),
                        "tipo": b.get("tipo"),
                    }
                )
            if "Y" in ejes and dy >= _MIN_DIST_HOJA:
                mems_y.append(
                    {
                        "cx": cx,
                        "cy": cy,
                        "tamaño": tam,
                        "dist_hoja": dy,
                        "clave": _valor_desde_hoja(vista, hoja, origen_y, cy),
                        "tipo": b.get("tipo"),
                    }
                )

        grupos_x = _agrupar_coincidencias(mems_x, vista)
        grupos_y = _agrupar_coincidencias(mems_y, vista)
        if not grupos_x and not grupos_y:
            continue
        print(
            f"    coincidencias estrictas: X={len(grupos_x)} Y={len(grupos_y)} "
            f"(tol≈{_tol_coincidencia_hoja(vista, barrenos):.4f} cm hoja)"
        )

        base_nombre = _base_despliegue_desde_frente(nombre)

        def _emitir(eje: str, grupos: list[dict]) -> None:
            nonlocal creadas_nombres
            for idx, grupo in enumerate(grupos, start=1):
                suf = f"{eje}CENTRO_TYP" if grupo["typ"] else f"{eje}CENTRO"
                nombre_nueva = (
                    f"{base_nombre}_{suf}"
                    if len(grupos) <= 1
                    else f"{base_nombre}_{suf}_{idx:02d}"
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
                if origen_n is None:
                    ox, _mx, oy, _my, _ = sil_n
                else:
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
                    print(
                        f"✅ {nombre_nueva}: {eje}centro={val_txt} "
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

        _emitir("X", grupos_x)
        _emitir("Y", grupos_y)

    print(
        f"✅ barrenos_xy_despliegue: {len(creadas_nombres)} hojas X/Y creadas"
    )
    return creadas_nombres
