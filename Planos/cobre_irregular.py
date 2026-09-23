# -*- coding: utf-8 -*-
"""
Cobre irregular (zapato / botella / escalón) — GIGA busbar.

Reglas (pizarrón):
  1. Slots / barrenos / cortes: misma lógica cobre (orilla hacia 0 + Ø).
  2. Ancla = punto del barreno más cercano al origen (por eje).
  3. Cotas de izquierda → derecha.
  4. Pieza en cuadrante superior-derecho.
  5. Origen = esquina inferior-izquierda REAL del contorno.
  6. Siempre diámetros (HOLE).

Extra: WIDTH1 = cuello (tramo más estrecho); WIDTH2 = pad si aporta;
WIDTH_TOTAL = span máximo (FRENTE_2).
"""
from __future__ import annotations

import math
import os
from typing import Any

# Rotaciones a probar (grados). DrawingView.Rotation está en radianes.
_ROTACIONES_DEG = (0.0, 90.0, 180.0, 270.0)

# Diferencia relativa de anchos L/R para detectar escalón.
_RATIO_ESCALON = 0.18
# Mínimo span hoja (cm) para considerar banda válida.
_MIN_BANDA_CM = 0.05


def _leer_rotacion_rad(vista) -> float:
    """DrawingView.Rotation (rad). Fallback 0."""
    try:
        return float(vista.Rotation)
    except Exception:
        return 0.0


def _aplicar_rotacion_rad(vista, rad: float) -> None:
    """Fija Rotation en radianes; si falla, usa RotateByAngle(delta)."""
    rad = float(rad)
    try:
        vista.Rotation = rad
        return
    except Exception:
        pass
    try:
        cur = _leer_rotacion_rad(vista)
        vista.RotateByAngle(rad - cur)
    except Exception as exc:
        raise RuntimeError(f"no se pudo rotar vista: {exc}") from exc


def es_candidato_nombre_irregular(nombre: str) -> bool:
    """
    Filtro rápido por nombre. Env ``COTAS_COBRE_IRREGULAR=1`` fuerza todas
    las cobre; ``=0`` desactiva. Default: cobre del catálogo (se confirma
    por geometría en la vista).
    """
    flag = os.environ.get("COTAS_COBRE_IRREGULAR", "").strip().lower()
    if flag in ("0", "false", "no", "off"):
        return False
    if flag in ("1", "true", "yes", "si", "on", "all"):
        return True
    try:
        from piezas_cobre import es_pieza_cobre

        return bool(es_pieza_cobre(nombre))
    except Exception:
        return False


def _silueta(vista):
    try:
        from diametro import _silueta_vista

        return _silueta_vista(vista, solo_lineas=True) or _silueta_vista(vista)
    except Exception:
        return None


def _segmentos_contorno(vista) -> list[tuple[float, float, float, float]]:
    """Segmentos (x0,y0,x1,y1) de líneas del contorno (sin círculos/arcos)."""
    segs: list[tuple[float, float, float, float]] = []
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return segs
    from diametro import _TIPOS_CIRCULO, _TIPOS_ARCO

    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            try:
                ct = int(curva.CurveType)
            except Exception:
                ct = None
            if ct in _TIPOS_CIRCULO or ct in _TIPOS_ARCO:
                continue
            sp, ep = curva.StartPoint, curva.EndPoint
            x0, y0 = float(sp.X), float(sp.Y)
            x1, y1 = float(ep.X), float(ep.Y)
            if abs(x1 - x0) < 1e-9 and abs(y1 - y0) < 1e-9:
                continue
            segs.append((x0, y0, x1, y1))
        except Exception:
            continue
    return segs


def _ys_en_x(
    segs: list[tuple[float, float, float, float]], x: float
) -> list[float]:
    """Y de intersección de la vertical X=x con cada segmento."""
    ys: list[float] = []
    for x0, y0, x1, y1 in segs:
        xa, xb = (x0, x1) if x0 <= x1 else (x1, x0)
        if x < xa - 1e-6 or x > xb + 1e-6:
            continue
        if abs(x1 - x0) < 1e-9:
            # Vertical: aporta ambos extremos
            ys.append(y0)
            ys.append(y1)
            continue
        t = (x - x0) / (x1 - x0)
        if -1e-6 <= t <= 1.0 + 1e-6:
            ys.append(y0 + t * (y1 - y0))
    return ys


def _anchos_por_banda(
    vista, n_bandas: int = 16
) -> list[tuple[float, float, float, float]]:
    """
    ``[(x_mid, y_span, x0, x1), ...]`` muestreando el contorno en bandas X.

    y_span = maxY-minY por intersección de la vertical central con aristas
    (no solo vértices: así el cuello de botella no se “come” el pad).
    """
    segs = _segmentos_contorno(vista)
    if len(segs) < 3:
        return []
    xs: list[float] = []
    for x0, _y0, x1, _y1 in segs:
        xs.extend((x0, x1))
    minx, maxx = min(xs), max(xs)
    span_x = maxx - minx
    if span_x < _MIN_BANDA_CM:
        return []
    out: list[tuple[float, float, float, float]] = []
    for i in range(n_bandas):
        a = minx + span_x * (i / n_bandas)
        b = minx + span_x * ((i + 1) / n_bandas)
        x_mid = 0.5 * (a + b)
        ys = _ys_en_x(segs, x_mid)
        if len(ys) < 2:
            continue
        y_span = max(ys) - min(ys)
        if y_span < _MIN_BANDA_CM:
            continue
        out.append((float(x_mid), float(y_span), float(a), float(b)))
    return out


def _agrupar_tramos_ancho(
    bandas: list[tuple[float, float, float, float]],
) -> list[dict[str, float]]:
    """Fusiona bandas contiguas de ancho similar → tramos reales."""
    if not bandas:
        return []
    grupos: list[list[tuple[float, float, float, float]]] = []
    for b in bandas:
        if not grupos:
            grupos.append([b])
            continue
        ref = grupos[-1][-1]
        rspan = max(float(ref[1]), 1e-9)
        if abs(float(b[1]) - float(ref[1])) / rspan <= max(_RATIO_ESCALON, 0.12):
            grupos[-1].append(b)
        else:
            grupos.append([b])

    tramos: list[dict[str, float]] = []
    for g in grupos:
        spans = sorted(float(x[1]) for x in g)
        y_span = spans[len(spans) // 2]
        near = [
            x
            for x in g
            if abs(float(x[1]) - y_span) / max(y_span, 1e-9) <= 0.08
        ]
        if not near:
            near = list(g)
        x0 = min(float(x[2]) for x in near)
        x1 = max(float(x[3]) for x in near)
        tramos.append(
            {
                "x_mid": 0.5 * (x0 + x1),
                "y_span": float(y_span),
                "x0": float(x0),
                "x1": float(x1),
            }
        )
    return tramos


def detectar_escalon_vista(vista) -> dict[str, Any] | None:
    """
    True geométrico si el flat tiene ≥2 anchos distintos (zapato/botella).

    Devuelve dict con ``width_left``, ``width_right``, ``ratio``, ``tramos``.
    """
    bandas = _anchos_por_banda(vista, n_bandas=16)
    if len(bandas) < 3:
        return None
    tramos = _agrupar_tramos_ancho(bandas)
    if len(tramos) < 2:
        left = bandas[:2]
        right = bandas[-2:]
        w_l = sum(b[1] for b in left) / len(left)
        w_r = sum(b[1] for b in right) / len(right)
        mayor = max(w_l, w_r, 1e-9)
        menor = min(w_l, w_r)
        ratio = (mayor - menor) / mayor
        if ratio < _RATIO_ESCALON:
            return None
        return {
            "width_left": w_l,
            "width_right": w_r,
            "ratio": ratio,
            "bandas": bandas,
            "tramos": [],
        }

    spans = [float(t["y_span"]) for t in tramos]
    w_min, w_max = min(spans), max(spans)
    ratio = (w_max - w_min) / max(w_max, 1e-9)
    if ratio < _RATIO_ESCALON:
        return None
    return {
        "width_left": float(tramos[0]["y_span"]),
        "width_right": float(tramos[-1]["y_span"]),
        "ratio": ratio,
        "bandas": bandas,
        "tramos": tramos,
    }


def es_cobre_irregular_vista(vista, nombre_pieza: str = "") -> bool:
    if nombre_pieza and not es_candidato_nombre_irregular(nombre_pieza):
        return False
    return detectar_escalon_vista(vista) is not None


def _origen_il_cobre(vista):
    """
    Esquina inferior-izquierda del extremo IZQUIERDO (heel/pad).

    No usa la banda del Y global mínimo: en zapatos con cuello offset el
    punto más bajo está a la DERECHA y eso rompe “acotar de izq → der”.
    Aquí: vértices del borde izquierdo → el más bajo de ese borde.
    """
    try:
        from diametro import _vertices_contorno_placa, _origen_flota_en_vacio
    except Exception:
        return None
    pts = _vertices_contorno_placa(vista)
    if len(pts) < 3:
        return None
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    minx, maxx = min(xs), max(xs)
    span_x = max(maxx - minx, 1e-6)
    # Banda izquierda (~12 % del largo o 3 mm hoja)
    tol = max(0.03, min(0.12 * span_x, 0.25))
    izq = [p for p in pts if p[0] <= minx + tol]
    if len(izq) < 2:
        izq = [p for p in pts if p[0] <= minx + 0.20 * span_x] or pts
    ox, oy = min(izq, key=lambda p: (p[1], p[0]))  # más bajo; desempate más izq
    if _origen_flota_en_vacio(vista, ox, oy):
        # Fallback: esquina IL clásica
        try:
            from diametro import _origen_il_pieza

            return _origen_il_pieza(vista)
        except Exception:
            return float(ox), float(oy)
    return float(ox), float(oy)


def _escuadra_cartesiana(vista) -> dict[str, Any]:
    """
    Como BCK-104: pieza entera en +X/+Y desde esquina IL real.
    Nada cuelga debajo ni a la izquierda del origen (sin “sobresalir”).
    """
    try:
        from diametro import _vertices_contorno_placa
    except Exception:
        return {"ok": False, "motivo": "sin vertices"}
    pts = _vertices_contorno_placa(vista)
    if len(pts) < 4:
        return {"ok": False, "motivo": "pocos pts"}
    sil = _silueta(vista)
    if not sil:
        return {"ok": False, "motivo": "sin silueta"}
    minx, maxx, miny, maxy, _ = sil
    span_x = float(maxx) - float(minx)
    span_y = float(maxy) - float(miny)
    length_h = span_x >= span_y * 0.98
    origen = _origen_il_cobre(vista)
    if origen is None:
        return {"ok": False, "motivo": "sin origen", "length_h": length_h}
    ox, oy = float(origen[0]), float(origen[1])
    tol = max(0.04, 0.01 * max(span_x, span_y))

    # 1) Origen en el borde izquierdo
    if abs(ox - float(minx)) > max(0.15, 0.08 * span_x):
        return {
            "ok": False,
            "motivo": "origen_no_izq",
            "length_h": length_h,
            "ox": ox,
            "oy": oy,
        }

    # 2) Nada sobresale a la izq. ni abajo del origen
    n_izq = sum(1 for x, _y in pts if x < ox - tol)
    n_abajo = sum(1 for _x, y in pts if y < oy - tol)
    if n_izq > 0 or n_abajo > 0:
        return {
            "ok": False,
            "motivo": f"sobresale izq={n_izq} abajo={n_abajo}",
            "length_h": length_h,
            "ox": ox,
            "oy": oy,
        }

    # 3) El suelo lo marca el pad izquierdo (1/3 izq): el cuello no cuelga
    corte = float(minx) + 0.33 * span_x
    ys_izq = [y for x, y in pts if x <= corte]
    if ys_izq:
        miny_izq = min(ys_izq)
        if float(miny) < miny_izq - tol:
            return {
                "ok": False,
                "motivo": "cuello_cuelga_bajo_pad",
                "length_h": length_h,
                "ox": ox,
                "oy": oy,
            }

    # 4) LENGTH horizontal
    if not length_h:
        return {
            "ok": False,
            "motivo": "length_no_horizontal",
            "length_h": False,
            "ox": ox,
            "oy": oy,
        }

    return {
        "ok": True,
        "motivo": "ok",
        "length_h": True,
        "ox": ox,
        "oy": oy,
    }


def _score_orientacion(vista, tg) -> float:
    """
    Mayor = mejor. Reglas forzosas:
      1) LENGTH = eje largo horizontal
      2) Origen = esquina inferior del borde IZQUIERDO (pad/heel)
      3) Material en +X/+Y desde ese origen
      4) Pad ancho a la IZQUIERDA; cuello a la derecha
      5) Cotas de izquierda → derecha
    """
    origen = _origen_il_cobre(vista)
    if origen is None:
        return -1e6
    try:
        from diametro import _origen_flota_en_vacio

        if _origen_flota_en_vacio(vista, origen[0], origen[1]):
            return -1e6
    except Exception:
        pass
    ox, oy = float(origen[0]), float(origen[1])
    sil = _silueta(vista)
    if not sil:
        return -1e5
    minx, maxx, miny, maxy, _span = sil
    span_x = float(maxx) - float(minx)
    span_y = float(maxy) - float(miny)
    dx = maxx - ox
    dy = maxy - oy
    if dx < _MIN_BANDA_CM or dy < _MIN_BANDA_CM:
        return -1e4

    score = 0.0

    # 1) Eje largo horizontal
    if span_x >= span_y * 0.98:
        score += 300.0
    else:
        score -= 800.0

    # 2) Origen en el borde izquierdo (no en el cuello derecho)
    if abs(ox - float(minx)) <= 0.12 * max(span_x, 1e-6) or abs(ox - float(minx)) <= 0.15:
        score += 300.0
    else:
        score -= 700.0  # origen a la derecha = MAL

    # Preferir origen cerca del fondo del tramo izquierdo (no del cuello colgante)
    score -= abs(ox - float(minx)) * 8.0

    # Material a la derecha y arriba del origen
    score += min(dx, 50.0) * 2.0
    score += min(max(dy, 0.0), 50.0) * 2.0
    if dy < -0.05:
        score -= 200.0  # material debajo del origen

    # Barrenos en +X/+Y respecto al origen IL izquierdo
    try:
        from barrenos_xy_despliegue import _centros_barrenos

        barrenos = _centros_barrenos(vista, tg) or []
    except Exception:
        barrenos = []
    n_pos = 0
    n_izq = 0
    for b in barrenos:
        try:
            cx_b = float(b["cx"])
            cy_b = float(b["cy"])
        except Exception:
            continue
        if cx_b >= ox - 0.02 and cy_b >= oy - 0.05:
            n_pos += 1
        if cx_b < ox - 0.05:
            n_izq += 1
    score += n_pos * 2.0
    score -= n_izq * 8.0

    # 4) Pad ancho a la IZQUIERDA
    if span_x >= span_y * 0.98:
        info = detectar_escalon_vista(vista)
        if info:
            w_l = float(info["width_left"])
            w_r = float(info["width_right"])
            if w_l >= w_r * 1.05:
                score += 400.0
            elif w_r >= w_l * 1.05:
                score -= 700.0
            else:
                score -= 50.0

    # 5) Escuadra cartesiana (como BCK-104): nada cuelga del origen IL
    esc = _escuadra_cartesiana(vista)
    if esc.get("ok"):
        score += 500.0
    else:
        score -= 600.0
        motivo = str(esc.get("motivo") or "")
        if "cuelga" in motivo or "sobresale" in motivo:
            score -= 400.0

    return score


def orientar_vista_cobre(vista, tg, log=print) -> dict[str, Any]:
    """
    Plano cartesiano cobre:
      LENGTH horizontal, origen IL, pad a la izquierda, cotas L→R.
    """
    base = _leer_rotacion_rad(vista)

    best_deg = 0.0
    best_score = -1e99
    for deg in _ROTACIONES_DEG:
        try:
            _aplicar_rotacion_rad(vista, base + math.radians(deg))
        except Exception:
            continue
        try:
            _ = int(vista.DrawingCurves.Count)
        except Exception:
            pass
        sc = _score_orientacion(vista, tg)
        if sc > best_score:
            best_score = sc
            best_deg = deg

    try:
        _aplicar_rotacion_rad(vista, base + math.radians(best_deg))
    except Exception as exc:
        log(f"  AVISO Rotate cobre: {exc}")
        return {"deg": 0.0, "score": best_score, "irregular": False}

    # Si sigue de pie → forzar horizontal
    sil = _silueta(vista)
    if sil:
        minx, maxx, miny, maxy, _ = sil
        if (float(maxy) - float(miny)) > (float(maxx) - float(minx)) * 1.02:
            cur = _leer_rotacion_rad(vista)
            best_extra, sc_best = 0.0, best_score
            for extra in (90.0, 270.0):
                try:
                    _aplicar_rotacion_rad(vista, cur + math.radians(extra))
                except Exception:
                    continue
                sc = _score_orientacion(vista, tg)
                if sc > sc_best:
                    sc_best, best_extra = sc, extra
            try:
                _aplicar_rotacion_rad(vista, cur + math.radians(best_extra))
                best_deg = (best_deg + best_extra) % 360.0
                best_score = sc_best
            except Exception:
                pass

    # Si el cuello quedó a la izquierda → +180 (pad debe ser heel izquierdo)
    info = detectar_escalon_vista(vista)
    if info and float(info["width_right"]) > float(info["width_left"]) * 1.05:
        cur = _leer_rotacion_rad(vista)
        try:
            _aplicar_rotacion_rad(vista, cur + math.radians(180.0))
            best_deg = (best_deg + 180.0) % 360.0
            best_score = _score_orientacion(vista, tg)
        except Exception:
            pass

    sil = _silueta(vista)
    length_h = False
    pad_izq = None
    origen_il_ok = False
    if sil:
        minx, maxx, miny, maxy, _ = sil
        length_h = (float(maxx) - float(minx)) >= (float(maxy) - float(miny)) * 0.98
        ori = _origen_il_cobre(vista)
        if ori is not None:
            origen_il_ok = abs(float(ori[0]) - float(minx)) <= max(
                0.15, 0.12 * (float(maxx) - float(minx))
            )
    info2 = detectar_escalon_vista(vista)
    if info2:
        pad_izq = float(info2["width_left"]) >= float(info2["width_right"]) * 0.95

    esc = _escuadra_cartesiana(vista)
    irregular = info2 is not None
    log(
        f"  cobre cartesiano: Rotation +{best_deg:.0f}° "
        f"(score={best_score:.1f}, escalon={irregular}, "
        f"LENGTH_h={'SI' if length_h else 'NO'}, "
        f"pad_izq={pad_izq}, origen_IL={'SI' if origen_il_ok else 'NO'}, "
        f"escuadra={'SI' if esc.get('ok') else 'NO:' + str(esc.get('motivo'))})"
    )
    return {
        "deg": best_deg,
        "score": best_score,
        "irregular": irregular,
        "length_horizontal": length_h,
        "pad_izquierda": pad_izq,
        "origen_il": origen_il_ok,
        "escuadra": bool(esc.get("ok")),
        "escuadra_motivo": esc.get("motivo"),
    }


# Compat: nombre histórico (solo irregular). Misma implementación cartesiana.
orientar_vista_irregular = orientar_vista_cobre


def tramos_width_irregular(vista) -> list[dict[str, float]]:
    """
    Tramos WIDTH tras orientación pad-izquierda:
      WIDTH1 = cuello (tramo estrecho a la DERECHA)
      WIDTH2 = pad solo si aporta y ≠ TOTAL (heel izquierdo)

    Ignora bandas de transición (chaflán) donde el y_span no es estable.
    """
    info = detectar_escalon_vista(vista)
    if not info:
        return []
    bandas = list(info.get("bandas") or _anchos_por_banda(vista, n_bandas=16))
    if len(bandas) < 4:
        return []

    # Descartar bandas de transición: vecinos con salto brusco a ambos lados
    estables: list[tuple[float, float, float, float]] = []
    for i, b in enumerate(bandas):
        if i == 0 or i == len(bandas) - 1:
            estables.append(b)
            continue
        prev_s, next_s = float(bandas[i - 1][1]), float(bandas[i + 1][1])
        cur_s = float(b[1])
        # Transición si difiere mucho de ambos vecinos en direcciones opuestas
        if (
            abs(cur_s - prev_s) / max(prev_s, 1e-9) > 0.12
            and abs(cur_s - next_s) / max(next_s, 1e-9) > 0.12
        ):
            continue
        estables.append(b)
    if len(estables) < 4:
        estables = bandas

    mid = len(estables) // 2
    izq, der = estables[: max(2, mid)], estables[min(len(estables) - 2, mid) :]
    if not izq or not der:
        return []

    def _tramo_de(grupo, pick: str) -> dict[str, float]:
        if pick == "min":
            b = min(grupo, key=lambda x: float(x[1]))
        else:
            b = max(grupo, key=lambda x: float(x[1]))
        # Agrupar vecinos de ancho similar alrededor de b
        near = [
            x
            for x in grupo
            if abs(float(x[1]) - float(b[1])) / max(float(b[1]), 1e-9) <= 0.10
        ]
        if not near:
            near = [b]
        x0 = min(float(x[2]) for x in near)
        x1 = max(float(x[3]) for x in near)
        spans = sorted(float(x[1]) for x in near)
        y_span = spans[len(spans) // 2]
        return {
            "x_mid": 0.5 * (x0 + x1),
            "y_span": float(y_span),
            "x0": float(x0),
            "x1": float(x1),
        }

    # Con pad a la izquierda: izq=pad (max), der=cuello (min)
    pad = _tramo_de(izq, "max")
    cuello = _tramo_de(der, "min")
    if float(cuello["y_span"]) >= float(pad["y_span"]) * 0.92:
        # Orientación aún mal o sin escalón real
        cuello = _tramo_de(estables, "min")
        pad = _tramo_de(estables, "max")

    if abs(float(pad["y_span"]) - float(cuello["y_span"])) / max(
        float(pad["y_span"]), 1e-9
    ) < _RATIO_ESCALON:
        return []

    out = [cuello]
    # Solo cuello + WIDTH_TOTAL (FRENTE_2). El pad izquierdo ≈ envelope
    # en zapatos offset; WIDTH2 generaba cotas basura (106.20 en chaflán).
    return out


def aplicar_orientacion_cobre(
    vista, part_name: str, tg, log=print
) -> bool:
    """
    Hook post-creación de DESPLIEGUE_FRENTE_* (todas las cobre).

    LENGTH horizontal, origen IL izquierdo, pad izq., escuadra +X/+Y
    (referencia: BCK-104 — nada cuelga del origen).
    """
    if not es_candidato_nombre_irregular(part_name):
        return False
    info = orientar_vista_cobre(vista, tg, log=log)
    return bool(info.get("escuadra") or info.get("length_horizontal"))


def necesita_flip_camara_cobre(vista) -> bool:
    """True si tras rotar sigue sin escuadra (hay que recrear la vista volteada)."""
    return not bool(_escuadra_cartesiana(vista).get("ok"))


def aplicar_orientacion_si_irregular(
    vista, part_name: str, tg, log=print
) -> bool:
    """Compat: alias de ``aplicar_orientacion_cobre`` (todas las cobre)."""
    return aplicar_orientacion_cobre(vista, part_name, tg, log=log)


def acotar_widths_tramos_irregular(
    nombres_frente_ok: list[str] | None = None,
    log=print,
) -> list[str]:
    """
    Crea hojas WIDTH1..N (tramos desde el origen) y marca FRENTE_2 como WIDTH_TOTAL.

    Parte de DESPLIEGUE_FRENTE_1 orientada; no toca XY/HOLE.
    """
    import win32com.client

    from inventor_com import conectar_inventor
    from barrenos_xy_despliegue import (
        _dibujar_cota_centro_sketch,
        _forzar_vista_lista,
        _limpiar_dims_y_sketches,
        _borrar_hojas_prefijo,
        _pieza_desde_hoja_despliegue,
    )
    from diametro import _origen_il_pieza

    inv = conectar_inventor()
    tg = inv.TransientGeometry
    try:
        plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
    except Exception:
        log("WIDTH irregular: sin DrawingDocument activo")
        return []

    creadas: list[str] = []
    permitidos = None
    if nombres_frente_ok:
        permitidos = {str(x).upper().rsplit(":", 1)[0] for x in nombres_frente_ok}

    piezas_irreg: set[str] = set()

    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            nombre = str(hoja.Name)
        except Exception:
            continue
        nu = nombre.upper()
        base = nombre.rsplit(":", 1)[0]
        if "DESPLIEGUE_FRENTE_1" not in nu:
            continue
        if any(
            t in nu
            for t in ("XMIN", "YMIN", "XMAX", "YMAX", "HOLE", "XCENTRO", "YCENTRO")
        ):
            continue
        if permitidos is not None and base.upper() not in permitidos:
            continue
        pieza = _pieza_desde_hoja_despliegue(nombre)
        if not es_candidato_nombre_irregular(pieza):
            continue
        if hoja.DrawingViews.Count < 1:
            continue
        vista = hoja.DrawingViews.Item(1)
        if not es_cobre_irregular_vista(vista, pieza):
            continue
        piezas_irreg.add(pieza.upper())
        tramos = tramos_width_irregular(vista)
        if not tramos:
            log(f"  {base}: irregular sin tramos WIDTH distintos (solo TOTAL)")
            continue
        # WIDTH1 = cuello (ya ordenado por tramos_width_irregular)
        segs = _segmentos_contorno(vista)
        log(f"  {base}: WIDTH1=cuello..N irregular n={len(tramos)}")
        stem = base
        if "_DESPLIEGUE_FRENTE_1" in base.upper():
            stem = base[: base.upper().rfind("_DESPLIEGUE_FRENTE_1")]
        for idx, tr in enumerate(tramos, start=1):
            x0, x1 = float(tr["x0"]), float(tr["x1"])
            # Medir en el interior del tramo (evitar chaflán)
            pad_x = max(0.02, 0.20 * (x1 - x0))
            xa, xb = x0 + pad_x, x1 - pad_x
            if xb <= xa:
                xa, xb = x0, x1
            x_mid = 0.5 * (xa + xb)
            spans = []
            for t in (0.3, 0.5, 0.7):
                ys_i = _ys_en_x(segs, xa + t * (xb - xa))
                if len(ys_i) >= 2:
                    spans.append((min(ys_i), max(ys_i)))
            if not spans:
                continue
            # Mediana de alturas; descartar si hay mucha variación (transición)
            spans.sort(key=lambda p: p[1] - p[0])
            y_bot, y_top = spans[len(spans) // 2]
            h_vals = [p[1] - p[0] for p in spans]
            if max(h_vals) - min(h_vals) > 0.12 * max(h_vals):
                log(f"  skip WIDTH{idx}: banda inestable (chaflán)")
                continue
            if abs(y_top - y_bot) < _MIN_BANDA_CM:
                continue
            nombre_nueva = f"{stem}_DESPLIEGUE_WIDTH{idx}"
            _borrar_hojas_prefijo(plano, nombre_nueva)
            try:
                nueva = hoja.CopyTo(plano)
            except Exception as exc:
                log(f"  WARN CopyTo WIDTH{idx} {nombre_nueva}: {exc}")
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
            _forzar_vista_lista(inv, nueva, vista_n)
            _limpiar_dims_y_sketches(nueva)
            segs_n = _segmentos_contorno(vista_n)
            ys_n = _ys_en_x(segs_n, x_mid)
            if len(ys_n) >= 2:
                y_bot, y_top = min(ys_n), max(ys_n)
            # Valor en cm de hoja → texto con unidad activa (mm)
            try:
                from cota_estilo import texto_cota_dibujo, asegurar_unidad_cota

                # Distancia en hoja / Scale = cm modelo (API Inventor)
                esc = abs(float(vista_n.Scale)) or 1.0
                dy_modelo_cm = abs(y_top - y_bot) / esc
                texto = texto_cota_dibujo(dy_modelo_cm, nueva) or asegurar_unidad_cota(
                    f"{dy_modelo_cm * 10.0:.2f}"
                )
            except Exception:
                texto = f"{abs(y_top - y_bot):.3f}"
            try:
                _dibujar_cota_centro_sketch(
                    nueva,
                    vista_n,
                    tg,
                    inv,
                    "Y",
                    x_mid,
                    y_bot,
                    x_mid,
                    y_top,
                    texto,
                )
            except Exception as exc:
                log(f"  WARN dim WIDTH{idx} {nombre_nueva}: {exc}")
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            creadas.append(nombre_nueva)
            tag = "cuello" if idx == 1 else "pad"
            log(f"  ✅ {nombre_nueva}: WIDTH{idx}={texto} ({tag})")

    # FRENTE_2 de piezas irregulares → export como WIDTH_TOTAL
    if piezas_irreg:
        os.environ["COTAS_WIDTH_TOTAL"] = "1"
        for i in range(1, int(plano.Sheets.Count) + 1):
            try:
                hoja = plano.Sheets.Item(i)
                nombre = str(hoja.Name)
            except Exception:
                continue
            nu = nombre.upper()
            if "DESPLIEGUE_FRENTE_2" not in nu and "DESPLIEGUE_ANCHO" not in nu:
                continue
            if any(t in nu for t in ("WIDTH1", "WIDTH2", "WIDTH_TOTAL", "XMIN", "YMIN")):
                continue
            pieza = _pieza_desde_hoja_despliegue(nombre)
            if pieza.upper() not in piezas_irreg:
                continue
            stem = nombre.rsplit(":", 1)[0]
            for tok in ("_DESPLIEGUE_FRENTE_2", "_DESPLIEGUE_ANCHO"):
                if tok in stem.upper():
                    # case-preserving replace via rfind
                    up = stem.upper()
                    idx = up.rfind(tok)
                    nuevo = stem[:idx] + "_DESPLIEGUE_WIDTH_TOTAL"
                    break
            else:
                continue
            try:
                hoja.Name = nuevo
                creadas.append(nuevo)
                log(f"  ✅ {nuevo}: WIDTH_TOTAL (span global)")
            except Exception as exc:
                log(f"  WARN rename WIDTH_TOTAL: {exc}")

    return creadas
