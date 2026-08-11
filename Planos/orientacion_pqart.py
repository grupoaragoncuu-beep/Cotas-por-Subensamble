"""
Marco de orientacion igual que Precesador STEP PQart 2.0:

  tapa / altura  -> +Y  (cover)
  cara principal -> +Z  (face)
  lateral        -> +X  (right)

Si el ensamble llego chueco, se leen normales de caras planas reales
(como hace PQart) y se arma ese marco. Las fotos de COTAS_CARAS usan
estas direcciones para quedar derechas.
"""

from __future__ import annotations

import math


K_PLANE_SURFACE = 5890
ANGULO_CLUSTER_DEG = 15.0
AREA_MINIMA_CARA_CM2 = 80.0


def _norm(v):
    length = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    if length < 1e-12:
        return (0.0, 0.0, 0.0)
    return (v[0] / length, v[1] / length, v[2] / length)


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a, b):
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _scale(v, s):
    return (v[0] * s, v[1] * s, v[2] * s)


def _abs_dot(a, b):
    return abs(_dot(a, b))


def _span_en_direccion(corners, direccion):
    vals = [_dot(c, direccion) for c in corners]
    return max(vals) - min(vals)


def _transformar_normal(n_local, matriz):
    try:
        x = (
            matriz.Cell(1, 1) * n_local[0]
            + matriz.Cell(1, 2) * n_local[1]
            + matriz.Cell(1, 3) * n_local[2]
        )
        y = (
            matriz.Cell(2, 1) * n_local[0]
            + matriz.Cell(2, 2) * n_local[1]
            + matriz.Cell(2, 3) * n_local[2]
        )
        z = (
            matriz.Cell(3, 1) * n_local[0]
            + matriz.Cell(3, 2) * n_local[1]
            + matriz.Cell(3, 3) * n_local[2]
        )
        return _norm((x, y, z))
    except Exception:
        return _norm(n_local)


def _recolectar_normales_caras(ensamble, log):
    muestras = []
    ocurrencias = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences

    for i in range(1, ocurrencias.Count + 1):
        try:
            occ = ocurrencias.Item(i)
            if occ.Suppressed:
                continue
            part = occ.Definition.Document
            if part is None:
                continue

            matriz = occ.Transformation
            bodies = part.ComponentDefinition.SurfaceBodies
            for b in range(1, bodies.Count + 1):
                body = bodies.Item(b)
                for f in range(1, body.Faces.Count + 1):
                    face = body.Faces.Item(f)
                    try:
                        if face.SurfaceType != K_PLANE_SURFACE:
                            continue
                        area = float(face.Evaluator.Area)
                        if area < AREA_MINIMA_CARA_CM2:
                            continue

                        n_asm = None
                        try:
                            proxy = occ.CreateGeometryProxy(face)
                            n = proxy.Geometry.Normal
                            n_asm = _norm((n.X, n.Y, n.Z))
                        except Exception:
                            n = face.Geometry.Normal
                            n_asm = _transformar_normal((n.X, n.Y, n.Z), matriz)

                        if _dot(n_asm, n_asm) < 0.5:
                            continue
                        muestras.append((n_asm, area))
                    except Exception:
                        continue
        except Exception:
            continue

    log(f"  Caras planas (estilo PQart): {len(muestras)}")
    return muestras


def _clusterizar_normales(muestras):
    cos_lim = math.cos(math.radians(ANGULO_CLUSTER_DEG))
    clusters = []

    for normal, area in muestras:
        n = _norm(normal)
        if _dot(n, n) < 0.5:
            continue

        colocado = False
        for cluster in clusters:
            ad = _dot(n, cluster["dir"])
            if abs(ad) >= cos_lim:
                if ad < 0:
                    n = _scale(n, -1.0)
                peso = cluster["area"]
                cluster["dir"] = _norm(
                    (
                        cluster["dir"][0] * peso + n[0] * area,
                        cluster["dir"][1] * peso + n[1] * area,
                        cluster["dir"][2] * peso + n[2] * area,
                    )
                )
                cluster["area"] += area
                colocado = True
                break

        if not colocado:
            clusters.append({"dir": n, "area": area})

    clusters.sort(key=lambda c: c["area"], reverse=True)
    return clusters


def marco_como_pqart(ensamble, bbox, log):
    """
    Devuelve (cover_+Y, face_+Z, right_+X) construidos como en PQart 2.0.
    """
    muestras = _recolectar_normales_caras(ensamble, log)
    clusters = _clusterizar_normales(muestras)

    if len(clusters) < 2:
        log("  AVISO: pocas caras; se usa marco mundial PQart (+Y/+Z/+X).")
        return (0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (1.0, 0.0, 0.0)

    for idx, c in enumerate(clusters[:6]):
        d = c["dir"]
        span = _span_en_direccion(bbox["corners"], d)
        log(
            f"  Normal #{idx + 1}: ({d[0]:.3f},{d[1]:.3f},{d[2]:.3f}) "
            f"area={c['area']:.0f} span={span:.1f}"
        )

    # COVER (+Y): siempre la tapa / altura, NO una pared.
    # Bug previo: premiar area hacia que una pared ~X ganara a la tapa ~Y
    # y el tanque salia acostado. Aqui se fuerza preferencia a mundo +Y.
    mundo_y = (0.0, 1.0, 0.0)
    cover = None
    mejor = -1.0
    for c in clusters[:12]:
        aline_y = _abs_dot(c["dir"], mundo_y)
        # Solo candidatos claramente verticales (tapa / piso).
        if aline_y < 0.55:
            continue
        score = aline_y * 20000.0 + math.sqrt(max(c["area"], 1.0))
        if score > mejor:
            mejor = score
            cover = c["dir"]

    if cover is None:
        # Fallback: el cluster mas alineado a +Y, sin mirar X/Z.
        for c in clusters[:12]:
            score = _abs_dot(c["dir"], mundo_y)
            if score > mejor:
                mejor = score
                cover = c["dir"]

    if cover is None:
        cover = mundo_y

    if _dot(cover, mundo_y) < 0:
        cover = _scale(cover, -1.0)
    cover = _norm(cover)

    # FACE (+Z): mayor area casi perpendicular a la tapa.
    face = None
    mejor_face = -1.0
    for c in clusters[:10]:
        if _abs_dot(c["dir"], cover) > 0.35:
            continue
        v = (
            c["dir"][0] - cover[0] * _dot(c["dir"], cover),
            c["dir"][1] - cover[1] * _dot(c["dir"], cover),
            c["dir"][2] - cover[2] * _dot(c["dir"], cover),
        )
        v = _norm(v)
        if _dot(v, v) < 0.5:
            continue
        if c["area"] > mejor_face:
            mejor_face = c["area"]
            face = v

    if face is None:
        tmp = (0.0, 0.0, 1.0)
        if _abs_dot(tmp, cover) > 0.9:
            tmp = (1.0, 0.0, 0.0)
        face = _norm(
            (
                tmp[0] - cover[0] * _dot(tmp, cover),
                tmp[1] - cover[1] * _dot(tmp, cover),
                tmp[2] - cover[2] * _dot(tmp, cover),
            )
        )

    # Ortogonalizar EXACTO: tapa / cara / lateral sin tilt residual.
    face = _norm(
        (
            face[0] - cover[0] * _dot(face, cover),
            face[1] - cover[1] * _dot(face, cover),
            face[2] - cover[2] * _dot(face, cover),
        )
    )
    right = _norm(_cross(cover, face))
    if _dot(right, right) < 0.5:
        right = _norm(_cross(cover, (1.0, 0.0, 0.0)))
    face = _norm(_cross(right, cover))
    cover = _norm(_cross(face, right))

    log(
        f"  Marco PQart: cover(+Y)=({cover[0]:.3f},{cover[1]:.3f},{cover[2]:.3f}) "
        f"face(+Z)=({face[0]:.3f},{face[1]:.3f},{face[2]:.3f}) "
        f"right(+X)=({right[0]:.3f},{right[1]:.3f},{right[2]:.3f})"
    )
    return cover, face, right
