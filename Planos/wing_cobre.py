# -*- coding: utf-8 -*-
"""
Cota WING en perfil Z (dos alas), como ABB-42-BCK-723.

El modelo GIGA está en pulgadas. La cota publicada es milímetros:
punta del ala hasta la cara de afuera del alma.
Una JPG por ala: ``WING01_<mm>mm`` y ``WING02_<mm>mm``.
En metal solo entran las dos alas largas; un labio del orden del espesor no.
"""
from __future__ import annotations

import os
import time

import win32com.client

import creador_vistas as cv
from cota_estilo import aplicar_estilo_cota, set_unidad_cota
from generador_caras_tanque import _encontrar_hoja_machote, _obtener_plano_activo
from generador_vistas import (
    ANCHO_EXPORTACION,
    ALTO_EXPORTACION,
    _recortar_exportacion_jpg,
    borrar_hojas_por_nombres,
    _nombre_hoja_machote,
)
from piezas_cobre import es_pieza_cobre

CM_TO_MM = 10.0
K_VERTICAL = 60163
K_HIDDEN = 32257


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _scale(a, s):
    return (a[0] * s, a[1] * s, a[2] * s)


def _cross(a, b):
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _norm(a):
    L = _dot(a, a) ** 0.5
    if L < 1e-12:
        return None
    return _scale(a, 1.0 / L)


def _xyz(p):
    return (float(p.X), float(p.Y), float(p.Z))


def _stem(doc) -> str:
    try:
        ruta = str(doc.FullFileName or "")
    except Exception:
        ruta = ""
    base = os.path.splitext(os.path.basename(ruta))[0] if ruta else ""
    if ":" in base:
        base = base.split(":", 1)[0]
    return base.strip()


def _sm(part):
    try:
        return win32com.client.CastTo(
            part.ComponentDefinition, "SheetMetalComponentDefinition"
        )
    except Exception:
        return None


def _bends(sm):
    out = []
    try:
        n = int(sm.Bends.Count)
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            b = sm.Bends.Item(i)
            ff = b.FrontFaces.Item(1).Geometry
            bf = b.BackFaces.Item(1).Geometry
            rf, rb = float(ff.Radius), float(bf.Radius)
            inner = ff if rf <= rb else bf
            ax = _norm(_xyz(inner.AxisVector))
            if ax is None:
                continue
            out.append(
                {
                    "ax": ax,
                    "bp": _xyz(inner.BasePoint),
                    "r_in": min(rf, rb),
                    "r_out": max(rf, rb),
                }
            )
        except Exception:
            continue
    return out


def _estaciones(bends):
    grupos = []
    for b in bends:
        puesto = False
        for g in grupos:
            ref = g[0]
            if abs(_dot(b["ax"], ref["ax"])) <= 0.97:
                continue
            along = _dot(b["bp"], ref["ax"])
            perp = _sub(b["bp"], _scale(ref["ax"], along))
            along0 = _dot(ref["bp"], ref["ax"])
            perp0 = _sub(ref["bp"], _scale(ref["ax"], along0))
            if _dot(_sub(perp, perp0), _sub(perp, perp0)) ** 0.5 < 0.8:
                g.append(b)
                puesto = True
                break
        if not puesto:
            grupos.append([b])
    return grupos


def medir_alas_mm(part) -> list[float] | None:
    """Dos alas en mm, de menor a mayor. None si no es perfil Z."""
    sm = _sm(part)
    if sm is None:
        return None
    try:
        thk = float(sm.Thickness.Value)
    except Exception:
        return None
    estaciones = _estaciones(_bends(sm))
    if len(estaciones) != 2:
        return None
    if abs(_dot(estaciones[0][0]["ax"], estaciones[1][0]["ax"])) <= 0.97:
        return None
    ax = estaciones[0][0]["ax"]
    r_in = min(b["r_in"] for g in estaciones for b in g)
    c1 = estaciones[0][0]["bp"]
    c2 = estaciones[1][0]["bp"]
    web = _sub(c2, c1)
    web = _sub(web, _scale(ax, _dot(web, ax)))
    wdir = _norm(web)
    if wdir is None:
        return None
    pdir = _norm(_cross(ax, wdir))
    if pdir is None:
        return None
    sep = _dot(_sub(c2, c1), wdir)
    if sep < 0:
        c1, c2 = c2, c1
        wdir = _scale(wdir, -1.0)
        pdir = _scale(pdir, -1.0)
        sep = -sep
    try:
        body = part.ComponentDefinition.SurfaceBodies.Item(1)
        n = int(body.Vertices.Count)
    except Exception:
        return None
    coords = []
    for i in range(1, n + 1):
        try:
            d = _sub(_xyz(body.Vertices.Item(i).Point), c1)
        except Exception:
            continue
        coords.append((_dot(d, wdir), _dot(d, pdir), _xyz(body.Vertices.Item(i).Point)))
    # Solo los cantos del alma, no las puntas de las alas.
    izq = [c for c in coords if -0.2 * sep < c[0] < 0.30 * sep]
    der = [c for c in coords if 0.70 * sep < c[0] < 1.20 * sep]
    candidatas = []
    for a in izq:
        for b in der:
            dw, dp = b[0] - a[0], b[1] - a[1]
            largo = (dw * dw + dp * dp) ** 0.5
            if largo < sep * 0.7 or abs(largo - sep) / sep > 0.08:
                continue
            if abs(dw) / largo < 0.95:
                continue
            ux, uy = dw / largo, dp / largo
            nx, ny = -uy, ux
            ns = [(c[0] - a[0]) * nx + (c[1] - a[1]) * ny for c in coords]
            hi, lo = max(ns), min(ns)
            web_ns = [
                (c[0] - a[0]) * nx + (c[1] - a[1]) * ny
                for c in coords
                if 0.12 * largo < (c[0] - a[0]) * ux + (c[1] - a[1]) * uy < 0.88 * largo
            ]
            if len(web_ns) < 2:
                web_ns = list(ns)
            web_ns = sorted(web_ns)
            niveles = []
            for v in web_ns:
                if not niveles or abs(v - niveles[-1]) > thk * 0.35:
                    niveles.append(v)
            pares = []
            for i in range(len(niveles)):
                for j in range(i + 1, len(niveles)):
                    if abs(abs(niveles[j] - niveles[i]) - thk) > thk * 0.3:
                        continue
                    cara_lo = min(niveles[i], niveles[j])
                    cara_hi = max(niveles[i], niveles[j])
                    if (hi - cara_hi) * (lo - cara_lo) >= 0:
                        continue
                    # Punta a la cara de afuera del alma, como el (3.500) y (6.824) del modelo.
                    ala_sup = (hi - cara_lo) * CM_TO_MM
                    ala_inf = (cara_hi - lo) * CM_TO_MM
                    pares.append(
                        (round(ala_inf, 2), round(ala_sup, 2), cara_lo, cara_hi)
                    )
            if not pares:
                continue
            # Misma normal: varias caras paralelas a un espesor. La del alma
            # es la del medio, no la que se corre un grosor hacia la punta.
            pares.sort(key=lambda t: max(t[0], t[1]))
            ala_inf, ala_sup, cara_lo, cara_hi = pares[len(pares) // 2]

            def _vert(n_obj, solo_alma, _a=a, _nx=nx, _ny=ny, _ux=ux, _uy=uy, _L=largo):
                mejor_v = None
                for c in coords:
                    nn = (c[0] - _a[0]) * _nx + (c[1] - _a[1]) * _ny
                    ww = (c[0] - _a[0]) * _ux + (c[1] - _a[1]) * _uy
                    if solo_alma and not (0.15 * _L < ww < 0.85 * _L):
                        continue
                    err = abs(nn - n_obj)
                    if mejor_v is None or err < mejor_v[0]:
                        mejor_v = (err, c[2])
                if mejor_v is None and solo_alma:
                    return _vert(n_obj, False)
                return None if mejor_v is None else mejor_v[1]

            p_lo, p_hi = _vert(lo, False), _vert(hi, False)
            f_lo, f_hi = _vert(cara_hi, True), _vert(cara_lo, True)
            if None in (p_lo, p_hi, f_lo, f_hi):
                continue
            if ala_inf <= ala_sup:
                puntos = ((p_lo, f_lo), (p_hi, f_hi))
                ala_a, ala_b = ala_inf, ala_sup
            else:
                puntos = ((p_hi, f_hi), (p_lo, f_lo))
                ala_a, ala_b = ala_sup, ala_inf
            normal = (
                nx * wdir[0] + ny * pdir[0],
                nx * wdir[1] + ny * pdir[1],
                nx * wdir[2] + ny * pdir[2],
            )
            candidatas.append((largo, ala_a, ala_b, puntos, normal))
    if not candidatas:
        return None
    corto = min(c[0] for c in candidatas)
    grupo = [c for c in candidatas if abs(c[0] - corto) < 0.02]
    grupo.sort(key=lambda t: max(t[1], t[2]))
    _largo, ala_a, ala_b, puntos, normal = grupo[len(grupo) // 2]
    return {
        "alas": [round(ala_a, 2), round(ala_b, 2)],
        "puntos": puntos,
        "ax": ax,
        "up": normal,
        "pdir": pdir,
        "c": c1,
    }


def _alas_largas_metal(part, alas) -> bool:
    """En metal, las dos alas tienen que ser largas. Un labio del espesor no."""
    if not alas or min(alas) < 15.0:
        return False
    sm = _sm(part)
    if sm is None:
        return False
    try:
        thk_mm = float(sm.Thickness.Value) * CM_TO_MM
    except Exception:
        return False
    return min(alas) >= 4.0 * thk_mm


def _piezas(asm):
    hojas = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
    vistos = set()
    out = []
    for i in range(1, int(hojas.Count) + 1):
        try:
            doc = hojas.Item(i).Definition.Document
            if int(doc.DocumentType) != 12290:
                continue
            ruta = str(doc.FullFileName or "")
        except Exception:
            continue
        if not ruta or ruta in vistos:
            continue
        vistos.add(ruta)
        stem = _stem(doc)
        out.append((stem, doc))
    out.sort(key=lambda t: t[0].upper())
    return out


def _asm(inv, ensamble=None):
    if ensamble is not None:
        try:
            if int(ensamble.DocumentType) == 12291:
                return win32com.client.CastTo(ensamble, "AssemblyDocument")
        except Exception:
            pass
    mejor = None
    for i in range(1, int(inv.Documents.Count) + 1):
        d = inv.Documents.Item(i)
        try:
            if int(d.DocumentType) != 12291:
                continue
            nombre = str(d.DisplayName or "").upper().replace(" ", "")
        except Exception:
            continue
        if "9919" in nombre and "BOARD" in nombre:
            return win32com.client.CastTo(d, "AssemblyDocument")
        mejor = d
    if mejor is None:
        return None
    return win32com.client.CastTo(mejor, "AssemblyDocument")


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
            out.append((float(s.X), float(s.Y), float(e.X), float(e.Y), c))
        except Exception:
            continue
    return out


def _curva_cerca(curvas, sx, sy):
    mejor = None
    for x1, y1, x2, y2, c in curvas:
        dx, dy = x2 - x1, y2 - y1
        largo2 = dx * dx + dy * dy
        if largo2 < 1e-12:
            d2 = (sx - x1) ** 2 + (sy - y1) ** 2
        else:
            t = max(0.0, min(1.0, ((sx - x1) * dx + (sy - y1) * dy) / largo2))
            qx, qy = x1 + t * dx, y1 + t * dy
            d2 = (sx - qx) ** 2 + (sy - qy) ** 2
        if mejor is None or d2 < mejor[0]:
            mejor = (d2, c)
    if mejor is None or mejor[0] > 0.2 * 0.2:
        return None
    return mejor[1]


def _acotar_ala(hoja, vista, tg, par, target, seguir_ala: bool = False) -> float | None:
    """Cota vertical: punta del ala a la cara de afuera del alma."""
    curvas = _curvas(vista)
    if len(curvas) < 2:
        return None
    try:
        s1 = vista.ModelToSheetSpace(tg.CreatePoint(*par[0]))
        s2 = vista.ModelToSheetSpace(tg.CreatePoint(*par[1]))
    except Exception:
        return None
    sx1, sy1 = float(s1.X), float(s1.Y)
    sx2, sy2 = float(s2.X), float(s2.Y)
    punta = _curva_cerca(curvas, sx1, sy1)
    if punta is None:
        return None
    # Solo cantos rectos del alma. El arco del doblez no vale.
    caras = []
    for x1, y1, x2, y2, c in curvas:
        if abs(y1 - y2) > 0.8:
            continue
        if abs(x1 - x2) < 0.05:
            continue
        caras.append((x1, y1, x2, y2, c))
    if not caras and not seguir_ala:
        return None
    if seguir_ala:
        cara = _curva_cerca(curvas, sx2, sy2)
        caras = [(sx2, sy2, sx2, sy2, cara)] if cara is not None else []
    else:
        # La cara de afuera es el canto largo mas lejos de la punta.
        # Un tramo corto puede medir el numero 3D y quedar fuera del alma.
        largas = [t for t in caras if abs(t[2] - t[0]) >= 1.2]
        pool = largas if largas else caras

        def _lejos(t):
            return abs((t[1] + t[3]) / 2.0 - sy1)

        tope = max(_lejos(t) for t in pool)
        caras = [t for t in pool if _lejos(t) >= tope - 0.08]
    # El texto queda por fuera del ala, como en el modelo.
    hacia_fuera = -1.8 if sx2 >= sx1 else 1.8
    x_txt = sx1 + hacia_fuera
    y_txt = (sy1 + sy2) / 2.0
    mejor = None
    for x1, y1, x2, y2, c in caras:
        try:
            int_a = hoja.CreateGeometryIntent(punta, tg.CreatePoint2d(sx1, sy1))
            px = x1 if abs(x1 - sx1) <= abs(x2 - sx1) else x2
            py = (y1 + y2) / 2.0
            int_b = hoja.CreateGeometryIntent(c, tg.CreatePoint2d(px, py))
            dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
                tg.CreatePoint2d(x_txt, y_txt), int_a, int_b, K_VERTICAL
            )
            mm = abs(float(dim.ModelValue)) * CM_TO_MM
        except Exception:
            continue
        err = abs(mm - target)
        if mejor is None or err < mejor[0]:
            if mejor is not None:
                try:
                    mejor[2].Delete()
                except Exception:
                    pass
            mejor = (err, mm, dim)
        else:
            try:
                dim.Delete()
            except Exception:
                pass
    if mejor is None or mejor[0] > 8.0:
        if mejor is not None:
            print(f"    candidato {mejor[1]:.2f} vs {target:.2f}")
            try:
                mejor[2].Delete()
            except Exception:
                pass
        return None
    try:
        aplicar_estilo_cota(mejor[2], hoja=hoja)
    except Exception:
        pass
    return mejor[1]


def _exportar(inv, hoja, ruta: str) -> bool:
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    hoja.Activate()
    try:
        inv.ActiveView.Update()
    except Exception:
        pass
    time.sleep(0.35)
    white = inv.TransientObjects.CreateColor(255, 255, 255)
    tmp = ruta + ".__tmp.jpg"
    try:
        inv.ActiveView.Camera.SaveAsBitmap(tmp, ANCHO_EXPORTACION, ALTO_EXPORTACION, white)
        time.sleep(0.1)
        _recortar_exportacion_jpg(hoja, tmp, ruta)
        return os.path.isfile(ruta) and os.path.getsize(ruta) > 1500
    finally:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except OSError:
            pass


def _limpiar(plano, hoja):
    try:
        mach = _nombre_hoja_machote(plano)
        borrar_hojas_por_nombres(
            plano,
            {str(hoja.Name).split(":")[0]},
            nombre_machote_protegido=mach,
        )
    except Exception:
        try:
            hoja.Delete()
        except Exception:
            pass


def _publicar(stem: str, tag: str, mm: float, src: str, job: str) -> str:
    fn = f"{job}__{stem}__{tag}_{mm:.2f}mm.jpg"
    rama = "Busbar" if es_pieza_cobre(stem) else "Metal"
    local = os.path.join(
        r"C:\Proyectos\COTAS\Planos\JPG",
        job,
        "PIEZAS_ACOTADAS",
        "Doblado",
        rama,
        stem,
        fn,
    )
    os.makedirs(os.path.dirname(local), exist_ok=True)
    if os.path.abspath(src) != os.path.abspath(local):
        import shutil

        shutil.copy2(src, local)
    try:
        from cotas_dossier_registro import registrar_jpg

        if not registrar_jpg(local, job=job):
            print(f"  NO DB {fn}")
    except Exception as exc:
        print(f"  AVISO DB {fn}: {exc}")
    print(f"  OK {fn}")
    return local


def acotar_wings_ensamble(inv, ensamble=None, plano=None) -> int:
    """Acota las alas del cobre en Z del ensamble. Devuelve JPG publicados."""
    set_unidad_cota("mm")
    asm = _asm(inv, ensamble)
    if asm is None:
        print("WING: no hay ensamble")
        return 0
    if plano is None:
        plano = _obtener_plano_activo(inv)
    base = _encontrar_hoja_machote(plano) or plano.Sheets.Item(1)
    try:
        from nomenclatura_capturas import nombre_job_desde_ensamble

        job = nombre_job_desde_ensamble(asm)
    except Exception:
        job = _stem(asm) or "JOB"
    try:
        from cotas_dossier_registro import iniciar_sesion_dossier

        iniciar_sesion_dossier(job)
    except Exception as exc:
        print(f"AVISO WING sesion: {exc}")
    tg = inv.TransientGeometry
    to = inv.TransientObjects
    n_ok = 0
    for stem, part in _piezas(asm):
        solo = os.environ.get("COTAS_WING_SOLO", "").strip().upper()
        if solo and not any(tok in stem.upper() for tok in solo.split(",")):
            continue
        info = medir_alas_mm(part)
        if not info:
            continue
        if es_pieza_cobre(stem) and os.environ.get("COTAS_COBRE_ABC", "").strip() == "1":
            continue
        if not es_pieza_cobre(stem) and not _alas_largas_metal(part, info["alas"]):
            continue
        alas = info["alas"]
        print(f"WING {stem}: {alas[0]:.2f} mm y {alas[1]:.2f} mm")
        box = part.ComponentDefinition.RangeBox
        cx = 0.5 * (float(box.MaxPoint.X) + float(box.MinPoint.X))
        cy = 0.5 * (float(box.MaxPoint.Y) + float(box.MinPoint.Y))
        cz = 0.5 * (float(box.MaxPoint.Z) + float(box.MinPoint.Z))
        eye = tg.CreateVector(*info["ax"])
        seguir_ala = os.environ.get("COTAS_WING_UP") == "ala"
        for idx, target in enumerate(alas, start=1):
            tag = f"WING{idx:02d}"
            if seguir_ala:
                p0, p1 = info["puntos"][idx - 1]
                up = tg.CreateVector(p0[0] - p1[0], p0[1] - p1[1], p0[2] - p1[2])
                target = (
                    (p0[0] - p1[0]) ** 2 + (p0[1] - p1[1]) ** 2 + (p0[2] - p1[2]) ** 2
                ) ** 0.5 * CM_TO_MM
            else:
                up_src = info["up"]
                if os.environ.get("COTAS_WING_UP") == "pdir" and info.get("pdir"):
                    up_src = info["pdir"]
                up = tg.CreateVector(*up_src)
            cam = cv.crear_camara(part, tg, to, cx, cy, cz, eye, up)
            nombre = cv.construir_nombre_hoja(plano, stem, tag)
            hoja = cv._crear_hoja_vista(plano, base, nombre)
            cv._limpiar_border_y_titleblock(hoja)
            px, py, aw, ah = cv._area_util_hoja(hoja)
            vista = cv._crear_vista_base(hoja, part, tg, to, px, py, cam, False)
            try:
                cv.escalar_vista(
                    plano, vista, tg, px, py, aw, ah, modo_cobre=es_pieza_cobre(stem)
                )
            except Exception:
                cv.escalar_vista(plano, vista, tg, px, py, aw, ah)
            try:
                vista.ViewStyle = K_HIDDEN
            except Exception:
                pass
            time.sleep(0.25)
            if len(_curvas(vista)) < 4:
                _limpiar(plano, hoja)
                print(f"  FAIL {stem} {tag}: sin curvas")
                continue
            medido = _acotar_ala(
                hoja, vista, tg, info["puntos"][idx - 1], target, seguir_ala=seguir_ala
            )
            if medido is None or abs(medido - target) > 8.0:
                _limpiar(plano, hoja)
                print(f"  FAIL {stem} {tag}: medido={medido} target={target:.2f}")
                continue
            tmp = os.path.join(
                r"C:\Proyectos\COTAS\.runtime", f"_wing_{stem}_{tag}.jpg"
            )
            if not _exportar(inv, hoja, tmp):
                _limpiar(plano, hoja)
                print(f"  FAIL export {stem} {tag}")
                continue
            _publicar(stem, tag, medido, tmp, job)
            try:
                os.remove(tmp)
            except OSError:
                pass
            _limpiar(plano, hoja)
            n_ok += 1
    print(f"WING publicados: {n_ok}")
    return n_ok
