"""
Instructivo de armado para ensambles independientes (kits OTC + Colorimetría).

(0,0) = esquina SUPERIOR-IZQUIERDA escuadrable del ANCLA (pieza de frente /
más cercana a la cámara, con dos rectas). Sin tangentes de arco.

Cada hijo útil (instancia) recibe cota X e Y respecto a ese origen.
HW solo posición. TYP misma familia. 6 vistas ViewCube. ``--limpiar``.

Salida (bajo PIEZAS_ACOTADAS)::

    Corte/Maquinado/Accesorios Sueltos/<kit>/<VISTA>/*.jpg
    Corte/Maquinado/Inspeccion Visual/<kit>/*.jpg
    Corte/Maquinado/Accesorios Sueltos por pieza/<pieza>/*.jpg

Regla iLogic: ``COTAS_ENSAMBLES_INDEPENDIENTES``.
Colorimetría: ``Clasificación = Ensambles Individuales`` en .iam.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import time
import traceback

import pythoncom

from inventor_com import conectar_inventor

# Legacy (migración / --limpiar de corridas viejas).
CARPETA_ENSAMBLES_LEGACY = "ENSAMBLES_INDEPENDIENTES"
CARPETA_PIEZAS = "PIEZAS_ACOTADAS"
# ViewCube completo (6 caras). Ya no se limita a FRONT/TOP/RIGHT.
VISTAS = ("FRONT", "BACK", "TOP", "BOTTOM", "RIGHT", "LEFT")
# Distancia mínima en hoja (cm) para ni siquiera considerar el extremo.
MIN_DIST_ORIGEN_CM = 0.08
# Umbral HW (ignorar espesores). Placas/piezas: más bajo para no perder
# posiciones de armado entre instancias cercanas.
MIN_COTA_INSTRUCTIVO_IN = 1.0
MIN_COTA_PLACA_IN = 0.25
IN_TO_CM = 2.54
TOL_TYP_CM = 0.12
TIPO_DOCUMENTO_PIEZA = 12290
TIPO_DOCUMENTO_ENSAMBLE = 12291

# Excluir del todo (no aportan posición de armado).
_RE_HIJO_EXCLUIR = re.compile(r"(_HOLE)", re.IGNORECASE)
# HW/tornillería: SÍ se acotan por POSICIÓN; nunca por espesor (< umbral).
_RE_ES_HW = re.compile(
    r"(^HW[-_])|(^DIN\d)|(^ISO\d)|"
    r"(WASHER|NUT|BOLT|SCREW|RIVET|FW[-_]|HN[-_]|HH[-_]|LOCK[\s_-]?WASHER)|"
    r"(\bSTACK\b)|(ISOLATOR)",
    re.IGNORECASE,
)
_RE_PLACA_P = re.compile(r"-P\d+", re.IGNORECASE)


def _log(msg):
    print(msg, flush=True)


def _parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Instructivo de ensambles independientes (6 vistas + origen)"
    )
    p.add_argument(
        "--seleccion",
        default="",
        help="JSON de picks TOP+SEGM1..4+BASE (mismo contrato que Subensamble)",
    )
    p.add_argument(
        "--solo",
        default="",
        help="Procesar solo kits cuyo nombre contenga este texto",
    )
    p.add_argument(
        "--max",
        type=int,
        default=0,
        help="Máximo de kits a procesar (0 = todos)",
    )
    p.add_argument(
        "--listar",
        action="store_true",
        help="Solo listar candidatos y salir",
    )
    p.add_argument(
        "--limpiar",
        action="store_true",
        help="Vaciar Accesorios Sueltos / Inspeccion Visual / por pieza antes de exportar",
    )
    return p.parse_args(argv)


def _exclusiones_desde_seleccion(ensamble, seleccion):
    """Nombres/rutas de TOP/SEGM/BASE resueltos por pick — no son kits."""
    from generador_caras_tanque import _segmento_desde_pick

    excl = set()

    def _agregar_seg(seg, etiqueta):
        if not seg:
            return
        antes = len(excl)
        occ = seg.get("occurrence")
        if occ is not None:
            try:
                excl.add(str(occ.Name).split(":")[0].upper())
            except Exception:
                pass
            try:
                doc = occ.Definition.Document
                ff = str(doc.FullFileName or "")
                if ff:
                    excl.add(ff.upper())
                    excl.add(os.path.splitext(os.path.basename(ff))[0].upper())
            except Exception:
                pass
        for key in ("occ_name", "nombre", "name"):
            val = seg.get(key) if isinstance(seg, dict) else None
            if val:
                excl.add(str(val).split(":")[0].upper())
        _log(f"  Cara excluida ({etiqueta}): +{len(excl) - antes} token(s)")

    top = _segmento_desde_pick(ensamble, seleccion["top"], "TOP")
    _agregar_seg(top, "TOP")
    # También el occ_name crudo del JSON por si el contenedor no resolvió.
    try:
        excl.add(str(seleccion["top"].get("occ_name") or "").split(":")[0].upper())
    except Exception:
        pass

    for i, pick in enumerate(seleccion.get("segmentos") or [], start=1):
        etiqueta = str(pick.get("etiqueta") or f"SEGM{i}").upper()
        if not etiqueta.startswith("SEGM"):
            etiqueta = f"SEGM{i}"
        seg = _segmento_desde_pick(ensamble, pick, etiqueta)
        _agregar_seg(seg, etiqueta)
        try:
            excl.add(str(pick.get("occ_name") or "").split(":")[0].upper())
        except Exception:
            pass

    if seleccion.get("base"):
        base = _segmento_desde_pick(ensamble, seleccion["base"], "BASE")
        _agregar_seg(base, "BASE")
        try:
            excl.add(
                str(seleccion["base"].get("occ_name") or "").split(":")[0].upper()
            )
        except Exception:
            pass

    excl.discard("")
    _log(f"Exclusiones por selección: {len(excl)} tokens")
    return excl


def _es_hijo_excluir(nombre):
    """Solo excluye HOLE sueltos (no aportan armado)."""
    return bool(_RE_HIJO_EXCLUIR.search(str(nombre or "")))


def _es_hw(nombre):
    return bool(_RE_ES_HW.search(str(nombre or "")))


def _es_placa_p(nombre):
    return bool(_RE_PLACA_P.search(str(nombre or "")))


def _prefijo_serie(nombre):
    """Primer bloque numérico tipo 62201 / 62134."""
    m = re.match(r"^(\d{4,6})", str(nombre or ""))
    return m.group(1) if m else ""


def _es_kit_instructivo(nombre, n_hijos):
    """
    Solo kits con sentido de armado.

    - Códigos OTC ``*-A##`` con ≥2 hijos útiles.
    - SP / otros: ≥2 hijos (el aislamiento OTC root ya filtró ruido).
    - Nunca placas Pxx sueltas.
    """
    from generador_caras_tanque import _parse_codigo_otc

    nom = str(nombre or "")
    up = nom.upper()
    info = _parse_codigo_otc(nom)
    if info and info["tipo"] == "A":
        return n_hijos >= 2
    if up.startswith("SP-"):
        return n_hijos >= 2
    if re.search(r"-P\d+", up):
        return False
    return n_hijos >= 2


def _nombres_iam_hijos(asm_doc):
    """Nombres base de sub-ensambles directos (para dedupe anidados)."""
    out = set()
    try:
        occs = asm_doc.ComponentDefinition.Occurrences
        n = int(occs.Count)
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            occ = occs.Item(i)
            if int(occ.DefinitionDocumentType) != TIPO_DOCUMENTO_ENSAMBLE:
                continue
            try:
                ff = occ.Definition.Document.FullFileName or ""
                base = os.path.splitext(os.path.basename(ff))[0]
            except Exception:
                base = str(occ.Name).split(":")[0]
            if base:
                out.add(base.upper())
        except Exception:
            continue
    return out


def _filtrar_kits(lista, solo="", max_n=0, serie_job=""):
    out = []
    solo_u = str(solo or "").strip().upper()
    serie = str(serie_job or "").strip()
    for item in lista:
        asm_doc, nombre, qty, hijos = item
        if not _es_kit_instructivo(nombre, hijos):
            _log(f"  Omitido (filtro instructivo): {nombre} hijos={hijos}")
            continue
        if solo_u and solo_u not in str(nombre).upper():
            continue
        # Cross-series: 62134-* dentro de job 62201-* → ruido.
        if serie:
            sk = _prefijo_serie(nombre)
            if sk and sk != serie:
                _log(f"  Omitido (cross-series {sk}!={serie}): {nombre}")
                continue
        out.append(item)

    nombres = {str(t[1]).upper() for t in out}
    anidados = set()
    for asm_doc, nombre, _qty, _h in out:
        for hijo_iam in _nombres_iam_hijos(asm_doc):
            if hijo_iam in nombres and hijo_iam != str(nombre).upper():
                anidados.add(hijo_iam)
    if anidados:
        _log(f"  Kits anidados omitidos como export propio: {sorted(anidados)}")
        out = [t for t in out if str(t[1]).upper() not in anidados]

    out.sort(key=lambda t: str(t[1]).upper())
    if max_n and max_n > 0:
        out = out[:max_n]
    return out


def _hijos_utiles(asm_doc):
    """Ocurrencias 1er nivel acotables. HW permitido (posición); HOLE no."""
    hijos = []
    try:
        occs = asm_doc.ComponentDefinition.Occurrences
        n = int(occs.Count)
    except Exception:
        return hijos
    vistos_base = {}
    for i in range(1, n + 1):
        try:
            occ = occs.Item(i)
            if occ.Suppressed:
                continue
            rb = occ.RangeBox
            occ_id = str(occ.Name).replace(":", "_")
            nombre = occ_id.split("_")[0] if occ_id else f"PIEZA_{i}"
            try:
                doc = occ.Definition.Document
                base = os.path.splitext(os.path.basename(doc.FullFileName or ""))[0]
                if base:
                    nombre = base
            except Exception:
                pass
            if _es_hijo_excluir(nombre) or _es_hijo_excluir(occ_id):
                continue
            pieza_id = occ_id if occ_id else nombre
            if nombre in vistos_base:
                vistos_base[nombre] += 1
                if pieza_id == nombre or pieza_id == nombre.replace(":", "_"):
                    pieza_id = f"{nombre}_{vistos_base[nombre]}"
            else:
                vistos_base[nombre] = 1
            hijos.append(
                {
                    "name": nombre,
                    "pieza_id": pieza_id,
                    "occ": occ,
                    "rb": rb,
                    "es_hw": _es_hw(nombre) or _es_hw(occ_id),
                }
            )
        except Exception:
            continue
    return hijos


def _proyeccion_rb(vista, tg, rb):
    """Proyecta las 8 esquinas del RangeBox 3D a hoja → bbox 2D."""
    try:
        xs = [float(rb.MinPoint.X), float(rb.MaxPoint.X)]
        ys = [float(rb.MinPoint.Y), float(rb.MaxPoint.Y)]
        zs = [float(rb.MinPoint.Z), float(rb.MaxPoint.Z)]
    except Exception:
        return None
    pts = []
    for x in xs:
        for y in ys:
            for z in zs:
                try:
                    p2 = vista.ModelToSheetSpace(tg.CreatePoint(x, y, z))
                    pts.append((float(p2.X), float(p2.Y)))
                except Exception:
                    continue
    if len(pts) < 2:
        return None
    minx = min(p[0] for p in pts)
    maxx = max(p[0] for p in pts)
    miny = min(p[1] for p in pts)
    maxy = max(p[1] for p in pts)
    if max(maxx - minx, maxy - miny) < 1e-4:
        return None
    return {
        "minx": minx,
        "maxx": maxx,
        "miny": miny,
        "maxy": maxy,
        "cx": (minx + maxx) * 0.5,
        "cy": (miny + maxy) * 0.5,
        "area": abs(maxx - minx) * abs(maxy - miny),
        "dx": abs(maxx - minx),
        "dy": abs(maxy - miny),
    }


def _occ_es_descendiente(leaf, raiz):
    actual = leaf
    for _ in range(24):
        if actual is None:
            return False
        try:
            if actual == raiz:
                return True
        except Exception:
            pass
        try:
            actual = actual.ParentOccurrence
        except Exception:
            return False
    return False


def _curvas_hlr_hijo(vista, occ):
    """Curvas HLR de un hijo del kit (occ o sus hojas). Vacío = no visible."""
    from generador_caras_tanque import _info_curva

    datos = []
    try:
        curves = vista.DrawingCurves(occ)
        for i in range(1, int(curves.Count) + 1):
            info = _info_curva(curves.Item(i))
            if info:
                datos.append(info)
    except Exception:
        pass
    if datos:
        return datos
    try:
        todas = vista.DrawingCurves
        total = int(todas.Count)
    except Exception:
        return datos
    for i in range(1, total + 1):
        try:
            curva = todas.Item(i)
            leaf = curva.ModelGeometry.ContainingOccurrence
            if not _occ_es_descendiente(leaf, occ):
                continue
            info = _info_curva(curva)
            if info:
                datos.append(info)
        except Exception:
            continue
    return datos


def _envolvente_hijo_vista(vista, tg, hijo):
    """
    Preferir silueta HLR (cota alineada a lo visible).

    Si el HLR es usable → toque real sin constructiva.
    Si hay curvas pero franja / incompleto → aún se guardan para el toque
    y se marca ``constructiva=True``.
    Si no hay HLR → RangeBox + constructiva.
    """
    from generador_caras_tanque import (
        _envolvente_hlr_de_datos,
        _hlr_sirve_para_inicio,
    )

    datos = _curvas_hlr_hijo(vista, hijo["occ"])
    env_hlr = _envolvente_hlr_de_datos(datos) if datos else None
    if env_hlr is not None and _hlr_sirve_para_inicio(datos):
        dato = dict(env_hlr)
        dato["curvas"] = list(datos)
        dato["nombre"] = hijo["name"]
        dato["dx"] = abs(float(dato["maxx"]) - float(dato["minx"]))
        dato["dy"] = abs(float(dato["maxy"]) - float(dato["miny"]))
        return dato, False
    if env_hlr is not None:
        # Franja / silueta pobre: igual anclar al HLR parcial + constructiva.
        dato = dict(env_hlr)
        dato["curvas"] = list(datos)
        dato["nombre"] = hijo["name"]
        dato["dx"] = abs(float(dato["maxx"]) - float(dato["minx"]))
        dato["dy"] = abs(float(dato["maxy"]) - float(dato["miny"]))
        return dato, True
    env_rb = _proyeccion_rb(vista, tg, hijo["rb"])
    if env_rb is None:
        return None, True
    dato = dict(env_rb)
    dato["curvas"] = []
    dato["nombre"] = hijo["name"]
    return dato, True


def _distancia_pulgadas_hoja(vista, a, b):
    """|a-b| en hoja → pulgadas de modelo (vista.Scale)."""
    try:
        escala = abs(float(vista.Scale))
        if escala <= 1e-12:
            return 0.0
        return abs(float(a) - float(b)) / escala / IN_TO_CM
    except Exception:
        return 0.0


def _origen_desde_ancla(ancla_env, exigir_esquina=True):
    """
    (0,0) = esquina SUPERIOR-IZQUIERDA escuadrable del ANCLA.

    Constante de piso: el operador mide desde arriba-izquierda, no desde
    tangentes de radio ni desde el rincón inferior del AABB.
    """
    if not ancla_env:
        return None, None, None
    from generador_caras_tanque import _origen_esquina_escuadrable_si

    dato = dict(ancla_env)
    dato.setdefault("curvas", list(ancla_env.get("curvas") or []))
    dato["nombre"] = ancla_env.get("nombre") or "ANCLA"
    esq = _origen_esquina_escuadrable_si(dato)
    if esq is not None:
        ox, oy = float(esq[0]), float(esq[1])
        dato["minx"] = ox
        dato["miny"] = oy  # en SI, "miny" del dato_origen guarda Y del origen
        dato["origen_y"] = oy
        dato["origen_escuadrable"] = True
        dato["origen_tipo"] = "SI"
        return ox, oy, dato
    if exigir_esquina and list(dato.get("curvas") or []):
        return None, None, None
    # Sin HLR: AABB superior-izquierda.
    ox = float(dato["minx"])
    oy = float(dato["maxy"])
    dato["minx"] = ox
    dato["miny"] = oy
    dato["origen_y"] = oy
    dato["origen_escuadrable"] = False
    dato["origen_tipo"] = "SI_AABB"
    return ox, oy, dato


def _profundidad_camara_hijo(vista, hijo):
    """Mayor = más cerca / de frente a la cámara de la vista."""
    from generador_caras_tanque import _vector_hacia_camara

    rb = hijo.get("rb")
    if rb is None or vista is None:
        return 0.0
    try:
        look, target = _vector_hacia_camara(vista)
        cx = 0.5 * (float(rb.MinPoint.X) + float(rb.MaxPoint.X))
        cy = 0.5 * (float(rb.MinPoint.Y) + float(rb.MaxPoint.Y))
        cz = 0.5 * (float(rb.MinPoint.Z) + float(rb.MaxPoint.Z))
        return (
            (cx - float(target[0])) * float(look[0])
            + (cy - float(target[1])) * float(look[1])
            + (cz - float(target[2])) * float(look[2])
        )
    except Exception:
        return 0.0


def _score_ancla_3d(hijo):
    """Score de ancla estable (3D), independiente de la vista."""
    nom = str(hijo.get("name") or "")
    tam = _tamano_3d_in(hijo)
    if tam is None:
        return 0.0
    sx, sy, sz = tam
    cara = max(sx * sy, sy * sz, sx * sz)
    vol = max(sx * sy * sz, 1e-6)
    espesor = max(min(sx, sy, sz), 1e-3)
    plate_ratio = cara / espesor
    score = cara * (1.0 + min(plate_ratio, 40.0) * 0.2) + vol * 0.05
    if hijo.get("es_hw") or _es_hw(nom):
        return score * 0.02
    if str(nom).upper().startswith("SP-"):
        return score * 0.08
    if _es_placa_p(nom):
        return score * 6.0
    return score


def _elegir_ancla_kit_3d(hijos):
    """Ancla del kit por geometría 3D (misma en FRONT/BACK/…)."""
    utiles = [h for h in (hijos or []) if h.get("rb") is not None]
    if not utiles:
        return None
    return max(utiles, key=_score_ancla_3d)


def _tamano_3d_in(hijo):
    """(dx, dy, dz) en pulgadas del RangeBox 3D; None si falla."""
    rb = hijo.get("rb")
    if rb is None:
        return None
    try:
        return (
            abs(float(rb.MaxPoint.X) - float(rb.MinPoint.X)) / IN_TO_CM,
            abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y)) / IN_TO_CM,
            abs(float(rb.MaxPoint.Z) - float(rb.MinPoint.Z)) / IN_TO_CM,
        )
    except Exception:
        return None


def _score_ancla(item):
    """
    Ranking de ancla en UNA vista.

    1) Esquina escuadrable (no arco).
    2) Pieza más cercana / de frente a la cámara (no confundir al operador).
    3) Placa ``*-P##`` grande.
    """
    from generador_caras_tanque import _tiene_esquina_escuadrable

    h, env = item
    nom = str(h.get("name") or "")
    dx = abs(float(env.get("dx") or (env.get("maxx", 0) - env.get("minx", 0))))
    dy = abs(float(env.get("dy") or (env.get("maxy", 0) - env.get("miny", 0))))
    area_2d = float(env.get("area") or (dx * dy))

    plate_ratio = 1.0
    tam = _tamano_3d_in(h)
    if tam is not None:
        sx, sy, sz = tam
        cara = max(sx * sy, sy * sz, sx * sz)
        espesor = max(min(sx, sy, sz), 1e-3)
        plate_ratio = cara / espesor

    score = area_2d * (1.0 + min(plate_ratio, 40.0) * 0.15)
    tiene_esquina = False
    try:
        tiene_esquina = _tiene_esquina_escuadrable(env)
    except Exception:
        tiene_esquina = False
    if list(env.get("curvas") or []):
        if tiene_esquina:
            score *= 2.5
            # Preferir esquina ya cercana a SI del env (arriba-izq).
            try:
                from generador_caras_tanque import _origen_esquina_escuadrable_si

                esq = _origen_esquina_escuadrable_si(env)
                if esq is not None:
                    span_x = max(dx, 1e-3)
                    span_y = max(dy, 1e-3)
                    # 1 = perfecto SI del bbox; 0 = lejos.
                    cercania_si = 1.0 - 0.5 * (
                        abs(float(esq[0]) - float(env["minx"])) / span_x
                        + abs(float(esq[1]) - float(env["maxy"])) / span_y
                    )
                    score *= 1.0 + max(0.0, cercania_si) * 0.35
            except Exception:
                pass
        else:
            score *= 0.04
    # De frente / cercana a la cámara de ESTA vista.
    prox = float(h.get("_prox_camara") or 0.0)
    score *= 1.0 + max(0.0, min(prox, 50.0)) * 0.04
    if h.get("_ancla_kit_3d"):
        score *= 3.0 if tiene_esquina or not list(env.get("curvas") or []) else 0.2
    if h.get("es_hw") or _es_hw(nom):
        return score * 0.02
    if str(nom).upper().startswith("SP-"):
        return score * 0.08
    if _es_placa_p(nom):
        return score * 5.0
    return score


def _rank_anclas(proy):
    """Anclas candidatas, mejor primero."""
    orden = sorted(proy, key=_score_ancla, reverse=True)
    vistos = set()
    out = []
    for item in orden:
        nom = item[0].get("name")
        if nom in vistos:
            continue
        vistos.add(nom)
        out.append(item)
    return out


def _elegir_ancla(proy):
    """Prefiere placa ``*-P##`` / cara grande; penaliza SP/HW."""
    return _rank_anclas(proy)[0]


def _posiciones_desde_hijos(
    vista, tg, hijos, origen_x, origen_y, ancla_pieza_id, ancla_name=None
):
    """
    Cotas H/V desde el origen SI del ancla a cada INSTANCIA útil.

    - Omite solo la ocurrencia ancla (``pieza_id``), no el resto del mismo
      part number (bug típico: 2×P48 → solo se acotaba P47).
    - Cada no-HW intenta X e Y (≥ MIN_COTA_PLACA_IN).
    - HW: solo eje dominante ≥ MIN_COTA_INSTRUCTIVO_IN.
    - Con origen SI: tips preferidos izq/sup (hacia el origen arriba-izq).
    """
    from generador_caras_tanque import _punto_toque_real_pieza

    pos_x = []
    pos_y = []
    n_const = 0
    # pieza_id → set ejes ya cubiertos
    ejes_ok = {}
    ancla_pid = str(ancla_pieza_id or "")

    def _min_eje(es_hw):
        return MIN_COTA_INSTRUCTIVO_IN if es_hw else MIN_COTA_PLACA_IN

    def _cands_eje(base, env, origen, eje, min_in):
        """Candidatos válidos: tip cerca del origen SI primero."""
        lados_keys = (
            (("izq", "minx"), ("der", "maxx"))
            if eje == "X"
            else (("sup", "maxy"), ("inf", "miny"))
        )
        out = []
        for lado, key in lados_keys:
            cand = {**base, "valor": float(env[key]), "lado": lado}
            if eje == "X":
                tx, _ = _punto_toque_real_pieza(cand, "X", preferir_recta=True)
                cand["valor"] = float(tx)
            else:
                _, ty = _punto_toque_real_pieza(cand, "Y", preferir_recta=True)
                cand["valor"] = float(ty)
            dist = _distancia_pulgadas_hoja(vista, cand["valor"], origen)
            if (
                abs(float(cand["valor"]) - float(origen)) >= MIN_DIST_ORIGEN_CM
                and dist >= float(min_in)
            ):
                out.append((dist, cand))
        # Más cerca del origen primero.
        out.sort(key=lambda t: t[0])
        return out

    for h in hijos:
        pieza_id = str(h.get("pieza_id") or h["name"])
        if pieza_id == ancla_pid:
            continue
        # Nunca excluir por solo coincidir el part number del ancla.
        env, constructiva = _envolvente_hijo_vista(vista, tg, h)
        if env is None:
            continue
        if constructiva:
            n_const += 1
        es_hw = bool(h.get("es_hw"))
        min_in = _min_eje(es_hw)
        base = {
            "dato": env,
            "clave": pieza_id,
            "pieza_id": pieza_id,
            "constructiva": bool(constructiva),
            "es_hw": es_hw,
        }
        cx = _cands_eje(base, env, origen_x, "X", min_in)
        cy = _cands_eje(base, env, origen_y, "Y", min_in)
        cubiertos = ejes_ok.setdefault(pieza_id, set())
        if es_hw:
            dx_in = cx[0][0] if cx else 0.0
            dy_in = cy[0][0] if cy else 0.0
            if dx_in >= dy_in and cx and "X" not in cubiertos:
                pos_x.append(cx[0][1])
                cubiertos.add("X")
            elif cy and "Y" not in cubiertos:
                pos_y.append(cy[0][1])
                cubiertos.add("Y")
        else:
            if cx and "X" not in cubiertos:
                pos_x.append(cx[0][1])
                cubiertos.add("X")
            if cy and "Y" not in cubiertos:
                pos_y.append(cy[0][1])
                cubiertos.add("Y")

    # Segunda pasada: completar eje faltante (umbral más permisivo).
    for h in hijos:
        pieza_id = str(h.get("pieza_id") or h["name"])
        if pieza_id == ancla_pid:
            continue
        cubiertos = ejes_ok.setdefault(pieza_id, set())
        faltan = [e for e in ("X", "Y") if e not in cubiertos]
        if not faltan:
            continue
        if bool(h.get("es_hw")) and cubiertos:
            continue  # HW: un solo eje basta
        env, constructiva = _envolvente_hijo_vista(vista, tg, h)
        if env is None:
            continue
        es_hw = bool(h.get("es_hw"))
        # Forzar con umbral de placa aunque sea HW sin nada aún.
        min_in = MIN_COTA_PLACA_IN
        base = {
            "dato": env,
            "clave": pieza_id,
            "pieza_id": pieza_id,
            "constructiva": bool(constructiva),
            "es_hw": es_hw,
        }
        for eje in faltan:
            origen = origen_x if eje == "X" else origen_y
            cands = _cands_eje(base, env, origen, eje, min_in)
            if not cands:
                continue
            dist, cand = cands[0]
            if eje == "X":
                pos_x.append(cand)
            else:
                pos_y.append(cand)
            cubiertos.add(eje)
            _log(
                f"    forzada {eje} ({cand.get('lado')}) → {pieza_id} "
                f"({dist:.3f} in)"
            )

    # Tercera pasada: tip más lejano aunque sea < MIN_COTA_PLACA (no-HW).
    # Evita piezas “fantasma” pegadas al ancla que igual deben acotarse.
    for h in hijos:
        pieza_id = str(h.get("pieza_id") or h["name"])
        if pieza_id == ancla_pid or bool(h.get("es_hw")):
            continue
        cubiertos = ejes_ok.setdefault(pieza_id, set())
        faltan = [e for e in ("X", "Y") if e not in cubiertos]
        if not faltan:
            continue
        env, constructiva = _envolvente_hijo_vista(vista, tg, h)
        if env is None:
            continue
        base = {
            "dato": env,
            "clave": pieza_id,
            "pieza_id": pieza_id,
            "constructiva": bool(constructiva),
            "es_hw": False,
        }
        for eje in faltan:
            origen = float(origen_x if eje == "X" else origen_y)
            lados_keys = (
                (("izq", "minx"), ("der", "maxx"))
                if eje == "X"
                else (("sup", "maxy"), ("inf", "miny"))
            )
            mejor = None
            for lado, key in lados_keys:
                cand = {**base, "valor": float(env[key]), "lado": lado}
                if eje == "X":
                    tx, _ = _punto_toque_real_pieza(cand, "X", preferir_recta=True)
                    cand["valor"] = float(tx)
                else:
                    _, ty = _punto_toque_real_pieza(cand, "Y", preferir_recta=True)
                    cand["valor"] = float(ty)
                if abs(float(cand["valor"]) - origen) < 1e-4:
                    continue
                dist = abs(float(cand["valor"]) - origen)
                if mejor is None or dist > mejor[0]:
                    mejor = (dist, cand)
            if mejor is None:
                continue
            _, cand = mejor
            if eje == "X":
                pos_x.append(cand)
            else:
                pos_y.append(cand)
            cubiertos.add(eje)
            _log(
                f"    forzada-extrema {eje} ({cand.get('lado')}) → {pieza_id}"
            )

    omitidos = []
    incompletos = []
    for h in hijos:
        pieza_id = str(h.get("pieza_id") or h["name"])
        if pieza_id == ancla_pid:
            continue
        cub = ejes_ok.get(pieza_id) or set()
        if not cub:
            omitidos.append(pieza_id)
        elif (not h.get("es_hw")) and cub != {"X", "Y"}:
            incompletos.append(f"{pieza_id}:{''.join(sorted(cub))}")
    if omitidos:
        _log(
            f"    sin cota (≥{MIN_COTA_PLACA_IN}in): "
            f"{', '.join(omitidos[:8])}"
        )
    if incompletos:
        _log(
            f"    incompletas X+Y: {', '.join(incompletos[:8])}"
        )
    if n_const:
        _log(f"    constructivas (HLR pobre/ausente): {n_const} pieza(s)")
    return (
        _agrupar_typ(pos_x, origen_x=origen_x, origen_y=origen_y, eje="X"),
        _agrupar_typ(pos_y, origen_x=origen_x, origen_y=origen_y, eje="Y"),
    )



def _token_pieza(nombre):
    from nomenclatura_capturas import limpiar_token_archivo

    return limpiar_token_archivo(nombre)[:48] or "PIEZA"


def _nombre_jpg_ensamble(
    nombre_job,
    ancla_id,
    pos,
    eje,
    etiqueta_vista,
    valor_cota,
    secuencia=0,
):
    """
    JOB__ITEM__VISTA_S##_XMIN[_TYP_N]_VALOR.jpg

    ITEM = pieza principal de la cota (clave de tipo / representante).
    Sin cadenas Item_X_Item ni listado de miembros TYP.
    Si es TYP, solo se añade ``_TYP_N`` con la cantidad.
    ``ancla_id`` se conserva en la firma por compatibilidad; ya no va al nombre
    (el kit/carpeta identifica el ensamble).
    """
    from generador_caras_tanque import _clave_tipo_pieza
    from nomenclatura_capturas import (
        limpiar_token_archivo,
        formatear_valor_en_nombre,
    )

    del ancla_id  # no va en el nombre; firma estable para callers
    miembros = list(pos.get("miembros") or [pos])
    ids = {
        str(m.get("pieza_id") or m.get("clave") or "").strip()
        for m in miembros
        if str(m.get("pieza_id") or m.get("clave") or "").strip()
    }
    raw = (
        pos.get("clave")
        or pos.get("pieza_id")
        or (miembros[0].get("pieza_id") if miembros else None)
        or (miembros[0].get("clave") if miembros else None)
        or "PIEZA"
    )
    tipo = _clave_tipo_pieza(raw) or str(raw)
    item = _token_pieza(tipo)

    medida = "XMIN" if str(eje).upper() == "X" else "YMIN"
    n_typ = max(len(ids), len(miembros), 1)
    if pos.get("typ") and n_typ >= 2:
        medida = f"{medida}_TYP_{n_typ}"

    job = limpiar_token_archivo(nombre_job)
    vista = limpiar_token_archivo(etiqueta_vista).upper()
    seq = f"S{int(secuencia):02d}_" if int(secuencia or 0) > 0 else ""
    valor = formatear_valor_en_nombre(valor_cota)
    return f"{job}__{item}__{vista}_{seq}{medida}_{valor}.jpg"


def _centro_bbox_kit(asm_doc):
    rb = asm_doc.ComponentDefinition.RangeBox
    return (
        (float(rb.MinPoint.X) + float(rb.MaxPoint.X)) / 2.0,
        (float(rb.MinPoint.Y) + float(rb.MaxPoint.Y)) / 2.0,
        (float(rb.MinPoint.Z) + float(rb.MaxPoint.Z)) / 2.0,
    )


def _camara_desde_eye_up(asm_doc, tg, to, cx, cy, cz, eye, up):
    import creador_vistas

    eye_v = tg.CreateVector(float(eye[0]), float(eye[1]), float(eye[2]))
    up_v = tg.CreateVector(float(up[0]), float(up[1]), float(up[2]))
    return creador_vistas.crear_camara(
        asm_doc, tg, to, cx, cy, cz, eye_v, up_v
    )


# ViewCube completo (6 caras).
_CAM_SPECS = {
    "FRONT": ((0, 0, 1), (0, 1, 0)),
    "BACK": ((0, 0, -1), (0, 1, 0)),
    "TOP": ((0, 1, 0), (0, 0, -1)),
    "BOTTOM": ((0, -1, 0), (0, 0, 1)),
    "RIGHT": ((1, 0, 0), (0, 0, 1)),
    "LEFT": ((-1, 0, 0), (0, 0, 1)),
}


def _camaras_ortogonales(asm_doc, tg, to):
    """FRONT/BACK/TOP/BOTTOM/RIGHT/LEFT = ViewCube Inventor."""
    cx, cy, cz = _centro_bbox_kit(asm_doc)
    cams = []
    for nombre in VISTAS:
        eye, up = _CAM_SPECS[nombre]
        cam = _camara_desde_eye_up(asm_doc, tg, to, cx, cy, cz, eye, up)
        cams.append((nombre, cam))
    return cams


def _ordenar_cotas_secuencia(pos_x, pos_y, origen_x, origen_y):
    """
    Orden de armado: cerca → lejos del origen (distancia en hoja).

    X e Y se intercalan por distancia para S01, S02… en la vista.
    """
    items = []
    for pos in pos_x or []:
        d = abs(float(pos["valor"]) - float(origen_x))
        items.append(("X", pos, d))
    for pos in pos_y or []:
        d = abs(float(pos["valor"]) - float(origen_y))
        items.append(("Y", pos, d))
    items.sort(key=lambda t: (t[2], t[0], abs(float(t[1]["valor"]))))
    return items


def _agrupar_typ(posiciones, tol=TOL_TYP_CM, origen_x=None, origen_y=None, eje=None):
    """
    Consolida cotas iguales SOLO entre la misma familia de pieza.

    P18_405 + P18_412 en la misma X → TYP. P17 + P18 en la misma X → NO.
    Miembros TYP ordenados cerca → lejos del origen (letras A/B/C).
    """
    from generador_caras_tanque import (
        _clave_tipo_pieza,
        _ordenar_miembros_typ_cerca_origen,
    )

    if not posiciones:
        return []
    orden = sorted(posiciones, key=lambda p: float(p["valor"]))
    grupos = []
    for pos in orden:
        tipo = _clave_tipo_pieza(pos.get("pieza_id") or pos.get("clave") or "")
        colocado = False
        for g in grupos:
            if abs(float(pos["valor"]) - float(g["valor"])) > tol:
                continue
            mismo_id = any(
                m.get("pieza_id") == pos.get("pieza_id")
                and m.get("lado") != pos.get("lado")
                for m in g["miembros"]
            )
            if mismo_id:
                continue
            tipos_g = {
                _clave_tipo_pieza(m.get("pieza_id") or m.get("clave") or "")
                for m in g["miembros"]
            }
            tipos_g.discard("")
            # Regla dura: no mezclar familias distintas en un TYP.
            if tipo and tipos_g and tipo not in tipos_g:
                continue
            if tipos_g and not tipo:
                continue
            g["miembros"].append(pos)
            g["qty"] = len(g["miembros"])
            if abs(float(pos["valor"])) >= abs(float(g["valor"])):
                g["valor"] = float(pos["valor"])
                g["dato"] = pos["dato"]
                g["lado"] = pos.get("lado", g.get("lado"))
                g["clave"] = pos.get("clave") or g.get("clave")
            colocado = True
            break
        if not colocado:
            grupos.append(
                {
                    "valor": float(pos["valor"]),
                    "lado": pos.get("lado", "izq"),
                    "dato": pos["dato"],
                    "miembros": [pos],
                    "qty": 1,
                    "typ": False,
                    "clave": pos.get("clave")
                    or pos["dato"].get("nombre", "PIEZA"),
                    "pieza_id": pos.get("pieza_id") or pos.get("clave"),
                }
            )

    consolidados = []
    for g in grupos:
        miembros = _ordenar_miembros_typ_cerca_origen(
            g["miembros"],
            origen_x=origen_x,
            origen_y=origen_y,
            eje=eje,
        )
        ids = {str(m.get("pieza_id") or m.get("clave") or "") for m in miembros}
        tipos = {
            _clave_tipo_pieza(m.get("pieza_id") or m.get("clave") or "")
            for m in miembros
        }
        tipos.discard("")
        es_typ = len(ids) > 1 and len(tipos) <= 1
        rep = dict(miembros[0])
        rep["valor"] = float(g["valor"])
        rep["lado"] = g.get("lado", rep.get("lado", "izq"))
        rep["dato"] = g.get("dato") or rep.get("dato")
        rep["miembros"] = miembros
        rep["qty"] = len(miembros)
        rep["typ"] = bool(es_typ)
        rep["clave"] = g.get("clave") or rep.get("clave")
        rep["constructiva"] = any(bool(m.get("constructiva")) for m in miembros)
        if es_typ and tipos:
            tipo_txt = next(iter(tipos))
            _log(
                f"    TYP {rep['lado'].upper()}×{len(miembros)} "
                f"tipo={tipo_txt} @ {rep['valor']:.3f}"
            )
        consolidados.append(rep)
    return consolidados


def _crear_hoja_vista(plano, base_sheet, asm_doc, nombre_hoja, cam, inv_app, tg, to):
    import creador_vistas

    hoja = creador_vistas._crear_hoja_vista(plano, base_sheet, nombre_hoja)
    try:
        creador_vistas._limpiar_border_y_titleblock(hoja)
    except Exception:
        pass
    px, py, ancho_util, alto_util = creador_vistas._area_util_hoja(hoja)
    vista = creador_vistas._crear_vista_base(
        hoja, asm_doc, tg, to, px, py, cam, False
    )
    if vista is None:
        try:
            hoja.Delete()
        except Exception:
            pass
        return None, None
    try:
        creador_vistas.escalar_vista(
            plano, vista, tg, px, py, ancho_util, alto_util
        )
    except Exception:
        pass
    try:
        inv_app.ActiveView.Update()
    except Exception:
        pass
    time.sleep(0.4)
    return hoja, vista


def _exportar_grupos(
    inv_app,
    plano,
    hoja,
    vista,
    tg,
    grupos,
    origen_x,
    origen_y,
    carpeta_vista,
    nombre_job,
    nombre_kit,
    etiqueta_vista,
    ancla_id=None,
    dato_origen=None,
):
    from generador_caras_tanque import (
        _borrar_sketches_cotas,
        _dibujar_cotas_hv_desde_origen,
        _bbox_foto_grupo,
        _recortar_jpg_caras,
    )
    from generador_vistas import ANCHO_EXPORTACION, ALTO_EXPORTACION

    os.makedirs(carpeta_vista, exist_ok=True)
    ancla_tok = ancla_id or nombre_kit
    # Secuencia de armado: cerca → lejos del origen (S01, S02…).
    secuencia = _ordenar_cotas_secuencia(
        grupos.get("x") or [],
        grupos.get("y") or [],
        origen_x,
        origen_y,
    )

    white = None
    try:
        white = inv_app.TransientObjects.CreateColor(255, 255, 255)
    except Exception:
        pass

    exportadas = 0
    try:
        hoja.Activate()
    except Exception:
        pass

    for indice, (eje, pos, _dist) in enumerate(secuencia, start=1):
        temporal = os.path.join(
            carpeta_vista, f"_tmp_{etiqueta_vista}_{indice}.jpg"
        )
        pos_x = [pos] if eje == "X" else []
        pos_y = [pos] if eje == "Y" else []
        try:
            _borrar_sketches_cotas(hoja)
            creadas, fallos, valores_txt = _dibujar_cotas_hv_desde_origen(
                hoja,
                vista,
                tg,
                inv_app,
                pos_x,
                pos_y,
                origen_x,
                origen_y,
                typ_un_anillo=True,
                dibujar_constructivas=True,
                anclar_toque_real=True,
                dato_origen=dato_origen,
            )
            if creadas < 1:
                _log(
                    f"    ⚠️ {etiqueta_vista} S{indice:02d}: sin cota "
                    f"(fallos={fallos})"
                )
                continue
            valor_cota = (
                str(valores_txt[0]).strip() if valores_txt else str(indice)
            )
            nombre_jpg = _nombre_jpg_ensamble(
                nombre_job,
                ancla_tok,
                pos,
                eje,
                etiqueta_vista,
                valor_cota,
                secuencia=indice,
            )
            salida = os.path.join(carpeta_vista, nombre_jpg)
            bbox_foto = _bbox_foto_grupo(
                vista,
                pos_x,
                pos_y,
                origen_x,
                origen_y,
            )
            inv_app.ActiveView.Camera.SaveAsBitmap(
                temporal,
                ANCHO_EXPORTACION,
                ALTO_EXPORTACION,
                white,
            )
            time.sleep(0.12)
            _recortar_jpg_caras(
                hoja,
                temporal,
                salida,
                vista,
                bbox_forzada=bbox_foto,
            )
            exportadas += 1
        except Exception as err:
            _log(f"    ERROR export {etiqueta_vista} S{indice:02d}: {err}")
        finally:
            try:
                _borrar_sketches_cotas(hoja)
            except Exception:
                pass
            try:
                if os.path.isfile(temporal):
                    os.remove(temporal)
            except OSError:
                pass
    return exportadas


def _cotas_de_vista(vista, tg, hijos, ancla_preferida=None):
    """
    Proyecta hijos y elige ancla: esquina SI + pieza de frente a la cámara.

    Omite solo la instancia ancla; el resto (mismo P## incluido) se acota
    en X e Y.
    """
    proy = []
    pref = str(ancla_preferida or "")
    # Normalizar profundidad de cámara (0..~span).
    profundidades = []
    for h in hijos:
        profundidades.append(_profundidad_camara_hijo(vista, h))
    pmin = min(profundidades) if profundidades else 0.0
    pmax = max(profundidades) if profundidades else 1.0
    pspan = max(pmax - pmin, 1e-6)

    for h, prox_raw in zip(hijos, profundidades):
        env, _const = _envolvente_hijo_vista(vista, tg, h)
        if env is None:
            continue
        if "area" not in env:
            env["area"] = abs(
                float(env.get("maxx", 0) - env.get("minx", 0))
            ) * abs(float(env.get("maxy", 0) - env.get("miny", 0)))
        env.setdefault(
            "dx", abs(float(env.get("maxx", 0) - env.get("minx", 0)))
        )
        env.setdefault(
            "dy", abs(float(env.get("maxy", 0) - env.get("miny", 0)))
        )
        hh = dict(h)
        hh["_prox_camara"] = (float(prox_raw) - pmin) / pspan * 10.0
        if pref and str(h.get("name") or "") == pref:
            hh["_ancla_kit_3d"] = True
        proy.append((hh, env))
    if len(proy) < 2:
        return None

    for idx, (ancla_h, ancla_env) in enumerate(_rank_anclas(proy)):
        origen_x, origen_y, dato_origen = _origen_desde_ancla(
            ancla_env, exigir_esquina=True
        )
        if origen_x is None:
            _log(
                f"    ancla descartada (sin esquina SI escuadrable): "
                f"{ancla_h.get('pieza_id') or ancla_h.get('name')}"
            )
            continue
        ancla_id = ancla_h.get("pieza_id") or ancla_h["name"]
        pos_x, pos_y = _posiciones_desde_hijos(
            vista,
            tg,
            hijos,
            origen_x,
            origen_y,
            ancla_pieza_id=ancla_id,
            ancla_name=ancla_h.get("name"),
        )
        if not pos_x and not pos_y:
            continue
        if idx > 0:
            _log(
                f"    ancla alternativa #{idx + 1}: {ancla_id} "
                f"(SI escuadrable + frente a cámara)"
            )
        return {
            "ancla_h": ancla_h,
            "ancla_id": ancla_id,
            "origen_x": origen_x,
            "origen_y": origen_y,
            "dato_origen": dato_origen,
            "pos_x": pos_x,
            "pos_y": pos_y,
            "n_proy": len(proy),
            "ancla_intento": idx + 1,
        }
    return None


def _rutas_corte_maquinado(carpeta_tanque):
    """Raíz PIEZAS_ACOTADAS/Corte/Maquinado y subcarpetas del proceso."""
    from generador_tanque_completo import (
        SUBCARPETA_ACCESORIOS_POR_PIEZA,
        SUBCARPETA_ACCESORIOS_SUELTOS,
        SUBCARPETA_INSPECCION_VISUAL,
    )

    piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS)
    corte_maq = os.path.join(piezas, "Corte", "Maquinado")
    return {
        "piezas": piezas,
        "corte_maq": corte_maq,
        "accesorios": os.path.join(corte_maq, SUBCARPETA_ACCESORIOS_SUELTOS),
        "inspeccion": os.path.join(corte_maq, SUBCARPETA_INSPECCION_VISUAL),
        "por_pieza": os.path.join(corte_maq, SUBCARPETA_ACCESORIOS_POR_PIEZA),
    }


def _limpiar_carpetas_ensambles(rutas, carpeta_tanque):
    """Vacía Accesorios/Inspeccion/por pieza y legacy ENSAMBLES_INDEPENDIENTES."""
    for key in ("accesorios", "inspeccion", "por_pieza"):
        ruta = rutas.get(key)
        if ruta and os.path.isdir(ruta):
            shutil.rmtree(ruta, ignore_errors=True)
    legacy = os.path.join(carpeta_tanque, CARPETA_ENSAMBLES_LEGACY)
    if os.path.isdir(legacy):
        shutil.rmtree(legacy, ignore_errors=True)


def _camara_isometrica_kit(asm_doc, tg, to):
    """Isométrica del kit (mismo criterio que ESTANIADO cobre, sobre .iam)."""
    import creador_vistas

    cx, cy, cz = _centro_bbox_kit(asm_doc)
    v_frente = tg.CreateVector(0.0, 0.0, 1.0)
    up_hint = tg.CreateVector(0.0, 1.0, 0.0)
    return creador_vistas.crear_camara_isometrica_desde_frente(
        asm_doc, tg, to, cx, cy, cz, v_frente, up_hint
    )


def _exportar_jpg_hoja(inv_app, plano, hoja, ruta_jpg):
    """Exporta la hoja activa a JPG (sin cotas / Inspección Visual)."""
    from generador_vistas import ANCHO_EXPORTACION, ALTO_EXPORTACION

    os.makedirs(os.path.dirname(ruta_jpg) or ".", exist_ok=True)
    try:
        hoja.Activate()
    except Exception:
        pass
    try:
        inv_app.ActiveView.Update()
    except Exception:
        pass
    time.sleep(0.35)
    white = None
    try:
        white = inv_app.TransientObjects.CreateColor(255, 255, 255)
    except Exception:
        pass
    try:
        plano.SaveAsBitmap(
            ruta_jpg,
            ANCHO_EXPORTACION,
            ALTO_EXPORTACION,
            white,
            white,
        )
        return os.path.isfile(ruta_jpg)
    except Exception as exc:
        _log(f"  AVISO export JPG Inspeccion Visual: {exc}")
        return False


def _procesar_inspeccion_visual(
    inv_app, plano, base_sheet, asm_doc, nombre_kit, carpeta_insp, nombre_job
):
    """Una captura isométrica sin cotas del ensamble completo."""
    from generador_vistas import borrar_hojas_por_nombres, _nombre_hoja_machote
    import creador_vistas
    from generador_tanque_completo import _nombre_carpeta_pieza

    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    try:
        cam = _camara_isometrica_kit(asm_doc, tg, to)
    except Exception as exc:
        _log(f"  {nombre_kit}: cámara isométrica falló ({exc})")
        return 0

    nombre_hoja = creador_vistas.construir_nombre_hoja(
        plano, nombre_kit, "ISO_INSP"
    )
    hoja, vista = _crear_hoja_vista(
        plano, base_sheet, asm_doc, nombre_hoja, cam, inv_app, tg, to
    )
    if hoja is None:
        return 0

    carpeta_kit = os.path.join(carpeta_insp, _nombre_carpeta_pieza(nombre_kit))
    if os.path.isdir(carpeta_kit):
        shutil.rmtree(carpeta_kit, ignore_errors=True)
    os.makedirs(carpeta_kit, exist_ok=True)
    nombre_jpg = f"{nombre_job}__{nombre_kit}__ISO_INSP_1.jpg"
    ruta_jpg = os.path.join(carpeta_kit, nombre_jpg)
    ok = _exportar_jpg_hoja(inv_app, plano, hoja, ruta_jpg)

    nombre_machote = _nombre_hoja_machote(plano)
    try:
        borrar_hojas_por_nombres(
            plano,
            {str(hoja.Name).split(":")[0]},
            nombre_machote_protegido=nombre_machote,
        )
    except Exception:
        try:
            hoja.Delete()
        except Exception:
            pass

    if ok:
        _log(f"  {nombre_kit}: Inspeccion Visual → {ruta_jpg}")
        return 1
    return 0


def _nombres_piezas_de_kits(lista_kits):
    """Nombres base únicos de .ipt leaf dentro de los kits."""
    nombres = set()
    for asm_doc, _nombre, _qty, _hijos in lista_kits:
        try:
            leaf = asm_doc.ComponentDefinition.Occurrences.AllLeafOccurrences
            n = int(leaf.Count)
        except Exception:
            continue
        for i in range(1, n + 1):
            try:
                occ = leaf.Item(i)
                if occ.Suppressed:
                    continue
                doc = occ.Definition.Document
                if int(doc.DocumentType) != TIPO_DOCUMENTO_PIEZA:
                    continue
                base = os.path.splitext(
                    os.path.basename(str(doc.FullFileName or ""))
                )[0]
                if base:
                    nombres.add(base)
            except Exception:
                continue
    return nombres


def _procesar_accesorios_por_pieza(
    inv_app, plano, ensamble, carpeta_piezas, nombres_piezas
):
    """
    Flujo Abigail (vistas+cotas+JPG) solo para piezas de los kits,
    reorganizado a Corte/Maquinado/Accesorios Sueltos por pieza/<PIEZA>/.
    """
    if not nombres_piezas:
        _log("  Accesorios Sueltos por pieza: sin .ipt en kits.")
        return 0

    from generador_vistas import ejecutar_flujo_desde_app
    from generador_tanque_completo import (
        SUBCARPETA_ACCESORIOS_POR_PIEZA,
        _reorganizar_piezas_por_clasificacion,
    )

    _log(
        f"  Accesorios Sueltos por pieza: {len(nombres_piezas)} pieza(s) "
        f"vía flujo Abigail"
    )
    staging = os.path.join(
        carpeta_piezas, "_STAGING_ACCESORIOS_POR_PIEZA"
    )
    if os.path.isdir(staging):
        shutil.rmtree(staging, ignore_errors=True)
    os.makedirs(staging, exist_ok=True)

    try:
        ok = ejecutar_flujo_desde_app(
            inv_app,
            ensamble,
            plano,
            carpeta_salida=staging,
            incremental=False,
            catalogo_piezas=set(nombres_piezas),
        )
    except Exception as exc:
        _log(f"  ERROR Accesorios Sueltos por pieza (flujo): {exc}")
        _log(traceback.format_exc())
        return 0

    mapa = {SUBCARPETA_ACCESORIOS_POR_PIEZA: set(nombres_piezas)}
    try:
        _reorganizar_piezas_por_clasificacion(staging, mapa)
    except Exception as exc:
        _log(f"  AVISO reorg por pieza: {exc}")

    # Mover árbol resultante a PIEZAS_ACOTADAS/Corte/Maquinado/...
    dest_root = os.path.join(
        carpeta_piezas,
        "Corte",
        "Maquinado",
        SUBCARPETA_ACCESORIOS_POR_PIEZA,
    )
    src_root = os.path.join(
        staging,
        "Corte",
        "Maquinado",
        SUBCARPETA_ACCESORIOS_POR_PIEZA,
    )
    # Si reorg dejó piezas en la raíz del staging con la clave sintética:
    if not os.path.isdir(src_root):
        alt = os.path.join(staging, SUBCARPETA_ACCESORIOS_POR_PIEZA)
        if os.path.isdir(alt):
            src_root = alt

    movidos = 0
    if os.path.isdir(src_root):
        os.makedirs(dest_root, exist_ok=True)
        for root, _dirs, files in os.walk(src_root):
            for fn in files:
                if not fn.lower().endswith(".jpg"):
                    continue
                src = os.path.join(root, fn)
                rel = os.path.relpath(src, src_root)
                dst = os.path.join(dest_root, rel)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                try:
                    if os.path.isfile(dst):
                        os.remove(dst)
                    shutil.move(src, dst)
                    movidos += 1
                except OSError as err:
                    _log(f"  AVISO move {fn}: {err}")
    else:
        # Fallback: cualquier JPG bajo staging → por pieza/<nombre>/
        from generador_tanque_completo import (
            _extraer_pieza_de_jpg,
            _nombre_carpeta_pieza,
        )

        os.makedirs(dest_root, exist_ok=True)
        for root, _dirs, files in os.walk(staging):
            if os.path.basename(root).casefold().startswith("_staging"):
                continue
            for fn in files:
                if not fn.lower().endswith(".jpg"):
                    continue
                pieza = _nombre_carpeta_pieza(_extraer_pieza_de_jpg(fn))
                dst_dir = os.path.join(dest_root, pieza)
                os.makedirs(dst_dir, exist_ok=True)
                src = os.path.join(root, fn)
                dst = os.path.join(dst_dir, fn)
                try:
                    if os.path.isfile(dst):
                        os.remove(dst)
                    shutil.move(src, dst)
                    movidos += 1
                except OSError as err:
                    _log(f"  AVISO move {fn}: {err}")

    shutil.rmtree(staging, ignore_errors=True)
    _log(
        f"  Accesorios Sueltos por pieza: {movidos} JPG → {dest_root} "
        f"(flujo ok={ok})"
    )
    return movidos


def _procesar_kit(
    inv_app, plano, base_sheet, asm_doc, nombre_kit, carpeta_kit, nombre_job
):
    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    hijos = _hijos_utiles(asm_doc)
    if len(hijos) < 2:
        _log(f"  {nombre_kit}: menos de 2 hijos útiles, se omite")
        return 0

    try:
        cams = _camaras_ortogonales(asm_doc, tg, to)
    except Exception as exc:
        _log(f"  {nombre_kit}: error cámaras ({exc})")
        return 0

    from generador_vistas import borrar_hojas_por_nombres, _nombre_hoja_machote
    import creador_vistas

    total_jpg = 0
    nombres_hojas = set()
    nombre_machote = _nombre_hoja_machote(plano)
    conteo_vista = {v: 0 for v in VISTAS}

    ancla_kit = _elegir_ancla_kit_3d(hijos)
    ancla_pref = ancla_kit["name"] if ancla_kit else ""
    if ancla_pref:
        _log(f"  {nombre_kit}: ancla 3D kit = {ancla_pref}")

    for etiqueta, cam in cams:
        nombre_hoja = creador_vistas.construir_nombre_hoja(
            plano, nombre_kit, etiqueta
        )
        try:
            hoja, vista = _crear_hoja_vista(
                plano, base_sheet, asm_doc, nombre_hoja, cam, inv_app, tg, to
            )
        except Exception as exc:
            _log(f"  {nombre_kit}/{etiqueta}: no se creó hoja ({exc})")
            continue
        if hoja is None:
            continue
        nombres_hojas.add(str(hoja.Name).split(":")[0])

        plan = _cotas_de_vista(
            vista, tg, hijos, ancla_preferida=ancla_pref
        )
        if plan is None:
            _log(f"  {nombre_kit}/{etiqueta}: sin cotas útiles, se omite")
            try:
                hoja.Delete()
            except Exception:
                pass
            continue

        _log(
            f"  {nombre_kit}/{etiqueta}: ancla={plan['ancla_id']} "
            f"origen_IL=({plan['origen_x']:.2f},{plan['origen_y']:.2f}) "
            f"hijos={plan['n_proy']} "
            f"cotas={len(plan['pos_x'])}+{len(plan['pos_y'])}"
            + (
                f" [alt#{plan.get('ancla_intento')}]"
                if int(plan.get("ancla_intento") or 1) > 1
                else ""
            )
        )

        carpeta_vista = os.path.join(carpeta_kit, etiqueta)
        n = _exportar_grupos(
            inv_app,
            plano,
            hoja,
            vista,
            tg,
            {"x": plan["pos_x"], "y": plan["pos_y"]},
            plan["origen_x"],
            plan["origen_y"],
            carpeta_vista,
            nombre_job,
            nombre_kit,
            etiqueta,
            ancla_id=plan["ancla_id"],
            dato_origen=plan["dato_origen"],
        )
        total_jpg += n
        conteo_vista[etiqueta] = n
        _log(f"  {nombre_kit}/{etiqueta}: {n} JPG")

    # Equilibrio: avisar vistas vacías (6 caras ViewCube).
    activas = [v for v in VISTAS if conteo_vista[v] > 0]
    if activas and len(activas) < len(VISTAS):
        vacias = [v for v in VISTAS if conteo_vista[v] == 0]
        _log(
            f"  {nombre_kit}: vistas con cotas={len(activas)}/6 "
            f"{dict(conteo_vista)} vacías={vacias}"
        )

    if nombres_hojas:
        try:
            borrar_hojas_por_nombres(
                plano, nombres_hojas, nombre_machote_protegido=nombre_machote
            )
        except Exception as exc:
            _log(f"  AVISO borrando hojas {nombre_kit}: {exc}")
    return total_jpg


def ejecutar(solo="", max_n=0, listar=False, limpiar=False, ruta_seleccion=""):
    _log("=" * 62)
    _log(" COTAS — ENSAMBLES INDEPENDIENTES (instructivo)")
    _log("=" * 62)

    pythoncom.CoInitialize()
    inv_app = None
    ok = False
    try:
        from generador_caras_tanque import (
            _cargar_seleccion_caras,
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from ensambles_independientes import recolectar_ensambles_independientes
        from nomenclatura_capturas import nombre_job_desde_ensamble
        from generador_tanque_completo import _nombre_carpeta_pieza

        inv_app = conectar_inventor()
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            _log("ERROR: abre el machote (.dwg/.idw) y el .iam del tanque.")
            return False

        from producto_tipo import (
            aplicar_unidad_producto,
            clasificar_producto,
            redirigir_si_board,
        )

        try:
            info = clasificar_producto(ensamble)
            aplicar_unidad_producto(ensamble=ensamble, info=info)
        except Exception as exc_u:
            _log(f"AVISO unidades: {exc_u}")

        desvio = redirigir_si_board(
            inv_app,
            plano,
            ensamble,
            origen_flujo="ENSAMBLES",
            gestionar_com_board=False,
            limpiar=bool(limpiar),
        )
        if desvio is not None:
            return bool(desvio)

        exclusiones = set()
        ruta_sel = str(ruta_seleccion or "").strip()
        if ruta_sel:
            if not os.path.isfile(ruta_sel):
                _log(f"ERROR: no existe --seleccion {ruta_sel}")
                return False
            _log(f"Selección caras: {ruta_sel}")
            seleccion = _cargar_seleccion_caras(ruta_sel)
            exclusiones = _exclusiones_desde_seleccion(ensamble, seleccion)
        else:
            _log(
                "AVISO: sin --seleccion; solo exclusiones OTC por nombre. "
                "La regla iLogic debe pasar TOP+SEGM+BASE."
            )

        try:
            job = nombre_job_desde_ensamble(ensamble)
        except Exception:
            job = os.path.splitext(os.path.basename(ensamble.FullFileName or "JOB"))[0]
        serie_job = _prefijo_serie(job)
        if serie_job:
            _log(f"Serie job (filtro cross-series): {serie_job}")

        lista = recolectar_ensambles_independientes(
            ensamble, exclusiones_extra=exclusiones
        )
        lista = _filtrar_kits(
            lista, solo=solo, max_n=max_n, serie_job=serie_job
        )
        _log(f"Kits instructivo a procesar: {len(lista)}")
        for _d, nom, qty, hijos in lista:
            _log(f"  - {nom}  qty={qty}  hijos={hijos}")

        if listar:
            return True
        if not lista:
            _log("Sin kits. Prueba --listar o revisa exclusiones.")
            return True

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        rutas = _rutas_corte_maquinado(carpeta_tanque)
        if limpiar:
            _limpiar_carpetas_ensambles(rutas, carpeta_tanque)
        for key in ("accesorios", "inspeccion", "por_pieza"):
            os.makedirs(rutas[key], exist_ok=True)

        try:
            base_sheet = _encontrar_hoja_machote(plano) or plano.Sheets.Item(1)
            base_sheet.Activate()
        except Exception:
            base_sheet = plano.Sheets.Item(1)

        total = 0
        total_insp = 0
        for asm_doc, nombre, qty, hijos in lista:
            _log(f"\n>>> Kit: {nombre} (qty={qty}, hijos={hijos})")
            carpeta_kit = os.path.join(
                rutas["accesorios"], _nombre_carpeta_pieza(nombre)
            )
            if os.path.isdir(carpeta_kit):
                shutil.rmtree(carpeta_kit, ignore_errors=True)
            os.makedirs(carpeta_kit, exist_ok=True)
            try:
                n = _procesar_kit(
                    inv_app, plano, base_sheet, asm_doc, nombre, carpeta_kit, job
                )
                total += n
            except Exception as exc:
                _log(f"ERROR kit {nombre}: {exc}")
                _log(traceback.format_exc())

            try:
                total_insp += _procesar_inspeccion_visual(
                    inv_app,
                    plano,
                    base_sheet,
                    asm_doc,
                    nombre,
                    rutas["inspeccion"],
                    job,
                )
            except Exception as exc:
                _log(f"ERROR Inspeccion Visual {nombre}: {exc}")
                _log(traceback.format_exc())

        nombres_piezas = _nombres_piezas_de_kits(lista)
        total_por_pieza = 0
        try:
            total_por_pieza = _procesar_accesorios_por_pieza(
                inv_app, plano, ensamble, rutas["piezas"], nombres_piezas
            )
        except Exception as exc:
            _log(f"ERROR Accesorios Sueltos por pieza: {exc}")
            _log(traceback.format_exc())

        try:
            from cotas_dossier_registro import publicar_y_sincronizar_dossier

            publicar_y_sincronizar_dossier(rutas["piezas"])
        except Exception as exc:
            _log(f"AVISO dossier: {exc}")

        _log(
            f"\nListo: instructivo={total} JPG | "
            f"inspeccion={total_insp} | por_pieza={total_por_pieza}"
        )
        _log(f"  Accesorios Sueltos → {rutas['accesorios']}")
        _log(f"  Inspeccion Visual → {rutas['inspeccion']}")
        _log(f"  Accesorios Sueltos por pieza → {rutas['por_pieza']}")
        ok = True
        return True
    except Exception as exc:
        _log(f"ERROR fatal: {exc}")
        _log(traceback.format_exc())
        return False
    finally:
        if inv_app is not None:
            try:
                from generador_caras_tanque import (
                    _encontrar_hoja_machote,
                    _obtener_plano_activo,
                )

                plano = _obtener_plano_activo(inv_app)
                hoja = _encontrar_hoja_machote(plano)
                if hoja is not None:
                    hoja.Activate()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        if ok:
            _log("PROCESO COMPLETO: instructivo de ensambles exportado.")


if __name__ == "__main__":
    args = _parse_args()
    sys.exit(
        0
        if ejecutar(
            solo=args.solo,
            max_n=args.max,
            listar=args.listar,
            limpiar=args.limpiar,
            ruta_seleccion=args.seleccion,
        )
        else 1
    )
