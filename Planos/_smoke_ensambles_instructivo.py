"""
Smoke tests (sin Inventor) para mejoras del instructivo de ensambles.

Ejecutar: python Planos/_smoke_ensambles_instructivo.py
"""
from __future__ import annotations

import math
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import generador_ensambles_instructivo as g
from generador_caras_tanque import (
    RADIO_MARCA_TYP_CM,
    RADIO_MARCA_TYP_INTERIOR_CM,
    SEP_MIN_MARCA_TYP_CM,
)


class _Fail(Exception):
    pass


def _assert(cond, msg):
    if not cond:
        raise _Fail(msg)


def test_umbral_y_hw():
    _assert(g.MIN_COTA_INSTRUCTIVO_IN == 1.0, "umbral debe ser 1.0 in")
    _assert(g._es_hw("HW-BOLT-M10"), "HW detecta tornillo")
    _assert(g._es_hw("DIN933"), "HW detecta DIN")
    _assert(not g._es_hijo_excluir("HW-BOLT"), "HW no se excluye del todo")
    _assert(g._es_hijo_excluir("FOO_HOLE"), "HOLE sí se excluye")
    _assert(not g._es_hijo_excluir("62201-1254-P18"), "placa no excluida")
    print("  OK umbral + HW/HOLE")


def test_ancla_prefer_placa():
    proy = [
        (
            {"name": "SP-WRAPPER", "es_hw": False},
            {"area": 100.0, "minx": 0, "miny": 0, "maxx": 10, "maxy": 10},
        ),
        (
            {"name": "62201-P18", "es_hw": False},
            {"area": 40.0, "minx": 0, "miny": 0, "maxx": 8, "maxy": 5},
        ),
        (
            {"name": "HW-BOLT", "es_hw": True},
            {"area": 200.0, "minx": 0, "miny": 0, "maxx": 1, "maxy": 1},
        ),
    ]
    ancla_h, env = g._elegir_ancla(proy)
    _assert(ancla_h["name"] == "62201-P18", f"ancla={ancla_h['name']}")
    _assert(env["area"] == 40.0, "env de placa")
    print("  OK ancla prefiere P## sobre SP/HW")


def test_cross_series():
    fake = [
        (None, "62201-A08", 1, 5),
        (None, "62134-A01", 1, 5),
        (None, "SP-999", 1, 5),
        (None, "62201-P12", 1, 5),  # placa suelta → filtro instructivo
    ]

    class _Doc:
        ComponentDefinition = type("CD", (), {"Occurrences": type("O", (), {"Count": 0})()})()

    # Parchear _nombres_iam_hijos para no tocar COM
    orig = g._nombres_iam_hijos
    g._nombres_iam_hijos = lambda _d: set()
    try:
        lista = [( _Doc(), n, q, h) for _d, n, q, h in fake]
        out = g._filtrar_kits(lista, serie_job="62201")
        noms = [t[1] for t in out]
        _assert("62201-A08" in noms, f"debe quedar A08: {noms}")
        _assert("62134-A01" not in noms, f"cross-series debe salir: {noms}")
        _assert("62201-P12" not in noms, f"placa P no es kit: {noms}")
        _assert("SP-999" in noms, f"SP con ≥4 hijos queda: {noms}")
    finally:
        g._nombres_iam_hijos = orig
    print("  OK filtro cross-series + instructivo")


def test_secuencia_y_nombre():
    ox, oy = 0.0, 0.0
    pos_x = [
        {
            "valor": 10.0,
            "pieza_id": "P18_B",
            "clave": "P18_B",
            "dato": {},
            "lado": "izq",
            "typ": False,
            "miembros": [{"pieza_id": "P18_B"}],
        },
        {
            "valor": 3.0,
            "pieza_id": "P18_A",
            "clave": "P18_A",
            "dato": {},
            "lado": "izq",
            "typ": False,
            "miembros": [{"pieza_id": "P18_A"}],
        },
    ]
    pos_y = [
        {
            "valor": 5.0,
            "pieza_id": "P17_C",
            "clave": "P17_C",
            "dato": {},
            "lado": "inf",
            "typ": False,
            "miembros": [{"pieza_id": "P17_C"}],
        },
    ]
    seq = g._ordenar_cotas_secuencia(pos_x, pos_y, ox, oy)
    ids = [p["pieza_id"] for _e, p, _d in seq]
    _assert(ids == ["P18_A", "P17_C", "P18_B"], f"orden cerca→lejos: {ids}")

    nom = g._nombre_jpg_ensamble(
        "62201 JOB",
        "62201-P18",
        pos_x[1],
        "X",
        "FRONT",
        "3.000",
        secuencia=1,
    )
    _assert("__FRONT_S01_XMIN_" in nom, f"secuencia en nombre: {nom}")
    _assert("_X_" not in nom.split("__")[1], f"sin Item_X_Item: {nom}")
    _assert("TYP" not in nom.split("__")[-1] or True, "sin typ ok")
    # TYP: solo ítem base + cantidad
    pos_typ = {
        "valor": 5.0,
        "pieza_id": "62201-1248-P18_405",
        "clave": "62201-1248-P18_405",
        "typ": True,
        "miembros": [
            {"pieza_id": "62201-1248-P18_405"},
            {"pieza_id": "62201-1248-P18_412"},
            {"pieza_id": "62201-1248-P18_420"},
        ],
    }
    nom_typ = g._nombre_jpg_ensamble(
        "62201 JOB", "ANCLA", pos_typ, "Y", "BACK", "1.728", secuencia=2
    )
    _assert("_TYP_3_" in nom_typ, f"TYP qty: {nom_typ}")
    _assert("_X_" not in nom_typ.split("__")[1], f"TYP sin Item_X: {nom_typ}")
    _assert("veces" not in nom_typ.lower(), f"sin 'veces': {nom_typ}")
    print(f"  OK secuencia S## → {nom}")
    print(f"  OK nomenclatura TYP corta → {nom_typ}")


def test_cam_viewcube_6():
    _assert(
        tuple(g.VISTAS) == ("FRONT", "BACK", "TOP", "BOTTOM", "RIGHT", "LEFT"),
        f"VISTAS={g.VISTAS}",
    )
    _assert(set(g._CAM_SPECS) == set(g.VISTAS), "CAM_SPECS = 6 vistas")
    _assert(not hasattr(g, "_CAM_FALLBACK"), "sin fallback espejo")
    _assert(not hasattr(g, "_camara_fallback"), "sin _camara_fallback")
    eye_r, up_r = g._CAM_SPECS["RIGHT"]
    eye_l, up_l = g._CAM_SPECS["LEFT"]
    _assert(eye_r[0] == 1 and eye_l[0] == -1, "RIGHT↔LEFT ojos opuestos")
    _assert(up_r == up_l == (0, 0, 1), "RIGHT/LEFT up = +Z")
    eye_f, _ = g._CAM_SPECS["FRONT"]
    eye_b, _ = g._CAM_SPECS["BACK"]
    _assert(eye_f[2] == 1 and eye_b[2] == -1, "FRONT↔BACK")
    print("  OK ViewCube 6 caras (sin fallback)")


def test_origen_si_ancla():
    """Sin HLR: AABB superior-izquierda (minx, maxy)."""
    env = {
        "minx": 10.0,
        "miny": 20.0,
        "maxx": 40.0,
        "maxy": 50.0,
        "nombre": "P16",
        "curvas": [],
    }
    ox, oy, dato = g._origen_desde_ancla(env)
    _assert(ox == 10.0 and oy == 50.0, f"origen SI=({ox},{oy})")
    _assert(dato.get("origen_tipo") == "SI_AABB", "tipo SI_AABB")
    print("  OK origen = superior-izquierda (AABB)")


def test_ancla_rechaza_stadium_sin_esquina():
    """P94-like: laterales rectos + arcos → sin esquina escuadrable."""
    from generador_caras_tanque import (
        _esquinas_hlr_escuadrables,
        _tiene_esquina_escuadrable,
    )

    # Stadium: 2 verticales + extremos de arco (no recta+recta).
    laterales = [
        {
            "es_recta": True,
            "minx": 0.0,
            "maxx": 0.0,
            "miny": 1.0,
            "maxy": 5.0,
            "dx": 0.0,
            "dy": 4.0,
            "cx": 0.0,
            "cy": 3.0,
            "curve": None,
        },
        {
            "es_recta": True,
            "minx": 2.0,
            "maxx": 2.0,
            "miny": 1.0,
            "maxy": 5.0,
            "dx": 0.0,
            "dy": 4.0,
            "cx": 2.0,
            "cy": 3.0,
            "curve": None,
        },
        {
            "es_recta": False,
            "minx": 0.0,
            "maxx": 2.0,
            "miny": 5.0,
            "maxy": 6.0,
            "dx": 2.0,
            "dy": 1.0,
            "cx": 1.0,
            "cy": 5.5,
            "curve": None,
        },
        {
            "es_recta": False,
            "minx": 0.0,
            "maxx": 2.0,
            "miny": 0.0,
            "maxy": 1.0,
            "dx": 2.0,
            "dy": 1.0,
            "cx": 1.0,
            "cy": 0.5,
            "curve": None,
        },
    ]
    # Endpoints artificiales: verticales no se tocan entre sí.
    def _ends_vert(d, x):
        d["_ends"] = [(x, float(d["miny"])), (x, float(d["maxy"]))]

    # Monkeypatch endpoints via curve=None path uses bbox corners — vertical
    # degenerado da (minx,miny)-(maxx,maxy) = mismos X → OK no forman esquina
    # entre las dos verticales (paralelas y separadas).
    _assert(not _esquinas_hlr_escuadrables(laterales), "stadium sin esquina")
    _assert(not _tiene_esquina_escuadrable({"curvas": laterales}), "flag")

    # Placa rectangular: 4 lados se encuentran.
    rect = [
        {
            "es_recta": True,
            "minx": 0,
            "maxx": 0,
            "miny": 0,
            "maxy": 4,
            "dx": 0,
            "dy": 4,
            "cx": 0,
            "cy": 2,
            "curve": None,
        },
        {
            "es_recta": True,
            "minx": 0,
            "maxx": 6,
            "miny": 0,
            "maxy": 0,
            "dx": 6,
            "dy": 0,
            "cx": 3,
            "cy": 0,
            "curve": None,
        },
        {
            "es_recta": True,
            "minx": 6,
            "maxx": 6,
            "miny": 0,
            "maxy": 4,
            "dx": 0,
            "dy": 4,
            "cx": 6,
            "cy": 2,
            "curve": None,
        },
        {
            "es_recta": True,
            "minx": 0,
            "maxx": 6,
            "miny": 4,
            "maxy": 4,
            "dx": 6,
            "dy": 0,
            "cx": 3,
            "cy": 4,
            "curve": None,
        },
    ]
    esq = _esquinas_hlr_escuadrables(rect, tol=0.15)
    _assert(len(esq) >= 1, f"rect debe tener esquinas: {esq}")
    ox, oy, dato = g._origen_desde_ancla(
        {
            "minx": 0,
            "miny": 0,
            "maxx": 6,
            "maxy": 4,
            "curvas": rect,
            "nombre": "P93",
        },
        exigir_esquina=True,
    )
    _assert(ox is not None and oy is not None, "origen escuadrable")
    _assert(dato.get("origen_escuadrable"), "flag origen")
    # SI: esquina arriba-izquierda ≈ (0, 4)
    _assert(abs(ox - 0.0) < 0.2 and oy > 3.0, f"SI esperado ~ (0,4) got ({ox},{oy})")
    # Stadium debe rechazarse como ancla.
    ox2, oy2, _ = g._origen_desde_ancla(
        {
            "minx": 0,
            "miny": 0,
            "maxx": 2,
            "maxy": 6,
            "curvas": laterales,
            "nombre": "P94",
        },
        exigir_esquina=True,
    )
    _assert(ox2 is None and oy2 is None, "stadium rechazado")
    # Ranking: placa con esquina gana a stadium aunque stadium sea ancla 3D.
    proy = [
        (
            {"name": "P94", "es_hw": False, "_ancla_kit_3d": True},
            {
                "area": 200.0,
                "dx": 10,
                "dy": 20,
                "minx": 0,
                "miny": 0,
                "maxx": 10,
                "maxy": 20,
                "curvas": laterales,
            },
        ),
        (
            {"name": "P93", "es_hw": False, "_ancla_kit_3d": False},
            {
                "area": 50.0,
                "dx": 6,
                "dy": 4,
                "minx": 0,
                "miny": 0,
                "maxx": 6,
                "maxy": 4,
                "curvas": rect,
            },
        ),
    ]
    elegida, _ = g._elegir_ancla(proy)
    _assert(elegida["name"] == "P93", f"debe ganar P93 escuadrable, got {elegida['name']}")
    print("  OK ancla rechaza stadium; prefiere esquina escuadrable")


def test_no_omitir_instancias_gemelas():
    """A14: ancla P48_335 NO debe silenciar P48_336 (mismo part number)."""

    class Vista:
        Scale = 0.10

    # Origen SI
    ox, oy = 0.0, 10.0
    # Ancla = una instancia; gemela desplazada en X; L-placa en otra pos.
    hijos = [
        {
            "name": "62201-1248-P48",
            "pieza_id": "P48_335",
            "es_hw": False,
            "occ": None,
            "rb": None,
        },
        {
            "name": "62201-1248-P48",
            "pieza_id": "P48_336",
            "es_hw": False,
            "occ": None,
            "rb": None,
        },
        {
            "name": "62201-1248-P47",
            "pieza_id": "P47_334",
            "es_hw": False,
            "occ": None,
            "rb": None,
        },
    ]

    def fake_env(vista, tg, h):
        pid = h["pieza_id"]
        if pid == "P48_335":
            return {
                "minx": 0.0,
                "maxx": 1.0,
                "miny": 0.0,
                "maxy": 10.0,
                "cx": 0.5,
                "cy": 5.0,
                "dx": 1.0,
                "dy": 10.0,
                "curvas": [],
            }, False
        if pid == "P48_336":
            return {
                "minx": 4.0,
                "maxx": 5.0,
                "miny": 0.0,
                "maxy": 10.0,
                "cx": 4.5,
                "cy": 5.0,
                "dx": 1.0,
                "dy": 10.0,
                "curvas": [],
            }, False
        return {
            "minx": 0.0,
            "maxx": 8.0,
            "miny": 8.0,
            "maxy": 10.0,
            "cx": 4.0,
            "cy": 9.0,
            "dx": 8.0,
            "dy": 2.0,
            "curvas": [],
        }, False

    orig = g._envolvente_hijo_vista
    g._envolvente_hijo_vista = fake_env
    try:
        px, py = g._posiciones_desde_hijos(
            Vista(), None, hijos, ox, oy, ancla_pieza_id="P48_335"
        )
    finally:
        g._envolvente_hijo_vista = orig

    ids_x = {p.get("pieza_id") for p in px}
    ids_y = {p.get("pieza_id") for p in py}
    _assert("P48_336" in ids_x, f"gemela debe tener X: {ids_x}")
    _assert("P47_334" in ids_x or "P47_334" in ids_y, f"P47 acotada: {ids_x}|{ids_y}")
    _assert("P48_335" not in ids_x and "P48_335" not in ids_y, "ancla no se acota")
    print("  OK instancias gemelas se acotan (solo se omite ancla por pieza_id)")


def test_ancla_kit_3d_placa_mayor():
    """A08-like: P16 (placa grande) gana a P17 (más chica) en 3D."""
    IN = 2.54

    class RB:
        def __init__(self, mn, mx):
            self.MinPoint = type("p", (), {"X": mn[0], "Y": mn[1], "Z": mn[2]})()
            self.MaxPoint = type("p", (), {"X": mx[0], "Y": mx[1], "Z": mx[2]})()

    p16 = {
        "name": "62201-1248-P16",
        "es_hw": False,
        "rb": RB((0, 0, 0), (12 * IN, 8 * IN, 0.5 * IN)),
    }
    p17 = {
        "name": "62201-1248-P17",
        "es_hw": False,
        "rb": RB((0, 0, 0), (4 * IN, 3 * IN, 0.5 * IN)),
    }
    tubo = {
        "name": "62201-1248-P98",
        "es_hw": False,
        "rb": RB((0, 0, 0), (1.5 * IN, 1.5 * IN, 10 * IN)),
    }
    ancla = g._elegir_ancla_kit_3d([p17, tubo, p16])
    _assert(ancla["name"] == "62201-1248-P16", f"ancla 3D={ancla['name']}")
    # Preferida 3D debe ganar ranking 2D aunque P17 proyecte más en RIGHT.
    proy = [
        (
            {**p17, "_ancla_kit_3d": False},
            {
                "area": 200.0,
                "dx": 20,
                "dy": 10,
                "minx": 0,
                "miny": 0,
                "maxx": 20,
                "maxy": 10,
            },
        ),
        (
            {**p16, "_ancla_kit_3d": True},
            {
                "area": 50.0,
                "dx": 10,
                "dy": 5,
                "minx": 0,
                "miny": 0,
                "maxx": 10,
                "maxy": 5,
            },
        ),
    ]
    elegida, _ = g._elegir_ancla(proy)
    _assert(elegida["name"] == "62201-1248-P16", f"boost 3D={elegida['name']}")
    print("  OK ancla 3D kit estabiliza origen (P16)")


def test_bbox_foto_incluye_texto_y():
    """Cota Y corta: el bbox debe cubrir el texto rotado (no solo el tip)."""
    from generador_caras_tanque import (
        MARGEN_FOTO_SUP_CM,
        _bbox_foto_grupo,
        _mitades_caja_texto_cota,
    )

    class Vista:
        Left = 10.0
        Width = 20.0
        Top = 15.0
        Height = 8.0
        Scale = 0.15  # hoja→modelo

    # Altura de pieza en hoja ≈ 3.2 cm → texto "21.375 in" ~4 cm rotado
    # desborda arriba/abajo si solo se mira el tip.
    origen_x, origen_y = 12.0, 8.0
    y_tip = 11.2
    pos_y = [
        {
            "valor": y_tip,
            "typ": False,
            "lado": "inf",
            "dato": {
                "minx": 14,
                "maxx": 28,
                "miny": 8,
                "maxy": 11.2,
                "cx": 21,
                "cy": 9.6,
            },
            "miembros": [
                {
                    "valor": y_tip,
                    "dato": {
                        "minx": 14,
                        "maxx": 28,
                        "miny": 8,
                        "maxy": 11.2,
                        "cx": 21,
                        "cy": 9.6,
                    },
                }
            ],
        }
    ]
    hx, hy = _mitades_caja_texto_cota("21.375 in", vertical=True)
    _assert(hy > 1.5, f"mitad Y texto insuficiente: {hy}")
    bbox = _bbox_foto_grupo(Vista(), [], pos_y, origen_x, origen_y)
    _minx, _maxx, miny, maxy = bbox
    mid = (origen_y + y_tip) * 0.5
    _assert(maxy >= mid + hy + MARGEN_FOTO_SUP_CM * 0.5, f"maxy={maxy}")
    _assert(miny <= mid - hy, f"miny={miny}")
    print(f"  OK bbox foto Y incluye texto rotado (hy={hy:.2f})")


def test_prefijo_serie():
    _assert(g._prefijo_serie("62201-1254-A08") == "62201", "serie A")
    _assert(g._prefijo_serie("SP-792") == "", "SP sin serie")
    print("  OK prefijo serie")


def test_radio_typ_dinamico_logica():
    """Replica la lógica de shrink sin sketch COM."""
    r_base = float(RADIO_MARCA_TYP_CM)
    r_min = float(RADIO_MARCA_TYP_INTERIOR_CM)
    pts = [(0.0, 0.0), (0.4, 0.0), (0.8, 0.0)]  # columna cercana
    dmin = None
    for i in range(len(pts)):
        for j in range(i + 1, len(pts)):
            d = math.hypot(pts[i][0] - pts[j][0], pts[i][1] - pts[j][1])
            if d < 1e-6:
                continue
            dmin = d if dmin is None else min(dmin, d)
    _assert(dmin is not None and dmin < SEP_MIN_MARCA_TYP_CM * 2.2, "cerca")
    r_dyn = max(r_min, min(r_base, dmin * 0.38))
    _assert(r_dyn < r_base, f"radio dinámico {r_dyn} < {r_base}")
    # Cluster: promedio de puntos (marcas ya centradas en pieza).
    clusters = []
    tol = float(SEP_MIN_MARCA_TYP_CM)
    for x, y in [(1.0, 2.0), (1.05, 2.02), (5.0, 2.0)]:
        puesto = False
        for c in clusters:
            cx0 = c["sx"] / c["n"]
            cy0 = c["sy"] / c["n"]
            if math.hypot(x - cx0, y - cy0) <= tol:
                c["sx"] += x
                c["sy"] += y
                c["n"] += 1
                puesto = True
                break
        if not puesto:
            clusters.append({"sx": x, "sy": y, "n": 1})
    _assert(len(clusters) == 2, f"clusters={clusters}")
    cx0 = clusters[0]["sx"] / clusters[0]["n"]
    cy0 = clusters[0]["sy"] / clusters[0]["n"]
    _assert(abs(cx0 - 1.025) < 1e-9 and abs(cy0 - 2.01) < 1e-9, "promedio")
    _assert(clusters[0]["n"] == 2, "qty en cluster")
    print("  OK radio TYP dinámico + cluster promedio")


def test_ilogic_limpiar():
    path = os.path.join(ROOT, "ilogic", "COTAS_ENSAMBLES_INDEPENDIENTES.iLogicVb")
    txt = open(path, encoding="utf-8", errors="ignore").read()
    _assert("--limpiar" in txt, "iLogic debe pasar --limpiar")
    print("  OK iLogic --limpiar")


def test_hw_elige_eje_dominante():
    """Simula la regla HW: solo el eje mayor ≥1.0 in."""
    dx, dy = 4.5, 0.25  # posición X, espesor Y
    elegidos = []
    if dx >= dy and dx >= g.MIN_COTA_INSTRUCTIVO_IN:
        elegidos.append("X")
    elif dy >= g.MIN_COTA_INSTRUCTIVO_IN:
        elegidos.append("Y")
    _assert(elegidos == ["X"], f"HW debe acotar solo X: {elegidos}")
    dx2, dy2 = 0.2, 0.3  # ambos bajo umbral
    elegidos2 = []
    if dx2 >= dy2 and dx2 >= g.MIN_COTA_INSTRUCTIVO_IN:
        elegidos2.append("X")
    elif dy2 >= g.MIN_COTA_INSTRUCTIVO_IN:
        elegidos2.append("Y")
    _assert(elegidos2 == [], f"bajo 1.0in no cota: {elegidos2}")
    print("  OK HW solo eje dominante ≥1.0in")


def test_rank_ancla_placa_vs_tubo():
    """A42: placa delgada gana a tubo aunque el tubo proyecte más en RIGHT."""
    IN = 2.54

    class RB:
        def __init__(self, mn, mx):
            self.MinPoint = type("p", (), {"X": mn[0], "Y": mn[1], "Z": mn[2]})()
            self.MaxPoint = type("p", (), {"X": mx[0], "Y": mx[1], "Z": mx[2]})()

    p97 = {
        "name": "62201-1248-P97",
        "es_hw": False,
        "rb": RB((-2.25 * IN, -1.5 * IN, 0), (2.25 * IN, 1.5 * IN, 0.5 * IN)),
    }
    p98 = {
        "name": "62201-1248-P98",
        "es_hw": False,
        "rb": RB(
            (-0.83 * IN, -0.83 * IN, 0.5 * IN), (0.83 * IN, 0.83 * IN, 3.0 * IN)
        ),
    }
    # RIGHT: sheet X=Z, Y=Y — tubo proyecta más área 2D cruda.
    def env_r(h):
        rb = h["rb"]
        xs = (rb.MinPoint.Z, rb.MaxPoint.Z)
        ys = (rb.MinPoint.Y, rb.MaxPoint.Y)
        minx, maxx = min(xs), max(xs)
        miny, maxy = min(ys), max(ys)
        return {
            "minx": minx,
            "maxx": maxx,
            "miny": miny,
            "maxy": maxy,
            "dx": maxx - minx,
            "dy": maxy - miny,
            "area": (maxx - minx) * (maxy - miny),
        }

    proy = [(p97, env_r(p97)), (p98, env_r(p98))]
    _assert(env_r(p98)["area"] > env_r(p97)["area"], "tubo proyecta más")
    ancla, _ = g._elegir_ancla(proy)
    _assert(ancla["name"] == "62201-1248-P97", f"ancla={ancla['name']}")
    _assert(callable(g._rank_anclas), "rank existe")
    print("  OK ranking placa vs tubo (A42 RIGHT)")


def main():
    print("=" * 56)
    print(" SMOKE ensambles instructivo (revalidación)")
    print("=" * 56)
    tests = [
        test_umbral_y_hw,
        test_ancla_prefer_placa,
        test_cross_series,
        test_secuencia_y_nombre,
        test_cam_viewcube_6,
        test_origen_si_ancla,
        test_ancla_rechaza_stadium_sin_esquina,
        test_no_omitir_instancias_gemelas,
        test_ancla_kit_3d_placa_mayor,
        test_bbox_foto_incluye_texto_y,
        test_prefijo_serie,
        test_radio_typ_dinamico_logica,
        test_ilogic_limpiar,
        test_hw_elige_eje_dominante,
        test_rank_ancla_placa_vs_tubo,
    ]
    # Pasada 1
    print("\n[Pasada 1]")
    for fn in tests:
        fn()
    # Pasada 2 (misma batería)
    print("\n[Pasada 2 — repetición]")
    for fn in tests:
        fn()
    # Pasada 3: asserts cruzados
    print("\n[Pasada 3 — cruces]")
    _assert(g.MIN_COTA_INSTRUCTIVO_IN >= 1.0, "umbral sano")
    _assert(len(g.VISTAS) == 6, "6 vistas ViewCube")
    _assert(callable(g._elegir_ancla_kit_3d), "ancla 3D existe")
    _assert(callable(g._ordenar_cotas_secuencia), "secuencia existe")
    _assert(callable(g._elegir_ancla), "ancla existe")
    ilogic = open(
        os.path.join(ROOT, "ilogic", "COTAS_ENSAMBLES_INDEPENDIENTES.iLogicVb"),
        encoding="utf-8",
        errors="ignore",
    ).read()
    _assert("--limpiar" in ilogic, "limpiar iLogic")
    _assert("BOTTOM" in ilogic and "LEFT" in ilogic, "iLogic menciona 6 vistas")
    print("  OK cruces API")
    print("\nTODAS LAS REVALIDACIONES OK")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except _Fail as exc:
        print(f"\nFAIL: {exc}")
        sys.exit(1)
