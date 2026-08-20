# -*- coding: utf-8 -*-
"""Auditoria ESTATICA de los .iam/.ipt en la carpeta Tanques del repo.

Objetivo: identificar patrones de nombres, piezas duplicadas entre tanques y
mapear cada pieza al resolver de THK.py que deberia atenderla. NO abre Inventor.
Sirve para anticipar donde puede fallar la cotacion antes de la corrida real.

Salida: AUDITORIA_TANQUES.md (misma carpeta que este script).
"""
from __future__ import annotations

import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

_PLANOS = Path(__file__).resolve().parent
_REPO = _PLANOS.parent
RAIZ_TANQUES = _REPO / "Tanques"
SALIDA_MD = _PLANOS / "AUDITORIA_TANQUES.md"
SALIDA_CSV = _PLANOS / "AUDITORIA_TANQUES.csv"

# ---------------------------------------------------------------------------
# Clasificacion por keyword
# ---------------------------------------------------------------------------
# Cada regla: (nombre_categoria, resolver_esperado_THK, patron_regex)
# resolver_esperado: uno de
#   - "circular_solid"   -> barra/varilla/pin/stud
#   - "circular_hollow"  -> tubo redondo hueco
#   - "rect_hollow"      -> HSS / tubo rectangular
#   - "prismatic_plate"  -> placa plana (chapa lisa)
#   - "prismatic_L"      -> angulo (L)
#   - "prismatic_U"      -> canal (U/C)
#   - "prismatic_semi"   -> chapa doblada semicircular
#   - "accesorio"        -> pieza de accesorio (lug/flange/bracket) -> variable
#   - "desconocido"      -> nombre no da pistas
REGLAS = [
    # -----------------------------------------------------------------
    # HARDWARE / FASTENERS / STANDARD PARTS -> no requieren THK (catalogo)
    # Aragon usa prefijos: HW- (Hardware), SF- (Standard Fastener),
    # GUNSTUD (soldadura), FT-PRD- (Fitting/Relief Device), FPP- (Ferrous Pipe Part),
    # SF-HC/CP/NP/etc.  Estas piezas son tornilleria/aditamentos estandar
    # que no se cotan porque vienen de catalogo.
    # -----------------------------------------------------------------
    ("hardware_estandar",     "no_aplica",       re.compile(r"^(HW-|SF-|SFHC|GUNSTUD|GUN\s*STUD|FT-PRD|FPP-|SMSP-|SMPS-|SMH-|MCM-|MSS-)", re.I)),
    ("stud_hardware",         "no_aplica",       re.compile(r"(?<![A-Z])STUD(?![A-Z])(?!.*(WELD|SOLD))|BOLT|SCREW|(?<![A-Z])NUT(?![A-Z])|WASHER|TUERCA|TORNILLO|ARANDELA|RONDANA|FASTENER", re.I)),

    # -----------------------------------------------------------------
    # SP-XXX - Standard Part de Aragon (nombre generico pero muy frecuente).
    # Podrian ser placas, tapas o hardware. En OTC casi todos son piezas
    # planas de acero (plate/bracket). Marcamos como probable prismatic_plate
    # con riesgo medio (mejor que "alto").
    # -----------------------------------------------------------------
    ("standard_part_SP",      "prismatic_plate", re.compile(r"^SP-\d", re.I)),

    # circulares
    ("varilla_redonda",  "circular_solid",  re.compile(r"(?<![A-Z])(VARILLA|TIERRA\s+REDONDA|GROUND\s+ROD|SHAFT|EJE|ROD)(?![A-Z])", re.I)),
    ("tubo_redondo",     "circular_hollow", re.compile(r"(?<![A-Z])(TUBO|PIPE|TUBE|NIPPLE)(?![A-Z])", re.I)),
    ("brida_pipe",       "circular_hollow", re.compile(r"(?<![A-Z])(PIPE\s*FLANE|PIPE\s*FLANGE|FLANGE|BRIDA|WELDNECK|RADVLV)(?![A-Z])", re.I)),
    # HSS
    ("tubo_rectangular", "rect_hollow",     re.compile(r"(?<![A-Z])HSS(?![A-Z])|(?<![A-Z])TUBING(?![A-Z])|RECT\s*TUB", re.I)),
    # perfiles laminados
    ("angulo_AISC_L",    "prismatic_L",     re.compile(r"(?<![A-Z])AISC(?![A-Z])|(?<![A-Z])ANGULO(?![A-Z])|(?<![A-Z])ANGLE(?![A-Z])|(?<![A-Z])L\d", re.I)),
    ("canal_U_C",        "prismatic_U",     re.compile(r"(?<![A-Z])CANAL(?![A-Z])|(?<![A-Z])CHANNEL(?![A-Z])|(?<![A-Z])U\d|(?<![A-Z])C\d\s*x\s*\d", re.I)),
    # placa / chapa
    ("placa_segmento",   "prismatic_plate", re.compile(r"(?<![A-Z])PLACA(?![A-Z])|(?<![A-Z])PLATE(?![A-Z])|(?<![A-Z])PANEL(?![A-Z])|(?<![A-Z])SOLERA(?![A-Z])|(?<![A-Z])SEG(MENT|MENTO)?(?![A-Z])|(?<![A-Z])SEG\s*\d", re.I)),
    ("chapa_preformada", "prismatic_semi",  re.compile(r"(?<![A-Z])PREFORMAD|(?<![A-Z])DOBLAD|(?<![A-Z])BEND|(?<![A-Z])ROLLED|(?<![A-Z])CURVED|SEMICIRCULAR|(?<![A-Z])TAB(?![A-Z])", re.I)),
    ("cover",            "prismatic_plate", re.compile(r"TOP\s*_?COVER|(?<![A-Z])COVER(?![A-Z])|(?<![A-Z])TAPA(?![A-Z])|(?<![A-Z])BASE(?![A-Z])|(?<![A-Z])BOTTOM(?![A-Z])|(?<![A-Z])FLOOR(?![A-Z])|(?<![A-Z])CAJA(?![A-Z])", re.I)),
    # accesorios estructurales (bracket / lug / patch / boss / gusset)
    ("bracket_lug",      "accesorio",       re.compile(r"(?<![A-Z])BRACKET(?![A-Z])|(?<![A-Z])LUG(?![A-Z])|LIFTINGLUG|(?<![A-Z])CARTAB|(?<![A-Z])BRACE(?![A-Z])|(?<![A-Z])GUSSET(?![A-Z])|(?<![A-Z])CLIP(?![A-Z])|(?<![A-Z])REFUERZO(?![A-Z])|(?<![A-Z])MARCO(?![A-Z])|(?<![A-Z])CUADRO(?![A-Z])", re.I)),
    ("pad_boss_patch",   "accesorio",       re.compile(r"(?<![A-Z])PADS?(?![A-Z])|(?<![A-Z])BOSS(?![A-Z])|(?<![A-Z])PATCH(?![A-Z])|(?<![A-Z])PARKING(?![A-Z])|(?<![A-Z])GAUGE(?![A-Z])|(?<![A-Z])BUSHING(?![A-Z])|(?<![A-Z])MANWAY(?![A-Z])|(?<![A-Z])JACKING(?![A-Z])|(?<![A-Z])GROUND(?![A-Z])|(?<![A-Z])TIERRA(?![A-Z])|(?<![A-Z])INSPECTION(?![A-Z])|(?<![A-Z])HANDLE(?![A-Z])|(?<![A-Z])MANIJA(?![A-Z])|(?<![A-Z])BISAGRA(?![A-Z])|(?<![A-Z])HINGE(?![A-Z])|(?<![A-Z])SWITCH(?![A-Z])|(?<![A-Z])CIERRE(?![A-Z])|(?<![A-Z])RADIATOR(?![A-Z])|(?<![A-Z])NOZZLE(?![A-Z])|(?<![A-Z])FITTING(?![A-Z])|(?<![A-Z])ORING(?![A-Z])|O-RING|(?<![A-Z])JUNTA(?![A-Z])|BOSS", re.I)),
    # generico critico (Compound/Solid sin descripcion)
    ("nombre_generico",  "desconocido",     re.compile(r"^(COMPOUND|SOLID|PART|COMPONENT)\d*$", re.I)),
]

# Piezas que sabemos por experiencia OTC estan bajo control
OTC_PIEZAS_CONOCIDAS = {
    "62176-1247-P04": ("tubo_ring",        "circular_hollow"),  # anillo tubo
    "62176-1247-P10": ("canal_U_C",        "prismatic_U"),
    "62176-1248-P10": ("canal_U_C",        "prismatic_U"),
    "62176-1248-P27": ("angulo_L",         "prismatic_L"),
    "62176-1248-P29": ("chapa_semicirc",   "prismatic_semi"),
    "62176-1248-P31": ("tubo_HSS",         "rect_hollow"),
    "62176-1248-P46": ("angulo_L_flange",  "prismatic_L"),
    "62176-1248-P47": ("angulo_L",         "prismatic_L"),
    "62176-1248-P25": ("placa",            "prismatic_plate"),
}


def clasificar(nombre_ipt: str):
    """Devuelve (categoria, resolver_THK_esperado, riesgo)."""
    limpio = nombre_ipt.strip()
    up = limpio.upper()

    # 1) codigos OTC conocidos
    for prefijo, (cat, resolver) in OTC_PIEZAS_CONOCIDAS.items():
        if prefijo in up:
            return (cat, resolver, "conocida_OTC")

    # 2) codigo tipo XXXXX-XXXX-PNN (OTC generico, incluyendo sufijos como
    # _Default_As Machined_, _TAB, _ROD, _Predeterminado, _numeroCatalogo,
    # _DefaultSM-FLAT-PATTERN, etc.)
    if re.search(r"\d{4,5}-\d{4}-P\d{2,3}", up):
        # Sub-clasificar por sufijo conocido:
        if re.search(r"_ROD(?![A-Z])", up):
            return ("OTC_rod", "circular_solid", "medio")
        if re.search(r"_TAB(?![A-Z])", up):
            return ("OTC_tab_plate", "prismatic_plate", "medio")
        if "FLAT-PATTERN" in up or "SM-FLAT" in up:
            return ("OTC_flat_pattern", "prismatic_plate", "medio")
        return ("OTC_codigo_generico", "prismatic_plate", "medio")

    # 3) reglas por keyword
    for cat, resolver, patron in REGLAS:
        if patron.search(limpio):
            riesgo = "bajo" if resolver.startswith("prismatic_") or resolver.startswith("circular_") else "medio"
            if resolver == "desconocido":
                riesgo = "alto"
            if resolver == "accesorio":
                riesgo = "medio_alto"
            if resolver == "no_aplica":
                riesgo = "no_aplica"
            return (cat, resolver, riesgo)

    # 4) sin match -> desconocido
    return ("sin_categoria", "desconocido", "alto")


def cargar_inventario(csv_path: Path):
    filas = []
    with csv_path.open("r", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            filas.append(row)
    return filas


def resumen_por_tanque(filas):
    """Devuelve dict: tanque -> {tipo -> count, resolver -> count, piezas...}."""
    tanques = defaultdict(lambda: {
        "iam": 0,
        "ipt": 0,
        "categorias": Counter(),
        "resolvers": Counter(),
        "riesgos": Counter(),
        "piezas": [],  # (nombre, cat, resolver, riesgo)
    })
    for row in filas:
        t = row["Tanque"]
        tipo = row["Tipo"].lower()
        nombre = row["Nombre"]
        if tipo == ".iam":
            tanques[t]["iam"] += 1
            continue
        tanques[t]["ipt"] += 1
        cat, resolver, riesgo = clasificar(nombre)
        tanques[t]["categorias"][cat] += 1
        tanques[t]["resolvers"][resolver] += 1
        tanques[t]["riesgos"][riesgo] += 1
        tanques[t]["piezas"].append((nombre, cat, resolver, riesgo))
    return tanques


def detectar_duplicados(filas):
    """Piezas .ipt cuyo nombre aparece en >1 tanque -> reutilizables."""
    piezas_por_nombre = defaultdict(list)
    for row in filas:
        if row["Tipo"].lower() != ".ipt":
            continue
        clave = row["Nombre"].strip().upper()
        # normalizar copias (_MIR, _1, _2, - Copy) para agrupar
        clave_norm = re.sub(r"(_MIR\d*|_\d+|_COPY|\s*-\s*COPY|\s*_\d+)$", "", clave)
        piezas_por_nombre[clave_norm].append(row["Tanque"])
    duplicados = {
        nom: sorted(set(tanks))
        for nom, tanks in piezas_por_nombre.items()
        if len(set(tanks)) >= 2
    }
    return duplicados


def piezas_alto_riesgo(tanques):
    salida = []
    for t, info in tanques.items():
        for nombre, cat, resolver, riesgo in info["piezas"]:
            if riesgo in ("alto", "medio_alto"):
                salida.append((t, nombre, cat, resolver, riesgo))
    return sorted(salida, key=lambda x: (x[4], x[0], x[1]))


def emitir_md(tanques, duplicados, altos):
    lines = []
    A = lines.append
    A("# AUDITORIA ESTATICA DE TANQUES - Planos Abigail")
    A("")
    A(f"Fecha de auditoria: analisis offline sobre .iam/.ipt en `{RAIZ_TANQUES}`.")
    A("")
    A("Este reporte NO abre Inventor. Clasifica cada pieza por keyword de su nombre y la mapea al ")
    A("resolver de `THK.py` que deberia atenderla. Sirve para anticipar piezas problematicas ")
    A("antes de correr el flujo real.")
    A("")
    A("---")
    A("")
    A("## 1. Resumen por tanque")
    A("")
    A("| Tanque | .iam | .ipt | resolvers principales | riesgo alto | riesgo medio_alto |")
    A("|---|---:|---:|---|---:|---:|")
    for t in sorted(tanques):
        info = tanques[t]
        top_res = ", ".join(f"{k}({v})" for k, v in info["resolvers"].most_common(4))
        A(f"| {t} | {info['iam']} | {info['ipt']} | {top_res} | {info['riesgos'].get('alto', 0)} | {info['riesgos'].get('medio_alto', 0)} |")
    A("")

    # 2) Distribucion por resolver
    A("## 2. Cobertura teorica por resolver de THK.py")
    A("")
    global_res = Counter()
    for info in tanques.values():
        global_res.update(info["resolvers"])
    total = sum(global_res.values()) or 1
    A("| Resolver esperado | Piezas | % del total | Estado en THK.py |")
    A("|---|---:|---:|---|")
    estados = {
        "circular_solid":   "OK - `_resolver_circular_solid`",
        "circular_hollow":  "OK - `_resolver_circular_hollow`",
        "rect_hollow":      "OK - `_resolver_rectangular_hollow` (nuevo)",
        "prismatic_plate":  "OK - `_resolver_prismatico` (rama plate)",
        "prismatic_L":      "OK - `_resolver_prismatico` + `_es_perfil_u_o_l` + ALTO",
        "prismatic_U":      "OK - `_resolver_prismatico` + `_es_perfil_u_o_l` + ALTO",
        "prismatic_semi":   "OK - `_resolver_prismatico` + `_es_perfil_semicircular` + ALTO",
        "accesorio":        "PARCIAL - depende de geometria del accesorio",
        "no_aplica":        "NO SE COTA - hardware/fastener/pieza de catalogo",
        "desconocido":      "SIN COBERTURA garantizada - requiere inspeccion en Inventor",
    }
    for res, cnt in global_res.most_common():
        pct = 100.0 * cnt / total
        est = estados.get(res, "?")
        A(f"| `{res}` | {cnt} | {pct:.1f}% | {est} |")
    A("")

    # 3) Duplicados entre tanques
    A("## 3. Piezas compartidas entre tanques (reutilizables)")
    A("")
    if not duplicados:
        A("_No se detectaron piezas con nombre normalizado repetido entre tanques._")
    else:
        A(f"Se detectaron **{len(duplicados)}** nombres de pieza que aparecen en 2+ tanques. ")
        A("Si el resolver funciona en uno, funcionara en todos.")
        A("")
        A("| Pieza (nombre normalizado) | Tanques |")
        A("|---|---|")
        for nom, tanks in sorted(duplicados.items()):
            tanks_short = ", ".join(t.split(" ")[0][:14] for t in tanks)
            A(f"| `{nom}` | {tanks_short} |")
    A("")

    # 4) Piezas de alto riesgo (a inspeccionar)
    A("## 4. Piezas de riesgo ALTO / MEDIO_ALTO (foco de la proxima corrida)")
    A("")
    if not altos:
        A("_Ninguna pieza clasificada como alto riesgo. La proxima corrida deberia cotarlas todas._")
    else:
        A(f"Total de piezas marcadas: **{len(altos)}**.")
        A("")
        A("| Tanque | Pieza | Categoria | Resolver esperado | Riesgo |")
        A("|---|---|---|---|---|")
        for t, nom, cat, res, riesgo in altos:
            t_short = t.split(" ")[0][:18]
            A(f"| {t_short} | `{nom}` | {cat} | `{res}` | **{riesgo}** |")
    A("")

    # 5) Detalle por tanque (todas las piezas)
    A("## 5. Detalle completo por tanque")
    A("")
    for t in sorted(tanques):
        info = tanques[t]
        A(f"### {t}")
        A("")
        A(f"- Ensambles (.iam): **{info['iam']}**")
        A(f"- Piezas (.ipt): **{info['ipt']}**")
        A(f"- Riesgo alto: **{info['riesgos'].get('alto', 0)}** | ")
        A(f"medio_alto: **{info['riesgos'].get('medio_alto', 0)}** | ")
        A(f"medio: **{info['riesgos'].get('medio', 0)}** | ")
        A(f"bajo: **{info['riesgos'].get('bajo', 0)}** | ")
        A(f"conocida_OTC: **{info['riesgos'].get('conocida_OTC', 0)}**")
        A("")
        A("<details><summary>Ver todas las piezas</summary>")
        A("")
        A("| Pieza | Categoria | Resolver esperado | Riesgo |")
        A("|---|---|---|---|")
        for nombre, cat, resolver, riesgo in sorted(info["piezas"]):
            A(f"| `{nombre}` | {cat} | `{resolver}` | {riesgo} |")
        A("")
        A("</details>")
        A("")

    # 6) Conclusiones
    A("---")
    A("")
    A("## 6. Conclusiones y proxima corrida")
    A("")
    n_altos = sum(1 for t, *_ in altos)
    total_ipt = sum(info["ipt"] for info in tanques.values())
    pct_alto = 100.0 * n_altos / (total_ipt or 1)
    A(f"- Total piezas .ipt inventariadas: **{total_ipt}**")
    A(f"- Piezas de riesgo alto o medio_alto: **{n_altos}** ({pct_alto:.1f}%)")
    A(f"- La mayoria de piezas ({100 - pct_alto:.1f}%) deberia acotarse OK con los resolvers actuales.")
    A("")
    A("### Acciones sugeridas antes de correr")
    A("")
    A("1. Revisar la tabla de la seccion 4 y, para las piezas de riesgo alto:")
    A("   - Abrir el `.ipt` en Inventor 1x1 para saber su geometria real.")
    A("   - Anotar el espesor esperado.")
    A("   - Correr el flujo pieza por pieza usando la variable de entorno `PIEZAS_FILTRO`.")
    A("")
    A("2. Correr el flujo completo por tanque, uno por uno. Al final, revisar el archivo ")
    A("   `piezas_sin_cotas.txt` que emite `generador_vistas.py` y contrastar con las piezas ")
    A("   marcadas aqui como riesgo alto.")
    A("")
    A("3. Si hay piezas con nombre `COMPOUND` / `SOLID` genericos, prioridad maxima: son las que")
    A("   no dan pista textual del tipo geometrico.")
    A("")

    SALIDA_MD.write_text("\n".join(lines), encoding="utf-8")


def emitir_csv(tanques):
    with SALIDA_CSV.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["tanque", "pieza", "categoria", "resolver_esperado", "riesgo"])
        for t, info in sorted(tanques.items()):
            for nombre, cat, resolver, riesgo in sorted(info["piezas"]):
                w.writerow([t, nombre, cat, resolver, riesgo])


def main():
    csv_inventario = _PLANOS / "_inventario_tanques.csv"
    if not csv_inventario.exists():
        print(f"ERROR: no existe {csv_inventario}. Corre primero el Get-ChildItem que genera el CSV.")
        return 1

    filas = cargar_inventario(csv_inventario)
    tanques = resumen_por_tanque(filas)
    duplicados = detectar_duplicados(filas)
    altos = piezas_alto_riesgo(tanques)

    emitir_md(tanques, duplicados, altos)
    emitir_csv(tanques)

    total = sum(info["ipt"] for info in tanques.values())
    n_alto = sum(1 for _ in altos)
    print(f"Tanques auditados: {len(tanques)}")
    print(f"Piezas .ipt totales: {total}")
    print(f"Piezas riesgo alto/medio_alto: {n_alto}")
    print(f"Reporte: {SALIDA_MD}")
    print(f"Detalle CSV: {SALIDA_CSV}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
