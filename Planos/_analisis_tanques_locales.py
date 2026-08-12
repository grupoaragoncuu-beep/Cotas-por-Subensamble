# -*- coding: utf-8 -*-
"""
Analisis rapido de los STEP de prueba en `..\\Tanques\\` sin abrir Inventor.

Uso: `python _analisis_tanques_locales.py`

Genera `_analisis_tanques_locales_resultado.txt` con:
  - Cantidad de PRODUCT por STEP.
  - Deteccion de subensambles candidatos a tapa (nombre TAPA/TOP/COVER/ROOF...).
  - Deteccion de segmentos con la convencion "SEGMENTO" / "SEGMENT".
  - Familia estimada (VANTRAN, SUNBELT, OTC, ...).
  - Muestreo de accesorios y casco.

Sirve como pre-diagnostico antes de correr el flujo real del generador para
saber cuales tanques deberian producir cara TOP y cuales caeran a fallback
geometrico.
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path


AQUI = Path(__file__).resolve().parent
CARPETA_TANQUES = AQUI.parent / "Tanques"
SALIDA = AQUI / "_analisis_tanques_locales_resultado.txt"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PRODUCT_RE = re.compile(
    r"PRODUCT\s*\(\s*'((?:[^']|'')*)'\s*,\s*'((?:[^']|'')*)'",
    re.IGNORECASE,
)

PALABRAS_TAPA = ("TAPA", "TOP", "COVER", "ROOF", "CUBIERTA", "TOP_COVER", "TOP COVER")
PALABRAS_SEGMENTO = ("SEGMENTO", "SEGMENT", "SEG-", "SEG_")
PALABRAS_ACCESORIO = (
    "LUG", "LIFT", "PAD", "FLANGE", "BRACKET", "NOZZLE", "PORT", "BOSS",
    "CLIP", "CLAMP", "SUPPORT", "MANWAY", "JACK", "OREJA", "BOCA", "GROUND",
    "TIERRA", "GAUGE", "NIPPLE", "PARKING", "VALVE", "BUSHING", "RADIATOR",
    "BREATHER", "THERMO", "PRESSURE", "HANDHOLE", "PIPE", "COUPLING",
    "ELBOW", "STUB", "PATCH", "RADIAL", "BUSH", "ARRESTER", "SURGE", "DRAIN",
    "FILL", "FILTER", "SAMPLER",
)
PALABRAS_CASCO = (
    "SHELL", "PLACA", "PLATE", "WALL", "CASCO", "BODY", "TANK WALL",
    "PANEL", "SIDE WALL", "END WALL", "BOTTOM", "FLOOR", "SOLERA", "BASE DE",
    "HEADIRON", "COMPARTMENT",
)
FAMILIAS = (
    "VANTRAN", "SUNBELT", "OTC", "SWE", "GIGA", "PTT", "ERMCO", "HOWARD",
    "PROLEC", "IEM", "EATON", "HITACHI", "WEG", "ABB", "PACIFIC", "SUBSTATION",
    "PADMOUNT", "BOARD",
)


def familia_desde_nombre(nombre_archivo: str) -> str:
    u = nombre_archivo.upper()
    for f in FAMILIAS:
        if f in u:
            return f
    m = re.match(r"^(\d{4,5})", nombre_archivo.strip())
    if m:
        return f"CODIGO_{m.group(1)}"
    return "OTRO"


def clasifica(nombre: str) -> str:
    u = nombre.upper()
    if any(k in u for k in PALABRAS_TAPA):
        return "TAPA"
    if any(k in u for k in PALABRAS_SEGMENTO):
        return "SEGMENTO"
    if any(k in u for k in PALABRAS_ACCESORIO):
        return "ACCESORIO"
    if any(k in u for k in PALABRAS_CASCO):
        return "CASCO"
    return "OTRO"


def extraer_productos(ruta: Path, limite: int = 20_000_000) -> tuple[list[str], str | None]:
    try:
        with ruta.open("r", errors="ignore") as fh:
            data = fh.read(min(ruta.stat().st_size, limite))
    except OSError as exc:
        return [], str(exc)
    nombres: list[str] = []
    vistos: set[str] = set()
    for m in PRODUCT_RE.finditer(data):
        for cand in (m.group(1), m.group(2)):
            cand = cand.replace("''", "'").strip()
            if not cand or cand.upper() in vistos:
                continue
            vistos.add(cand.upper())
            nombres.append(cand)
    return nombres, None


def analizar_tanque(ruta: Path) -> dict:
    familia = familia_desde_nombre(ruta.name)
    productos, err = extraer_productos(ruta)
    if err:
        return {"archivo": ruta.name, "error": err}
    buckets = Counter(clasifica(p) for p in productos)
    tapas = [p for p in productos if clasifica(p) == "TAPA"]
    segmentos = [p for p in productos if clasifica(p) == "SEGMENTO"]
    accesorios = [p for p in productos if clasifica(p) == "ACCESORIO"]
    casco = [p for p in productos if clasifica(p) == "CASCO"]
    return {
        "archivo": ruta.name,
        "familia": familia,
        "MB": round(ruta.stat().st_size / 1e6, 2),
        "productos": len(productos),
        "buckets": dict(buckets),
        "tapa_candidatos": tapas[:15],
        "segmentos": segmentos[:15],
        "accesorios_muestra": accesorios[:12],
        "casco_muestra": casco[:8],
        "preview_todos": productos[:20],
    }


def main() -> int:
    if not CARPETA_TANQUES.is_dir():
        print(f"ERROR: no existe {CARPETA_TANQUES}")
        return 1
    steps = [
        p
        for p in sorted(CARPETA_TANQUES.iterdir())
        if p.is_file() and p.suffix.lower() in (".stp", ".step")
    ]
    if not steps:
        print(f"AVISO: no hay STEP en {CARPETA_TANQUES}")
        return 1

    lineas: list[str] = []

    def w(s: str = "") -> None:
        lineas.append(s)
        print(s)

    w(f"Carpeta: {CARPETA_TANQUES}")
    w(f"STEPs encontrados: {len(steps)}")
    w("=" * 78)

    todos_familias: Counter[str] = Counter()
    resumen: list[tuple[str, str, bool, int]] = []
    for step in steps:
        w(f"\n### {step.name}")
        info = analizar_tanque(step)
        if "error" in info:
            w(f"  ERROR: {info['error']}")
            continue
        todos_familias[info["familia"]] += 1
        tiene_tapa = bool(info["tapa_candidatos"])
        segs = len(info["segmentos"])
        w(f"  familia estimada : {info['familia']}")
        w(f"  peso             : {info['MB']} MB")
        w(f"  productos totales: {info['productos']}")
        w(f"  buckets          : {info['buckets']}")
        w(f"  candidatos TAPA  : {len(info['tapa_candidatos'])}")
        if info["tapa_candidatos"]:
            for t in info["tapa_candidatos"]:
                w(f"    - {t}")
        w(f"  segmentos nombre : {segs}")
        if info["segmentos"]:
            for s in info["segmentos"][:6]:
                w(f"    - {s}")
        if info["accesorios_muestra"]:
            w(f"  accesorios (top) : {info['accesorios_muestra']}")
        if info["casco_muestra"]:
            w(f"  casco (top)      : {info['casco_muestra']}")
        w(f"  preview          : {info['preview_todos'][:10]}")
        resumen.append((step.name, info["familia"], tiene_tapa, segs))

    w("")
    w("=" * 78)
    w("RESUMEN")
    w("=" * 78)
    w(f"Familias detectadas: {dict(todos_familias)}")
    w("")
    w(f"{'archivo':<45} {'familia':<12} {'tapa?':<6} {'segs'}")
    w("-" * 78)
    for archivo, fam, tapa, segs in resumen:
        w(f"{archivo[:44]:<45} {fam[:11]:<12} {'SI' if tapa else 'no':<6} {segs}")

    w("")
    w("Interpretacion:")
    w("  - 'tapa? SI' = el STEP conserva un nombre reconocible como tapa;")
    w("    en Inventor lo detectara `_es_nombre_tapa`.")
    w("  - 'tapa? no' = tocara fallback geometrico +cover para TOP.")
    w("    Si el flujo omite TOP con 'no se detecto subensamble de tapa',")
    w("    revisar aqui si el STEP realmente contiene una tapa como IAM raiz.")
    w("  - 'segs > 0' = tanque estilo Vantran con segmentos nombrados;")
    w("    para OTC lo esperado es segs=0 y el flujo cae a deteccion geometrica.")

    SALIDA.write_text("\n".join(lineas), encoding="utf-8")
    print(f"\nGuardado: {SALIDA}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
