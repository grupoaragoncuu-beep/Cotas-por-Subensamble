# -*- coding: utf-8 -*-
"""Analisis de STEPs de tanque en OPs reales (solo carpetas numeradas)."""
from __future__ import annotations

import os
import re
import sys
from collections import Counter
from pathlib import Path

OUT = Path(__file__).with_name("_analisis_steps_op_resultado.txt")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PRODUCT_RE = re.compile(
    r"PRODUCT\s*\(\s*'((?:[^']|'')*)'\s*,\s*'((?:[^']|'')*)'",
    re.IGNORECASE,
)

KEYWORDS_ACCES = (
    "LUG", "LIFT", "PAD", "FLANGE", "FLANE", "BRACKET", "NOZZLE", "FITTING",
    "PORT", "BOSS", "CLIP", "CLAMP", "SUPPORT", "MANWAY", "JACK", "OREJA",
    "BOCA", "ACCES", "GROUND", "TIERRA", "SWITCH", "GAUGE", "NIPPLE",
    "PARKING", "VALVE", "BUSHING", "RADIATOR", "CONSERV", "BREATHER",
    "THERMO", "PRESSURE", "INDICATOR", "HANGER", "HINGE", "DOOR",
    "HANDHOLE", "THROAT", "PIPE", "COUPLING", "ELBOW", "STUB", "PATCH",
    "HALF", "WELDNECK", "RADIAL", "BUSH", "CT ", "CURRENT", "HV ", "LV ",
    "ARRESTER", "SURGE", "DRAIN", "FILL", "FILTER", "SAMPLER", "PRD",
)
KEYWORDS_CASCO = (
    "SHELL", "PLACA", "PLATE", "WALL", "CASCO", "BODY", "TANK WALL",
    "SEGMENT", "SEGMENTO", "COVER", "TAPA", "SOLERA", "BASE DE",
    "PANEL", "SIDE WALL", "END WALL", "BOTTOM", "TOP COVER", "FLOOR",
    "HEADIRON", "COMPARTMENT",
)
KEYWORDS_SEGMENT = ("SEGMENT", "SEGMENTO", "SEG-", "SEG_")


def log(msg=""):
    print(msg, flush=True)


def op_root() -> Path:
    shared = None
    for child in Path("Z:/").iterdir():
        if child.is_dir() and "ARGA" in child.name.upper():
            shared = child
            break
    if shared is None:
        raise SystemExit("No ARGA en Z:")
    tik = None
    for d in (shared / "BIENVENIDO").iterdir():
        if d.is_dir() and "TIK" in d.name.upper():
            tik = d
            break
    if tik is None:
        raise SystemExit("No TIK")
    op = tik / "ORDENES DE PRODUCCION"
    if not op.is_dir():
        raise SystemExit("No OP")
    return op


def is_real_order(name: str) -> bool:
    # 2229- VANTRAN..., 1882 - SUNBELT..., GAE 783...
    if name.upper().startswith("ORDENES DE PRODUCCION"):
        return False
    if name.upper() in ("COMPONENTES ESTANDAR", "MACHOTE"):
        return False
    return bool(re.match(r"^\d{3,}", name.strip()) or name.upper().startswith("GAE"))


def brand(name: str) -> str:
    u = name.upper()
    for b in (
        "VANTRAN", "SUNBELT", "OTC", "SWE", "GIGA", "PTT", "ERMCO", "HOWARD",
        "PROLEC", "IEM", "EATON", "HITACHI", "WEG", "ABB", "ARGA", "GAE",
        "PACIFIC", "ONCOR", "SUBSTATION", "PADMOUNT",
    ):
        if b in u:
            return b
    parts = re.split(r"[\s\-_]+", name)
    for p in parts:
        if len(p) >= 3 and not p.isdigit() and not re.fullmatch(r"X\d+", p.upper()):
            return p.upper()[:24]
    return "OTRO"


def bucket(name: str) -> str:
    u = name.upper()
    if any(k in u for k in KEYWORDS_SEGMENT):
        return "SEGMENTO"
    if any(k in u for k in KEYWORDS_ACCES):
        return "ACCESORIO"
    if any(k in u for k in KEYWORDS_CASCO):
        return "CASCO"
    return "OTRO"


def extract_names(path: Path, limit=5_000_000):
    try:
        with path.open("r", errors="ignore") as fh:
            data = fh.read(min(path.stat().st_size, limit))
    except OSError as exc:
        return [], str(exc)
    names, seen = [], set()
    for m in PRODUCT_RE.finditer(data):
        for cand in (m.group(1), m.group(2)):
            cand = cand.replace("''", "'").strip()
            if not cand or cand.upper() in seen:
                continue
            seen.add(cand.upper())
            names.append(cand)
    return names, None


def find_steps(order: Path):
    """Busca STEPs de tanque sin rglob infinito."""
    solidos = None
    try:
        for child in order.iterdir():
            if child.is_dir() and (
                "SOLIDO" in child.name.upper()
                or re.match(r"^10[\.\s]", child.name.strip())
            ):
                solidos = child
                break
    except OSError:
        return []

    if solidos is None:
        return []

    found = []

    def add_stp(folder: Path, depth=0):
        if depth > 3:
            return
        try:
            for p in folder.iterdir():
                try:
                    if p.is_file() and p.suffix.lower() in (".stp", ".step") and p.stat().st_size > 100_000:
                        found.append(p)
                    elif p.is_dir() and p.name.upper() not in ("OLD", "OLDVERSIONS", "OLD VERSIONS"):
                        # priorizar TANQUE
                        if "TANQUE" in p.name.upper() or "TANK" in p.name.upper() or depth == 0:
                            add_stp(p, depth + 1)
                except OSError:
                    continue
        except OSError:
            pass

    # primero TANQUE directo
    for sub in solidos.iterdir() if solidos.is_dir() else []:
        try:
            if sub.is_dir() and ("TANQUE" in sub.name.upper() or sub.name.upper() == "TANK"):
                add_stp(sub, 0)
        except OSError:
            continue

    if not found:
        add_stp(solidos, 0)

    # unicos por tamaño desc
    uniq = {}
    for p in found:
        uniq[str(p).lower()] = p
    return sorted(uniq.values(), key=lambda x: x.stat().st_size, reverse=True)


def main():
    lines = []
    def w(s=""):
        lines.append(s)
        log(s)

    op = op_root()
    w(f"OP root: ...\\{op.parent.name}\\{op.name}")

    orders = [d for d in op.iterdir() if d.is_dir() and is_real_order(d.name)]
    orders = sorted(orders, key=lambda p: p.name, reverse=True)
    max_scan = int(os.environ.get("MAX_OP_SCAN", "45"))
    orders = orders[:max_scan]
    w(f"Ordenes numericas a revisar: {len(orders)}")

    brand_c = Counter()
    bucket_c = Counter()
    kw_c = Counter()
    name_c = Counter()
    path_c = Counter()
    structural = Counter()
    samples = []
    analyzed = 0
    no_step = 0
    brands_with_step = Counter()

    for i, order in enumerate(orders, 1):
        b = brand(order.name)
        brand_c[b] += 1
        log(f"[{i}/{len(orders)}] {order.name[:75]}")

        # estructura Solidos
        try:
            kids = [c.name for c in order.iterdir() if c.is_dir()]
            has10 = any("SOLIDO" in k.upper() or k.strip().startswith("10") for k in kids)
            structural["tiene_10_solidos" if has10 else "sin_10_solidos"] += 1
        except OSError:
            structural["error_list"] += 1

        steps = find_steps(order)
        if not steps:
            no_step += 1
            log("  (sin STEP en 10. Solidos)")
            continue

        # analizar hasta 2 STEPs grandes por OP (diversidad)
        for step in steps[:2]:
            try:
                rel = str(step.relative_to(order))
            except Exception:
                rel = step.name
            path_c["/".join(Path(rel).parts[:3])] += 1
            names, err = extract_names(step)
            if err:
                log(f"  read err {step.name}: {err}")
                continue
            analyzed += 1
            brands_with_step[b] += 1
            mb = step.stat().st_size / 1e6
            log(f"  STEP {step.name[:55]} ({mb:.1f}MB) products={len(names)}")

            for n in names:
                bucket_c[bucket(n)] += 1
                name_c[n.upper()[:70]] += 1
                uu = n.upper()
                for k in KEYWORDS_ACCES + KEYWORDS_CASCO + KEYWORDS_SEGMENT:
                    if k in uu:
                        kw_c[k] += 1

            if len(samples) < 50:
                samples.append({
                    "order": order.name,
                    "brand": b,
                    "step": step.name,
                    "mb": round(mb, 2),
                    "rel": rel,
                    "n": len(names),
                    "buckets": dict(Counter(bucket(x) for x in names)),
                    "acces": [x for x in names if bucket(x) == "ACCESORIO"][:12],
                    "segs": [x for x in names if bucket(x) == "SEGMENTO"][:10],
                    "casco": [x for x in names if bucket(x) == "CASCO"][:8],
                    "preview": names[:20],
                })

    w()
    w(f"STEPs analizados={analyzed} | OPs sin STEP={no_step}")
    w(f"Estructura: {dict(structural)}")
    w()
    w("=== Familias en OPs escaneadas ===")
    for k, v in brand_c.most_common():
        w(f"  {k}: {v} OPs (con STEP leido: {brands_with_step.get(k,0)})")

    w()
    w("=== Rutas STEP (hasta 3 niveles) ===")
    for k, v in path_c.most_common(20):
        w(f"  {k}: {v}")

    w()
    w("=== Buckets PRODUCT ===")
    tot = sum(bucket_c.values()) or 1
    for k, v in bucket_c.most_common():
        w(f"  {k}: {v} ({100*v/tot:.1f}%)")

    w()
    w("=== Keywords frecuentes ===")
    for k, v in kw_c.most_common(50):
        w(f"  {k}: {v}")

    w()
    w("=== PRODUCT repetidos (>=2 tanques) ===")
    for k, v in name_c.most_common(80):
        if v < 2:
            break
        w(f"  [{v}x] {k}")

    w()
    w("=== Muestras por familia ===")
    seen = set()
    for s in samples:
        if s["brand"] in seen:
            continue
        seen.add(s["brand"])
        w(f"\n[{s['brand']}] OP: {s['order'][:80]}")
        w(f"  {s['step'][:60]} ({s['mb']}MB)")
        w(f"  ruta: {s['rel'][:110]}")
        w(f"  products={s['n']} buckets={s['buckets']}")
        if s["segs"]:
            w(f"  SEGMENTOS: {s['segs']}")
        if s["acces"]:
            w(f"  ACCESORIOS: {s['acces']}")
        if s["casco"]:
            w(f"  CASCO: {s['casco']}")
        w(f"  preview: {s['preview']}")

    w()
    w("=== IMPLICACIONES COTAS ABIGAIL (piso real) ===")
    w("1. Clientes reales: VANTRAN, SUNBELT, OTC, SWE, GIGA, PTT, GAE... (no solo Vantran).")
    w("2. Estructura OP estable: 10. Solidos\\TANQUE (a veces subcarpetas REV/ULTIMO).")
    w("3. No todos los OP tienen STEP listo; Planos/Inventor puede vivir aparte.")
    w("4. Nombres PRODUCT mezclan ingles, espanol, codigos VT/OTC/SWE — keywords solas NO bastan.")
    w("5. 'Assembly Segmento N' no es universal; muchos tanques no usan esa convencion.")
    w("6. Por eso acotar = catalogo del subensamble/segmento ABIERTO + pieza mas de frente a la camara.")
    w("7. Excluir casco: shell/plate/wall/tapa/cover/solera/headiron/compartment/segment body.")
    w("8. Accesorios tipicos transversales: flange, nipple, lug, pad, ground, gauge, manway, jack, parking, patch, pipe, bushing, valve.")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    w(f"\nGuardado: {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
