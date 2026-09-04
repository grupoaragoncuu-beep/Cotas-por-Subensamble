# -*- coding: utf-8 -*-
"""Inventario tipológico de PDFs en '4. Planos' de TODAS las OPs (Z:).

No se limita a OTC/Vantran: clasifica por marca, patrón de nombre y
palabras clave para ver qué tipos de planos maneja el piso real.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

OUT_TXT = Path(__file__).with_name("_analisis_planos_global_resultado.txt")
OUT_JSON = Path(__file__).with_name("_analisis_planos_global_resultado.json")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BRANDS = (
    "VANTRAN", "SUNBELT", "OTC", "SWE", "GIGA", "PTT", "ERMCO", "HOWARD",
    "PROLEC", "IEM", "EATON", "HITACHI", "WEG", "ABB", "ARGA", "GAE",
    "PACIFIC", "ONCOR", "SUBSTATION", "PADMOUNT", "COOPER", "SCHNEIDER",
    "GE ", "SIEMENS", "VIRGINIA", "CENTRAL", "SPX", "HPS", "CG POWER",
)

# Tipología de plano por palabras en nombre / carpeta / texto PDF.
TIPO_RULES = (
    ("GA_ENSAMBLE", (
        "GENERAL ASSEMBLY", "GEN ASSY", "GA ", " G.A", "ENSAMBLE GENERAL",
        "OVERALL", "LAYOUT TANK", "TANK ASSEMBLY", "ASSY TANK",
    )),
    ("TANQUE_COMPLETO", (
        "TANK", "TANQUE", "TNK", "VESSEL",
    )),
    ("TAPA_TOP", (
        "TOP COVER", "COVER", "TAPA", "LID", "ROOF", "HEAD COVER",
    )),
    ("CASCO_SHELL", (
        "SHELL", "CASCO", "WALL", "SEGMENT", "SEGMENTO", "BODY",
        "SIDE WALL", "END WALL",
    )),
    ("BASE_BOTTOM", (
        "BASE", "BOTTOM", "FLOOR", "SOLERA", "SKID", "UNDERFRAME",
    )),
    ("ACCESORIO_NOZZLE", (
        "NOZZLE", "FLANGE", "NIPPLE", "BUSHING", "BOSS", "PORT",
        "MANWAY", "HANDHOLE", "THROAT", "PIPE", "FITTING", "COUPLING",
    )),
    ("ACCESORIO_LUG_PAD", (
        "LUG", "LIFT", "PAD", "BRACKET", "SUPPORT", "JACK", "PARKING",
        "HINGE", "CLIP", "CLAMP", "GROUND", "TIERRA",
    )),
    ("INSTRUMENTOS", (
        "GAUGE", "THERMO", "PRESSURE", "INDICATOR", "SWITCH", "PRD",
        "BREATHER", "CONSERV", "RADIATOR", "FILTER", "SAMPLER",
    )),
    ("DETALLE", (
        "DETAIL", "DETALLE", "DET ", "SECTION", "SECCION", "VISTA",
        "VIEW ", "CUT ",
    )),
    ("SOLDADURA_WELD", (
        "WELD", "SOLDAD", "WPS", "PQR", "JOINT",
    )),
    ("MATERIALES_BOM", (
        "BOM", "BILL OF", "MATERIAL", "PART LIST", "LISTA DE",
        "CUT LIST", "NESTING",
    )),
    ("FABRICACION", (
        "FAB", "BLANK", "FLAT", "DESARROLLO", "BEND", "PUNCH",
        "LASER", "PLASMA", "CNC", "CORTE",
    )),
    ("ELECTRICO", (
        "WIRING", "ELECTR", "CT ", "HV ", "LV ", "BUSHING HV",
        "TERMINAL", "CONNECTION",
    )),
    ("CONTROL_REV", (
        "REV", "ECO", "ECN", "CHANGE", "MARKUP", "RED LINE",
    )),
    ("CERT_SPEC", (
        "SPEC", "SPECIFICATION", "DATA SHEET", "DATASHEET",
        "CERTIFICATE", "CERT ", "ITP", "QAP",
    )),
)

OTC_PDF_RE = re.compile(
    r"(?P<proj>\d{4,6})\s*[-_]\s*(?P<code>\d{2,4})(?:\.(?P<sub>\d{2}))?",
    re.I,
)
IAM_RE = re.compile(
    r"(?P<proj>\d{4,6})-(?P<fam>\d{3,5})-(?P<tipo>[AP])(?P<num>\d{2,3})",
    re.I,
)
SWE_RE = re.compile(r"TNK[\.\s\-]*1PH[\.\s\-]*\d+|1PH[\.\s]*TNK[\.\s\-]*\d+", re.I)
VT_RE = re.compile(r"(?:VT|VANTRAN)[\s\-_]*\d{4,}|^\d{6}(?:[\s\-_].*)?$", re.I)


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
    if name.upper().startswith("ORDENES DE PRODUCCION"):
        return False
    if name.upper() in ("COMPONENTES ESTANDAR", "MACHOTE"):
        return False
    return bool(re.match(r"^\d{3,}", name.strip()) or name.upper().startswith("GAE"))


def brand(name: str) -> str:
    u = name.upper()
    for b in BRANDS:
        if b.strip() in u:
            return b.strip()
    parts = re.split(r"[\s\-_&]+", name)
    for p in parts:
        if len(p) >= 3 and not p.isdigit() and not re.fullmatch(r"X\d+", p.upper()):
            return p.upper()[:24]
    return "OTRO"


def find_planos_dir(order: Path) -> Path | None:
    try:
        kids = list(order.iterdir())
    except OSError:
        return None
    exact, soft = [], []
    for c in kids:
        if not c.is_dir():
            continue
        n = c.name.strip().upper()
        if re.match(r"^4[\.\s\-_]*PLANOS?", n):
            exact.append(c)
        elif "PLANO" in n and "OLD" not in n:
            soft.append(c)
    if exact:
        return exact[0]
    if soft:
        return soft[0]
    return None


def list_drawings(planos: Path, depth=0, max_depth=3, hard_cap=800):
    out = []
    if depth > max_depth:
        return out
    skip_dirs = {
        "OLD", "OLDVERSIONS", "OLD VERSIONS", "BAK", "BACKUP",
        "OBSOLETO", "ARCHIVO",
    }
    try:
        for p in planos.iterdir():
            if len(out) >= hard_cap:
                break
            try:
                if p.is_file() and p.suffix.lower() in (
                    ".pdf", ".dwg", ".idw", ".dwf", ".tif", ".tiff", ".png", ".jpg",
                ):
                    out.append(p)
                elif p.is_dir() and p.name.upper() not in skip_dirs:
                    out.extend(
                        list_drawings(
                            p, depth + 1, max_depth, hard_cap - len(out)
                        )
                    )
            except OSError:
                continue
    except OSError:
        pass
    return out[:hard_cap]


def classify_tipo(text: str) -> list[str]:
    u = " " + re.sub(r"\s+", " ", text.upper()) + " "
    hits = []
    for tipo, kws in TIPO_RULES:
        if any(k in u for k in kws):
            hits.append(tipo)
    return hits or ["SIN_CLASIFICAR"]


def naming_scheme(stem: str, brand_name: str) -> str:
    s = stem.strip()
    if OTC_PDF_RE.search(s.replace(" ", "")) or OTC_PDF_RE.search(s):
        return "CODIGO_NUM_XX_YY"  # 62201-48.00
    if IAM_RE.search(s.replace(" ", "")):
        return "IAM_ESTILO_A_P"  # 62176-1248-A01
    if SWE_RE.search(s) or "SWE" in brand_name:
        if re.search(r"TNK|1PH|TANK", s, re.I):
            return "SWE_TNK_1PH"
    if brand_name == "VANTRAN" or VT_RE.search(s):
        return "VANTRAN_O_NUMERICO"
    if re.search(r"GA\b|GENERAL|ASSY|ENSAMBLE", s, re.I):
        return "DESCRIPCION_GA"
    if re.search(r"[A-Z]{2,}\s*-?\s*\d+", s):
        return "ALFA_NUM"
    return "LIBRE_TEXTO"


def try_pdf_text(path: Path, max_pages=2, max_chars=4000) -> str:
    """Extrae texto de las primeras páginas si hay librería PDF."""
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        try:
            from PyPDF2 import PdfReader  # type: ignore
        except Exception:
            return ""
    try:
        reader = PdfReader(str(path))
        chunks = []
        for i, page in enumerate(reader.pages[:max_pages]):
            try:
                t = page.extract_text() or ""
            except Exception:
                t = ""
            if t.strip():
                chunks.append(t)
            if sum(len(c) for c in chunks) >= max_chars:
                break
        return "\n".join(chunks)[:max_chars]
    except Exception:
        return ""


def main():
    lines = []

    def w(s=""):
        lines.append(s)
        log(s)

    root = op_root()
    w(f"OP root: {root.parent.name}\\{root.name}")

    orders = [d for d in root.iterdir() if d.is_dir() and is_real_order(d.name)]
    orders = sorted(orders, key=lambda p: p.name)
    max_scan = int(os.environ.get("MAX_OP_SCAN", "200"))
    orders = orders[:max_scan]
    w(f"Órdenes a escanear: {len(orders)}")

    brand_ops = Counter()
    brand_with_planos = Counter()
    brand_pdfs = Counter()
    tipo_global = Counter()
    tipo_por_marca = defaultdict(Counter)
    scheme_global = Counter()
    scheme_por_marca = defaultdict(Counter)
    ext_c = Counter()
    subfolder_c = Counter()
    no_planos = []
    empty_planos = []
    samples_by_brand = defaultdict(list)
    samples_by_tipo = defaultdict(list)
    pdf_text_samples = []
    # No guardar todos los records (OPs grandes en red); solo conteos + muestras.

    extract_text = os.environ.get("EXTRACT_PDF_TEXT", "1") != "0"
    text_budget = int(os.environ.get("PDF_TEXT_SAMPLES", "28"))
    max_files_per_op = int(os.environ.get("MAX_FILES_PER_OP", "400"))

    for i, order in enumerate(orders, 1):
        b = brand(order.name)
        brand_ops[b] += 1
        log(f"[{i}/{len(orders)}] {b} | {order.name[:85]}")

        planos = find_planos_dir(order)
        if planos is None:
            no_planos.append(order.name)
            log("  (sin 4. Planos)")
            continue
        brand_with_planos[b] += 1

        drawings = list_drawings(planos, hard_cap=max_files_per_op)
        log(f"  archivos={len(drawings)} en {planos.name}")
        if not drawings:
            empty_planos.append(order.name)
            continue

        pdfs_sorted = sorted(
            [p for p in drawings if p.suffix.lower() == ".pdf"],
            key=lambda p: p.stat().st_size if p.exists() else 0,
            reverse=True,
        )

        for p in drawings:
            try:
                rel = str(p.relative_to(planos))
            except Exception:
                rel = p.name
            stem = p.stem
            ext_c[p.suffix.lower()] += 1
            brand_pdfs[b] += 1

            parts = Path(rel).parts
            if len(parts) > 1:
                subfolder_c[f"{b}|{'/'.join(parts[:-1])[:80]}"] += 1
            else:
                subfolder_c[f"{b}|."] += 1

            blob = f"{rel} {stem}"
            tipos = classify_tipo(blob)
            scheme = naming_scheme(stem, b)
            scheme_global[scheme] += 1
            scheme_por_marca[b][scheme] += 1

            for t in tipos:
                tipo_global[t] += 1
                tipo_por_marca[b][t] += 1
                if len(samples_by_tipo[t]) < 8:
                    samples_by_tipo[t].append(f"[{b}] {rel}")

            if len(samples_by_brand[b]) < 12:
                samples_by_brand[b].append(f"{tipos[0]} | {rel}")

        if extract_text and len(pdf_text_samples) < text_budget:
            for p in pdfs_sorted[:2]:
                if len(pdf_text_samples) >= text_budget:
                    break
                try:
                    if p.stat().st_size > 40_000_000:
                        continue
                except OSError:
                    continue
                txt = try_pdf_text(p)
                if not txt.strip():
                    continue
                tipos_txt = classify_tipo(txt[:2000] + " " + p.stem)
                pdf_text_samples.append({
                    "brand": b,
                    "order": order.name[:60],
                    "file": p.name[:90],
                    "tipos_nombre": classify_tipo(p.stem),
                    "tipos_texto": tipos_txt,
                    "preview": re.sub(r"\s+", " ", txt)[:280],
                })
                log(f"  PDF text OK: {p.name[:55]}")

        log(f"  OK marca={b} acumulado_archivos={brand_pdfs[b]}")

    w()
    w("=" * 72)
    w("RESUMEN GLOBAL — 4. Planos")
    w("=" * 72)
    w(f"OPs escaneadas: {len(orders)}")
    w(f"OPs con carpeta Planos: {sum(brand_with_planos.values())}")
    w(f"OPs SIN carpeta Planos: {len(no_planos)}")
    w(f"OPs con Planos vacía: {len(empty_planos)}")
    w(f"Archivos de dibujo totales: {sum(ext_c.values())}")
    w()
    w("--- Extensiones ---")
    for k, v in ext_c.most_common():
        w(f"  {k}: {v}")

    w()
    w("--- Marcas (OPs / con Planos / archivos) ---")
    for b, n in brand_ops.most_common():
        w(
            f"  {b}: {n} OPs | con Planos={brand_with_planos.get(b, 0)} | "
            f"archivos={brand_pdfs.get(b, 0)}"
        )

    w()
    w("--- Tipología por nombre/ruta (un archivo puede tener >1) ---")
    tot_t = sum(tipo_global.values()) or 1
    for k, v in tipo_global.most_common():
        w(f"  {k}: {v} ({100 * v / tot_t:.1f}%)")

    w()
    w("--- Esquemas de nombrado ---")
    for k, v in scheme_global.most_common():
        w(f"  {k}: {v}")

    w()
    w("--- Tipología por marca (top) ---")
    for b, _ in brand_pdfs.most_common():
        w(f"\n[{b}] archivos={brand_pdfs[b]}")
        for t, v in tipo_por_marca[b].most_common(10):
            w(f"  {t}: {v}")
        w("  esquemas: " + ", ".join(
            f"{k}={v}" for k, v in scheme_por_marca[b].most_common(5)
        ))
        w("  muestras:")
        for s in samples_by_brand[b][:8]:
            w(f"    - {s}")

    w()
    w("--- Muestras por tipo de plano ---")
    for t, _ in tipo_global.most_common():
        w(f"\n{t}:")
        for s in samples_by_tipo[t][:6]:
            w(f"  - {s}")

    w()
    w("--- Subcarpetas frecuentes (marca|ruta) ---")
    for k, v in subfolder_c.most_common(40):
        w(f"  [{v}] {k}")

    if no_planos:
        w()
        w("--- OPs sin carpeta 4. Planos ---")
        for n in no_planos[:30]:
            w(f"  - {n[:90]}")

    if pdf_text_samples:
        w()
        w("--- Muestras de texto PDF (título/contenido) ---")
        for s in pdf_text_samples:
            w(f"\n[{s['brand']}] {s['file']}")
            w(f"  nombre→ {s['tipos_nombre']} | texto→ {s['tipos_texto']}")
            w(f"  preview: {s['preview']}")
    else:
        w()
        w("(Sin texto PDF: instala pypdf o PyPDF2, o hay PDFs escaneados.)")

    w()
    w("=" * 72)
    w("IMPLICACIONES PARA COTAS ABIGAIL")
    w("=" * 72)
    w("1. Hay varias familias de producto (OTC/SWE/Vantran/GIGA/PTT/…), no un solo estándar de PDF.")
    w("2. OTC suele usar códigos NN.00 (46 tanque / 47 tapa / 48 casco); otras marcas usan texto libre o códigos propios.")
    w("3. En '4. Planos' conviven GA, casco, tapa, detalles, BOM, soldadura y a veces fab — Abigail solo cubre cotas geométricas de pieza/cara.")
    w("4. Tipología útil para el flujo: TANQUE/TAPA/CASCO/BASE/ACCESORIO vs DETALLE/BOM/WELD/SPEC (estos últimos no son entrada directa del generador).")
    w("5. Validar Abigail por marca: mismo motor, distinto naming y densidad de accesorios.")
    w("6. Si un PDF es escaneo, el nombre de archivo/carpeta es la mejor señal (texto embebido vacío).")

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    payload = {
        "orders": len(orders),
        "with_planos": sum(brand_with_planos.values()),
        "no_planos": no_planos,
        "empty_planos": empty_planos,
        "ext": dict(ext_c),
        "brand_ops": dict(brand_ops),
        "brand_pdfs": dict(brand_pdfs),
        "tipo_global": dict(tipo_global),
        "tipo_por_marca": {k: dict(v) for k, v in tipo_por_marca.items()},
        "scheme_global": dict(scheme_global),
        "scheme_por_marca": {k: dict(v) for k, v in scheme_por_marca.items()},
        "pdf_text_samples": pdf_text_samples,
        "records_n": sum(brand_pdfs.values()),
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    w(f"\nGuardado: {OUT_TXT.name}")
    w(f"Guardado: {OUT_JSON.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
