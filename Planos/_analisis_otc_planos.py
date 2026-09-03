# -*- coding: utf-8 -*-
"""Analisis global de Planos OTC en ORDENES DE PRODUCCION (Z:).

Recorre todas las OPs con 'OTC' en el nombre, inventaria '4. Planos'
y extrae patrones de codigos de dibujo / ensamble para generalizar
el mapeo TOP/BASE/SEGM del generador.
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

OUT_TXT = Path(__file__).with_name("_analisis_otc_planos_resultado.txt")
OUT_JSON = Path(__file__).with_name("_analisis_otc_planos_resultado.json")
OUT_MD = Path(__file__).resolve().parents[1] / "PATRONES_OTC_GLOBAL.md"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# PDF tipico: 62201-48.00 Rev 01.PDF  |  62176-47.00.pdf
PDF_CODE_RE = re.compile(
    r"(?P<proj>\d{4,6})\s*[-_]\s*(?P<code>\d{2,4})(?:\.(?P<sub>\d{2}))?",
    re.IGNORECASE,
)
# IAM/STEP tipico: 62201-1248-A03 / 62176-1247-P01
IAM_CODE_RE = re.compile(
    r"(?P<proj>\d{4,6})-(?P<fam>\d{3,5})-(?P<tipo>[AP])(?P<num>\d{2,3})",
    re.IGNORECASE,
)
PRODUCT_RE = re.compile(
    r"PRODUCT\s*\(\s*'((?:[^']|'')*)'\s*,\s*'((?:[^']|'')*)'",
    re.IGNORECASE,
)

# Roles de dibujo por ultimos digitos del codigo de plano (xx.00)
# Confirmados en piso OTC: 46 tanque, 47 tapa, 48 casco
DRAWING_ROLE_HINT = {
    "46": "TANQUE_GENERAL",
    "47": "TOP_COVER",
    "48": "CASCO_SHELL",
    "50": "DETALLE_O_SUB",
    "51": "DETALLE_O_SUB",
    "52": "DETALLE_O_SUB",
    "53": "DETALLE_O_SUB",
    "54": "DETALLE_O_SUB",
    "60": "DETALLE_O_SUB",
}


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


def find_planos_dir(order: Path) -> Path | None:
    """Busca carpeta '4. Planos' (tolerante a variantes de nombre)."""
    try:
        kids = list(order.iterdir())
    except OSError:
        return None
    exact = []
    soft = []
    for c in kids:
        if not c.is_dir():
            continue
        n = c.name.strip().upper()
        if re.match(r"^4[\.\s\-_]*PLANOS?", n):
            exact.append(c)
        elif "PLANO" in n and not n.startswith("OLD"):
            soft.append(c)
    if exact:
        return exact[0]
    if soft:
        return soft[0]
    return None


def find_solidos_dir(order: Path) -> Path | None:
    try:
        for c in order.iterdir():
            if not c.is_dir():
                continue
            n = c.name.upper()
            if "SOLIDO" in n or re.match(r"^10[\.\s]", c.name.strip()):
                return c
    except OSError:
        return None
    return None


def list_pdfs(planos: Path, depth=0, max_depth=3):
    out = []
    if depth > max_depth:
        return out
    try:
        for p in planos.iterdir():
            try:
                if p.is_file() and p.suffix.lower() in (".pdf", ".dwg", ".idw"):
                    if "OLD" in p.parts:
                        continue
                    out.append(p)
                elif p.is_dir() and p.name.upper() not in (
                    "OLD",
                    "OLDVERSIONS",
                    "OLD VERSIONS",
                    "BAK",
                    "BACKUP",
                ):
                    # STD a veces tiene planos utiles; incluir un nivel
                    out.extend(list_pdfs(p, depth + 1, max_depth))
            except OSError:
                continue
    except OSError:
        pass
    return out


def find_otc_steps(order: Path, limit=3):
    solidos = find_solidos_dir(order)
    if solidos is None:
        return []
    found = []

    def walk(folder: Path, depth=0):
        if depth > 3 or len(found) >= 20:
            return
        try:
            for p in folder.iterdir():
                try:
                    if p.is_file() and p.suffix.lower() in (".stp", ".step"):
                        if p.stat().st_size > 80_000:
                            found.append(p)
                    elif p.is_dir() and p.name.upper() not in (
                        "OLD",
                        "OLDVERSIONS",
                        "OLD VERSIONS",
                    ):
                        if (
                            depth == 0
                            or "TANQUE" in p.name.upper()
                            or "TANK" in p.name.upper()
                            or IAM_CODE_RE.search(p.name)
                        ):
                            walk(p, depth + 1)
                except OSError:
                    continue
        except OSError:
            pass

    # priorizar TANQUE
    try:
        for sub in solidos.iterdir():
            if sub.is_dir() and (
                "TANQUE" in sub.name.upper() or sub.name.upper() == "TANK"
            ):
                walk(sub, 0)
    except OSError:
        pass
    if not found:
        walk(solidos, 0)
    uniq = {str(p).lower(): p for p in found}
    return sorted(uniq.values(), key=lambda x: x.stat().st_size, reverse=True)[
        :limit
    ]


def extract_products(path: Path, limit_bytes=4_000_000):
    try:
        with path.open("r", errors="ignore") as fh:
            data = fh.read(min(path.stat().st_size, limit_bytes))
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


def familia_rol(familia: str) -> str:
    digitos = re.sub(r"\D", "", str(familia or ""))
    if len(digitos) >= 2:
        return digitos[-2:]
    return digitos


def classify_iam(nombre: str):
    m = IAM_CODE_RE.search(nombre.upper().replace(" ", ""))
    if not m:
        return None
    fam = m.group("fam")
    rol = familia_rol(fam)
    tipo = m.group("tipo").upper()
    num = int(m.group("num"))
    role = "OTRO"
    if rol == "46" and tipo == "A" and num == 1:
        role = "TANQUE"
    elif rol == "47" and tipo == "A":
        role = "TOP"
    elif rol == "48" and tipo == "A" and num == 1:
        role = "CASCO_SHELL"
    elif rol == "48" and tipo == "A" and num == 2:
        role = "BASE"
    elif rol == "48" and tipo == "A" and num in (3, 4, 5, 6):
        role = "SEGM_WALL"
    elif rol == "48" and tipo == "A" and num >= 7:
        role = "SUB_A_CASCO"
    elif rol == "48" and tipo == "P":
        role = "PIEZA_CASCO"
    elif rol == "47" and tipo == "P":
        role = "PIEZA_TOP"
    elif rol == "46" and tipo == "P":
        role = "PIEZA_TANQUE"
    return {
        "proj": m.group("proj"),
        "fam": fam,
        "tipo": tipo,
        "num": num,
        "rol_familia": rol,
        "role": role,
        "raw": m.group(0),
    }


def parse_drawing_name(name: str):
    stem = Path(name).stem
    # quitar Rev XX
    stem_clean = re.sub(r"\s*Rev\.?\s*\d+.*$", "", stem, flags=re.I).strip()
    m = PDF_CODE_RE.search(stem_clean.replace(" ", ""))
    if not m:
        # intentar con espacios: 62201 - 48.00
        m = PDF_CODE_RE.search(stem_clean)
    if not m:
        return None
    code = m.group("code")
    # normalizar: si code es 1248 usar ultimos 2; si es 48 usar 48
    if len(code) >= 4:
        drawing_key = code[-2:]
        fam_full = code
    else:
        drawing_key = code.zfill(2)[-2:]
        fam_full = code
    return {
        "proj": m.group("proj"),
        "code": code,
        "sub": m.group("sub") or "00",
        "drawing_key": drawing_key,
        "fam_full": fam_full,
        "hint": DRAWING_ROLE_HINT.get(drawing_key, "OTRO_O_DETALLE"),
        "stem": stem_clean,
    }


def main():
    lines = []

    def w(s=""):
        lines.append(s)
        log(s)

    root = op_root()
    w(f"OP root: {root}")
    orders = sorted(
        [
            d
            for d in root.iterdir()
            if d.is_dir() and "OTC" in d.name.upper()
        ],
        key=lambda p: p.name,
    )
    w(f"OPs OTC encontradas: {len(orders)}")
    for o in orders:
        w(f"  - {o.name}")

    tanks = []
    drawing_key_c = Counter()
    drawing_hint_c = Counter()
    iam_role_c = Counter()
    fam_role_c = Counter()  # (rol_familia, tipo, num_bucket) 
    a48_nums = Counter()  # A0x bajo familia 48
    a47_nums = Counter()
    proj_to_drawings = defaultdict(list)
    sp_codes = Counter()
    typ_suffix = Counter()  # cuantas piezas con _NNN
    wall_patterns = Counter()

    for i, order in enumerate(orders, 1):
        w()
        w(f"=== [{i}/{len(orders)}] {order.name} ===")
        # proyecto desde nombre OP
        mproj = re.search(r"OTC\s+(\d{4,6})", order.name, re.I)
        proyecto = mproj.group(1) if mproj else "?"

        planos = find_planos_dir(order)
        entry = {
            "order": order.name,
            "proyecto": proyecto,
            "planos_path": str(planos) if planos else None,
            "pdfs": [],
            "drawings": [],
            "steps": [],
            "iam_roles": Counter(),
            "products_sample": [],
        }

        if planos is None:
            w("  SIN carpeta 4. Planos")
        else:
            w(f"  Planos: {planos.name}")
            pdfs = list_pdfs(planos)
            entry["pdfs"] = [p.name for p in pdfs]
            w(f"  Archivos planos (.pdf/.dwg/.idw): {len(pdfs)}")
            for p in sorted(pdfs, key=lambda x: x.name.lower()):
                parsed = parse_drawing_name(p.name)
                if parsed:
                    drawing_key_c[parsed["drawing_key"]] += 1
                    drawing_hint_c[parsed["hint"]] += 1
                    entry["drawings"].append(parsed)
                    proj_to_drawings[parsed["proj"]].append(parsed)
                    w(
                        f"    PDF {p.name[:70]} -> {parsed['proj']}-{parsed['code']}."
                        f"{parsed['sub']} [{parsed['hint']}]"
                    )
                else:
                    w(f"    PDF {p.name[:70]} -> (sin codigo tipico)")

        steps = find_otc_steps(order, limit=2)
        if not steps:
            w("  SIN STEP grande en 10. Solidos")
        for step in steps:
            mb = step.stat().st_size / 1e6
            names, err = extract_products(step)
            entry["steps"].append(
                {
                    "name": step.name,
                    "mb": round(mb, 2),
                    "n_products": len(names) if not err else 0,
                    "error": err,
                }
            )
            if err:
                w(f"  STEP {step.name[:55]} ERR: {err}")
                continue
            w(f"  STEP {step.name[:55]} ({mb:.1f}MB) products={len(names)}")

            roles_local = Counter()
            for n in names:
                # SP / L845 / SF-
                nu = n.upper()
                msp = re.match(r"^(SP-\d+|SF-[A-Z]+-\d+|L\d+\.[A-Z]+|PT-[A-Z]+-\d+|FT-[A-Z]+-\d+)", nu)
                if msp:
                    sp_codes[msp.group(1)] += 1
                if re.search(r"_\d+$", nu.split(":")[0]):
                    typ_suffix["con_sufijo_numerico"] += 1
                else:
                    typ_suffix["sin_sufijo"] += 1

                info = classify_iam(n)
                if info is None:
                    continue
                iam_role_c[info["role"]] += 1
                roles_local[info["role"]] += 1
                fam_role_c[
                    (info["rol_familia"], info["tipo"], info["num"])
                ] += 1
                if info["rol_familia"] == "48" and info["tipo"] == "A":
                    a48_nums[info["num"]] += 1
                if info["rol_familia"] == "47" and info["tipo"] == "A":
                    a47_nums[info["num"]] += 1
                if info["role"] == "SEGM_WALL":
                    wall_patterns[info["raw"]] += 1

            entry["iam_roles"] = dict(roles_local)
            # muestra de contenedores clave
            sample = []
            for n in names:
                info = classify_iam(n)
                if info and info["role"] in (
                    "TANQUE",
                    "TOP",
                    "CASCO_SHELL",
                    "BASE",
                    "SEGM_WALL",
                ):
                    sample.append(f"{info['raw']}={info['role']}")
            # unicos preservando orden
            seen = set()
            uniq_sample = []
            for s in sample:
                if s not in seen:
                    seen.add(s)
                    uniq_sample.append(s)
            entry["products_sample"] = uniq_sample[:30]
            w(f"  roles IAM: {dict(roles_local)}")
            w(f"  contenedores: {uniq_sample[:20]}")

        # serializable
        entry["iam_roles"] = dict(entry["iam_roles"])
        tanks.append(entry)

    w()
    w("=" * 60)
    w("RESUMEN GLOBAL OTC")
    w("=" * 60)
    w(f"Tanques OTC: {len(tanks)}")
    con_planos = sum(1 for t in tanks if t["planos_path"])
    con_pdf = sum(1 for t in tanks if t["pdfs"])
    con_step = sum(1 for t in tanks if t["steps"])
    w(f"Con 4. Planos: {con_planos} | con PDF/DWG: {con_pdf} | con STEP: {con_step}")

    w()
    w("--- Codigos de dibujo (ultimos 2 digitos) ---")
    for k, v in drawing_key_c.most_common():
        hint = DRAWING_ROLE_HINT.get(k, "OTRO")
        w(f"  *-{k}.xx : {v} archivos  [{hint}]")

    w()
    w("--- Roles IAM en STEPs ---")
    for k, v in iam_role_c.most_common():
        w(f"  {k}: {v}")

    w()
    w("--- Ensamble *-xx48-A0N (casco) ---")
    for num, v in sorted(a48_nums.items()):
        label = {
            1: "CASCO_SHELL",
            2: "BASE",
            3: "SEGM",
            4: "SEGM",
            5: "SEGM",
            6: "SEGM",
        }.get(num, "SUB_A")
        w(f"  A{num:02d}: {v} menciones  [{label}]")

    w()
    w("--- Ensamble *-xx47-A0N (tapa) ---")
    for num, v in sorted(a47_nums.items()):
        w(f"  A{num:02d}: {v}")

    w()
    w("--- Codigos estandar frecuentes (SP/SF/L...) ---")
    for k, v in sp_codes.most_common(40):
        w(f"  {k}: {v}")

    w()
    w("--- Sufijos _NNN ---")
    for k, v in typ_suffix.most_common():
        w(f"  {k}: {v}")

    # Consistencia del patron 46/47/48
    proyectos_con_46 = set()
    proyectos_con_47 = set()
    proyectos_con_48 = set()
    for proj, drawings in proj_to_drawings.items():
        keys = {d["drawing_key"] for d in drawings}
        if "46" in keys:
            proyectos_con_46.add(proj)
        if "47" in keys:
            proyectos_con_47.add(proj)
        if "48" in keys:
            proyectos_con_48.add(proj)

    w()
    w("--- Cobertura de planos 46/47/48 por proyecto ---")
    w(f"  con *-46.xx: {sorted(proyectos_con_46)}")
    w(f"  con *-47.xx: {sorted(proyectos_con_47)}")
    w(f"  con *-48.xx: {sorted(proyectos_con_48)}")

    # Reglas propuestas
    rules = {
        "drawing_to_assembly": {
            "46": "familia ...46 / tanque *-1246-A01 (o *-*46-A01)",
            "47": "familia ...47 / TOP *-1247-A01",
            "48": "familia ...48 / casco *-1248-A01 con A02=BASE, A03-A06=paredes",
        },
        "assembly_tree": {
            "TANQUE": "*-*46-A01",
            "TOP": "*-*47-A0N (casi siempre A01)",
            "CASCO_SHELL": "*-*48-A01  (NO acotar como cara)",
            "BASE": "*-*48-A02",
            "SEGM": "*-*48-A03..A06 (4 paredes tipicas)",
            "SUB": "*-*48-A07+  subensambles de pared",
        },
        "grouping": "mismo nombre hasta '_' (quitar _\\d+$)",
        "nesting": "placas Pxx_1 / Pxx_2 = mitades; origen = placa madre completa",
    }

    # Markdown report
    md = []
    md.append("# Patrones globales OTC (Planos + ensamble)")
    md.append("")
    md.append(
        "Fuente: `Z:\\...\\ORDENES DE PRODUCCION\\*OTC*\\4. Planos` "
        "+ STEPs en `10. Solidos` cuando existen."
    )
    md.append("")
    md.append(f"**OPs OTC analizadas:** {len(tanks)}")
    md.append("")
    md.append("## 1. Arbol de ensamble tipico (estable entre tanques)")
    md.append("")
    md.append("```")
    md.append("*-*46-A01.iam          ← TANQUE (plano *-*46.00)")
    md.append("├── *-*47-A01          ← TOP COVER (plano *-*47.00 / 47.01)")
    md.append("└── *-*48-A01          ← CASCO / shell (plano *-*48.00) — NO es una cara")
    md.append("    ├── *-*48-A02      ← BASE")
    md.append("    ├── *-*48-A03      ← pared / SEGM")
    md.append("    ├── *-*48-A04      ← pared / SEGM")
    md.append("    ├── *-*48-A05      ← pared / SEGM")
    md.append("    └── *-*48-A06      ← pared / SEGM")
    md.append("```")
    md.append("")
    md.append(
        "La familia numerica usa los **ultimos 2 digitos**: "
        "`1246`→46 tanque, `1247`→47 tapa, `1248`→48 casco. "
        "El prefijo (`12`, etc.) puede variar; el rol lo dan `46/47/48`."
    )
    md.append("")
    md.append("## 2. Codigos de dibujo en 4. Planos")
    md.append("")
    md.append("| Codigo PDF | Rol | Frecuencia |")
    md.append("|------------|-----|------------|")
    for k, v in drawing_key_c.most_common():
        hint = DRAWING_ROLE_HINT.get(k, "OTRO_O_DETALLE")
        md.append(f"| `*-{k}.xx` | {hint} | {v} |")
    md.append("")
    md.append("## 3. Reglas para el generador (globalizar)")
    md.append("")
    md.append("1. **Contenedor de cara** = IAM `A02`…`A06` (48) o `47-A*`, nunca `48-A01` ni `46-A01`.")
    md.append("2. **TOP** = familia rol `47`; **BASE** = `48-A02`; **SEGM** = `48-A03`…`A06`.")
    md.append("3. **Agrupar piezas** por nombre hasta `_` (`SP-852_1` = `SP-852_2`).")
    md.append("4. **Nesteo** `Pxx_1`/`Pxx_2`: excluir ambas mitades como 'pared madre'; origen = placa completa.")
    md.append("5. Piezas `SP-*`, `SF-*`, `L845.*`, `PT-*`, `FT-*` son accesorios tipicos OTC (no casco).")
    md.append("6. Planos `50`–`60` suelen ser detalles/subarmados; no redefinir TOP/BASE/SEGM.")
    md.append("")
    md.append("## 4. Por tanque")
    md.append("")
    for t in tanks:
        md.append(f"### {t['order']}")
        md.append("")
        md.append(f"- Proyecto: `{t['proyecto']}`")
        md.append(
            f"- Planos: {'si' if t['planos_path'] else 'no'} "
            f"({len(t['pdfs'])} archivos)"
        )
        if t["drawings"]:
            codes = sorted({f"{d['code']}.{d['sub']}" for d in t["drawings"]})
            md.append(f"- Codigos PDF: {', '.join(codes)}")
        if t["steps"]:
            for s in t["steps"]:
                md.append(
                    f"- STEP: `{s['name']}` ({s['mb']} MB, "
                    f"{s.get('n_products', 0)} products)"
                )
        if t["products_sample"]:
            md.append(f"- Contenedores: `{', '.join(t['products_sample'][:15])}`")
        if t["iam_roles"]:
            md.append(f"- Roles IAM: `{t['iam_roles']}`")
        md.append("")

    md.append("## 5. Implicacion")
    md.append("")
    md.append(
        "No hardcodear solo `62201-1248-A0x`. Usar regex "
        "`(?P<proj>\\d+)-(?P<fam>\\d+)-(?P<tipo>[AP])(?P<num>\\d+)` "
        "y rol = ultimos 2 digitos de `fam` ∈ {46,47,48}."
    )
    md.append("")

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    OUT_JSON.write_text(
        json.dumps(
            {
                "tanks": tanks,
                "drawing_key_c": dict(drawing_key_c),
                "drawing_hint_c": dict(drawing_hint_c),
                "iam_role_c": dict(iam_role_c),
                "a48_nums": {str(k): v for k, v in a48_nums.items()},
                "a47_nums": {str(k): v for k, v in a47_nums.items()},
                "sp_codes": dict(sp_codes.most_common(60)),
                "typ_suffix": dict(typ_suffix),
                "rules": rules,
                "proyectos_46": sorted(proyectos_con_46),
                "proyectos_47": sorted(proyectos_con_47),
                "proyectos_48": sorted(proyectos_con_48),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    w()
    w(f"Guardado TXT: {OUT_TXT}")
    w(f"Guardado JSON: {OUT_JSON}")
    w(f"Guardado MD: {OUT_MD}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
