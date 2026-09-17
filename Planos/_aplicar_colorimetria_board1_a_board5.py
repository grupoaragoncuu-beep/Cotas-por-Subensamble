# -*- coding: utf-8 -*-
"""Copia colorimetría (apariencia + iProperty Clasificación) de Board 1 a Board 5.

Regla de referencia:
  Colorimetria Sub Assembly + Norman.iLogicVb

Mapeo apariencia <-> clasificación:
  Almacén         -> Dark Green / Verde oscuro
  Corte           -> Rojo naranja
  Maquinado       -> Blue - Wall Paint - Glossy / Azul - Pintura mural - Brillante
  Doblado         -> Violeta
  Plasma          -> Yellow - Wall Paint - Glossy / Amarillo...
  Plasma Doblado  -> White - Wall Paint - Glossy / Blanco...

Estrategia:
  1) Leer Clasificación real de cada .ipt leaf en Board 1 (vía COM).
  2) Para cada .ipt leaf en Board 5, si el nombre base coincide (casefold),
     escribir Clasificación + apariencia igual que Board 1.
  3) Guardar documentos modificados.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import win32com.client

from inventor_com import conectar_inventor, localizar_documento

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def _as_assembly(doc):
    """Cast seguro a AssemblyDocument (Document genérico no expone ComponentDefinition)."""
    if doc is None:
        return None
    try:
        return win32com.client.CastTo(doc, "AssemblyDocument")
    except Exception:
        pass
    try:
        if int(doc.DocumentType) == kAssemblyDocumentObject:
            return doc
    except Exception:
        pass
    return None


def _as_part(doc):
    try:
        return win32com.client.CastTo(doc, "PartDocument")
    except Exception:
        return doc

PROPSET = "Inventor User Defined Properties"
PROP_CLS = "Clasificación"

VALIDAS = (
    "Almacén",
    "Corte",
    "Maquinado",
    "Doblado",
    "Plasma",
    "Plasma Doblado",
)

APARIENCIAS = {
    "Almacén": ("Dark Green", "Verde oscuro"),
    "Corte": ("Rojo naranja",),
    "Maquinado": (
        "Blue - Wall Paint - Glossy",
        "Azul - Pintura mural - Brillante",
    ),
    "Doblado": ("Violeta",),
    "Plasma": (
        "Yellow - Wall Paint - Glossy",
        "Amarillo - Pintura mural - Brillante",
        "Yellow",
        "Amarillo",
    ),
    "Plasma Doblado": (
        "White - Wall Paint - Glossy",
        "Blanco - Pintura mural - Brillante",
        "White",
        "Blanco",
    ),
}

# DocumentTypeEnum
kPartDocumentObject = 12290
kAssemblyDocumentObject = 12291

OUT = Path(__file__).with_name("_colorimetria_board1_a_board5_log.json")
DRY_RUN = os.environ.get("DRY_RUN", "0").strip() == "1"
APPLY = os.environ.get("APPLY", "1").strip() != "0"


def log(msg=""):
    print(msg, flush=True)


def _norm_cls(val):
    if val is None:
        return None
    t = str(val).strip()
    if not t:
        return None
    for v in VALIDAS:
        if t.casefold() == v.casefold():
            return v
    return None


def _leer_cls(part_doc):
    try:
        props = part_doc.PropertySets.Item(PROPSET)
    except Exception:
        return None
    try:
        return _norm_cls(props.Item(PROP_CLS).Value)
    except Exception:
        return None


def _escribir_cls(part_doc, clasificacion):
    props = part_doc.PropertySets.Item(PROPSET)
    try:
        prop = props.Item(PROP_CLS)
        prop.Value = clasificacion
    except Exception:
        props.Add(clasificacion, PROP_CLS)


def _basename(doc_or_path):
    try:
        if hasattr(doc_or_path, "FullFileName"):
            path = doc_or_path.FullFileName or ""
        else:
            path = str(doc_or_path or "")
    except Exception:
        path = ""
    if not path:
        return ""
    return os.path.splitext(os.path.basename(path))[0]


def _find_top_iam(inv, needle):
    """Busca el .iam top-level cuyo DisplayName empiece exactamente por needle."""
    needle_u = needle.strip().upper()
    # 1) ActiveDocument si coincide
    try:
        ad = inv.ActiveDocument
        if ad and (ad.DisplayName or "").upper().startswith(needle_u):
            if str(ad.FullFileName or "").lower().endswith(".iam"):
                return ad
    except Exception:
        pass
    # 2) localizar por display name exacto
    for cand in (f"{needle}.iam", needle):
        doc = localizar_documento(inv, display_name=cand)
        if doc is not None:
            try:
                if str(doc.FullFileName or "").lower().endswith(".iam"):
                    return doc
            except Exception:
                pass
    # 3) Escaneo acotado: solo documentos cuyo DisplayName empieza por needle
    #    (evita recorrer FullFileName de 500 docs).
    try:
        n = int(inv.Documents.Count)
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            d = inv.Documents.Item(i)
            name = (d.DisplayName or "").strip()
        except Exception:
            continue
        if not name.upper().startswith(needle_u):
            continue
        try:
            fn = d.FullFileName or ""
        except Exception:
            fn = ""
        if fn.lower().endswith(".iam") and name.upper().startswith(needle_u):
            # preferir el que es exactamente "9919-Board N.iam" / "9919-Board N"
            base = os.path.splitext(name)[0].upper()
            if base == needle_u:
                return d
    return None


def _iter_leaf_parts(asm_doc):
    """Yield (occ, part_doc, basename_upper) de AllLeafOccurrences que son .ipt."""
    asm = _as_assembly(asm_doc)
    if asm is None:
        log("  ERROR: documento no es AssemblyDocument")
        return
    try:
        leaves = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
        total = int(leaves.Count)
    except Exception as exc:
        log(f"  ERROR enumerando leaves: {exc}")
        return
    for i in range(1, total + 1):
        try:
            occ = leaves.Item(i)
            if occ.Suppressed:
                continue
            doc = occ.Definition.Document
            if int(doc.DocumentType) != kPartDocumentObject:
                continue
            part = _as_part(doc)
            name = _basename(part)
            if not name:
                name = str(occ.Name).split(":")[0].strip()
            yield occ, part, name.upper()
        except Exception:
            continue


def _ensure_assets(asm_doc, inv):
    """Copia apariencias de la biblioteca Autodesk al ensamble (como el iLogic)."""
    asm = _as_assembly(asm_doc) or asm_doc
    lib = None
    for lib_name in (
        "Autodesk Biblioteca de aspecto",
        "Autodesk Appearance Library",
    ):
        try:
            lib = inv.AssetLibraries.Item(lib_name)
            break
        except Exception:
            continue
    if lib is None:
        log("  AVISO: no se encontró biblioteca de apariencias")
        return

    needed = []
    for names in APARIENCIAS.values():
        needed.extend(names)

    existing = set()
    try:
        for a in asm.Assets:
            try:
                existing.add(a.DisplayName)
            except Exception:
                pass
    except Exception:
        pass

    for nombre in needed:
        if nombre in existing:
            continue
        try:
            asset = lib.AppearanceAssets.Item(nombre)
            asset.CopyTo(asm)
            existing.add(nombre)
            log(f"  + apariencia cargada: {nombre}")
        except Exception:
            continue


def _resolve_asset(asm_doc, clasificacion):
    asm = _as_assembly(asm_doc) or asm_doc
    names = APARIENCIAS.get(clasificacion) or ()
    try:
        for a in asm.Assets:
            if a.DisplayName in names:
                return a
    except Exception:
        pass
    return None


def _aplicar_apariencia(occ, part_doc, asset):
    if asset is None:
        return False
    ok = False
    try:
        occ.Appearance = asset
        ok = True
    except Exception:
        pass
    try:
        part_doc.ComponentDefinition.Appearance = asset
        ok = True
    except Exception:
        pass
    return ok


def _mapa_desde_asm(asm_doc, label):
    """{basename_upper: clasificacion} — primera clase no-nula gana."""
    mapa = {}
    counts = Counter()
    sin = 0
    total = 0
    for _occ, doc, base in _iter_leaf_parts(asm_doc):
        total += 1
        cls = _leer_cls(doc)
        if cls:
            counts[cls] += 1
            if base not in mapa:
                mapa[base] = cls
        else:
            sin += 1
            mapa.setdefault(base, None)
    log(
        f"[{label}] leaves_ipt={total} unicos={len(mapa)} "
        f"clasificadas={sum(1 for v in mapa.values() if v)} "
        f"sin={sin} | {dict(counts)}"
    )
    return mapa


def _mcmaster_code(name: str) -> str:
    """Extrae código tipo 90126A029 / 91280A100 del nombre (también truncados)."""
    u = (name or "").upper()
    m = re.match(r"^(\d{5}[A-Z]\d+)", u)
    if m:
        return m.group(1)
    # HW-..._91280A100
    m = re.search(r"_(\d{5}[A-Z]\d+)\b", u)
    if m:
        return m.group(1)
    return ""


def _build_b1_indexes(mapa_b1):
    """Índices auxiliares derivados del mapa vivo de Board 1."""
    by_code = {}
    for name, cls in mapa_b1.items():
        if not cls:
            continue
        code = _mcmaster_code(name)
        if code and code not in by_code:
            by_code[code] = cls
    return by_code


# Familias con voto unánime (o casi) en Board 1 → regla segura.
_FAMILY_RULES = (
    (re.compile(r"^ABB-.*BCK", re.I), "Corte"),
    (re.compile(r"^HW-", re.I), "Almacén"),
    (re.compile(r"^FB-", re.I), "Doblado"),
    (re.compile(r"^GENE-BCU", re.I), "Corte"),
    (re.compile(r"^GENE-BKS", re.I), "Doblado"),
    (re.compile(r"^GENE-BKT", re.I), "Doblado"),
    (re.compile(r"^GENE-DF", re.I), "Doblado"),
    (re.compile(r"^GENE-FCU", re.I), "Corte"),
    (re.compile(r"^GENE-HFM", re.I), "Doblado"),
    (re.compile(r"^GENE-VFM", re.I), "Doblado"),
    (re.compile(r"^GENE-SIHC", re.I), "Doblado"),
    (re.compile(r"^GENE-SIVC", re.I), "Doblado"),
    (re.compile(r"^GEN1-OP", re.I), "Corte"),
    (re.compile(r"^RLG-J-", re.I), "Corte"),  # misma familia chapa plana tipo Corte
)

# GENE-OP / GENE-GS ambiguos en Board1: usar sufijos observados.
_OP_CORTE_SUFFIX = re.compile(r"-(115|116|117)$", re.I)
_OP_DOBLADO_SUFFIX = re.compile(r"-(114|211|311|308)$", re.I)

_ALMACEN_EXACT_PREFIXES = (
    "CITEL",
    "CT200",
    "EXISCAN",
    "CONNECTEURS",
    "CORPUS",
    "ECRAN",
    "ISO1-",
    "ISO2-",
    "1SDH",
    "2-250T",
    "NEXT_",
)


def _resolver_cls(base: str, mapa_b1: dict, by_code: dict):
    """
    Resuelve clasificación para una pieza Board5.
    Prioridad: exacto Board1 → McMaster → familia Board1 → None.

    NO forzar todo ABB/GENE/RLG a Corte: Board1 clasifica muchos GENE
    (BKT/DF/HFM…) como Doblado. Solo se copia lo de Board1.
    """
    if not base:
        return None, "none"

    # 1) exacto
    cls = mapa_b1.get(base)
    if cls:
        return cls, "exact"
    # 1b) exacto tolerando truncado
    for b1_name, b1_cls in mapa_b1.items():
        if not b1_cls:
            continue
        if b1_name.startswith(base) or base.startswith(b1_name):
            if min(len(base), len(b1_name)) >= 12:
                return b1_cls, "prefix_name"

    # 2) código McMaster compartido
    code = _mcmaster_code(base)
    if code and code in by_code:
        return by_code[code], "mcmaster"

    # 3) Almacén por prefijos de comprados (como en Board1)
    bu = base.upper()
    for p in _ALMACEN_EXACT_PREFIXES:
        if bu.startswith(p.upper()):
            return "Almacén", "almacen_prefix"

    # 4) GENE-OP / GENE-GS por sufijo observado en Board1
    if bu.startswith("GENE-OP") or bu.startswith("GENE-GS"):
        if _OP_CORTE_SUFFIX.search(bu):
            return "Corte", "op_suffix_corte"
        if _OP_DOBLADO_SUFFIX.search(bu):
            return "Doblado", "op_suffix_doblado"
        if bu.endswith("-404") or bu.endswith("-403"):
            return "Doblado", "gs_suffix"
        if bu.endswith("-708"):
            return "Corte", "gs_suffix"

    # 5) familias unánimes Board1
    for rx, cls_fam in _FAMILY_RULES:
        if rx.search(base):
            return cls_fam, "family"

    # 6) McMaster genérico → Almacén
    if code or re.match(r"^\d{5}[A-Z]", bu):
        return "Almacén", "mcmaster_default"

    # 7) RG_SWBD / breakers comprados → Almacén
    if bu.startswith(("RG_SWBD", "INEXT_", "NEXT_")):
        return "Almacén", "almacen_prefix"

    return None, "unresolved"


def main():
    inv = conectar_inventor()
    log(f"Inventor {inv.SoftwareVersion.DisplayVersion} | docs={inv.Documents.Count}")
    try:
        log(f"Active: {inv.ActiveDocument.DisplayName}")
    except Exception:
        pass

    board1_raw = _find_top_iam(inv, "9919-Board 1")
    board5_raw = _find_top_iam(inv, "9919-Board 5")
    if board1_raw is None:
        raise SystemExit("No encontré 9919-Board 1.iam abierto en Inventor")
    if board5_raw is None:
        raise SystemExit("No encontré 9919-Board 5.iam abierto en Inventor")

    board1 = _as_assembly(board1_raw)
    board5 = _as_assembly(board5_raw)
    if board1 is None or board5 is None:
        raise SystemExit("No pude castear Board1/Board5 a AssemblyDocument")

    log(f"Board1: {board1.DisplayName} | {board1.FullFileName}")
    log(f"Board5: {board5.DisplayName} | {board5.FullFileName}")

    mapa_b1 = _mapa_desde_asm(board1, "Board1")
    mapa_b5_antes = _mapa_desde_asm(board5, "Board5-ANTES")
    by_code = _build_b1_indexes(mapa_b1)
    log(f"Índice McMaster Board1: {len(by_code)} códigos")

    _ensure_assets(board5, inv)

    # Index asset cache
    asset_by_cls = {c: _resolve_asset(board5, c) for c in VALIDAS}
    for c, a in asset_by_cls.items():
        log(f"  asset[{c}] = {a.DisplayName if a else 'MISSING'}")

    # Aplicar: por archivo único (misma lógica PorArchivo del iLogic)
    by_file = defaultdict(list)  # fullpath -> [(occ, doc, base)]
    for occ, doc, base in _iter_leaf_parts(board5):
        try:
            fp = (doc.FullFileName or "").lower()
        except Exception:
            fp = ""
        if not fp:
            continue
        by_file[fp].append((occ, doc, base))

    stats = Counter()
    method_c = Counter()
    cambios = []
    sin_resolver = []
    docs_a_guardar = {}

    for fp, items in by_file.items():
        base = items[0][2]
        cls_src, method = _resolver_cls(base, mapa_b1, by_code)
        method_c[method] += 1
        if not cls_src:
            stats["sin_resolver"] += 1
            sin_resolver.append(base)
            continue

        cls_actual = _leer_cls(items[0][1])
        asset = asset_by_cls.get(cls_src)
        need_cls = cls_actual != cls_src

        if DRY_RUN or not APPLY:
            stats["dry_would_apply"] += 1
            cambios.append(
                {
                    "pieza": base,
                    "de": cls_actual,
                    "a": cls_src,
                    "method": method,
                    "qty_occs": len(items),
                }
            )
            continue

        doc = items[0][1]
        try:
            if need_cls or cls_actual is None:
                _escribir_cls(doc, cls_src)
                stats["iprop_escrito"] += 1
            else:
                stats["iprop_ya_ok"] += 1
        except Exception as exc:
            stats["iprop_error"] += 1
            log(f"  ERR iProp {base}: {exc}")
            continue

        app_ok = 0
        for occ, part_doc, _ in items:
            if _aplicar_apariencia(occ, part_doc, asset):
                app_ok += 1
        stats["apariencia_ok"] += app_ok
        if app_ok == 0:
            stats["apariencia_fail"] += 1

        # gen_py de Inventor a veces no expone ReadOnly; intentar Save siempre.
        docs_a_guardar[fp] = doc

        cambios.append(
            {
                "pieza": base,
                "de": cls_actual,
                "a": cls_src,
                "method": method,
                "qty_occs": len(items),
                "app_ok": app_ok,
            }
        )
        stats["aplicadas"] += 1

    # Guardar
    saved = 0
    if APPLY and not DRY_RUN:
        for fp, doc in docs_a_guardar.items():
            try:
                ro = False
                try:
                    ro = bool(doc.ReadOnly)
                except Exception:
                    ro = False
                if ro:
                    stats["save_readonly"] += 1
                    continue
                doc.Save()
                saved += 1
            except Exception as exc:
                stats["save_error"] += 1
                log(f"  ERR save {os.path.basename(fp)}: {exc}")
        try:
            board5.Save()
            saved += 1
            log("  Board5.iam guardado")
        except Exception as exc:
            log(f"  ERR save board5.iam: {exc}")
            # Fallback: Save As no; pedir Save vía Documents
            try:
                inv.SilentOperation = True
                board5.Save2(True)  # Save and dependent docs
                saved += 1
                log("  Board5.iam Save2 OK")
            except Exception as exc2:
                log(f"  ERR Save2 board5: {exc2}")

    mapa_b5_despues = (
        _mapa_desde_asm(board5, "Board5-DESPUES")
        if APPLY and not DRY_RUN
        else mapa_b5_antes
    )

    b5_bases = set(mapa_b5_antes)
    classified_b5_after = sum(1 for v in mapa_b5_despues.values() if v)

    resumen = {
        "dry_run": DRY_RUN,
        "apply": APPLY,
        "board1": board1.FullFileName,
        "board5": board5.FullFileName,
        "board1_clasificadas_unicas": sum(1 for v in mapa_b1.values() if v),
        "board5_unicas": len(mapa_b5_antes),
        "methods": dict(method_c),
        "sin_resolver": sorted(sin_resolver),
        "sin_resolver_count": len(sin_resolver),
        "stats": dict(stats),
        "saved_docs": saved,
        "board5_clasificadas_despues": classified_b5_after,
        "cambios_sample": cambios[:80],
        "cambios_total": len(cambios),
        "mapa_b5_despues_por_clase": {
            c: sorted(b for b, v in mapa_b5_despues.items() if v == c)
            for c in VALIDAS
        },
    }
    OUT.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"\n=== RESUMEN ===")
    log(f"Métodos: {dict(method_c)}")
    log(f"Sin resolver: {len(sin_resolver)} → {sin_resolver[:20]}")
    log(f"Stats: {dict(stats)}")
    log(f"Docs guardados: {saved}")
    log(f"Board5 clasificadas después: {classified_b5_after}/{len(b5_bases)}")
    log(f"Log: {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
