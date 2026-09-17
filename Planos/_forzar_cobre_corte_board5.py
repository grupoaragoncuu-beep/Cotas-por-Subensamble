# -*- coding: utf-8 -*-
"""Fuerza colorimetría Corte en todo cobre/busbar (ABB/GENE/RLG) de Board 5.

Regla de producto: lo que el flujo de cotas trata como cobre siempre lleva
iProperty Clasificación = 'Corte' y apariencia Rojo naranja.
"""
from __future__ import annotations

import json
import os
import sys
from collections import Counter
from pathlib import Path

import win32com.client

from inventor_com import conectar_inventor, localizar_documento
from piezas_cobre import es_pieza_cobre, prefijo_cobre

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

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
APARIENCIAS_CORTE = ("Rojo naranja",)
K_PART = 12290
OUT = Path(__file__).with_name("_audit_cobre_corte_board5.json")
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


def _leer_cls(doc):
    try:
        return _norm_cls(doc.PropertySets.Item(PROPSET).Item(PROP_CLS).Value)
    except Exception:
        return None


def _escribir_cls(doc, clasificacion):
    props = doc.PropertySets.Item(PROPSET)
    try:
        props.Item(PROP_CLS).Value = clasificacion
    except Exception:
        props.Add(clasificacion, PROP_CLS)


def _basename(doc):
    try:
        return os.path.splitext(os.path.basename(doc.FullFileName or ""))[0]
    except Exception:
        return ""


def _asset_corte(asm):
    for a in asm.Assets:
        if a.DisplayName in APARIENCIAS_CORTE:
            return a
    # cargar de biblioteca si falta
    return None


def _ensure_rojo(asm, inv):
    asset = _asset_corte(asm)
    if asset is not None:
        return asset
    for lib_name in (
        "Autodesk Biblioteca de aspecto",
        "Autodesk Appearance Library",
    ):
        try:
            lib = inv.AssetLibraries.Item(lib_name)
            a = lib.AppearanceAssets.Item("Rojo naranja")
            a.CopyTo(asm)
            break
        except Exception:
            continue
    return _asset_corte(asm)


def main():
    inv = conectar_inventor()
    raw = localizar_documento(inv, display_name="9919-Board 5.iam")
    if raw is None:
        raise SystemExit("Board 5 no abierto")
    asm = win32com.client.CastTo(raw, "AssemblyDocument")
    asset = _ensure_rojo(asm, inv)
    log(f"Asset Corte: {asset.DisplayName if asset else 'MISSING'}")

    leaves = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
    by_file = {}
    for i in range(1, int(leaves.Count) + 1):
        try:
            occ = leaves.Item(i)
            if occ.Suppressed:
                continue
            doc = occ.Definition.Document
            if int(doc.DocumentType) != K_PART:
                continue
            fp = (doc.FullFileName or "").lower()
            if not fp:
                continue
            name = _basename(doc)
            if not es_pieza_cobre(name):
                continue
            by_file.setdefault(fp, {"doc": doc, "name": name, "occs": []})
            by_file[fp]["occs"].append(occ)
        except Exception:
            continue

    antes = Counter()
    mal = []
    ok = []
    for info in by_file.values():
        cls = _leer_cls(info["doc"])
        antes[cls or "SIN"] += 1
        info["cls"] = cls
        if cls != "Corte":
            mal.append(info)
        else:
            ok.append(info)

    log(f"Cobre únicos ABB/GENE/RLG: {len(by_file)}")
    log(f"Por clasificación ANTES: {dict(antes)}")
    log(f"Ya Corte OK: {len(ok)} | a corregir: {len(mal)}")
    for info in mal[:40]:
        log(f"  MAL {info['name']}: {info['cls']} → Corte")

    cambios = []
    saved = 0
    if APPLY and mal:
        for info in mal:
            doc = win32com.client.CastTo(info["doc"], "PartDocument")
            try:
                _escribir_cls(doc, "Corte")
            except Exception as exc:
                log(f"  ERR iProp {info['name']}: {exc}")
                continue
            app_ok = 0
            if asset is not None:
                for occ in info["occs"]:
                    try:
                        occ.Appearance = asset
                        app_ok += 1
                    except Exception:
                        pass
                try:
                    doc.ComponentDefinition.Appearance = asset
                except Exception:
                    pass
            try:
                doc.Save()
                saved += 1
            except Exception as exc:
                log(f"  ERR save {info['name']}: {exc}")
            cambios.append(
                {
                    "pieza": info["name"],
                    "prefijo": prefijo_cobre(info["name"]),
                    "de": info["cls"],
                    "a": "Corte",
                    "occs": len(info["occs"]),
                    "app_ok": app_ok,
                }
            )
        try:
            asm.Save()
            log("Board5.iam guardado")
        except Exception as exc:
            log(f"ERR asm save: {exc}")

    # Releer verificación
    despues = Counter()
    residual = []
    for info in by_file.values():
        # re-open from dict doc
        cls = _leer_cls(info["doc"])
        despues[cls or "SIN"] += 1
        if cls != "Corte":
            residual.append({"pieza": info["name"], "cls": cls})

    resumen = {
        "apply": APPLY,
        "cobre_unicos": len(by_file),
        "antes": dict(antes),
        "despues": dict(despues),
        "corregidos": len(cambios),
        "saved": saved,
        "residual_no_corte": residual,
        "cambios": cambios,
        "ok_ya_corte_sample": [i["name"] for i in ok[:30]],
    }
    OUT.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"\nDESPUÉS: {dict(despues)}")
    log(f"Residual no-Corte: {len(residual)}")
    log(f"Guardados: {saved} | Log: {OUT}")
    return 0 if not residual else 2


if __name__ == "__main__":
    raise SystemExit(main())
