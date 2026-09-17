# -*- coding: utf-8 -*-
"""Guarda piezas Board 5 que ya tienen iProperty Clasificación (post colorimetría)."""
from __future__ import annotations

import os
import sys

import win32com.client

from inventor_com import conectar_inventor, localizar_documento

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PROPSET = "Inventor User Defined Properties"
PROP_CLS = "Clasificación"
VALIDAS = {
    "Almacén",
    "Corte",
    "Maquinado",
    "Doblado",
    "Plasma",
    "Plasma Doblado",
}
K_PART = 12290


def log(msg=""):
    print(msg, flush=True)


def main():
    inv = conectar_inventor()
    raw = localizar_documento(inv, display_name="9919-Board 5.iam")
    if raw is None:
        raise SystemExit("Board 5 no abierto")
    asm = win32com.client.CastTo(raw, "AssemblyDocument")
    leaves = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
    seen = set()
    saved = 0
    errors = 0
    classified = 0
    for i in range(1, int(leaves.Count) + 1):
        try:
            occ = leaves.Item(i)
            if occ.Suppressed:
                continue
            doc = occ.Definition.Document
            if int(doc.DocumentType) != K_PART:
                continue
            fp = (doc.FullFileName or "").lower()
            if not fp or fp in seen:
                continue
            seen.add(fp)
            try:
                val = str(doc.PropertySets.Item(PROPSET).Item(PROP_CLS).Value).strip()
            except Exception:
                continue
            if not any(val.casefold() == v.casefold() for v in VALIDAS):
                continue
            classified += 1
            part = win32com.client.CastTo(doc, "PartDocument")
            try:
                part.Save()
                saved += 1
            except Exception as exc:
                errors += 1
                if errors <= 15:
                    log(f"  ERR {os.path.basename(fp)}: {exc}")
        except Exception:
            continue
    try:
        asm.Save()
        log("Board5.iam Save OK")
    except Exception as exc:
        log(f"Board5.iam Save ERR: {exc}")
        try:
            asm.Save2(True)
            log("Board5.iam Save2 OK")
        except Exception as exc2:
            log(f"Board5.iam Save2 ERR: {exc2}")
    log(f"Clasificadas vistas={classified} | guardadas={saved} | errores={errors}")
    return 0 if errors == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
