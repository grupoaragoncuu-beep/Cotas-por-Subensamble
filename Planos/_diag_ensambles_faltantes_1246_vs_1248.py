# -*- coding: utf-8 -*-
"""Compara roots 62223-1246-A01 vs 62223-1248-A01; lista faltantes excl. TOP."""
import re

import pythoncom
import win32com.client

pythoncom.CoInitialize()
inv = win32com.client.Dispatch("Inventor.Application")

docs = {}
for i in range(1, inv.Documents.Count + 1):
    d = inv.Documents.Item(i)
    if d.DocumentType != 12291:
        continue
    key = d.DisplayName.replace(".iam", "").replace(".IAM", "").strip()
    if key == "62223-1246-A01":
        docs["tank"] = win32com.client.CastTo(d, "AssemblyDocument")
    elif key == "62223-1248-A01":
        docs["shell"] = win32com.client.CastTo(d, "AssemblyDocument")

print("ACTIVE:", inv.ActiveDocument.DisplayName)
print("TANK:", docs["tank"].FullFileName)
print("SHELL:", docs["shell"].FullFileName)


def root_occs(asm):
    out = []
    for i in range(1, asm.ComponentDefinition.Occurrences.Count + 1):
        o = asm.ComponentDefinition.Occurrences.Item(i)
        try:
            dt = o.DefinitionDocumentType
        except Exception:
            dt = None
        base = o.Name.split(":")[0]
        code = re.sub(r"_\d+$", "", base)
        out.append(
            {
                "name": o.Name,
                "base": base,
                "code": code,
                "dtype": dt,
            }
        )
    return out


def es_top_cover(code: str) -> bool:
    u = code.upper()
    if any(x in u for x in ("COVER", "TAPA", "TOP COVER", "ROOF")):
        return True
    m = re.match(r"(\d+)-(\d+)-A(\d+)", u)
    if not m:
        return False
    fam = m.group(2)
    # OTC: familia *47 = top cover
    return fam.endswith("47") or fam[-2:] == "47"


tank_roots = root_occs(docs["tank"])
shell_roots = root_occs(docs["shell"])

print("\n=== ROOT TANK 1246-A01 (%d) ===" % len(tank_roots))
for r in tank_roots:
    kind = {12291: "IAM", 12290: "IPT"}.get(r["dtype"], str(r["dtype"]))
    print("  %-48s %s" % (r["name"], kind))

print("\n=== ROOT SHELL 1248-A01 (%d) ===" % len(shell_roots))
for r in shell_roots:
    kind = {12291: "IAM", 12290: "IPT"}.get(r["dtype"], str(r["dtype"]))
    print("  %-48s %s" % (r["name"], kind))

shell_codes = {r["code"].upper() for r in shell_roots}
# También códigos de todo el árbol del shell (por si el casco anida lo mismo)
shell_all_codes = set(shell_codes)


def walk_codes(occ, into, depth=0, max_depth=6):
    base = occ.Name.split(":")[0]
    code = re.sub(r"_\d+$", "", base).upper()
    into.add(code)
    if depth >= max_depth:
        return
    try:
        if occ.DefinitionDocumentType != 12291:
            return
        for j in range(1, occ.SubOccurrences.Count + 1):
            walk_codes(occ.SubOccurrences.Item(j), into, depth + 1, max_depth)
    except Exception:
        return


for i in range(1, docs["shell"].ComponentDefinition.Occurrences.Count + 1):
    walk_codes(docs["shell"].ComponentDefinition.Occurrences.Item(i), shell_all_codes)

print("\n=== ENSAMBLES IAM en root 1246 ausentes del root 1248 (excl. TOP) ===")
faltan_iam = []
top_skipped = []
for r in tank_roots:
    if r["dtype"] != 12291:
        continue
    code = r["code"].upper()
    if code in shell_codes or code.startswith("62223-1248-A01"):
        continue
    if es_top_cover(code):
        top_skipped.append(r["name"])
        continue
    faltan_iam.append(r)
    print(" ", r["name"])

print("TOP excluidos:", top_skipped)

print(
    "\n=== IAM root 1246 ausentes también del ÁRBOL 1248 (excl. TOP) ==="
)
faltan_deep = []
for r in faltan_iam:
    if r["code"].upper() in shell_all_codes:
        continue
    faltan_deep.append(r)
    print(" ", r["name"])

print("\n=== IPT root 1246 ausentes del root 1248 (excl. TOP) ===")
for r in tank_roots:
    if r["dtype"] != 12290:
        continue
    code = r["code"].upper()
    if code in shell_codes:
        continue
    if es_top_cover(code):
        continue
    print(" ", r["name"])

print(
    "\nRESUMEN: IAM independientes candidatos (root 1246 \\ arbol 1248, sin TOP) = %d"
    % len(faltan_deep)
)
for r in faltan_deep:
    print(" -", r["name"])
