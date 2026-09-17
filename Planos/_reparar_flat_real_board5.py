# -*- coding: utf-8 -*-
"""Repara Flat REAL en piezas DOBLADAS del ensamble activo (Board 5).

- No acepta Unfold() clásico como éxito.
- Obliga Unfold2(cara) tras medir espesor.
- Valida: bends>0 Ó flat con espesor ≈ medido y huella coherente.
- Save + Close de cada .ipt (no saturar Inventor).
"""
from __future__ import annotations

import importlib.util
import json
import math
import os
import sys
import time
from collections import Counter
from pathlib import Path

import pythoncom
import win32com.client

from inventor_com import conectar_inventor, localizar_documento
from piezas_cobre import es_pieza_cobre

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

K_PART = 12290
# Inventor 2027: Plane=5890, Cylinder=5891 (el 17921 del generativo viejo YA NO aplica)
K_PLANE_SURFACE = 5890
K_PLANE_SURFACE_LEGACY = 17921
SUB_SM_MODERN = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
SUB_SM_CLASSIC = "{9C464203-7BAE-11D3-8BAD-006008198D01}"
SM_MARKERS = ("9C464203-7BAE", "9C464203-9BAE", "9C464203")

OUT = Path(__file__).with_name("_reparar_flat_real_log.json")
SM_SCRIPT = Path(
    r"Z:\♦♦GRUPO ARGA CARPETAS COMPARTIDAS♦♦\BIENVENIDO"
    r"\Departamentos _antes TIK\8. Ingeniería\HUGO CHAVEZ"
    r"\Diseño Parametrico Generativo\InventorGenerativo"
    r"\_run_sin_grosor_regla.py"
)

PROPSET = "Inventor User Defined Properties"
PROP_CLS = "Clasificación"
CLS_CHAPA = {"Corte", "Doblado", "Plasma", "Plasma Doblado", "Maquinado"}


def log(msg=""):
    print(msg, flush=True)


def _load_sm():
    spec = importlib.util.spec_from_file_location("sm_orig", SM_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def _as_part(doc):
    return win32com.client.CastTo(doc, "PartDocument")


def _as_asm(doc):
    return win32com.client.CastTo(doc, "AssemblyDocument")


def _is_sm(doc) -> bool:
    try:
        st = str(doc.SubType or "").upper().replace("{", "").replace("}", "")
        return any(m.upper() in st for m in SM_MARKERS)
    except Exception:
        return False


def _sm_def(doc):
    return win32com.client.CastTo(
        doc.ComponentDefinition, "SheetMetalComponentDefinition"
    )


def _bbox_cm(body_or_doc):
    try:
        if hasattr(body_or_doc, "RangeBox"):
            rb = body_or_doc.RangeBox
        else:
            rb = body_or_doc.ComponentDefinition.SurfaceBodies.Item(1).RangeBox
        return sorted(
            [
                abs(rb.MaxPoint.X - rb.MinPoint.X),
                abs(rb.MaxPoint.Y - rb.MinPoint.Y),
                abs(rb.MaxPoint.Z - rb.MinPoint.Z),
            ]
        )
    except Exception:
        return None


def _leer_cls(doc):
    try:
        v = str(doc.PropertySets.Item(PROPSET).Item(PROP_CLS).Value).strip()
        for c in CLS_CHAPA | {"Almacén"}:
            if v.casefold() == c.casefold():
                return c
    except Exception:
        pass
    return None


def _es_candidata(nombre: str, cls: str | None) -> bool:
    """Piezas de chapa/cobre (no tornillería Almacén genérica)."""
    if es_pieza_cobre(nombre):
        return True
    if cls in CLS_CHAPA:
        return True
    u = nombre.upper()
    if u.startswith(("FB-", "GEN1-")):
        return True
    return False


def _parece_doblada(dims) -> bool:
    """True si el sólido no es una placa casi plana (candidato a unfold real)."""
    if not dims or len(dims) < 3:
        return False
    a, b, c = dims  # a<=b<=c
    if a <= 1e-6:
        return False
    # Placa: espesor mucho menor que las otras dos
    if a < 0.6 and b / max(a, 1e-9) > 8:
        return False
    if a < 1.2 and b / max(a, 1e-9) > 12:
        return False
    # Dos dimensiones "gruesas" → típicamente L/U/bracket
    return b / max(a, 1e-9) < 15 and a > 0.25


def _count_bends(smd) -> int:
    try:
        return int(smd.FlatPattern.FlatBendResults.Count)
    except Exception:
        pass
    try:
        return int(smd.FlatPattern.FlatPatternBends.Count)
    except Exception:
        return 0


def _flat_es_real(smd, espesor_cm: float, fold_dims) -> tuple[bool, str]:
    if not smd.HasFlatPattern:
        return False, "sin_HasFlatPattern"
    bends = _count_bends(smd)
    try:
        flat_dims = _bbox_cm(smd.FlatPattern.Body)
    except Exception:
        flat_dims = None
    if bends > 0:
        return True, f"bends={bends}"
    # Sin bends: solo OK si flat luce como placa del espesor medido
    if flat_dims and espesor_cm > 0:
        fa, fb, fc = flat_dims
        # espesor flat ≈ medido
        if abs(fa - espesor_cm) / max(espesor_cm, 1e-6) < 0.35:
            # huella plana >= huella plegada (algo se abrió) o ya era placa
            fold_area = (fold_dims[1] * fold_dims[2]) if fold_dims else 0
            flat_area = fb * fc
            if flat_area >= fold_area * 0.95:
                # si era doblada (fold_dims[0] >> espesor), exigir crecimiento
                if fold_dims and fold_dims[0] > espesor_cm * 1.8:
                    if flat_area < fold_area * 1.05 and bends == 0:
                        return False, f"fake_rotacion flat={flat_dims} fold={fold_dims}"
                return True, f"placa_espesor ok flat={flat_dims}"
        return False, f"fake bends=0 flat={flat_dims} thk={espesor_cm:.4f}"
    return False, "fake_sin_metricas"


def _open_part(inv, fp: str):
    doc = inv.Documents.Open(fp, True)
    return _as_part(doc)


def _ensure_sm(inv, part):
    if _is_sm(part):
        return _sm_def(part), part
    part = _open_part(inv, part.FullFileName)
    try:
        cmd = inv.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
        cmd.Execute()
        part.Update()
    except Exception:
        pass
    if not _is_sm(part):
        for g in (SUB_SM_MODERN, SUB_SM_CLASSIC):
            try:
                part.SubType = g
                part.Update()
                if _is_sm(part):
                    break
            except Exception:
                continue
    if not _is_sm(part):
        raise RuntimeError("no SM")
    return _sm_def(part), part


def _dists_planos_paralelos(smd) -> list[tuple[float, int]]:
    """[(espesor_cm, votos)] ordenado: más votos, luego más fino."""
    body = smd.SurfaceBodies.Item(1)
    planes = []
    for i in range(1, body.Faces.Count + 1):
        face = body.Faces.Item(i)
        try:
            if not _es_cara_plana(face):
                continue
            g = face.Geometry
            n = g.Normal
            pt = face.PointOnFace
            area = float(face.Evaluator.Area)
            planes.append((n.X, n.Y, n.Z, pt.X, pt.Y, pt.Z, area))
        except Exception:
            continue
    dists = []
    for i, a in enumerate(planes):
        ax, ay, az, px, py, pz, aa = a
        for b in planes[i + 1 :]:
            bx, by, bz, qx, qy, qz, ba = b
            dot = ax * bx + ay * by + az * bz
            if abs(dot + 1.0) > 0.05:
                continue
            dist = abs((qx - px) * ax + (qy - py) * ay + (qz - pz) * az)
            if 0.08 <= dist <= 2.6 and min(aa, ba) > 0.15:
                dists.append(round(dist, 4))
    if not dists:
        return []
    bucket = Counter(round(d, 3) for d in dists)
    return sorted(bucket.items(), key=lambda kv: (-kv[1], kv[0]))


def _candidatos_espesor(sm_mod, smd, inv, part) -> list[float]:
    """Varios espesores a probar (chapa real ≠ moda de huecos del U)."""
    out: list[float] = []
    for d, _n in _dists_planos_paralelos(smd)[:8]:
        out.append(float(d))
    try:
        out.append(float(sm_mod.get_thick_from_max_face(smd.SurfaceBodies.Item(1), inv)))
    except Exception:
        pass
    try:
        out.append(float(smd.Thickness.Value))
    except Exception:
        pass
    dims = _bbox_cm(part)
    if dims:
        out.append(float(dims[0]))
    # calibres cobre típicos (cm)
    out.extend([0.3175, 0.47625, 0.635, 0.9525, 1.27, 0.25, 0.3, 0.6, 0.9])
    seen = set()
    uniq: list[float] = []
    for t in sorted(out):
        k = round(t, 3)
        if k in seen or k < 0.08 or k > 2.6:
            continue
        seen.add(k)
        uniq.append(float(t))
    return uniq


def _aplicar_regla(sm_mod, part, smd, espesor_cm: float):
    espesor_in = espesor_cm * 0.3937007874
    regla = ""
    unfold = ""
    for r in sm_mod.CATALOGO:
        if r.minimo <= espesor_in <= r.maximo:
            regla, unfold = r.regla_sm, r.regla_unfold
            break
    try:
        smd.UseSheetMetalStyleThickness = False
        smd.Thickness.Value = espesor_cm
        try:
            smd.BendRadius.Value = max(espesor_cm, 0.1)
        except Exception:
            pass
        part.Update()
    except Exception:
        pass
    if regla:
        try:
            smd = sm_mod.aplicar_regla_chapa_interno(part, smd, regla, espesor_cm)
            smd.UseSheetMetalStyleThickness = False
            smd.Thickness.Value = espesor_cm
            part.Update()
        except Exception:
            pass
    return smd, regla, unfold


def _es_cara_plana(face) -> bool:
    try:
        st = int(face.SurfaceType)
        if st in (K_PLANE_SURFACE, K_PLANE_SURFACE_LEGACY):
            return True
    except Exception:
        pass
    try:
        return "Plane" in type(face.Geometry).__name__
    except Exception:
        return False


def _caras_planas_por_area(smd, max_n=10):
    body = smd.SurfaceBodies.Item(1)
    lista = []
    for i in range(1, body.Faces.Count + 1):
        face = body.Faces.Item(i)
        try:
            if not _es_cara_plana(face):
                continue
            area = float(face.Evaluator.Area)
            if area > 0:
                lista.append((area, face))
        except Exception:
            continue
    lista.sort(key=lambda x: -x[0])
    return [f for _, f in lista[:max_n]]


def _unfold2_real(sm_mod, part, smd, espesor_cm: float, unfold_name: str) -> tuple[bool, str, int]:
    """Devuelve (ok_has_flat, how, bends). Solo ok útil si bends>0."""
    try:
        part.Activate()
    except Exception:
        pass

    # limpiar flat basura
    try:
        if smd.HasFlatPattern:
            smd.FlatPattern.Delete()
        while smd.ASideDefinitions.Count > 0:
            smd.ASideDefinitions.Item(1).Delete()
        part.Update()
    except Exception:
        pass

    try:
        e = max(espesor_cm, 0.01)
        smd.BendReliefWidth.Value = e
        smd.BendReliefDepth.Value = e * 1.5
        smd.CornerReliefSize.Value = e
    except Exception:
        pass

    um = sm_mod.obtener_unfold_method(smd, unfold_name)
    if um is None:
        um = sm_mod.obtener_o_crear_unfold_factor_k(smd)
    if um is not None:
        try:
            smd.UnfoldMethod = um
            part.Update()
        except Exception:
            pass

    caras = _caras_planas_por_area(smd, 20)
    if not caras:
        return False, "sin_caras_planas_5890", 0

    last = "sin_ok"
    best_bends = 0
    for cara in caras:
        try:
            if smd.HasFlatPattern:
                smd.FlatPattern.Delete()
            smd.Unfold2(cara)
            if smd.HasFlatPattern:
                try:
                    smd.FlatPattern.ExitEdit()
                except Exception:
                    pass
                part.Update()
                bends = _count_bends(smd)
                if bends > best_bends:
                    best_bends = bends
                if bends > 0:
                    return True, "Unfold2_plane", bends
                # flat fake → borrar y seguir
                try:
                    smd.FlatPattern.Delete()
                except Exception:
                    pass
                last = "Unfold2_fake_0bends"
        except Exception as ex:
            last = str(ex)
            try:
                if smd.HasFlatPattern:
                    smd.FlatPattern.Delete()
            except Exception:
                pass

    # Factor K + Unfold2
    try:
        k = sm_mod.obtener_o_crear_unfold_factor_k(smd)
        if k is not None:
            smd.UnfoldMethod = k
            part.Update()
        for cara in caras[:8]:
            if smd.HasFlatPattern:
                smd.FlatPattern.Delete()
            smd.Unfold2(cara)
            if smd.HasFlatPattern:
                try:
                    smd.FlatPattern.ExitEdit()
                except Exception:
                    pass
                part.Update()
                bends = _count_bends(smd)
                if bends > 0:
                    return True, "Unfold2_K_plane", bends
                try:
                    smd.FlatPattern.Delete()
                except Exception:
                    pass
    except Exception as ex:
        last = str(ex)

    return False, last, best_bends


def _borrar_flat_si_fake(smd, part):
    try:
        if not smd.HasFlatPattern:
            return
        if _count_bends(smd) > 0:
            return
        smd.FlatPattern.Delete()
        part.Update()
    except Exception:
        pass


def _close_part(inv, fp: str):
    fp_l = (fp or "").lower()
    for di in range(inv.Documents.Count, 0, -1):
        try:
            d = inv.Documents.Item(di)
            if (d.FullFileName or "").lower() == fp_l and int(d.DocumentType) == K_PART:
                d.Close(True)  # SkipSave: ya guardamos
                return True
        except Exception:
            continue
    return False


def _cerrar_extras(inv, asm_fp: str):
    """Evita acumulación de .ipt abiertos (crash)."""
    asm_l = (asm_fp or "").lower()
    for di in range(inv.Documents.Count, 0, -1):
        try:
            d = inv.Documents.Item(di)
            fp = (d.FullFileName or "").lower()
            if fp == asm_l:
                continue
            if int(d.DocumentType) == K_PART:
                d.Close(True)
        except Exception:
            continue


def main() -> int:
    sm_mod = _load_sm()
    inv = conectar_inventor()
    try:
        inv.SilentOperation = True
    except Exception:
        pass

    raw = localizar_documento(inv, display_name="9919-Board 5.iam") or inv.ActiveDocument
    asm = _as_asm(raw)
    asm_path = asm.FullFileName
    log(f"Ensamble: {asm.DisplayName}")
    _cerrar_extras(inv, asm_path)

    leaves = asm.ComponentDefinition.Occurrences.AllLeafOccurrences
    by_fp = {}
    for i in range(1, int(leaves.Count) + 1):
        try:
            occ = leaves.Item(i)
            if occ.Suppressed:
                continue
            doc = occ.Definition.Document
            if int(doc.DocumentType) != K_PART:
                continue
            fp = (doc.FullFileName or "").lower()
            if not fp or fp in by_fp:
                continue
            name = os.path.splitext(os.path.basename(fp))[0]
            cls = _leer_cls(doc)
            if not _es_candidata(name, cls):
                continue
            dims = _bbox_cm(doc)
            if not _parece_doblada(dims):
                continue
            by_fp[fp] = {"name": name, "cls": cls, "dims": dims}
        except Exception:
            continue

    log(f"Candidatas DOBLADAS a reparar: {len(by_fp)}")
    stats = Counter()
    filas = []

    for n_done, (fp, info) in enumerate(by_fp.items(), 1):
        nombre = info["name"]
        try:
            part = _open_part(inv, fp)
            try:
                part.Activate()
            except Exception:
                pass
            fold_dims = _bbox_cm(part) or info["dims"]
            smd, part = _ensure_sm(inv, part)

            # Si ya tiene flat REAL, skip
            espesor0 = None
            cands = _candidatos_espesor(sm_mod, smd, inv, part)
            if cands:
                espesor0 = cands[0]
            ok0, why0 = _flat_es_real(smd, espesor0 or 0.3, fold_dims)
            if ok0 and _count_bends(smd) > 0:
                stats["ya_ok"] += 1
                log(f"[{n_done}/{len(by_fp)}] SKIP real {nombre} ({why0})")
                _close_part(inv, fp)
                continue

            ok = False
            why = "sin_intento"
            how = ""
            espesor = cands[0] if cands else 0.3
            regla = ""

            for thk in (cands or [0.3])[:10]:
                smd, regla, unfold = _aplicar_regla(sm_mod, part, smd, thk)
                ok_u, how, bends = _unfold2_real(sm_mod, part, smd, thk, unfold)
                smd = _sm_def(part)
                if bends > 0:
                    espesor = thk
                    ok, why = True, f"bends={bends}"
                    break
                _borrar_flat_si_fake(smd, part)
                why = f"fake/no bend via={how}"

            if ok:
                stats["ok"] += 1
                log(
                    f"[{n_done}/{len(by_fp)}] OK {nombre} thk={espesor:.4f}cm "
                    f"regla={regla or '-'} via={how} | {why}"
                )
                try:
                    part.Save()
                    stats["saved"] += 1
                except Exception as ex:
                    stats["save_err"] += 1
                    log(f"  save err: {ex}")
            else:
                stats["fail"] += 1
                # NUNCA guardar flat falso; sí dejar SM + espesor medido
                _borrar_flat_si_fake(smd, part)
                try:
                    if cands:
                        smd.UseSheetMetalStyleThickness = False
                        smd.Thickness.Value = float(cands[0])
                        part.Update()
                    part.Save()
                    stats["saved_sm_sin_flat"] += 1
                except Exception as ex:
                    stats["save_err"] += 1
                    log(f"  save err: {ex}")
                log(
                    f"[{n_done}/{len(by_fp)}] FAIL {nombre} thk={espesor:.4f}cm "
                    f"| {why} (SM guardado, flat fake borrado)"
                )

            filas.append(
                {
                    "pieza": nombre,
                    "espesor_cm": espesor,
                    "regla": regla,
                    "ok": ok,
                    "why": why,
                    "how": how,
                }
            )
            _close_part(inv, fp)
            if n_done % 5 == 0:
                _cerrar_extras(inv, asm_path)
                try:
                    inv.Documents.Open(asm_path, True)
                except Exception:
                    pass
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass

            if n_done % 10 == 0:
                OUT.write_text(
                    json.dumps(
                        {"stats": dict(stats), "filas": filas},
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
        except Exception as ex:
            stats["fail"] += 1
            log(f"[{n_done}/{len(by_fp)}] ERR {nombre}: {ex}")
            try:
                _close_part(inv, fp)
            except Exception:
                pass

    try:
        _cerrar_extras(inv, asm_path)
        inv.Documents.Open(asm_path, True)
        _as_asm(inv.ActiveDocument).Save()
        log("Ensamble guardado")
    except Exception as ex:
        log(f"Asm save: {ex}")

    OUT.write_text(
        json.dumps(
            {"candidatas": len(by_fp), "stats": dict(stats), "filas": filas},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    log(f"FIN stats={dict(stats)} log={OUT}")
    return 0 if stats["fail"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
