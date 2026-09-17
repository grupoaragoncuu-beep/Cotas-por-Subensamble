# -*- coding: utf-8 -*-
"""Sheet Metal + Flat Pattern para TODAS las piezas del ensamble activo.

Inventor 2027 usa SubType chapa distinto al clásico 7BAE; además Activate()
falla en muchas piezas abiertas desde ensamble → se usa Documents.Open.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import time
from collections import Counter
from pathlib import Path

import pythoncom
import win32com.client

from inventor_com import conectar_inventor

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

K_PART = 12290
K_ASSEMBLY = 12291

# GUIDs conocidos de Sheet Metal (2020…2027)
_SM_GUID_MARKERS = (
    "9C464203-7BAE-11D3-8BAD-006008198D01",
    "9C464203-9BAE-11D3-8BAD-0060B0CE6BB4",
    "9C464203",  # familia sheet metal
)

OUT = Path(__file__).with_name("_sheet_metal_board_log.json")
SM_SCRIPT = Path(
    r"Z:\♦♦GRUPO ARGA CARPETAS COMPARTIDAS♦♦\BIENVENIDO"
    r"\Departamentos _antes TIK\8. Ingeniería\HUGO CHAVEZ"
    r"\Diseño Parametrico Generativo\InventorGenerativo"
    r"\_run_sin_grosor_regla.py"
)

# Reabrir ensamble cada N piezas únicas (evita perder contexto / crash)
REACTIVATE_EVERY = 8
PROGRESS = Path(__file__).with_name("_sheet_metal_board_progress.json")


def log(msg=""):
    print(msg, flush=True)


def _has_flat(part_doc) -> bool:
    try:
        smd = _sm_def(part_doc)
        return bool(smd.HasFlatPattern)
    except Exception:
        return False


def _write_progress(payload: dict) -> None:
    try:
        PROGRESS.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def _load_sm():
    spec = importlib.util.spec_from_file_location("run_sin_grosor_regla", SM_SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


def _as_assembly(doc):
    return win32com.client.CastTo(doc, "AssemblyDocument")


def _as_part(doc):
    return win32com.client.CastTo(doc, "PartDocument")


def _is_sheet_metal(part_doc) -> bool:
    try:
        st = str(part_doc.SubType or "").upper().replace("{", "").replace("}", "")
        for m in _SM_GUID_MARKERS:
            if m.upper() in st:
                return True
    except Exception:
        pass
    try:
        _ = part_doc.ComponentDefinition.SheetMetalStyles
        return True
    except Exception:
        return False


def _sm_def(part_doc):
    cdef = part_doc.ComponentDefinition
    try:
        return win32com.client.CastTo(cdef, "SheetMetalComponentDefinition")
    except Exception:
        return cdef


def _activar_pieza(inv, part_doc):
    """Activa la pieza de forma fiable (Open > Activate)."""
    fp = part_doc.FullFileName
    if not fp:
        raise RuntimeError("Pieza sin FullFileName")
    # Preferir Open(visible)
    try:
        doc = inv.Documents.Open(fp, True)
        return _as_part(doc)
    except Exception:
        pass
    try:
        part_doc.Activate()
        return part_doc
    except Exception as ex:
        raise RuntimeError(f"No se pudo activar pieza: {ex}") from ex


def _convertir_a_sheet_metal(inv, part_doc):
    part_doc = _as_part(part_doc)
    if _is_sheet_metal(part_doc):
        return _sm_def(part_doc), part_doc

    part_doc = _activar_pieza(inv, part_doc)

    # 1) comando UI
    try:
        cmd = inv.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
        cmd.Execute()
        try:
            part_doc.Update()
        except Exception:
            pass
    except Exception:
        pass

    if _is_sheet_metal(part_doc):
        return _sm_def(part_doc), part_doc

    # 2) SubType GUID moderno / clásico
    for guid in (
        "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}",
        "{9C464203-7BAE-11D3-8BAD-006008198D01}",
    ):
        try:
            part_doc.SubType = guid
            part_doc.Update()
            if _is_sheet_metal(part_doc):
                return _sm_def(part_doc), part_doc
        except Exception:
            continue

    raise RuntimeError(
        f"No convertida a chapa (SubType={getattr(part_doc, 'SubType', '?')})"
    )


def _basename(doc):
    try:
        return os.path.splitext(os.path.basename(doc.FullFileName or ""))[0]
    except Exception:
        return "?"


def main() -> int:
    sm = _load_sm()
    # Parchear detección/conversión del módulo original
    sm.is_sheet_metal = _is_sheet_metal
    sm.convertir_a_sheet_metal = lambda part_doc, inv_app, forzar_reconvertir=False: (
        _convertir_a_sheet_metal(inv_app, part_doc)[0]
    )

    inv = conectar_inventor()
    activo = inv.ActiveDocument
    if activo is None or int(activo.DocumentType) != K_ASSEMBLY:
        from inventor_com import localizar_documento

        activo = localizar_documento(inv, display_name="9919-Board 5.iam")
        if activo is None:
            raise SystemExit("Activa 9919-Board 5.iam (u otro ensamble).")

    asm = _as_assembly(activo)
    asm_path = asm.FullFileName
    log(f"Ensamble: {asm.DisplayName}")
    log(f"Ruta: {asm_path}")

    try:
        inv.SilentOperation = False  # el comando chapa necesita UI
    except Exception:
        pass

    todas = []
    sm.recorrer_ocurrencias(asm.ComponentDefinition.Occurrences, todas)
    log(f"Ocurrencias: {len(todas)}")
    log("===== SHEET METAL RESUME (salta SM+Flat OK) + SAVE =====")

    docs_ya: set[str] = set()
    stats = Counter()
    filas = []
    saved = 0
    unicas_ok = 0
    unicas_skip_ok = 0
    procesadas_nuevas = 0

    for idx, occ in enumerate(todas, 1):
        try:
            if occ.Suppressed:
                stats["suppressed"] += 1
                continue
            try:
                part0 = occ.Definition.Document
                if int(part0.DocumentType) != K_PART:
                    stats["no_part"] += 1
                    continue
                key = (part0.FullFileName or part0.DisplayName or "").lower()
            except Exception as ex:
                stats["fail"] += 1
                log(f"-> [{idx}] occ err: {ex}")
                continue

            if not key:
                stats["fail"] += 1
                continue
            if key in docs_ya:
                stats["dup"] += 1
                continue
            docs_ya.add(key)

            nombre = _basename(part0)

            # Resume: ya SM + Flat → no tocar
            try:
                if _is_sheet_metal(part0) and _has_flat(part0):
                    unicas_skip_ok += 1
                    stats["skip_ya_ok"] += 1
                    if unicas_skip_ok <= 5 or unicas_skip_ok % 20 == 0:
                        log(f"-> skip OK SM+Flat: {nombre}")
                    continue
            except Exception:
                pass

            # SM sin flat → solo intentar desplegar vía procesar_pieza
            local_seen: set[str] = set()
            res = sm.procesar_pieza(occ, inv, local_seen)
            if res is None:
                stats["skip"] += 1
                continue
            _n, espesor, regla_sm, estado = res
            procesadas_nuevas += 1
            log(
                f"-> [new {procesadas_nuevas} | uniq {len(docs_ya)} | "
                f"occ {idx}/{len(todas)}] "
                f"{nombre:<32} | {espesor:<12} | {regla_sm:<28} | {estado}"
            )
            filas.append(
                {
                    "pieza": nombre,
                    "espesor": espesor,
                    "regla": regla_sm,
                    "estado": estado,
                }
            )
            okish = (
                str(estado).startswith("✔")
                or "Desplegado" in str(estado)
                or "Optimizado" in str(estado)
            )
            if okish:
                stats["ok"] += 1
                unicas_ok += 1
            elif str(estado).startswith("⚠"):
                stats["warn"] += 1
            else:
                stats["fail"] += 1

            # Verificar Flat; si es SM sin flat, reintentar unfold
            try:
                part_chk = _as_part(occ.Definition.Document)
                if _is_sheet_metal(part_chk) and not _has_flat(part_chk):
                    log(f"   ! sin Flat → reintento unfold: {nombre}")
                    part_chk = _activar_pieza(inv, part_chk)
                    smd = _sm_def(part_chk)
                    try:
                        estado2 = sm.intentar_desplegar_a_flat(
                            part_chk, smd, "", 0.1
                        )
                        log(f"   reintento: {estado2}")
                        if (
                            "Desplegado" in str(estado2)
                            or str(estado2).startswith("✔")
                        ):
                            estado = estado2
                            if not okish:
                                stats["ok"] += 1
                                unicas_ok += 1
                                stats["fail"] = max(0, stats["fail"] - 1)
                    except Exception as exu:
                        log(f"   reintento fail: {exu}")
            except Exception as exv:
                log(f"   verify flat err: {exv}")

            # Guardar y CERRAR la pieza (evitar saturar Inventor / crash)
            part_fp_close = None
            try:
                part_now = occ.Definition.Document
                part_now = _as_part(part_now)
                try:
                    part_fp_close = part_now.FullFileName
                except Exception:
                    part_fp_close = None
                part_now.Save()
                saved += 1
            except Exception as exc:
                stats["save_err"] += 1
                if stats["save_err"] <= 25:
                    log(f"   ! save {nombre}: {exc}")
            try:
                # Close(SkipSave=True): ya guardamos
                closed_ok = False
                if part_fp_close:
                    for di in range(inv.Documents.Count, 0, -1):
                        try:
                            dclose = inv.Documents.Item(di)
                            if (dclose.FullFileName or "").lower() == part_fp_close.lower():
                                if int(dclose.DocumentType) == K_PART:
                                    dclose.Close(True)
                                    closed_ok = True
                                break
                        except Exception:
                            continue
                if not closed_ok:
                    try:
                        part_now.Close(True)
                        closed_ok = True
                    except Exception:
                        pass
                if closed_ok:
                    stats["closed"] += 1
                else:
                    stats["close_err"] += 1
            except Exception as exc:
                stats["close_err"] += 1
                if stats["close_err"] <= 15:
                    log(f"   ! close {nombre}: {exc}")

            # Siempre volver al ensamble tras cerrar la pieza
            try:
                inv.Documents.Open(asm_path, True)
            except Exception:
                try:
                    asm.Activate()
                except Exception:
                    pass
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(0.08)

            # Progress disco (para retomar tras crash)
            if procesadas_nuevas % 5 == 0:
                _write_progress(
                    {
                        "procesadas_nuevas": procesadas_nuevas,
                        "unicas_ok_nuevas": unicas_ok,
                        "skip_ya_ok": unicas_skip_ok,
                        "stats": dict(stats),
                        "last": nombre,
                        "filas_tail": filas[-10:],
                    }
                )

        except Exception as ex:
            stats["fail"] += 1
            log(f"-> [{idx}] X {ex}")
            # Intentar volver al ensamble tras error
            try:
                inv.Documents.Open(asm_path, True)
            except Exception:
                pass

    # Volver al ensamble y guardar
    try:
        inv.Documents.Open(asm_path, True)
        asm = _as_assembly(inv.ActiveDocument)
        asm.Save()
        log("Ensamble guardado OK")
    except Exception as exc:
        log(f"Ensamble save ERR: {exc}")

    resumen = {
        "ensamble": asm_path,
        "ocurrencias": len(todas),
        "unicas_vistas": len(docs_ya),
        "skip_ya_ok": unicas_skip_ok,
        "procesadas_nuevas": procesadas_nuevas,
        "unicas_ok_nuevas": unicas_ok,
        "stats": dict(stats),
        "saved_ipt": saved,
        "filas": filas,
    }
    OUT.write_text(json.dumps(resumen, ensure_ascii=False, indent=2), encoding="utf-8")
    _write_progress(resumen)
    log("\n===== FIN =====")
    log(
        f"skip_ya_ok={unicas_skip_ok} nuevas={procesadas_nuevas} "
        f"ok={unicas_ok} fail={stats['fail']} warn={stats['warn']} "
        f"saved={saved} closed={stats['closed']}"
    )
    log(f"Log: {OUT}")
    return 0 if stats["fail"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
