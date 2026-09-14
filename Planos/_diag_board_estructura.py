# -*- coding: utf-8 -*-
"""
Diagnóstico COM del ensamble abierto (BOARD/GIGA).

Attach a Inventor → clasifica producto/unidades → muestra árbol 1er nivel
y cobertura del filtro de cobre. Solo lectura (no cotas / no JPG).
"""
from __future__ import annotations

import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client
from inventor_com import conectar_inventor
from producto_tipo import clasificar_producto, unidad_para_producto, aplicar_unidad_producto
from piezas_cobre import es_pieza_cobre, prefijo_cobre
from cota_estilo import get_unidad_cota, set_unidad_cota
from generador_caras_tanque import _obtener_ensamble_principal, _como_ensamble


def _base_occ(occ) -> str:
    try:
        doc = occ.Definition.Document
        ff = str(getattr(doc, "FullFileName", "") or "")
        if ff:
            return os.path.splitext(os.path.basename(ff))[0]
    except Exception:
        pass
    try:
        return str(occ.Name or "").split(":", 1)[0]
    except Exception:
        return "?"


def main() -> int:
    pythoncom.CoInitialize()
    prev_u = get_unidad_cota()
    try:
        inv = conectar_inventor()
        print(f"Inventor: OK | docs={inv.Documents.Count}", flush=True)

        ensamble = _obtener_ensamble_principal(inv)
        if ensamble is None:
            try:
                ad = inv.ActiveDocument
                if int(ad.DocumentType) == 12291:
                    ensamble = _como_ensamble(ad)
            except Exception:
                ensamble = None
        if ensamble is None:
            print("ERROR: no hay .iam abierto.")
            return 1
        ensamble = _como_ensamble(ensamble)

        print(f"Ensamble: {ensamble.DisplayName}", flush=True)
        try:
            print(f"Ruta: {ensamble.FullFileName}", flush=True)
        except Exception:
            pass

        info = clasificar_producto(ensamble, escanear_hijos=True)
        u = unidad_para_producto(ensamble=ensamble, info=info)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        print(
            f"Clasificación: tipo={info.get('tipo')} familia={info.get('familia')}",
            flush=True,
        )
        print(f"Motivo: {info.get('motivo')}", flush=True)
        print(f"Unidad cota: {u} (activa={get_unidad_cota()})", flush=True)
        if info.get("senales"):
            print("Señales:", ", ".join(info["senales"][:12]), flush=True)

        occs = ensamble.ComponentDefinition.Occurrences
        n = int(occs.Count)
        print(f"Ocurrencias 1er nivel: {n}", flush=True)

        cobre_ok: list[tuple[str, str]] = []
        no_cobre: list[str] = []
        prefijos = Counter()
        suppressed = 0
        limite = min(n, 400)

        for i in range(1, limite + 1):
            try:
                occ = occs.Item(i)
            except Exception:
                continue
            try:
                if getattr(occ, "Suppressed", False):
                    suppressed += 1
                    continue
            except Exception:
                pass
            base = _base_occ(occ)
            pre = prefijo_cobre(base)
            if es_pieza_cobre(base):
                cobre_ok.append((base, pre or "?"))
                prefijos[pre or "?"] += 1
            else:
                no_cobre.append(base)

        print(f"Muestreadas: {limite} | suppressed: {suppressed}", flush=True)
        print(f"Cobre detectado: {len(cobre_ok)} | no-cobre: {len(no_cobre)}", flush=True)
        print("Prefijos cobre:", dict(prefijos.most_common(20)), flush=True)
        print("\n--- Muestra COBRE (20) ---", flush=True)
        for nom, pre in cobre_ok[:20]:
            print(f"  [{pre}] {nom}", flush=True)
        print("\n--- Muestra NO-cobre (25) ---", flush=True)
        for nom in no_cobre[:25]:
            print(f"  {nom}", flush=True)

        # Candidatos sospechosos: 9919 / GEN / ABB / bus / barra en no-cobre
        sospechosos = [
            x
            for x in no_cobre
            if any(
                t in x.upper()
                for t in (
                    "9919",
                    "GEN",
                    "ABB",
                    "BUS",
                    "BAR",
                    "CU",
                    "COBRE",
                    "COPPER",
                    "FCU",
                    "BCK",
                )
            )
        ]
        print(f"\n--- Sospechosos omitidos por filtro ({len(sospechosos)}) ---", flush=True)
        for nom in sospechosos[:30]:
            print(f"  ? {nom}", flush=True)

        print("\nDIAG_BOARD_OK", flush=True)
        return 0
    except Exception as exc:
        print(f"DIAG_FAIL: {type(exc).__name__}: {exc}", flush=True)
        import traceback

        traceback.print_exc()
        return 2
    finally:
        try:
            set_unidad_cota(prev_u)
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
