# -*- coding: utf-8 -*-
"""
Diagnóstico COM: geometría de piezas cobre GIGA (ABB/GENE/RLG)
vs heurísticas del flujo Abigail (frente=área máx, LADO=área mín ortogonal).

Solo lectura. No crea vistas ni cotas.
"""
from __future__ import annotations

import math
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

from inventor_com import conectar_inventor
from generador_caras_tanque import _obtener_ensamble_principal, _como_ensamble
from producto_tipo import clasificar_producto, aplicar_unidad_producto
from piezas_cobre import es_pieza_cobre, prefijo_cobre
from cota_estilo import set_unidad_cota, get_unidad_cota
from creador_vistas import (
    set_nombre_pieza_completo,
    obtener_nombre_base_corto,
)

CM_TO_MM = 10.0
K_PLANE = 11058  # kPlaneSurface
SM_GUID = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

MAX_MUESTRA = int(os.environ.get("DIAG_COBRE_N", "45"))


def _cast_part(doc):
    try:
        return win32com.client.CastTo(doc, "PartDocument")
    except Exception:
        return doc


def _bbox_mm(part_doc):
    try:
        rg = part_doc.ComponentDefinition.RangeBox
        dx = abs(float(rg.MaxPoint.X) - float(rg.MinPoint.X)) * CM_TO_MM
        dy = abs(float(rg.MaxPoint.Y) - float(rg.MinPoint.Y)) * CM_TO_MM
        dz = abs(float(rg.MaxPoint.Z) - float(rg.MinPoint.Z)) * CM_TO_MM
        dims = sorted([dx, dy, dz], reverse=True)
        return dims  # L, W, T approx
    except Exception:
        return None


def _stats_caras(part_doc):
    n_body = 0
    n_faces = 0
    n_plane = 0
    areas = []
    try:
        cdef = part_doc.ComponentDefinition
        n_body = int(cdef.SurfaceBodies.Count)
        for i in range(1, n_body + 1):
            body = cdef.SurfaceBodies.Item(i)
            for j in range(1, int(body.Faces.Count) + 1):
                face = body.Faces.Item(j)
                n_faces += 1
                try:
                    if int(face.SurfaceType) == K_PLANE:
                        n_plane += 1
                        areas.append(float(face.Evaluator.Area))
                except Exception:
                    pass
    except Exception as exc:
        return {"err": str(exc)}
    areas.sort(reverse=True)
    ratio = None
    if len(areas) >= 2 and areas[0] > 1e-12:
        ratio = areas[1] / areas[0]
    return {
        "bodies": n_body,
        "faces": n_faces,
        "planes": n_plane,
        "area_max_cm2": areas[0] if areas else None,
        "area2_cm2": areas[1] if len(areas) > 1 else None,
        "ratio_a2_a1": ratio,
        "n_areas": len(areas),
    }


def _es_sheet_metal(part_doc) -> bool:
    try:
        return str(part_doc.SubType) == SM_GUID
    except Exception:
        return False


def _thk_sm_mm(part_doc):
    try:
        sm = win32com.client.CastTo(
            part_doc.ComponentDefinition, "SheetMetalComponentDefinition"
        )
        return float(sm.Thickness.Value) * CM_TO_MM
    except Exception:
        return None


def _clasificar_forma(dims):
    """Heurística: placa / barra / bloque / irregular."""
    if not dims or min(dims) <= 0:
        return "desconocida"
    L, W, T = dims
    if L / max(T, 1e-9) >= 8 and W / max(T, 1e-9) >= 3:
        return "placa_plana"
    if L / max(W, 1e-9) >= 5 and W / max(T, 1e-9) <= 3:
        return "barra_alargada"
    if L / max(T, 1e-9) <= 4:
        return "bloque_compacto"
    return "intermedia"


def _recolectar_cobre(ensamble, limite=800):
    """Leaf parts cobres únicas por ruta (muestra acotada)."""
    out = []
    vistos = set()
    try:
        leaf = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
        total = int(leaf.Count)
    except Exception as exc:
        print(f"ERROR leaf: {exc}")
        return out, 0

    i = 1
    scanned = 0
    while i <= total and len(out) < limite:
        try:
            occ = leaf.Item(i)
        except Exception:
            i += 1
            continue
        scanned += 1
        i += 1
        try:
            if occ.Suppressed:
                continue
        except Exception:
            continue
        try:
            doc = occ.Definition.Document
            ruta = str(getattr(doc, "FullFileName", "") or "")
            if not ruta or ruta in vistos:
                continue
            nombre = os.path.splitext(os.path.basename(ruta))[0]
            if not es_pieza_cobre(nombre):
                continue
            vistos.add(ruta)
            out.append((doc, nombre, ruta))
        except Exception:
            continue
    return out, total


def main() -> int:
    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        print(f"Inventor OK | docs={inv.Documents.Count}", flush=True)
        ensamble = _obtener_ensamble_principal(inv)
        if ensamble is None:
            print("ERROR: sin ensamble")
            return 1
        ensamble = _como_ensamble(ensamble)
        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        set_nombre_pieza_completo(True)
        print(
            f"Ensamble: {ensamble.DisplayName} | {info.get('tipo')} | "
            f"unidad={get_unidad_cota()} | nombre_completo=ON",
            flush=True,
        )

        cobre, total_leaf = _recolectar_cobre(ensamble, limite=500)
        pref = Counter(prefijo_cobre(n) for _, n, _ in cobre)
        print(
            f"Leaf total≈{total_leaf} | cobre únicos ABB/GENE/RLG: {len(cobre)} "
            f"| prefijos={dict(pref)}",
            flush=True,
        )

        # Colisiones de nombre corto vs completo
        cortos = Counter()
        for _, n, _ in cobre:
            set_nombre_pieza_completo(False)
            cortos[obtener_nombre_base_corto(n)] += 1
        set_nombre_pieza_completo(True)
        colisiones = {k: v for k, v in cortos.items() if v > 1}
        print(
            f"Colisiones si truncáramos a 3 segmentos: "
            f"{len(colisiones)} grupos (peor: "
            f"{max(colisiones.values()) if colisiones else 0} piezas/carpeta)",
            flush=True,
        )
        for k, v in sorted(colisiones.items(), key=lambda x: -x[1])[:8]:
            print(f"  {k}: {v} piezas distintas", flush=True)

        muestra = cobre[:MAX_MUESTRA]
        print(f"\n=== Análisis geométrico muestra n={len(muestra)} ===", flush=True)

        formas = Counter()
        sm_n = 0
        multi = 0
        thin = 0
        frente_ambiguo = 0  # 2ª cara > 70% de la 1ª
        sin_lado_claro = 0  # pocas planas
        filas = []

        for doc, nombre, _ruta in muestra:
            part = _cast_part(doc)
            is_sm = _es_sheet_metal(part)
            if is_sm:
                sm_n += 1
            dims = _bbox_mm(part)
            forma = _clasificar_forma(dims)
            formas[forma] += 1
            st = _stats_caras(part)
            if st.get("bodies", 0) > 1:
                multi += 1
            if dims and dims[2] < 8.0:  # espesor bbox < 8 mm
                thin += 1
            if st.get("ratio_a2_a1") is not None and st["ratio_a2_a1"] >= 0.70:
                frente_ambiguo += 1
            if st.get("planes", 0) < 3:
                sin_lado_claro += 1
            thk = _thk_sm_mm(part) if is_sm else None

            # Simular elegir_frente: ¿área máx es razonable vs bbox?
            riesgo = []
            if forma == "barra_alargada":
                riesgo.append("barra: frente=área máx puede ser canto largo")
            if st.get("ratio_a2_a1") and st["ratio_a2_a1"] >= 0.85:
                riesgo.append("2 caras casi iguales → frente inestable")
            if not is_sm and forma == "placa_plana":
                riesgo.append("sólido placa (no SM): THK sin Thickness")
            if st.get("bodies", 0) > 1:
                riesgo.append("multi-body")

            filas.append(
                {
                    "nombre": nombre,
                    "pre": prefijo_cobre(nombre),
                    "sm": is_sm,
                    "dims": dims,
                    "forma": forma,
                    "planes": st.get("planes"),
                    "faces": st.get("faces"),
                    "thk_mm": thk,
                    "riesgo": riesgo,
                }
            )

        print(f"Sheet Metal: {sm_n}/{len(muestra)}", flush=True)
        print(f"Formas: {dict(formas)}", flush=True)
        print(f"Multi-body: {multi} | bbox delgado (<8mm): {thin}", flush=True)
        print(
            f"Frente ambiguo (A2/A1≥0.70): {frente_ambiguo} | "
            f"pocas planas (<3): {sin_lado_claro}",
            flush=True,
        )

        print("\n--- Muestra con riesgos ---", flush=True)
        con_riesgo = [f for f in filas if f["riesgo"]]
        for f in (con_riesgo or filas)[:20]:
            d = f["dims"]
            ds = (
                f"{d[0]:.1f}x{d[1]:.1f}x{d[2]:.1f} mm"
                if d
                else "?"
            )
            print(
                f"  [{f['pre']}] {f['nombre'][:55]} | "
                f"{'SM' if f['sm'] else 'sol'} {f['forma']} {ds} "
                f"planas={f['planes']} "
                f"{('| ' + '; '.join(f['riesgo'])) if f['riesgo'] else ''}",
                flush=True,
            )

        # JPG export previo
        jpg_root = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "JPG",
            "9919-Board 1",
            "PIEZAS_ACOTADAS",
        )
        if os.path.isdir(jpg_root):
            folders = []
            for root, dirs, files in os.walk(jpg_root):
                jpgs = [x for x in files if x.lower().endswith(".jpg")]
                if jpgs:
                    folders.append((root, len(jpgs)))
            # carpetas con muchos JPG (posible conglomerado viejo)
            big = sorted(folders, key=lambda x: -x[1])[:8]
            print("\n--- JPG export previos (top carpetas) ---", flush=True)
            for r, n in big:
                rel = os.path.relpath(r, jpg_root)
                print(f"  {n:3d} jpg | {rel}", flush=True)
            # ¿aún hay carpetas truncadas ABB-42-BCK?
            trunc = [
                rel
                for rel, n in (
                    (os.path.relpath(r, jpg_root), n) for r, n in folders
                )
                if os.path.basename(rel).upper() == "ABB-42-BCK"
            ]
            if trunc:
                print(
                    "  AVISO: aún existe carpeta conglomerada ABB-42-BCK "
                    "(corrida anterior; re-exportar con nombre completo).",
                    flush=True,
                )

        print("\n=== OPORTUNIDADES (heurística) ===", flush=True)
        print(
            "1. Frente=cara de mayor área: en barras/cobre doblado el 'frente' "
            "útil para nest/CypTube suele ser la cara plana principal, pero LADO "
            "elige la plana MÁS CHICA ortogonal → en placas delgadas el canto "
            "es minúsculo y THK falla o se inventa.",
            flush=True,
        )
        print(
            "2. Cobre GIGA casi nunca es 'Jacking Pad': no se fuerza "
            "_orientacion_lado_doblado; si son sólidos (no SM) el LADO puede "
            "mirar mal.",
            flush=True,
        )
        print(
            "3. Cotas siguen lógica tanque (LENGTH/BROAD/THK en in del machote "
            "UOM). GIGA pide mm: hay que forzar texto desde ModelValue→mm "
            "(ya cableado) y verificar que la última corrida lo usó.",
            flush=True,
        )
        print(
            "4. Nombre completo ON evita fusionar ABB-42-BCK-*; la carpeta "
            f"vieja con {max(colisiones.values()) if colisiones else 0} "
            "variantes por truncado debe regenerarse.",
            flush=True,
        )
        print(
            "5. Posible mejora: para ABB/GENE/RLG orientar FRENTE al plano de "
            "mayor área cuyo normal ≈ eje menor del bbox (placa), y LADO al "
            "canto (espesor bbox), en vez de min-área genérica.",
            flush=True,
        )
        print("DIAG_GEO_OK", flush=True)
        return 0
    except Exception as exc:
        print(f"DIAG_FAIL: {type(exc).__name__}: {exc}", flush=True)
        import traceback

        traceback.print_exc()
        return 2
    finally:
        try:
            set_unidad_cota("in")
            set_nombre_pieza_completo(False)
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
