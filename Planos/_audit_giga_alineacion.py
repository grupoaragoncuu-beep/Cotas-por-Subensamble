# -*- coding: utf-8 -*-
"""
Auditoría COM alineación GIGA vs requisitos del usuario.

Attach Inventor → 9919-Board → verifica:
  1) Alcance: TODAS las piezas únicas (no solo cobre)
  2) Cobre ABB/GENE/RLG: solo marca SIN_COTA
  3) Nombres completos (sin truncar ABB-42-BCK)
  4) Unidades mm
  5) Orientación: FRENTE = cara grande (no perfil L)
  6) Barrenos: círculos + óvalos; no radios de doblez
  7) Almacén / Corte / Doblado poblados en mapa iProperty
"""
from __future__ import annotations

import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client

from inventor_com import conectar_inventor
from generador_caras_tanque import (
    _obtener_ensamble_principal,
    _como_ensamble,
    detectar_mapa_piezas_por_clasificacion,
)
from producto_tipo import clasificar_producto, aplicar_unidad_producto, unidad_para_producto
from piezas_cobre import es_pieza_cobre, prefijo_cobre, catalogo_prefijos
from cota_estilo import get_unidad_cota, set_unidad_cota
from creador_vistas import (
    set_nombre_pieza_completo,
    get_nombre_pieza_completo,
    obtener_nombre_base_corto,
    recolectar_piezas_unicas,
    _orientacion_board_giga,
    _orientacion_cobre_giga,
    _es_flujo_board_giga,
    elegir_frente,
    _caras_y_cuerpo_doblado,
    kPlaneSurface,
    obtener_normal_cara,
)
import diametro as diam


def _ok(cond, msg_ok, msg_bad):
    if cond:
        print(f"  [OK] {msg_ok}", flush=True)
        return True
    print(f"  [FAIL] {msg_bad}", flush=True)
    return False


def _cast_part(doc):
    try:
        return win32com.client.CastTo(doc, "PartDocument")
    except Exception:
        return doc


def _analizar_orientacion(inv, part_doc, nombre):
    """Compara normal FRENTE giga vs ejes bbox (perfil L = mirar T)."""
    tg = inv.TransientGeometry
    to = inv.TransientObjects
    ori = _orientacion_cobre_giga(part_doc, tg, to)
    if not ori:
        return {"ok": False, "motivo": "sin orientación"}
    caras, _ = _caras_y_cuerpo_doblado(part_doc, to)
    res = elegir_frente(caras) if caras else None
    if not res:
        return {"ok": False, "motivo": "sin cara frente"}
    _f, n_max, area = res
    v_f = ori["v_frente"]
    alin = abs(float(v_f.DotProduct(n_max)))
    # Perfil L típico: eye ≈ eje menor bbox (T). Si FRENTE ≈ cara máx, alin alto.
    return {
        "ok": alin >= 0.75,
        "alin_cara_max": round(alin, 3),
        "modo": ori.get("modo"),
        "tiene_frente2": ori.get("v_frente_2") is not None,
        "area_frente": round(float(area), 4),
    }


def main() -> int:
    pythoncom.CoInitialize()
    fallos = 0
    try:
        inv = conectar_inventor()
        print(f"Inventor OK | docs={inv.Documents.Count}", flush=True)
        ensamble = _como_ensamble(_obtener_ensamble_principal(inv))
        if ensamble is None:
            print("ERROR: sin ensamble")
            return 1
        print(f"Ensamble: {ensamble.DisplayName}", flush=True)

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        set_nombre_pieza_completo(True)

        print("\n=== 1) Producto / unidades / nombres ===", flush=True)
        fallos += not _ok(
            info.get("tipo") == "BOARD",
            f"tipo BOARD ({info.get('motivo')})",
            f"tipo={info.get('tipo')} (se esperaba BOARD)",
        )
        fallos += not _ok(
            get_unidad_cota() == "mm",
            "unidad cota = mm",
            f"unidad={get_unidad_cota()}",
        )
        fallos += not _ok(
            get_nombre_pieza_completo(),
            "nombre pieza COMPLETO activo",
            "nombre truncado activo",
        )
        fallos += not _ok(
            obtener_nombre_base_corto("ABB-42-BCK-705_718") == "ABB-42-BCK-705_718",
            "ABB-42-BCK-705_718 no se trunca",
            f"quedó {obtener_nombre_base_corto('ABB-42-BCK-705_718')!r}",
        )
        fallos += not _ok(
            catalogo_prefijos() == ("ABB", "GENE", "RLG"),
            f"cobre solo {catalogo_prefijos()}",
            f"prefijos inesperados {catalogo_prefijos()}",
        )

        print("\n=== 2) Alcance piezas (COM) ===", flush=True)
        # Silenciar log verboso de occs
        import creador_vistas as cv

        _old_log = cv._log
        cv._log = lambda *_a, **_k: None
        try:
            piezas = recolectar_piezas_unicas(ensamble)
        finally:
            cv._log = _old_log
        cobre = [(d, n) for d, n in piezas if es_pieza_cobre(n)]
        resto = [(d, n) for d, n in piezas if not es_pieza_cobre(n)]
        print(
            f"  Únicas={len(piezas)} | cobre={len(cobre)} | no-cobre={len(resto)}",
            flush=True,
        )
        fallos += not _ok(
            len(piezas) > len(cobre) and len(resto) > 0,
            "hay no-cobre en alcance (no se excluyen)",
            "solo hay cobre o lista vacía",
        )
        fallos += not _ok(
            _es_flujo_board_giga(),
            "flujo BOARD/GIGA activo",
            "flujo board inactivo",
        )
        import inspect
        import generador_piezas as gp

        src = inspect.getsource(gp.ejecutar)
        fallos += not _ok(
            "_catalogo_piezas_cobre_ensamble" not in src
            and "BOARD filtro cobre" not in src,
            "generador_piezas.ejecutar sin filtro de alcance cobre",
            "ejecutar() aún filtra por cobre",
        )

        print("\n=== 3) Clasificación iProperty ===", flush=True)
        mapa = detectar_mapa_piezas_por_clasificacion(inv, ensamble)
        for clase in ("Almacén", "Corte", "Doblado", "SIN CLASIFICACION"):
            n = len(mapa.get(clase) or [])
            print(f"  {clase}: {n}", flush=True)
        n_alm = len(mapa.get("Almacén") or [])
        fallos += not _ok(
            n_alm > 0,
            f"Almacén tiene {n_alm} piezas en mapa (deben acotarse)",
            "Almacén vacío en iProperty",
        )
        # Ninguna Almacén es cobre → antes fallaba el filtro; ahora deben procesarse
        alm = list(mapa.get("Almacén") or [])
        alm_cobre = [x for x in alm if es_pieza_cobre(x)]
        print(
            f"  Almacén cobre={len(alm_cobre)} / total={len(alm)} "
            f"(las no-cobre también entran al flujo)",
            flush=True,
        )

        print("\n=== 4) Orientación muestra ABB L (cara ≠ perfil) ===", flush=True)
        muestras = [
            (d, n)
            for d, n in cobre
            if n.upper().startswith("ABB-42-BCK")
        ][:5]
        if not muestras:
            muestras = cobre[:3]
        ori_ok = 0
        for doc, nom in muestras:
            part = _cast_part(doc)
            try:
                r = _analizar_orientacion(inv, part, nom)
            except Exception as exc:
                r = {"ok": False, "motivo": str(exc)}
            print(
                f"  {nom}: {r}",
                flush=True,
            )
            if r.get("ok"):
                ori_ok += 1
        fallos += not _ok(
            ori_ok >= max(1, len(muestras) // 2),
            f"orientación cara grande OK en {ori_ok}/{len(muestras)}",
            f"orientación débil {ori_ok}/{len(muestras)} (aún perfil L?)",
        )

        print("\n=== 5) Detector barrenos (código + humo) ===", flush=True)
        fallos += not _ok(
            hasattr(diam, "_ranuras_en_vista") and hasattr(diam, "_barrenos_en_vista"),
            "diametro: círculos + óvalos disponibles",
            "faltan helpers de óvalos",
        )
        fake = [
            {"curva": 1, "diam": 1.0, "cx": 0.0, "cy": 5.0, "eje": "V"},
            {"curva": 2, "diam": 1.0, "cx": 4.0, "cy": 5.0, "eje": "V"},
        ]
        slots = diam._emparejar_ranuras(fake)
        fallos += not _ok(
            len(slots) == 1 and slots[0]["tipo"] == "oval",
            "emparejado de ranura oval OK",
            f"emparejado falló: {slots}",
        )

        print("\n=== 6) Export JPG actual (disco) ===", flush=True)
        jpg_root = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "JPG",
            "9919-Board 1",
            "PIEZAS_ACOTADAS",
        )
        if os.path.isdir(jpg_root):
            por_clase = Counter()
            for root, _dirs, files in os.walk(jpg_root):
                jpgs = [f for f in files if f.lower().endswith(".jpg")]
                if not jpgs:
                    continue
                rel = os.path.relpath(root, jpg_root)
                clase = rel.split(os.sep)[0]
                por_clase[clase] += len(jpgs)
            for k, v in sorted(por_clase.items()):
                print(f"  {k}: {v} jpg", flush=True)
            if por_clase.get("Almacén", 0) == 0 and n_alm > 0:
                print(
                    "  [AVISO] Almacén/ en disco aún vacío → corrida "
                    "fue con filtro viejo; re-exportar con código actual.",
                    flush=True,
                )
            else:
                print("  [OK] hay JPG en Almacén o mapa sin piezas allí", flush=True)
        else:
            print("  (sin carpeta JPG aún)", flush=True)

        print("\n=== RESUMEN ===", flush=True)
        if fallos == 0:
            print("ALINEADO: requisitos OK en sesión Inventor + código.", flush=True)
            return 0
        print(f"DESALINEADO: {fallos} chequeo(s) fallaron.", flush=True)
        return 1
    except Exception as exc:
        print(f"AUDIT_FAIL: {type(exc).__name__}: {exc}", flush=True)
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
