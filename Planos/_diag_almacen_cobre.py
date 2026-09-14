# -*- coding: utf-8 -*-
"""Diagnóstico: clasificación Almacén vs filtro cobre ABB/GENE/RLG (COM)."""
from __future__ import annotations

import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom

from inventor_com import conectar_inventor
from generador_caras_tanque import (
    _obtener_ensamble_principal,
    _como_ensamble,
    detectar_mapa_piezas_por_clasificacion,
    CLASIFICACIONES_VALIDAS,
    CLASIFICACION_SIN_ASIGNAR,
)
from piezas_cobre import es_pieza_cobre, prefijo_cobre
from producto_tipo import aplicar_unidad_producto, clasificar_producto
from cota_estilo import set_unidad_cota
from creador_vistas import set_nombre_pieza_completo


def main() -> int:
    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        ensamble = _como_ensamble(_obtener_ensamble_principal(inv))
        if ensamble is None:
            print("ERROR: sin ensamble")
            return 1
        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        set_nombre_pieza_completo(True)
        print(f"Ensamble: {ensamble.DisplayName}", flush=True)

        print("Leyendo iProperty de clasificación...", flush=True)
        mapa = detectar_mapa_piezas_por_clasificacion(inv, ensamble)
        conteo = {k: len(v or []) for k, v in (mapa or {}).items()}
        print("Conteo por clasificación:", dict(sorted(conteo.items())), flush=True)

        almac = set(mapa.get("Almacén") or mapa.get("Almacen") or [])
        # buscar clave con acento normalizado
        for k, v in (mapa or {}).items():
            if "almac" in str(k).casefold():
                almac |= set(v or [])
                print(f"Clave Almacén encontrada: {k!r} → {len(v or [])} nombres", flush=True)

        cobre_almac = sorted(n for n in almac if es_pieza_cobre(n))
        no_cobre_almac = sorted(n for n in almac if not es_pieza_cobre(n))
        print(
            f"Almacén total={len(almac)} | cobre ABB/GENE/RLG={len(cobre_almac)} | "
            f"NO cobre (filtradas)={len(no_cobre_almac)}",
            flush=True,
        )
        pref = Counter(prefijo_cobre(n) or "?" for n in cobre_almac)
        print("Prefijos cobre en Almacén:", dict(pref), flush=True)

        print("\n--- Muestra Almacén COBRE (25) ---", flush=True)
        for n in cobre_almac[:25]:
            print(f"  + {n}", flush=True)
        print("\n--- Muestra Almacén NO-cobre (40) ---", flush=True)
        for n in no_cobre_almac[:40]:
            print(f"  - {n}", flush=True)

        # ¿Dónde cayeron los cobre exportados?
        jpg_root = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "JPG",
            "9919-Board 1",
            "PIEZAS_ACOTADAS",
        )
        en_disco = defaultdict(list)
        if os.path.isdir(jpg_root):
            for root, _dirs, files in os.walk(jpg_root):
                jpgs = [f for f in files if f.lower().endswith(".jpg")]
                if not jpgs:
                    continue
                rel = os.path.relpath(root, jpg_root)
                parts = rel.split(os.sep)
                clase = parts[0] if parts else "?"
                pieza = parts[1] if len(parts) > 1 else os.path.basename(root)
                en_disco[clase].append(pieza)

        print("\n--- Carpetas JPG por clasificación ---", flush=True)
        for clase, piezas in sorted(en_disco.items()):
            uniq = sorted(set(piezas))
            print(f"  {clase}: {len(uniq)} carpetas pieza", flush=True)

        # Cruce: cobre Almacén que SÍ tiene carpeta en Corte u otra
        todas_piezas_disco = {}
        for clase, piezas in en_disco.items():
            for p in set(piezas):
                todas_piezas_disco[p.upper()] = clase

        cruzados = []
        for n in cobre_almac:
            clave = n.upper()
            # match exacto o prefijo carpeta
            hit = todas_piezas_disco.get(clave)
            if not hit:
                for disk, clase in todas_piezas_disco.items():
                    if clave.startswith(disk) or disk.startswith(clave):
                        hit = clase
                        break
            if hit:
                cruzados.append((n, hit))

        print(
            f"\nCobre Almacén con JPG en disco: {len(cruzados)}/{len(cobre_almac)}",
            flush=True,
        )
        dest = Counter(c for _, c in cruzados)
        print("Destino real de esos JPG:", dict(dest), flush=True)
        for n, c in cruzados[:15]:
            print(f"  {n} → {c}/", flush=True)

        sin_jpg = [n for n in cobre_almac if n.upper() not in {x[0].upper() for x in cruzados}]
        print(f"\nCobre Almacén SIN jpg: {len(sin_jpg)}", flush=True)
        for n in sin_jpg[:20]:
            print(f"  ? {n}", flush=True)

        print("\n=== CONCLUSIÓN ===", flush=True)
        if len(almac) == 0:
            print(
                "Ninguna pieza del ensamble tiene iProperty = Almacén "
                "(o no se leyó). Por eso la carpeta queda vacía.",
                flush=True,
            )
        elif len(cobre_almac) == 0:
            print(
                "Hay piezas Almacén, pero NINGUNA inicia con ABB/GENE/RLG. "
                "El filtro cobre de BOARD no las exporta → Almacén/ vacío.",
                flush=True,
            )
        elif dest.get("Almacén", 0) == 0 and sum(dest.values()) > 0:
            print(
                "Sí hay cobre Almacén, pero al reorganizar cayeron en otra "
                "carpeta (p.ej. Corte) por match de mapa/clave.",
                flush=True,
            )
        print("DIAG_ALMACEN_OK", flush=True)
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
