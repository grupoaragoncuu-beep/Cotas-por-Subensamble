# -*- coding: utf-8 -*-
"""
Limpia basura de la reestructuración en JPGS (ruta canónica).

- Fusiona Estanado Busbar → Estañado Busbar
- Borra vacíos: Estanado Busbar, Corte/Estañado|Estanado|Corte, Plasma Busbar vacío
- No toca Nueva carpeta ni _STAGING_* con contenido útil
"""
from __future__ import annotations

import os
import shutil
import sys

ROOT_JPGS = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
# Gemelo erróneo que también ensucié; limpio igual (no es canónico).
ROOT_TWIN = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\9919-BOARD2_2"
)

DRY = "--dry-run" in sys.argv


def _es_basura_archivo(fn: str) -> bool:
    u = fn.lower()
    return u in ("thumbs.db", "desktop.ini", ".ds_store") or u.startswith("~$")


def _solo_basura_o_vacio(path: str) -> bool:
    if not os.path.isdir(path):
        return True
    for dp, _dns, fs in os.walk(path):
        for fn in fs:
            if not _es_basura_archivo(fn):
                return False
    return True


def _rmtree(path: str, motivo: str) -> None:
    if not os.path.exists(path):
        return
    print(f"  DEL [{motivo}]: {path}")
    if DRY:
        return
    if os.path.isfile(path):
        os.remove(path)
        return
    # Borrar basura suelta primero
    for dp, _dns, fs in os.walk(path, topdown=False):
        for fn in fs:
            try:
                os.remove(os.path.join(dp, fn))
            except OSError as e:
                print(f"    AVISO rm file {fn}: {e}")
        try:
            os.rmdir(dp)
        except OSError:
            try:
                shutil.rmtree(dp, ignore_errors=True)
            except Exception:
                pass


def _merge_dir(src: str, dst: str) -> int:
    """Mueve JPG de src a dst; devuelve cantidad movida."""
    if not os.path.isdir(src):
        return 0
    n = 0
    if not DRY:
        os.makedirs(dst, exist_ok=True)
    for dp, _dns, fs in os.walk(src):
        for fn in fs:
            if _es_basura_archivo(fn):
                continue
            s = os.path.join(dp, fn)
            rel = os.path.relpath(s, src)
            d = os.path.join(dst, rel)
            print(f"  MOVE {s} -> {d}")
            if DRY:
                n += 1
                continue
            os.makedirs(os.path.dirname(d), exist_ok=True)
            if os.path.exists(d):
                os.remove(d)
            shutil.move(s, d)
            n += 1
    return n


def limpiar_raiz(raiz: str, etiqueta: str) -> None:
    print(f"\n=== limpiar {etiqueta} ===")
    print(raiz)
    if not os.path.isdir(raiz):
        print("  (no existe)")
        return

    buen_est = os.path.join(raiz, "Estañado Busbar")
    mal_est = os.path.join(raiz, "Estanado Busbar")
    if os.path.isdir(mal_est):
        n = _merge_dir(mal_est, buen_est)
        print(f"  merge Estanado→Estañado: {n}")
        if _solo_basura_o_vacio(mal_est):
            _rmtree(mal_est, "Estanado Busbar vacío/duplicado")

    # Legacy bajo Corte/
    for rel in (
        r"Corte\Estañado",
        r"Corte\Estanado",
        r"Corte\Corte",
        r"Corte\Doblado",
    ):
        p = os.path.join(raiz, rel)
        if os.path.isdir(p) and _solo_basura_o_vacio(p):
            _rmtree(p, f"legacy vacío {rel}")

    # Hojas vacías del árbol nuevo
    for rel in (
        r"Corte\Plasma y Laser\Corte Busbar",
        r"Corte\Maquinado\Maquinados metal",
    ):
        p = os.path.join(raiz, rel)
        if os.path.isdir(p) and _solo_basura_o_vacio(p):
            _rmtree(p, f"hoja vacía {rel}")

    # Asegurar Estañado Busbar existe
    if not DRY:
        os.makedirs(buen_est, exist_ok=True)


def main() -> int:
    print("Limpiar cagadero", "DRY" if DRY else "APPLY")
    limpiar_raiz(ROOT_JPGS, "JPGS (canónico)")
    limpiar_raiz(ROOT_TWIN, "twin (no canónico)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
