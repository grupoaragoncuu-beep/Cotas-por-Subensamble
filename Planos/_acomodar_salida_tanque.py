# -*- coding: utf-8 -*-
"""
Acomoda y verifica la salida JPG de un tanque (PIEZAS / CARAS / ENSAMBLES).

- Reorganiza PIEZAS_ACOTADAS por cara + clasificación (mapas .json).
- Vacía staging residual (_STAGING_*).
- Resume conteos y avisa huecos (caras sin JPG, JPG sueltos en raíz).

Uso:
  python _acomodar_salida_tanque.py
  python _acomodar_salida_tanque.py --job 62223-1246-A01
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys


ROOT = os.path.dirname(os.path.abspath(__file__))
JPG_ROOT = os.path.join(ROOT, "JPG")


def _log(msg: str) -> None:
    print(msg, flush=True)


def _contar_jpgs(carpeta: str) -> int:
    n = 0
    if not os.path.isdir(carpeta):
        return 0
    for dirpath, _dirs, files in os.walk(carpeta):
        for f in files:
            if f.lower().endswith(".jpg"):
                n += 1
    return n


def _jpgs_sueltos_raiz(carpeta: str) -> list[str]:
    if not os.path.isdir(carpeta):
        return []
    out = []
    for nombre in os.listdir(carpeta):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isfile(ruta) and nombre.lower().endswith(".jpg"):
            out.append(nombre)
    return out


def _cargar_json(ruta: str) -> dict:
    if not os.path.isfile(ruta):
        return {}
    try:
        with open(ruta, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception as exc:
        _log(f"AVISO: no se pudo leer {ruta}: {exc}")
        return {}


def _resolver_job(job: str) -> str:
    job = str(job or "").strip()
    if job:
        ruta = os.path.join(JPG_ROOT, job)
        if os.path.isdir(ruta):
            return ruta
        raise SystemExit(f"ERROR: no existe job {ruta}")
    # Más reciente con PIEZAS_ACOTADAS o mapas.
    candidatos = []
    if not os.path.isdir(JPG_ROOT):
        raise SystemExit(f"ERROR: no existe {JPG_ROOT}")
    for nombre in os.listdir(JPG_ROOT):
        ruta = os.path.join(JPG_ROOT, nombre)
        if not os.path.isdir(ruta):
            continue
        if nombre.upper() in ("DOSSIER FILES",):
            continue
        marca = 0.0
        for f in (
            ".piezas_por_cara.json",
            ".piezas_por_clasificacion.json",
            "PIEZAS_ACOTADAS",
        ):
            p = os.path.join(ruta, f)
            if os.path.exists(p):
                try:
                    marca = max(marca, os.path.getmtime(p))
                except OSError:
                    pass
        if marca:
            candidatos.append((marca, ruta))
    if not candidatos:
        raise SystemExit("ERROR: no hay jobs JPG para acomodar.")
    candidatos.sort(reverse=True)
    return candidatos[0][1]


def _aplanar_jpgs_a_raiz(carpeta_piezas: str) -> int:
    """Trae todos los JPG anidados a la raíz (para re-aplicar cara+clase)."""
    movidos = 0
    if not os.path.isdir(carpeta_piezas):
        return 0
    pendientes = []
    for dirpath, _dirs, files in os.walk(carpeta_piezas):
        if os.path.abspath(dirpath) == os.path.abspath(carpeta_piezas):
            continue
        for nombre in files:
            if nombre.lower().endswith(".jpg"):
                pendientes.append(os.path.join(dirpath, nombre))
    for ruta in pendientes:
        nombre = os.path.basename(ruta)
        destino = os.path.join(carpeta_piezas, nombre)
        try:
            if os.path.abspath(ruta) == os.path.abspath(destino):
                continue
            if os.path.exists(destino):
                # Misma cota ya en raíz → descartar duplicado (no crear __dup).
                try:
                    if os.path.getsize(ruta) == os.path.getsize(destino):
                        os.remove(ruta)
                        continue
                except OSError:
                    pass
                base, ext = os.path.splitext(nombre)
                # Quitar __dup previo para no encadenar __dup1__dup2.
                base = re.sub(r"__dup\d+$", "", base, flags=re.IGNORECASE)
                k = 1
                while os.path.exists(destino):
                    destino = os.path.join(
                        carpeta_piezas, f"{base}__dup{k}{ext}"
                    )
                    k += 1
            shutil.move(ruta, destino)
            movidos += 1
        except OSError as err:
            _log(f"  AVISO aplanar '{nombre}': {err}")
    return movidos


def _limpiar_staging(carpeta_piezas: str) -> None:
    """Borra staging solo si ya está vacío (nunca con JPG dentro)."""
    for staging in ("_STAGING_DESPLIEGUE", "_STAGING_ESTANIADO"):
        ruta = os.path.join(carpeta_piezas, staging)
        if not os.path.isdir(ruta):
            continue
        n = _contar_jpgs(ruta)
        if n:
            _log(
                f"  AVISO: {staging} aún tiene {n} JPG; "
                "NO se borra (debe reorganizar antes)."
            )
            continue
        try:
            shutil.rmtree(ruta, ignore_errors=True)
        except OSError as err:
            _log(f"  AVISO: no se pudo borrar {staging}: {err}")


def _resumen_arbol(carpeta: str, max_nivel: int = 2) -> None:
    if not os.path.isdir(carpeta):
        _log(f"  (no existe) {os.path.basename(carpeta)}")
        return
    total = _contar_jpgs(carpeta)
    _log(f"  {os.path.basename(carpeta)}: {total} JPG")
    try:
        hijos = sorted(os.listdir(carpeta))
    except OSError:
        return
    for nombre in hijos:
        ruta = os.path.join(carpeta, nombre)
        if not os.path.isdir(ruta):
            continue
        n = _contar_jpgs(ruta)
        _log(f"    {nombre}/ → {n} JPG")
        if max_nivel <= 1:
            continue
        try:
            sub = sorted(os.listdir(ruta))
        except OSError:
            continue
        for sn in sub:
            sr = os.path.join(ruta, sn)
            if os.path.isdir(sr):
                _log(f"      {sn}/ → {_contar_jpgs(sr)} JPG")


def acomodar(job_dir: str) -> int:
    sys.path.insert(0, ROOT)
    from generador_tanque_completo import (
        SUBCARPETAS_CARA_SELECCION,
        STAGING_DESPLIEGUE,
        STAGING_ESTANIADO,
        _limpiar_carpetas_cara_vacias_y_legacy,
        _reorganizar_piezas_por_cara,
        _reorganizar_piezas_por_clasificacion,
    )
    from generador_piezas import _reorganizar_clasificacion_dentro_caras

    _log("=" * 62)
    _log(f" ACOMODAR SALIDA: {os.path.basename(job_dir)}")
    _log("=" * 62)

    mapa_cara = _cargar_json(os.path.join(job_dir, ".piezas_por_cara.json"))
    mapa_clase = _cargar_json(
        os.path.join(job_dir, ".piezas_por_clasificacion.json")
    )
    if "mapa" in mapa_cara and isinstance(mapa_cara["mapa"], dict):
        mapa_cara = mapa_cara["mapa"]
    if "mapa" in mapa_clase and isinstance(mapa_clase["mapa"], dict):
        mapa_clase = mapa_clase["mapa"]

    carpeta_piezas = os.path.join(job_dir, "PIEZAS_ACOTADAS")
    carpeta_caras = os.path.join(job_dir, "COTAS_POR_REFERENCIA")
    carpeta_ens = os.path.join(job_dir, "ENSAMBLES_INDEPENDIENTES")

    ok = 0
    if os.path.isdir(carpeta_piezas):
        sueltos = _jpgs_sueltos_raiz(carpeta_piezas)
        n_staging = sum(
            _contar_jpgs(os.path.join(carpeta_piezas, s))
            for s in (STAGING_DESPLIEGUE, STAGING_ESTANIADO)
        )
        _log(
            f"PIEZAS_ACOTADAS antes: {_contar_jpgs(carpeta_piezas)} JPG "
            f"({len(sueltos)} sueltos en raíz, {n_staging} en staging)"
        )
        if not mapa_cara and not mapa_clase:
            _log(
                "  ERROR: sin mapas .json; NO se aplana PIEZAS "
                "(evita dejar JPG sueltos). Regenera mapas primero."
            )
            ok = 1
        else:
            aplanados = _aplanar_jpgs_a_raiz(carpeta_piezas)
            if aplanados:
                _log(
                    f"  Aplanados a raíz: {aplanados} JPG "
                    "(re-organización limpia)"
                )
            # Staging raíz primero (DESPLIEGUE → Corte/…), nunca borrar con JPG.
            if n_staging and mapa_clase:
                _log("  Reorganizando staging raíz por clasificación...")
                _reorganizar_piezas_por_clasificacion(carpeta_piezas, mapa_clase)
            elif n_staging and mapa_cara:
                _log("  Reorganizando staging raíz por cara...")
                _reorganizar_piezas_por_cara(carpeta_piezas, mapa_cara)

            if mapa_cara:
                _log("  Reorganizando por cara (SEGM/TOP/BASE)...")
                _reorganizar_piezas_por_cara(carpeta_piezas, mapa_cara)
                if mapa_clase:
                    _log("  Anidando clasificación dentro de cada cara...")
                    _reorganizar_clasificacion_dentro_caras(
                        carpeta_piezas, mapa_clase
                    )
            elif mapa_clase:
                _log("  Reorganizando solo por clasificación...")
                _reorganizar_piezas_por_clasificacion(carpeta_piezas, mapa_clase)

            _limpiar_staging(carpeta_piezas)
            _limpiar_carpetas_cara_vacias_y_legacy(carpeta_piezas)
            sueltos_despues = _jpgs_sueltos_raiz(carpeta_piezas)
            if sueltos_despues:
                _log(
                    f"  AVISO: quedan {len(sueltos_despues)} JPG sueltos en raíz."
                )
                ok = 1
            n_staging_despues = sum(
                _contar_jpgs(os.path.join(carpeta_piezas, s))
                for s in (STAGING_DESPLIEGUE, STAGING_ESTANIADO)
            )
            if n_staging_despues:
                _log(f"  AVISO: staging residual con {n_staging_despues} JPG.")
                ok = 1
            for cara in SUBCARPETAS_CARA_SELECCION:
                n = _contar_jpgs(os.path.join(carpeta_piezas, cara))
                if n == 0 and mapa_cara:
                    piezas = mapa_cara.get(cara) or []
                    if piezas:
                        _log(f"  AVISO: {cara} en mapa pero 0 JPG.")
                        ok = 1
            _resumen_arbol(carpeta_piezas, max_nivel=2)
    else:
        _log("AVISO: no hay PIEZAS_ACOTADAS.")
        ok = 1

    _log("")
    _resumen_arbol(carpeta_caras, max_nivel=2)
    _log("")
    _resumen_arbol(carpeta_ens, max_nivel=2)

    total = (
        _contar_jpgs(carpeta_piezas)
        + _contar_jpgs(carpeta_caras)
        + _contar_jpgs(carpeta_ens)
    )
    _log("")
    _log(f"TOTAL JPG job: {total}")
    _log("ACOMODAR listo." if ok == 0 else "ACOMODAR listo con avisos.")
    return ok


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--job", default="", help="Nombre carpeta bajo Planos/JPG")
    args = p.parse_args()
    job_dir = _resolver_job(args.job)
    return acomodar(job_dir)


if __name__ == "__main__":
    sys.exit(main())
