# -*- coding: utf-8 -*-
"""Ángulos de doblez del metal, en el mismo paso que las piezas.

El arco queda en el bolsillo interior. Si la pieza ya tiene sus ANGLE,
no se reexporta. Cobre no entra aquí: sus alas van por wing_cobre.
"""
from __future__ import annotations

import importlib.util
import os
from pathlib import Path

_MOD = None


def _motor():
    global _MOD
    if _MOD is not None:
        return _MOD
    ruta = Path(r"C:\Proyectos\COTAS\.runtime\_probar_angulo_doblez_metal.py")
    spec = importlib.util.spec_from_file_location("angulo_doblez_metal", ruta)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"no pude cargar {ruta}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    _MOD = mod
    return mod


def acotar_angulos_ensamble(inv, ensamble=None, plano=None) -> int:
    """Un ANGLE por doblez de cada pieza de metal. Devuelve JPG nuevos."""
    from piezas_cobre import es_pieza_cobre
    from wing_cobre import _asm, _piezas, _sm, _stem

    mod = _motor()
    asm = _asm(inv, ensamble)
    if asm is None:
        print("ANGLE: no hay ensamble")
        return 0
    if plano is None:
        from generador_caras_tanque import _obtener_plano_activo

        plano = _obtener_plano_activo(inv)
    from generador_caras_tanque import _encontrar_hoja_machote

    base = _encontrar_hoja_machote(plano) or plano.Sheets.Item(1)
    try:
        from nomenclatura_capturas import nombre_job_desde_ensamble

        job = nombre_job_desde_ensamble(asm)
    except Exception:
        job = _stem(asm) or "JOB"
    local = Path(r"C:\Proyectos\COTAS\Planos\JPG") / job / "PIEZAS_ACOTADAS" / "Doblado" / "Metal"
    mod.JOB = job
    mod.OUT_ROOT = local
    root = os.environ.get("COTAS_VSM_JOB_ROOT", "").strip()
    if root:
        mod.Y_ROOT = (
            Path(root) / "DOSSIER FILES" / "JPGS" / "PIEZAS_ACOTADAS" / "Doblado" / "Metal"
        )
    forzar = os.environ.get("COTAS_ANGLE_FORCE", "").strip().lower() in ("1", "true", "yes")
    try:
        from cotas_dossier_registro import iniciar_sesion_dossier, registrar_jpg

        iniciar_sesion_dossier(job)
    except Exception as exc:
        registrar_jpg = None
        print(f"AVISO ANGLE sesion: {exc}")

    n_ok = 0
    for stem, part in _piezas(asm):
        if es_pieza_cobre(stem):
            continue
        if _sm(part) is None or mod.bends_count(part) < 1:
            continue
        if mod.es_multibody(part):
            continue
        dest = local / stem
        if not forzar and dest.is_dir() and any(dest.glob("*__ANGLE*.jpg")):
            continue
        creados = mod.process_part(inv, plano, base, part, stem)
        for ruta in creados:
            if registrar_jpg is not None and not registrar_jpg(ruta, job=job):
                print(f"  NO DB {os.path.basename(ruta)}")
            n_ok += 1
    print(f"ANGLE publicados: {n_ok}")
    return n_ok
