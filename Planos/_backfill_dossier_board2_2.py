# -*- coding: utf-8 -*-
"""Backfill cotas_dossier desde JPGS ya subidos a mano (9919-BOARD2_2)."""
from __future__ import annotations

import os
import sys

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
JOB = "9919-BOARD2_2"

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cotas_dossier_registro import (  # noqa: E402
    asegurar_tabla_cotas_dossier,
    clasificar_type_y_spoteos,
    iniciar_sesion_dossier,
    insertar_evidencia,
    proceso_desde_ruta,
    producto_cliente_desde_job_root,
    resolver_cliente_producto,
    _ya_registrada,
)


def main() -> int:
    print("ROOT exists:", os.path.isdir(ROOT))
    print("ROOT:", ROOT)
    if not os.path.isdir(ROOT):
        return 1

    job_root = os.path.dirname(os.path.dirname(ROOT))  # .../9919-BOARD2_2
    ctx = iniciar_sesion_dossier(JOB, job_root)
    print("contexto:", ctx)

    cli, prod, fuente = resolver_cliente_producto(JOB, job_root)
    # Forzar desde ruta VSM si quedó SIN_CLIENTE
    if not cli or cli == "SIN_CLIENTE":
        prod2, cli2 = producto_cliente_desde_job_root(job_root)
        if cli2:
            cli, prod, fuente = cli2, prod2, "ruta_job"
    print(f"cliente={cli!r} producto={prod!r} fuente={fuente}")

    asegurar_tabla_cotas_dossier()

    archivos = []
    for dirpath, _dirs, files in os.walk(ROOT):
        for fn in files:
            if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                archivos.append(os.path.join(dirpath, fn))
    archivos.sort()
    print(f"capturas encontradas: {len(archivos)}")

    insertados = 0
    omitidos = 0
    errores = 0
    por_type = {}
    por_proceso = {}

    for ruta in archivos:
        nombre = os.path.basename(ruta)
        ruta_abs = os.path.abspath(ruta)
        if _ya_registrada(JOB, ruta_abs, nombre):
            omitidos += 1
            continue
        tipo, n_spot = clasificar_type_y_spoteos(nombre)
        proc = proceso_desde_ruta(ruta_abs) or proceso_desde_ruta(ruta)
        eid = insertar_evidencia(
            cliente=cli,
            producto=prod,
            job=JOB,
            type_=tipo,
            cantidad_spoteos=n_spot,
            nombre_archivo=nombre,
            ruta=ruta_abs,
            clasificacion=proc,
        )
        if eid:
            insertados += 1
            por_type[tipo] = por_type.get(tipo, 0) + 1
            key = proc or "(sin proceso)"
            por_proceso[key] = por_proceso.get(key, 0) + 1
            print(f"  + id={eid} [{proc}] {tipo} x{n_spot} | {nombre}")
        else:
            errores += 1
            print(f"  ! FAIL {nombre}")

    print("---")
    print(f"insertados={insertados} omitidos_ya_en_db={omitidos} errores={errores}")
    print("por type:", por_type)
    print("por proceso:", por_proceso)
    return 0 if errores == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
