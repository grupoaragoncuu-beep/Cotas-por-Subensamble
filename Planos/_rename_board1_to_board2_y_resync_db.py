# -*- coding: utf-8 -*-
"""
One-shot: renombra JOB en JPG ``9919-Board 1`` → ``9919-Board 2`` bajo
DOSSIER FILES\\JPGS del board 9919-BOARD2_2, borra filas previas en
cotas_dossier y reinserta con las rutas nuevas.
"""
from __future__ import annotations

import os
import sys

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
OLD_JOB_TOKEN = "9919-Board 1"
NEW_JOB_TOKEN = "9919-Board 2"
# Filas a limpiar en DB (nomenclatura vieja / token de ensamble).
JOBS_A_BORRAR = (
    "9919-BOARD2_2",
    "9919-Board 2",
    "9919-Board 1",
    "9919-BOARD 2_2",
    "9919-BOARD2-2",
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cotas_dossier_registro import (  # noqa: E402
    asegurar_tabla_cotas_dossier,
    clasificar_type_y_spoteos,
    guardar_contexto_dossier,
    insertar_evidencia,
    proceso_desde_ruta,
    producto_cliente_desde_job_root,
    resolver_cliente_producto,
    _db_nesting,
)


def _renombrar_archivos(root: str) -> tuple[int, int, int]:
    renombrados = 0
    ya_ok = 0
    errores = 0
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if OLD_JOB_TOKEN not in fn:
                if fn.startswith(NEW_JOB_TOKEN + "__") or fn.startswith(
                    NEW_JOB_TOKEN + "_"
                ):
                    ya_ok += 1
                continue
            nuevo = fn.replace(OLD_JOB_TOKEN, NEW_JOB_TOKEN)
            src = os.path.join(dirpath, fn)
            dst = os.path.join(dirpath, nuevo)
            try:
                if os.path.abspath(src) == os.path.abspath(dst):
                    continue
                if os.path.exists(dst):
                    # Evitar colisión: si destino ya existe, quitar origen.
                    os.remove(src)
                    print(f"  colisión: eliminado origen (destino existe): {fn}")
                    renombrados += 1
                    continue
                os.rename(src, dst)
                renombrados += 1
            except OSError as exc:
                errores += 1
                print(f"  ERROR rename {fn}: {exc}")
    return renombrados, ya_ok, errores


def _borrar_filas_job() -> int:
    try:
        import psycopg2
    except ImportError:
        print("ERROR: falta psycopg2")
        return -1
    borradas = 0
    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            # Exactos
            cur.execute(
                """
                DELETE FROM public.cotas_dossier
                WHERE job = ANY(%s)
                RETURNING id
                """,
                (list(JOBS_A_BORRAR),),
            )
            borradas += len(cur.fetchall() or [])
            # Por si quedó con espacios/guiones distintos
            cur.execute(
                """
                DELETE FROM public.cotas_dossier
                WHERE REPLACE(REPLACE(UPPER(job), ' ', ''), '_', '')
                      LIKE '9919BOARD2%%'
                   OR REPLACE(REPLACE(UPPER(job), ' ', ''), '_', '')
                      LIKE '9919BOARD1%%'
                RETURNING id
                """
            )
            borradas += len(cur.fetchall() or [])
        conn.commit()
    return borradas


def _reinsertar(root: str, job: str, cli: str, prod: str) -> tuple[int, int]:
    asegurar_tabla_cotas_dossier()
    insertados = 0
    errores = 0
    for dirpath, _dirs, files in os.walk(root):
        for fn in files:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            ruta = os.path.join(dirpath, fn)
            ruta_abs = os.path.abspath(ruta)
            # Preferir token del nombre si viene JOB__ITEM__…
            job_fila = job
            if "__" in fn:
                job_fila = fn.split("__", 1)[0].strip() or job
            tipo, n_spot = clasificar_type_y_spoteos(fn)
            proc = proceso_desde_ruta(ruta_abs) or proceso_desde_ruta(ruta)
            eid = insertar_evidencia(
                cliente=cli,
                producto=prod,
                job=job_fila,
                type_=tipo,
                cantidad_spoteos=n_spot,
                nombre_archivo=fn,
                ruta=ruta_abs,
                clasificacion=proc,
            )
            if eid:
                insertados += 1
            else:
                errores += 1
    return insertados, errores


def main() -> int:
    print("ROOT:", ROOT)
    print("exists:", os.path.isdir(ROOT))
    if not os.path.isdir(ROOT):
        return 1

    print("\n=== 1) Renombrar JPG:", OLD_JOB_TOKEN, "->", NEW_JOB_TOKEN)
    ren, ya, err_r = _renombrar_archivos(ROOT)
    print(f"  renombrados={ren} ya_ok={ya} errores={err_r}")

    # Conteo post-rename
    quedan_old = 0
    total = 0
    board2 = 0
    for dirpath, _dirs, files in os.walk(ROOT):
        for fn in files:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            total += 1
            if OLD_JOB_TOKEN in fn:
                quedan_old += 1
            if NEW_JOB_TOKEN in fn:
                board2 += 1
    print(f"  total_img={total} con_Board2={board2} quedan_Board1={quedan_old}")

    print("\n=== 2) Borrar filas DB Board1/Board2/BOARD2_2")
    borradas = _borrar_filas_job()
    print(f"  filas_borradas={borradas}")
    if borradas < 0:
        return 2

    job_root = os.path.dirname(os.path.dirname(ROOT))  # …/9919-BOARD2_2
    cli, prod, fuente = resolver_cliente_producto(NEW_JOB_TOKEN, job_root)
    if not cli or cli == "SIN_CLIENTE":
        prod2, cli2 = producto_cliente_desde_job_root(job_root)
        if cli2:
            cli, prod, fuente = cli2, prod2, "ruta_job"
    print(f"  cliente={cli!r} producto={prod!r} fuente={fuente}")
    guardar_contexto_dossier(
        job=NEW_JOB_TOKEN,
        cliente=cli,
        producto=prod,
        job_root=job_root,
        dossier_files=os.path.dirname(ROOT),
        dossier_jpgs=ROOT,
    )

    print("\n=== 3) Reinsertar evidencias desde JPGS")
    ins, err_i = _reinsertar(ROOT, NEW_JOB_TOKEN, cli, prod)
    print(f"  insertados={ins} errores={err_i}")
    print("LISTO.")
    return 0 if quedan_old == 0 and err_r == 0 else 3


if __name__ == "__main__":
    raise SystemExit(main())
