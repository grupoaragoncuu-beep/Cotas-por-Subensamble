# -*- coding: utf-8 -*-
"""Limpia residuales Board1 en DB y re-sincroniza los 987 JPG Board2."""
from __future__ import annotations

import os
import sys

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
JOB = "9919-Board 2"

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cotas_dossier_registro import (  # noqa: E402
    clasificar_type_y_spoteos,
    insertar_evidencia,
    proceso_desde_ruta,
    producto_cliente_desde_job_root,
    _db_nesting,
)


def main() -> int:
    import psycopg2

    job_root = os.path.dirname(os.path.dirname(ROOT))
    prod, cli = producto_cliente_desde_job_root(job_root)
    if not cli:
        cli, prod = "GIGA", "ENCLOSURES NEMA 1"
    print("cliente", cli, "producto", prod)

    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                DELETE FROM public.cotas_dossier
                WHERE job = %s
                   OR nombre_archivo LIKE %s
                RETURNING id
                """,
                ("9919-Board 1", "9919-Board 1__%"),
            )
            n1 = len(cur.fetchall() or [])
            cur.execute(
                """
                DELETE FROM public.cotas_dossier
                WHERE job = %s
                   OR job = %s
                   OR job = %s
                RETURNING id
                """,
                (JOB, "9919-BOARD2_2", "9919-Board 2"),
            )
            n2 = len(cur.fetchall() or [])
        conn.commit()
    print("borradas Board1-ish:", n1, "borradas Board2-ish:", n2)

    archivos = []
    for dp, _, fs in os.walk(ROOT):
        for fn in fs:
            if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                archivos.append(os.path.join(dp, fn))
    print("jpg en share:", len(archivos))

    ins = err = 0
    for ruta in archivos:
        fn = os.path.basename(ruta)
        job = fn.split("__", 1)[0].strip() if "__" in fn else JOB
        tipo, n = clasificar_type_y_spoteos(fn)
        eid = insertar_evidencia(
            cliente=cli,
            producto=prod,
            job=job,
            type_=tipo,
            cantidad_spoteos=n,
            nombre_archivo=fn,
            ruta=os.path.abspath(ruta),
            clasificacion=proceso_desde_ruta(ruta),
        )
        if eid:
            ins += 1
        else:
            err += 1
    print("insertados", ins, "errores", err)

    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT job, COUNT(*) AS n
                FROM public.cotas_dossier
                WHERE job ILIKE %s OR job ILIKE %s OR nombre_archivo ILIKE %s
                GROUP BY job
                ORDER BY n DESC
                """,
                ("%9919%Board%2%", "%9919%BOARD%2%", "9919-Board 2__%"),
            )
            print("db resumen:")
            for row in cur.fetchall():
                print(" ", row)
            cur.execute(
                "SELECT COUNT(*) FROM public.cotas_dossier WHERE job = %s",
                (JOB,),
            )
            print("count exact", JOB, cur.fetchone()[0])
            cur.execute(
                """
                SELECT clasificacion, COUNT(*)
                FROM public.cotas_dossier WHERE job = %s
                GROUP BY clasificacion ORDER BY 2 DESC
                """,
                (JOB,),
            )
            print("por clasificacion:")
            for row in cur.fetchall():
                print(" ", row)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
