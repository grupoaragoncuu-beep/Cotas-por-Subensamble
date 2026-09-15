# -*- coding: utf-8 -*-
"""Reinserta rutas JPGS Board2 (excluye staging y Nueva carpeta)."""
from __future__ import annotations

import os
import sys

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
JOB = "9919-Board 2"
_SKIP_TOP = {
    "nueva carpeta",
    "_staging_despliegue",
    "_staging_estanado",
    "_staging_estañado",
}

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cotas_dossier_registro import (  # noqa: E402
    clasificar_type_y_spoteos,
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

    archivos = []
    for dp, dns, fs in os.walk(ROOT):
        rel = os.path.relpath(dp, ROOT)
        top = rel.split(os.sep)[0].casefold() if rel != "." else ""
        if top in _SKIP_TOP:
            dns[:] = []
            continue
        dns[:] = [d for d in dns if d.casefold() not in _SKIP_TOP]
        for fn in fs:
            if fn.lower().endswith((".jpg", ".jpeg", ".png")):
                archivos.append(os.path.join(dp, fn))
    print("jpg:", len(archivos), "cliente", cli, "producto", prod)

    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                DELETE FROM public.cotas_dossier
                WHERE job = %s
                   OR job = %s
                   OR job = %s
                   OR nombre_archivo LIKE %s
                   OR ruta ILIKE %s
                """,
                (
                    JOB,
                    "9919-BOARD2_2",
                    "9919-Board 1",
                    "9919-Board 2__%",
                    "%9919-BOARD2_2%DOSSIER FILES%JPGS%",
                ),
            )
            print("borradas:", cur.rowcount)
            ins = 0
            for ruta in archivos:
                fn = os.path.basename(ruta)
                ruta_abs = os.path.abspath(ruta)
                tipo, n_spot = clasificar_type_y_spoteos(fn)
                clase = proceso_desde_ruta(ruta_abs) or ""
                cur.execute(
                    """
                    INSERT INTO public.cotas_dossier
                        (cliente, producto, job, type, cantidad_spoteos,
                         nombre_archivo, ruta, clasificacion)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        cli,
                        prod,
                        JOB,
                        tipo,
                        n_spot,
                        fn,
                        ruta_abs,
                        clase,
                    ),
                )
                ins += 1
            conn.commit()
            print("insertadas:", ins)
            cur.execute(
                "SELECT COUNT(*) FROM public.cotas_dossier WHERE job = %s",
                (JOB,),
            )
            print("count job:", cur.fetchone()[0])
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
