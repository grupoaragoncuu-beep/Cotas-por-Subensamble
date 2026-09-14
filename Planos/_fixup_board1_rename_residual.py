# -*- coding: utf-8 -*-
"""Retry residual Board1 renames + verify DB counts."""
from __future__ import annotations

import os
import sys
import time

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
OLD, NEW = "9919-Board 1", "9919-Board 2"

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cotas_dossier_registro import (  # noqa: E402
    cargar_contexto_dossier,
    clasificar_type_y_spoteos,
    insertar_evidencia,
    proceso_desde_ruta,
    _db_nesting,
)


def main() -> int:
    left = []
    for dp, _, fs in os.walk(ROOT):
        for fn in fs:
            if fn.lower().endswith((".jpg", ".jpeg", ".png")) and OLD in fn:
                left.append(os.path.join(dp, fn))
    print("quedan Board1:", len(left))
    ctx = cargar_contexto_dossier()
    cli = ctx.get("cliente") or "GIGA"
    prod = ctx.get("producto") or "ENCLOSURES NEMA 1"
    for src in left:
        dst = src.replace(OLD, NEW)
        ok = False
        for i in range(10):
            try:
                if os.path.exists(dst):
                    os.remove(src)
                    ok = True
                    print("colision ok", os.path.basename(src))
                    break
                os.rename(src, dst)
                ok = True
                print("renamed", os.path.basename(dst))
                break
            except OSError as exc:
                print("retry", i, exc)
                time.sleep(1.5)
        if ok and os.path.isfile(dst):
            fn = os.path.basename(dst)
            job = fn.split("__", 1)[0]
            tipo, n = clasificar_type_y_spoteos(fn)
            eid = insertar_evidencia(
                cliente=cli,
                producto=prod,
                job=job,
                type_=tipo,
                cantidad_spoteos=n,
                nombre_archivo=fn,
                ruta=os.path.abspath(dst),
                clasificacion=proceso_desde_ruta(dst),
            )
            print("db id", eid, fn)

    b1 = b2 = tot = 0
    for dp, _, fs in os.walk(ROOT):
        for fn in fs:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            tot += 1
            if OLD in fn:
                b1 += 1
            if NEW in fn:
                b2 += 1
    print("total", tot, "board1", b1, "board2", b2)

    import psycopg2

    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT job, COUNT(*) AS n
                FROM public.cotas_dossier
                WHERE REPLACE(REPLACE(UPPER(job), ' ', ''), '_', '')
                      LIKE '9919BOARD%'
                GROUP BY job
                ORDER BY n DESC
                """
            )
            print("db jobs:")
            for row in cur.fetchall():
                print(" ", row)
            cur.execute(
                """
                SELECT COUNT(*) FROM public.cotas_dossier
                WHERE job = %s
                """,
                (NEW,),
            )
            print("db job exact", NEW, "->", cur.fetchone()[0])
    return 0 if b1 == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
