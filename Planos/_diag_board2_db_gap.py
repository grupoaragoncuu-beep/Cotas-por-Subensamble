# -*- coding: utf-8 -*-
from __future__ import annotations

import collections
import os
import sys

ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS"
)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cotas_dossier_registro import _db_nesting  # noqa: E402


def main() -> int:
    jobs = collections.Counter()
    names = []
    for dp, _, fs in os.walk(ROOT):
        for fn in fs:
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            names.append(fn)
            if "__" in fn:
                jobs[fn.split("__", 1)[0]] += 1
            else:
                jobs["(sin__)"] += 1
    print("archivos", len(names))
    print("tokens:")
    for k, v in jobs.most_common():
        print(repr(k), v)

    import psycopg2

    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT job, COUNT(*) AS n
                FROM public.cotas_dossier
                WHERE job ILIKE %s OR nombre_archivo ILIKE %s OR ruta ILIKE %s
                GROUP BY job
                ORDER BY n DESC
                """,
                ("%9919%", "9919%", "%9919-BOARD2_2%JPGS%"),
            )
            print("db jobs:")
            for row in cur.fetchall():
                print(row)
            cur.execute(
                """
                SELECT COUNT(*) FROM public.cotas_dossier
                WHERE ruta ILIKE %s
                """,
                ("%9919-BOARD2_2%DOSSIER FILES%JPGS%",),
            )
            print("db por ruta share:", cur.fetchone()[0])
            cur.execute(
                """
                SELECT COUNT(DISTINCT nombre_archivo) FROM public.cotas_dossier
                WHERE ruta ILIKE %s
                """,
                ("%9919-BOARD2_2%DOSSIER FILES%JPGS%",),
            )
            print("distinct nombres ruta:", cur.fetchone()[0])
            # Archivos en disco no presentes en DB
            cur.execute(
                """
                SELECT nombre_archivo FROM public.cotas_dossier
                WHERE job = %s
                """,
                ("9919-Board 2",),
            )
            en_db = {r[0] for r in cur.fetchall()}
    faltan = [n for n in names if n not in en_db]
    print("faltan en DB job Board2:", len(faltan))
    for n in faltan[:20]:
        print(" ", n)
    # duplicados en disco?
    c = collections.Counter(names)
    dups = [(n, k) for n, k in c.items() if k > 1]
    print("nombres duplicados en disco:", len(dups))
    for n, k in dups[:10]:
        print(" ", k, n)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
