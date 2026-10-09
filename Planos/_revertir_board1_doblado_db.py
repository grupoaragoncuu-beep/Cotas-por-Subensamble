# -*- coding: utf-8 -*-
"""Revert DB: Doblado/Metal -> Corte/Plasma y Laser/Corte metal for Board1 metal flat pieces."""
from __future__ import annotations

import json
import re
import sys

sys.path.insert(0, r"c:\Proyectos\COTAS\Planos")
import psycopg2
from psycopg2.extras import RealDictCursor
from cotas_dossier_registro import _db_nesting

MAPA = r"c:\Proyectos\COTAS\Planos\JPG\9919-Board 1\.piezas_por_clasificacion.json"
# Pieces I wrongly moved (exclude FB which stayed in Metal)
piezas = set(json.load(open(MAPA, encoding="utf-8"))["Doblado"]) - {"FB-10-10-A"}

OLD = re.compile(r"(?i)Doblado[\\/]+Metal[\\/]+")
NEW = r"Corte\\Plasma y Laser\\Corte metal\\"


def main() -> int:
    with psycopg2.connect(**_db_nesting(), cursor_factory=RealDictCursor) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT id, job, clasificacion, ruta, nombre_archivo
                FROM public.cotas_dossier
                WHERE (ruta ILIKE %s OR ruta ILIKE %s OR job ILIKE %s OR job ILIKE %s)
                  AND ruta ILIKE %s
                """,
                (
                    "%9919-Board 1%",
                    "%9919-BOARD1%",
                    "%Board 1%",
                    "%BOARD1%",
                    "%Doblado%Metal%",
                ),
            )
            rows = cur.fetchall()
            updated = 0
            skipped = 0
            for r in rows:
                ruta = r["ruta"] or ""
                na = r["nombre_archivo"] or ""
                hit = None
                norm = ruta.replace("/", "\\")
                for p in piezas:
                    if f"\\{p}\\" in norm or norm.rstrip("\\").endswith(f"\\{p}"):
                        hit = p
                        break
                    if p in na or f"__{p}__" in na:
                        hit = p
                        break
                if not hit:
                    skipped += 1
                    continue
                new_ruta = OLD.sub(NEW, norm)
                if "/" in ruta and "\\" not in ruta:
                    new_ruta = new_ruta.replace("\\", "/")
                if new_ruta == ruta:
                    skipped += 1
                    continue
                cur.execute(
                    """
                    UPDATE public.cotas_dossier
                    SET ruta = %s,
                        clasificacion = %s
                    WHERE id = %s
                    """,
                    (new_ruta, "Corte/Plasma y Laser/Corte metal", r["id"]),
                )
                updated += 1
            conn.commit()
            print(f"REVERT DB UPDATED={updated} SKIPPED={skipped} (rows scanned={len(rows)})")

            cur.execute(
                """
                SELECT COUNT(*) AS n FROM public.cotas_dossier
                WHERE (ruta ILIKE %s OR job = %s)
                  AND ruta ILIKE %s
                """,
                ("%9919-Board 1%", "9919-Board 1", "%Doblado%Metal%"),
            )
            print(f"Remain Doblado/Metal job Board 1: {cur.fetchone()['n']}")
            cur.execute(
                """
                SELECT COUNT(*) AS n FROM public.cotas_dossier
                WHERE job = %s AND ruta ILIKE %s
                """,
                ("9919-Board 1", "%Corte%Plasma%Corte metal%"),
            )
            print(f"Corte metal job Board 1: {cur.fetchone()['n']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
