# -*- coding: utf-8 -*-
"""Actualiza rutas DB tras mover Doblado metal Board 1 Corte→Doblado/Metal."""
from __future__ import annotations

import json
import re
import sys

sys.path.insert(0, r"c:\Proyectos\COTAS\Planos")

import psycopg2
from psycopg2.extras import RealDictCursor

from cotas_dossier_registro import _db_nesting

MAPA = r"c:\Proyectos\COTAS\Planos\JPG\9919-Board 1\.piezas_por_clasificacion.json"
OLD = re.compile(
    r"(?i)(Corte[\\/]+Plasma y Laser[\\/]+Corte metal[\\/]+)"
)
NEW = r"Doblado\\Metal\\"
# Also normalize forward slashes variants in stored paths
OLD_FWD = re.compile(
    r"(?i)(Corte/+Plasma y Laser/+Corte metal/+)"
)


def main() -> int:
    piezas = set(json.load(open(MAPA, encoding="utf-8"))["Doblado"])
    print(f"Piezas Doblado mapa: {len(piezas)}")

    with psycopg2.connect(**_db_nesting(), cursor_factory=RealDictCursor) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT job, COUNT(*) AS n
                FROM public.cotas_dossier
                WHERE job ILIKE %s
                   OR job ILIKE %s
                   OR ruta ILIKE %s
                   OR ruta ILIKE %s
                GROUP BY job
                ORDER BY n DESC
                LIMIT 20
                """,
                ("%Board 1%", "%BOARD1%", "%9919-Board 1%", "%9919-BOARD1%"),
            )
            print("JOBS:")
            for r in cur.fetchall():
                print(f"  {r['job']}: {r['n']}")

            cur.execute(
                """
                SELECT id, job, clasificacion, ruta, nombre_archivo
                FROM public.cotas_dossier
                WHERE (
                    ruta ILIKE %s OR ruta ILIKE %s
                    OR job ILIKE %s OR job ILIKE %s
                )
                  AND (
                    ruta ILIKE %s
                    OR ruta ILIKE %s
                  )
                """,
                (
                    "%9919-Board 1%",
                    "%9919-BOARD1%",
                    "%Board 1%",
                    "%BOARD1%",
                    "%Corte%Plasma%Corte metal%",
                    "%Corte/Plasma y Laser/Corte metal%",
                ),
            )
            rows = cur.fetchall()
            print(f"Filas en Corte metal (Board1-ish): {len(rows)}")

            updated = 0
            skipped = 0
            for r in rows:
                ruta = r["ruta"] or ""
                # Solo piezas del mapa Doblado
                hit = None
                for p in piezas:
                    if f"\\{p}\\" in ruta.replace("/", "\\") or ruta.replace(
                        "/", "\\"
                    ).endswith(f"\\{p}") or f"/{p}/" in ruta:
                        hit = p
                        break
                    # nombre_archivo contains piece
                    na = r["nombre_archivo"] or ""
                    if p in na or p in ruta:
                        # verify path segment
                        norm = ruta.replace("/", "\\")
                        if f"\\{p}\\" in norm or norm.rstrip("\\").endswith(f"\\{p}"):
                            hit = p
                            break
                        if f"__{p}__" in na or na.startswith(p) or f" {p}" in na:
                            hit = p
                            break
                if not hit:
                    skipped += 1
                    continue

                new_ruta = OLD.sub(NEW, ruta.replace("/", "\\"))
                # keep original slash style roughly: if ruta used /, convert back
                if "/" in ruta and "\\" not in ruta:
                    new_ruta = new_ruta.replace("\\", "/")
                elif "/" in ruta:
                    # mixed — prefer backslash UNC style as stored often uses \
                    pass

                if new_ruta == ruta:
                    skipped += 1
                    continue

                # clasificacion: prefer Doblado/Metal or Doblado
                clase = (r["clasificacion"] or "").strip()
                if "corte" in clase.lower() or "plasma" in clase.lower() or not clase:
                    clase_new = "Doblado/Metal"
                elif clase.lower() in ("doblado",):
                    clase_new = "Doblado/Metal"
                else:
                    clase_new = clase

                cur.execute(
                    """
                    UPDATE public.cotas_dossier
                    SET ruta = %s,
                        clasificacion = %s
                    WHERE id = %s
                    """,
                    (new_ruta, clase_new, r["id"]),
                )
                updated += 1

            conn.commit()
            print(f"UPDATED={updated} SKIPPED={skipped}")

            # verify remaining
            cur.execute(
                """
                SELECT COUNT(*) AS n
                FROM public.cotas_dossier
                WHERE (ruta ILIKE %s OR ruta ILIKE %s OR job ILIKE %s OR job ILIKE %s)
                  AND ruta ILIKE %s
                """,
                (
                    "%9919-Board 1%",
                    "%9919-BOARD1%",
                    "%Board 1%",
                    "%BOARD1%",
                    "%Corte%Plasma%Corte metal%",
                ),
            )
            print(f"Remain Corte metal Board1-ish: {cur.fetchone()['n']}")

            cur.execute(
                """
                SELECT COUNT(*) AS n
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
            print(f"Now Doblado/Metal Board1-ish: {cur.fetchone()['n']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
