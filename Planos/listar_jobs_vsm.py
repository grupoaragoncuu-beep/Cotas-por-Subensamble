# -*- coding: utf-8 -*-
"""Lista jobs en proceso del VSM, una linea TSV por job."""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cotas_dossier_registro import listar_jobs_vsm_activos


def main() -> int:
    for row in listar_jobs_vsm_activos():
        campos = [
            row.get("job_number") or "",
            row.get("client") or "",
            row.get("product") or "",
            row.get("status") or "",
            row.get("job_root") or "",
            row.get("dossier_jpgs") or "",
        ]
        print("\t".join(campos))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
