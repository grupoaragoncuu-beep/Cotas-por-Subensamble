# -*- coding: utf-8 -*-
"""Tras Subensamble actual: rehacer SEGM2 con A08 + acomodar."""
from __future__ import annotations

import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(ROOT, "error_log_segm2_retry.txt")
SELECCION = os.path.join(ROOT, "seleccion_caras.json")
JOB = "62223-1246-A01"


def _log(msg: str) -> None:
    print(msg, flush=True)
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def _caras_activo() -> bool:
    try:
        out = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
                "Where-Object { $_.CommandLine -match 'generador_caras_tanque' } | "
                "Select-Object -ExpandProperty ProcessId",
            ],
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return bool(str(out or "").strip())
    except Exception:
        return False


def _correr(titulo, args):
    _log(f"INICIO: {titulo}")
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["COTAS_DOSSIER"] = "0"
    env["COTAS_DOSSIER_PUBLISH"] = "0"
    env["COTAS_DOSSIER_PICKER"] = "0"
    env["ENSAMBLES_INDEPENDIENTES"] = "0"
    t0 = time.time()
    code = subprocess.run(
        [sys.executable, "-u", *args], cwd=ROOT, env=env
    ).returncode
    _log(f"FIN: {titulo} exit={code} {(time.time()-t0)/60:.1f}min")
    return code


def main():
    open(LOG, "w", encoding="utf-8").write(
        f"SEGM2 retry @ {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
    )
    _log("Esperando fin de generador_caras_tanque...")
    while _caras_activo():
        time.sleep(20)
    time.sleep(3)
    segm2 = os.path.join(ROOT, "JPG", JOB, "COTAS_POR_REFERENCIA", "SEGM2")
    n = 0
    if os.path.isdir(segm2):
        for _r, _d, files in os.walk(segm2):
            n += sum(1 for x in files if x.lower().endswith(".jpg"))
    _log(f"SEGM2 actual: {n} JPG")
    codes = []
    if n < 10:
        codes.append(
            _correr(
                "SEGM2 via A08",
                [
                    "generador_caras_tanque.py",
                    "--seleccion",
                    SELECCION,
                    "--solo",
                    "SEGM2",
                ],
            )
        )
    else:
        _log("SEGM2 ya tiene JPG suficientes; no reintento.")
    codes.append(
        _correr(
            "Acomodar",
            ["_acomodar_salida_tanque.py", "--job", JOB],
        )
    )
    _log(f"RESUMEN {codes}")
    return 0 if all(c == 0 for c in codes) else 1


if __name__ == "__main__":
    sys.exit(main())
