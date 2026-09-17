# -*- coding: utf-8 -*-
"""
Espera a que termine Abigail (generador_piezas) y encadena:
  1) acomodar PIEZAS
  2) Subensamble (COTAS_POR_REFERENCIA) con fix SEGM2
  3) acomodar de nuevo

No relanza Ensambles (ya OK). Sin dossier.
BASE puede quedar vacía (nada que acotar).
"""
from __future__ import annotations

import os
import subprocess
import sys
import time


ROOT = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(ROOT, "error_log_continuar_subensamble.txt")
SELECCION = os.path.join(ROOT, "seleccion_caras.json")
JOB = "62223-1246-A01"
TERM_ABIGAIL = os.path.normpath(
    os.path.join(
        os.path.dirname(ROOT),
        "..",
        "..",
        "Users",
        "jose_rosales",
        ".cursor",
        "projects",
        "c-Proyectos-COTAS-ABIGAIL-COTAS-ABIGAIL",
        "terminals",
        "287866.txt",
    )
)
# Ruta Cursor local típica:
TERM_CANDIDATES = [
    r"C:\Users\jose_rosales\.cursor\projects\c-Proyectos-COTAS-ABIGAIL-COTAS-ABIGAIL\terminals\287866.txt",
]


def _log(msg: str) -> None:
    linea = str(msg)
    print(linea, flush=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
            f.flush()
    except Exception:
        pass


def _python_generador_activo() -> bool:
    try:
        import subprocess as sp

        out = sp.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
                "Where-Object { $_.CommandLine -match 'generador_piezas' } | "
                "Select-Object -ExpandProperty ProcessId",
            ],
            text=True,
            stderr=sp.DEVNULL,
        )
        return bool(str(out or "").strip())
    except Exception:
        return False


def _abigail_exit_code() -> int | None:
    for path in TERM_CANDIDATES:
        if not os.path.isfile(path):
            continue
        try:
            with open(path, encoding="utf-8", errors="replace") as f:
                text = f.read()
        except Exception:
            continue
        if "exit_code:" in text:
            for line in text.strip().splitlines()[-15:]:
                if line.startswith("exit_code:"):
                    try:
                        return int(line.split(":", 1)[1].strip())
                    except Exception:
                        return None
        # Header still running?
        head = "\n".join(text.splitlines()[:8])
        if "status: running" in head:
            return None
    return None


def _correr(titulo: str, args: list[str]) -> int:
    _log("")
    _log("=" * 64)
    _log(f" INICIO: {titulo}")
    _log("=" * 64)
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["COTAS_DOSSIER"] = "0"
    env["COTAS_DOSSIER_PUBLISH"] = "0"
    env["COTAS_DOSSIER_PICKER"] = "0"
    env["ENSAMBLES_INDEPENDIENTES"] = "0"
    env["PIEZAS_INCREMENTAL"] = "0"
    t0 = time.time()
    proc = subprocess.run(
        [sys.executable, "-u", *args], cwd=ROOT, env=env
    )
    _log(
        f" FIN: {titulo} | exit={proc.returncode} | "
        f"{(time.time() - t0) / 60.0:.1f} min"
    )
    return int(proc.returncode)


def main() -> int:
    try:
        with open(LOG, "w", encoding="utf-8") as f:
            f.write(
                f"Continuar Subensamble @ {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            )
    except Exception:
        pass

    _log("Esperando fin de Abigail (generador_piezas)...")
    while True:
        code = _abigail_exit_code()
        activo = _python_generador_activo()
        if code is not None and not activo:
            _log(f"Abigail terminó con exit={code}")
            if code != 0:
                _log("ERROR: Abigail no OK; no se lanza Subensamble.")
                return code
            break
        if not activo and code is None:
            # Proceso ya no corre; dar un respiro y re-chequear terminal.
            time.sleep(5)
            code2 = _abigail_exit_code()
            if code2 is not None:
                if code2 != 0:
                    _log(f"ERROR: Abigail exit={code2}")
                    return code2
                break
            # Sin terminal claro pero sin proceso: asumir listo si hay JPG.
            piezas = os.path.join(ROOT, "JPG", JOB, "PIEZAS_ACOTADAS")
            n = 0
            if os.path.isdir(piezas):
                for _r, _d, files in os.walk(piezas):
                    n += sum(1 for x in files if x.lower().endswith(".jpg"))
            if n > 50:
                _log(f"Abigail parece OK ({n} JPG en PIEZAS); continúo.")
                break
            _log("AVISO: sin proceso Abigail y pocos JPG; reintento espera...")
        time.sleep(30)

    codes = []
    codes.append(
        _correr("Acomodar PIEZAS post-Abigail", ["_acomodar_salida_tanque.py", "--job", JOB])
    )
    codes.append(
        _correr(
            "Subensamble (COTAS_POR_REFERENCIA) + fix SEGM2",
            ["generador_caras_tanque.py", "--seleccion", SELECCION],
        )
    )
    # Si SEGM2 quedó flojo (A08 no estaba en el fix aún), rehacer solo SEGM2.
    segm2_dir = os.path.join(ROOT, "JPG", JOB, "COTAS_POR_REFERENCIA", "SEGM2")
    n_segm2 = 0
    if os.path.isdir(segm2_dir):
        for _r, _d, files in os.walk(segm2_dir):
            n_segm2 += sum(1 for x in files if x.lower().endswith(".jpg"))
    if n_segm2 < 10:
        _log(
            f"SEGM2 con solo {n_segm2} JPG → reintento --solo SEGM2 "
            "(A08 como 4ª pared)."
        )
        codes.append(
            _correr(
                "Subensamble solo SEGM2 (A08)",
                [
                    "generador_caras_tanque.py",
                    "--seleccion",
                    SELECCION,
                    "--solo",
                    "SEGM2",
                ],
            )
        )
    codes.append(
        _correr("Acomodar final", ["_acomodar_salida_tanque.py", "--job", JOB])
    )
    _log(f"RESUMEN exit codes: {codes}")
    # Ensambles no se tocan (ya OK). BASE vacía es aceptable.
    return 0 if all(c == 0 for c in codes) else 1


if __name__ == "__main__":
    sys.exit(main())
