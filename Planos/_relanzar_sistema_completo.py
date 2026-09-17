# -*- coding: utf-8 -*-
"""
Relanza el sistema completo (Abigail + Subensamble + Ensambles instructivo).

Requiere Inventor abierto con machote + ensamble del tanque.
Uso:
  python _relanzar_sistema_completo.py
  python _relanzar_sistema_completo.py --seleccion seleccion_caras.json
"""
from __future__ import annotations

import os
import subprocess
import sys
import time


ROOT = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(ROOT, "error_log_relanzar_completo.txt")


def _log(msg: str) -> None:
    linea = str(msg)
    print(linea, flush=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(linea + "\n")
            f.flush()
    except Exception:
        pass


def _activar_contexto() -> bool:
    sys.path.insert(0, ROOT)
    import pythoncom
    from inventor_com import conectar_inventor
    from generador_caras_tanque import (
        _obtener_ensamble_principal,
        _obtener_plano_activo,
    )

    pythoncom.CoInitialize()
    inv = conectar_inventor()
    plano = _obtener_plano_activo(inv)
    ensamble = _obtener_ensamble_principal(inv)
    if plano is None or ensamble is None:
        _log("ERROR: falta machote o ensamble principal abierto.")
        return False
    try:
        plano.Activate()
    except Exception:
        pass
    _log(f"Plano: {plano.DisplayName}")
    _log(f"Ensamble: {ensamble.DisplayName}")
    return True


def _correr(titulo: str, args: list[str]) -> int:
    _log("")
    _log("=" * 64)
    _log(f" INICIO: {titulo}")
    _log("=" * 64)
    t0 = time.time()
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PIEZAS_INCREMENTAL"] = os.environ.get("PIEZAS_INCREMENTAL", "0")
    # Instructivo iLogic (no MVP bbox legacy).
    env["ENSAMBLES_INDEPENDIENTES"] = "0"
    # Sin share/DB Nesting: solo JPG locales (velocidad).
    env["COTAS_DOSSIER"] = "0"
    env["COTAS_DOSSIER_PUBLISH"] = "0"
    env["COTAS_DOSSIER_PICKER"] = "0"
    proc = subprocess.run(
        [sys.executable, "-u", *args],
        cwd=ROOT,
        env=env,
    )
    dt = time.time() - t0
    _log(
        f" FIN: {titulo} | exit={proc.returncode} | "
        f"{dt / 60.0:.1f} min"
    )
    return int(proc.returncode)


def _job_desde_seleccion(seleccion: str) -> str:
    try:
        import json

        with open(seleccion, encoding="utf-8") as f:
            data = json.load(f)
        ens = str(data.get("ensamble") or "").strip()
        if ens.lower().endswith(".iam"):
            ens = ens[:-4]
        return ens
    except Exception:
        return ""


def main() -> int:
    seleccion = os.path.join(ROOT, "seleccion_caras.json")
    argv = sys.argv[1:]
    for i, a in enumerate(argv):
        if a == "--seleccion" and i + 1 < len(argv):
            seleccion = argv[i + 1]
        elif a.startswith("--seleccion="):
            seleccion = a.split("=", 1)[1]

    if not os.path.isfile(seleccion):
        _log(f"ERROR: no existe selección {seleccion}")
        return 1

    job = _job_desde_seleccion(seleccion)

    try:
        with open(LOG, "w", encoding="utf-8") as f:
            f.write(f"Relanzar completo @ {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"seleccion={seleccion}\n")
            f.write(f"job={job}\n")
    except Exception:
        pass

    if not _activar_contexto():
        return 1

    pasos = [
        (
            "Abigail (PIEZAS_ACOTADAS)",
            ["generador_piezas.py", "--seleccion", seleccion],
        ),
        (
            "Subensamble (COTAS_POR_REFERENCIA)",
            ["generador_caras_tanque.py", "--seleccion", seleccion],
        ),
        (
            "Ensambles independientes (instructivo)",
            [
                "generador_ensambles_instructivo.py",
                "--seleccion",
                seleccion,
                "--limpiar",
            ],
        ),
    ]

    codigos = []
    for titulo, args in pasos:
        code = _correr(titulo, args)
        codigos.append(code)
        if code != 0:
            _log(f"AVISO: {titulo} terminó con error; se continúa.")

    _log("")
    _log(" Acomodando carpetas / fotos de salida...")
    ac_args = ["_acomodar_salida_tanque.py"]
    if job:
        ac_args.extend(["--job", job])
    code_ac = _correr("Acomodar salida JPG", ac_args)
    codigos.append(code_ac)

    _log("")
    _log(f"RESUMEN exit codes: {codigos}")
    return 0 if all(c == 0 for c in codigos) else 1


if __name__ == "__main__":
    sys.exit(main())
