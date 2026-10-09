# -*- coding: utf-8 -*-
"""Reacotar THK P87 + FPP-PELSUE con Plan A (Sheet Metal Thickness manda)."""
from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

JOB = "62223-1246-A01"
JOB_ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
)
THK_TARGETS = ("62223-1248-P87", "FPP-PELSUE")


def _log(msg: str) -> None:
    print(msg, flush=True)


def main() -> int:
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["COTAS_DOSSIER"] = "1"
    os.environ["PIEZAS_FILTRO"] = ",".join(THK_TARGETS)
    os.environ["PIEZAS_INCREMENTAL"] = "1"
    os.environ["PIEZAS_TRAZABLE"] = "0"
    os.environ["PIEZAS_WIPE_JOB"] = "0"
    os.environ["ENSAMBLES_INDEPENDIENTES"] = "0"
    os.environ["SKIP_ENSAMBLES_IND"] = "1"
    os.environ.pop("SOLO_FLAT_CORTE", None)

    import pythoncom
    from inventor_com import conectar_inventor
    from generador_caras_tanque import (
        _carpeta_salida_tanque,
        _obtener_ensamble_principal,
        _obtener_plano_activo,
    )
    from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS
    from generador_vistas import ejecutar_flujo_desde_app
    from cotas_dossier_registro import (
        iniciar_sesion_dossier,
        publicar_pieza_carpeta,
        sincronizar_lote_carpetas,
    )
    from nomenclatura_capturas import nombre_job_desde_ensamble
    from producto_tipo import clasificar_producto, aplicar_unidad_producto
    import creador_vistas as cv

    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            _log("ERROR: sin plano/ensamble en Inventor")
            return 1
        try:
            plano.Activate()
        except Exception:
            pass

        job_tok = nombre_job_desde_ensamble(ensamble)
        iniciar_sesion_dossier(job_tok, JOB_ROOT)

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        cv.configurar_producto_flujo("TANQUE")
        # Chapas maquinadas: THK por LADO (Thickness), no flat-only.
        cv.configurar_piezas_doblado(set(THK_TARGETS))
        cv.configurar_piezas_corte(set())

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)
        _log(f"filtro={os.environ['PIEZAS_FILTRO']}")
        _log(f"salida={carpeta_piezas}")

        # Borrar THK viejos (0.35 / 0.75) para no mezclar.
        root = Path(carpeta_piezas)
        for piece in THK_TARGETS:
            for p in root.rglob(f"*__{piece}__THK_*.jp*"):
                _log(f"DEL old {p.relative_to(root)}")
                try:
                    p.unlink()
                except OSError as exc:
                    _log(f"  AVISO unlink: {exc}")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=False,
                catalogo_piezas=set(THK_TARGETS),
            )
        )
        _log(f"flujo ok={ok}")

        dest_base = root / "Corte" / "Maquinado" / "Maquinados metal"
        published = []
        for piece in THK_TARGETS:
            dest = dest_base / piece
            dest.mkdir(parents=True, exist_ok=True)
            for p in list(root.glob(f"*__{piece}__THK_*.jp*")):
                target = dest / p.name
                if p.resolve() != target.resolve():
                    if target.exists():
                        target.unlink()
                    p.replace(target)
                    _log(f"MOVE {p.name} -> {dest.relative_to(root)}")
            for st in ("_STAGING_DESPLIEGUE", "_STAGING_ESTANIADO"):
                sd = root / st
                if sd.is_dir():
                    for p in sd.glob(f"*__{piece}__THK_*.jp*"):
                        target = dest / p.name
                        if target.exists():
                            target.unlink()
                        p.replace(target)
                        _log(f"MOVE staging {p.name}")
            thks = sorted(dest.glob(f"*__{piece}__THK_*.jp*"))
            _log(f"{piece} THK files: {[t.name for t in thks]}")
            if thks:
                published.append(str(dest))

        for d in published:
            try:
                publicar_pieza_carpeta(d, job=JOB)
            except Exception as exc:
                _log(f"AVISO publish {d}: {exc}")
        if published:
            try:
                sincronizar_lote_carpetas(published, job=JOB)
            except Exception as exc:
                _log(f"AVISO sync DB: {exc}")

        return 0 if ok else 2
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
