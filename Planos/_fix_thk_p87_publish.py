# -*- coding: utf-8 -*-
"""Reacotar+publicar solo THK P87 (Thickness=2.00 in) sin sync masivo."""
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
PIECE = "62223-1248-P87"
Y_JPGS = (
    r"Y:\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
    r"\DOSSIER FILES\JPGS"
)


def _log(msg: str) -> None:
    print(msg, flush=True)


def main() -> int:
    # Sin dossier durante el flujo (evita sync de todo el arbol).
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["COTAS_DOSSIER"] = "0"
    os.environ["PIEZAS_FILTRO"] = PIECE
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
        guardar_contexto_dossier,
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
            _log("ERROR: sin plano/ensamble")
            return 1
        try:
            plano.Activate()
        except Exception:
            pass

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        cv.configurar_producto_flujo("TANQUE")
        cv.configurar_piezas_doblado({PIECE})
        cv.configurar_piezas_corte(set())

        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        root = Path(carpeta)
        dest = root / "Corte" / "Maquinado" / "Maquinados metal" / PIECE
        dest.mkdir(parents=True, exist_ok=True)

        for p in root.rglob(f"*__{PIECE}__THK_*.jp*"):
            _log(f"DEL local {p.relative_to(root)}")
            try:
                p.unlink()
            except OSError as e:
                _log(f"  {e}")
        for base in (Path(Y_JPGS), Path(JOB_ROOT) / "DOSSIER FILES" / "JPGS"):
            if not base.exists():
                continue
            for p in base.rglob(f"*__{PIECE}__THK_*.jp*"):
                _log(f"DEL share {p}")
                try:
                    p.unlink()
                except OSError as e:
                    _log(f"  {e}")

        _log(f"flujo filtro={PIECE}")
        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta,
                incremental=False,
                catalogo_piezas={PIECE},
            )
        )
        _log(f"flujo ok={ok}")

        staging = root / "_STAGING_DESPLIEGUE"
        moved = []
        for p in list(root.glob(f"*__{PIECE}__THK_*.jp*")) + (
            list(staging.glob(f"*__{PIECE}__THK_*.jp*")) if staging.is_dir() else []
        ):
            target = dest / p.name
            if target.exists():
                target.unlink()
            p.replace(target)
            moved.append(target.name)
            _log(f"MOVE {target.name}")

        thks = sorted(dest.glob(f"*__{PIECE}__THK_*.jp*"))
        _log(f"THK en dest: {[t.name for t in thks]}")
        if not thks:
            _log("ERROR: no salio THK")
            return 2
        # Esperado: THK_2.00in (Grosor del IPT). 1.50 = brazo, incorrecto.
        bad = [t for t in thks if "THK_1.50" in t.name or "THK_0.35" in t.name]
        if bad:
            _log(f"ERROR: THK incorrecto aun: {[b.name for b in bad]}")
            return 3

        # Publicar solo esta carpeta
        os.environ["COTAS_DOSSIER"] = "1"
        iniciar_sesion_dossier(nombre_job_desde_ensamble(ensamble), JOB_ROOT)
        try:
            guardar_contexto_dossier(cliente="OTC", producto="TANQUE")
        except Exception:
            pass
        unc, n = publicar_pieza_carpeta(str(dest))
        _log(f"publish unc={unc} n={n}")
        n_db = sincronizar_lote_carpetas([str(dest)], job=JOB)
        _log(f"sync db={n_db}")

        # Espejo Y si existe y no es el mismo UNC
        y_dest = (
            Path(Y_JPGS)
            / "PIEZAS_ACOTADAS"
            / "Corte"
            / "Maquinado"
            / "Maquinados metal"
            / PIECE
        )
        if Path(Y_JPGS).exists():
            y_dest.mkdir(parents=True, exist_ok=True)
            import shutil

            for t in thks:
                shutil.copy2(t, y_dest / t.name)
                _log(f"Y copy {t.name}")

        return 0
    finally:
        try:
            cv.configurar_piezas_doblado([])
            cv.configurar_piezas_corte([])
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
