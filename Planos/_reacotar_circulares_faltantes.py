# -*- coding: utf-8 -*-
"""Re-corrida rapida: piezas circulares con Ø/HOLE/THK faltantes (62223)."""
from __future__ import annotations

import os
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

JOB = "62223-1246-A01"
JOB_ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
)
Y_JPGS = (
    r"Y:\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
    r"\DOSSIER FILES\JPGS"
)
TARGETS = (
    "SP-767_2",
)

EXPECTED_OUTER_OD = {
    "SP-767_2": 8.0,
}


def _log(msg: str) -> None:
    print(msg, flush=True)


def _dedupe_circular_tags(carpeta_piezas: str) -> int:
    """Limpia etiquetas basura tipicas de bridas/discos.

    - HEIGHT con mismo valor que un OD existente → borrar (copia del Ø).
    - HEIGHT mayor que OD existentes → renombrar a OD.
    - LEG con mismo valor que THK → borrar (mismo JPG de espesor).
    - HOLE con mismo valor que OD/ID, o Ø grande de perfil → borrar/renombrar ID.
    - HOLE02+ duplicando HOLE01 mismo valor → borrar extras.
    """
    root = Path(carpeta_piezas)
    n = 0
    for piece in TARGETS:
        ods = list(root.rglob(f"*__{piece}__OD_*.jp*"))
        ids = list(root.rglob(f"*__{piece}__ID_*.jp*"))
        thks = list(root.rglob(f"*__{piece}__THK_*.jp*"))
        od_vals = []
        for p in ods:
            m = re.search(r"OD_([0-9.]+)in", p.name, re.I)
            if m:
                od_vals.append(float(m.group(1)))
        id_vals = []
        for p in ids:
            m = re.search(r"ID_([0-9.]+)in", p.name, re.I)
            if m:
                id_vals.append(float(m.group(1)))
        thk_vals = []
        for p in thks:
            m = re.search(r"THK_([0-9.]+)in", p.name, re.I)
            if m:
                thk_vals.append(float(m.group(1)))
        max_od = max(od_vals) if od_vals else 0.0

        for p in list(root.rglob(f"*__{piece}__HEIGHT_*.jp*")):
            m = re.search(r"HEIGHT_([0-9.]+)in", p.name, re.I)
            if not m:
                continue
            val = float(m.group(1))
            if any(abs(val - ov) <= 0.05 for ov in od_vals):
                p.unlink(missing_ok=True)
                _log(f"DEL dup HEIGHT==OD {p.name}")
                n += 1
                continue
            if od_vals and val > max(od_vals) + 0.05:
                new_name = re.sub(r"__HEIGHT_", "__OD_", p.name, count=1, flags=re.I)
                dest = p.with_name(new_name)
                if dest.exists():
                    p.unlink(missing_ok=True)
                    _log(f"DEL dup HEIGHT {p.name}")
                else:
                    p.rename(dest)
                    _log(f"REN {p.name} -> {dest.name}")
                    od_vals.append(val)
                    max_od = max(od_vals)
                n += 1

        for p in list(root.rglob(f"*__{piece}__LEG_*.jp*")):
            m = re.search(r"LEG_([0-9.]+)in", p.name, re.I)
            if not m:
                continue
            val = float(m.group(1))
            if any(abs(val - tv) <= 0.05 for tv in thk_vals):
                p.unlink(missing_ok=True)
                _log(f"DEL dup LEG==THK {p.name}")
                n += 1

        # HOLE mal etiquetados / duplicados TYP
        holes = list(root.rglob(f"*__{piece}__HOLE*.jp*"))
        by_val: dict[float, list] = {}
        for p in holes:
            m = re.search(r"HOLE\d*_([0-9.]+)in", p.name, re.I)
            if not m:
                continue
            val = round(float(m.group(1)), 2)
            by_val.setdefault(val, []).append(p)
        for val, group in by_val.items():
            is_big = (max_od and val >= max_od * 0.35) or (
                val > 1.5 and max_od and val >= max_od * 0.25
            )
            if any(abs(val - ov) <= 0.06 for ov in od_vals) or any(
                abs(val - iv) <= 0.06 for iv in id_vals
            ):
                for p in group:
                    p.unlink(missing_ok=True)
                    _log(f"DEL HOLE==OD/ID {p.name}")
                    n += 1
                continue
            if is_big:
                first = sorted(group, key=lambda x: x.name)[0]
                new_name = re.sub(r"__HOLE\d+_", "__ID_", first.name, count=1, flags=re.I)
                dest = first.with_name(new_name)
                if dest.exists():
                    first.unlink(missing_ok=True)
                    _log(f"DEL big HOLE (ID exists) {first.name}")
                else:
                    first.rename(dest)
                    _log(f"REN big HOLE->ID {dest.name}")
                    id_vals.append(val)
                n += 1
                for p in group:
                    if p.exists() and p != dest:
                        p.unlink(missing_ok=True)
                        _log(f"DEL extra big HOLE {p.name}")
                        n += 1
                continue
            # TYP chico: dejar un solo HOLE01
            group_sorted = sorted(
                group, key=lambda x: (0 if "HOLE01" in x.name.upper() else 1, x.name)
            )
            keep = group_sorted[0]
            if "HOLE01" not in keep.name.upper():
                new_name = re.sub(r"__HOLE\d+_", "__HOLE01_", keep.name, count=1, flags=re.I)
                dest = keep.with_name(new_name)
                if not dest.exists():
                    keep.rename(dest)
                    _log(f"REN {keep.name} -> {dest.name}")
                    n += 1
                    keep = dest
            for p in group_sorted[1:]:
                if p.exists():
                    p.unlink(missing_ok=True)
                    _log(f"DEL dup TYP HOLE {p.name}")
                    n += 1
    return n


def _wipe_target_jpgs(carpeta_piezas: str) -> int:
    """Borra JPG viejos de las piezas target (local) para no mezclar basura."""
    root = Path(carpeta_piezas)
    n = 0
    for piece in TARGETS:
        for p in list(root.rglob(f"*__{piece}__*.jp*")):
            try:
                p.unlink()
                n += 1
            except Exception:
                pass
    y_root = Path(Y_JPGS) / "PIEZAS_ACOTADAS"
    if y_root.exists():
        for piece in TARGETS:
            for p in list(y_root.rglob(f"*__{piece}__*.jp*")):
                try:
                    p.unlink()
                    n += 1
                except Exception:
                    pass
    return n


def _audit_expected_od(carpeta_piezas: str) -> list[str]:
    """Devuelve piezas cuyo OD max sigue por debajo del esperado."""
    root = Path(carpeta_piezas)
    mal = []
    for piece, exp in EXPECTED_OUTER_OD.items():
        vals = []
        for p in root.rglob(f"*__{piece}__OD_*.jp*"):
            m = re.search(r"OD_([0-9.]+)in", p.name, re.I)
            if m:
                vals.append(float(m.group(1)))
        if not vals or max(vals) + 0.15 < exp:
            mal.append(f"{piece}: OD={vals or 'NADA'} esperaba>={exp}")
    return mal


def main() -> int:
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["COTAS_DOSSIER"] = "0"  # evitar sync masivo durante el flujo
    os.environ["PIEZAS_FILTRO"] = ",".join(TARGETS)
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
        # Forzar ensamble 62223-1246-A01 (CastTo AssemblyDocument).
        from generador_caras_tanque import _como_ensamble

        ensamble = None
        for i in range(1, int(inv.Documents.Count) + 1):
            try:
                d = inv.Documents.Item(i)
                nom = str(getattr(d, "DisplayName", "") or "")
                if nom.upper().startswith("62223-1246-A01") and nom.upper().endswith(
                    ".IAM"
                ):
                    ensamble = _como_ensamble(d)
                    try:
                        ensamble.Activate()
                    except Exception:
                        pass
                    break
            except Exception:
                continue
        if ensamble is None:
            ensamble = _obtener_ensamble_principal(inv)
        plano = _obtener_plano_activo(inv)
        if plano is None or ensamble is None:
            _log("ERROR: sin plano/ensamble")
            return 1
        try:
            plano.Activate()
        except Exception:
            pass
        _log(f"ensamble={getattr(ensamble, 'DisplayName', '?')}")
        _log(f"plano={getattr(plano, 'DisplayName', '?')}")
        # Verificar ComponentDefinition antes de wipe/flujo
        try:
            _ = ensamble.ComponentDefinition
        except Exception as exc:
            _log(f"ERROR: ensamble sin ComponentDefinition: {exc}")
            return 1

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        cv.configurar_producto_flujo("TANQUE")
        # Circulares/bridas: permitir LADO+DESPLIEGUE (THK canto + Ø cara)
        cv.configurar_piezas_doblado(set(TARGETS))
        cv.configurar_piezas_corte(set())

        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        _log(f"filtro={os.environ['PIEZAS_FILTRO']}")
        _log(f"salida={carpeta}")

        n_wipe = _wipe_target_jpgs(carpeta)
        _log(f"wipe jpgs previos: {n_wipe}")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta,
                incremental=False,
                catalogo_piezas=set(TARGETS),
            )
        )
        _log(f"flujo ok={ok}")

        n_ren = _dedupe_circular_tags(carpeta)
        _log(f"dedupe circular tags: {n_ren}")
        mal_od = _audit_expected_od(carpeta)
        if mal_od:
            _log("AUDIT OD FALTANTE:")
            for line in mal_od:
                _log(f"  XX {line}")
        else:
            _log("AUDIT OD: todas las piezas alcanzan Ø exterior esperado")

        # Mover JPG sueltos / staging a carpetas de pieza.
        # Preferir Maquinado (revisor OTC) sobre Doblado si ambas existen.
        root = Path(carpeta)
        staging = root / "_STAGING_DESPLIEGUE"
        for piece in TARGETS:
            existing = [
                p
                for p in root.rglob(piece)
                if p.is_dir() and piece == p.name and "_STAGING" not in str(p)
            ]

            def _rank(p: Path) -> tuple:
                s = str(p).lower()
                if "maquinado" in s:
                    return (0, s)
                if "doblado" in s:
                    return (2, s)
                return (1, s)

            if existing:
                existing = sorted(existing, key=_rank)
                dest = existing[0]
            else:
                dest = root / "Corte" / "Maquinado" / "Maquinados metal" / piece
                dest.mkdir(parents=True, exist_ok=True)
                existing = [dest]
            for src in list(root.glob(f"*__{piece}__*.jp*")) + (
                list(staging.glob(f"*__{piece}__*.jp*")) if staging.is_dir() else []
            ):
                target = dest / src.name
                if target.exists():
                    target.unlink()
                src.replace(target)
                _log(f"MOVE {src.name} -> {dest.relative_to(root)}")
            # Espejo a TODAS las carpetas de la pieza (Maquinado + Doblado)
            # para que el revisor no vea carpeta vacía.
            for other in existing[1:]:
                for f in dest.glob("*.jp*"):
                    shutil.copy2(f, other / f.name)
                _log(f"MIRROR {dest.name} -> {other.relative_to(root)}")

        # Publicar solo carpetas de las piezas target
        os.environ["COTAS_DOSSIER"] = "1"
        iniciar_sesion_dossier(nombre_job_desde_ensamble(ensamble), JOB_ROOT)
        try:
            guardar_contexto_dossier(cliente="OTC", producto="TANQUE")
        except Exception:
            pass

        publish_dirs = []
        for piece in TARGETS:
            for p in root.rglob(piece):
                if p.is_dir() and p.name == piece and "_STAGING" not in str(p):
                    jpgs = list(p.glob("*.jp*"))
                    if jpgs:
                        publish_dirs.append(str(p))
        publish_dirs = sorted(set(publish_dirs))
        _log(f"publish dirs={len(publish_dirs)}")
        for d in publish_dirs:
            unc, n = publicar_pieza_carpeta(d)
            _log(f"  publish {Path(d).name}: n={n} unc={bool(unc)}")
            # espejo Y
            try:
                rel = Path(d).relative_to(root)
                y_dest = Path(Y_JPGS) / "PIEZAS_ACOTADAS" / rel
                y_dest.mkdir(parents=True, exist_ok=True)
                for f in Path(d).glob("*.jp*"):
                    shutil.copy2(f, y_dest / f.name)
            except Exception as exc:
                _log(f"  AVISO Y: {exc}")

        n_db = sincronizar_lote_carpetas(publish_dirs, job=JOB)
        _log(f"sync db={n_db}")
        if mal_od:
            return 3
        return 0 if ok else 2
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
