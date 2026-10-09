"""
Reparación hallazgos auditoría ATC 62223 (prioridad):
1) Reacotar THK de 62223-1248-P09 y 62223-1248-P28_Predeterminado
2) Quitar hardware mal acotado en SIN CLASIFICACION (share + DB)
3) Consolidar rutas Corte/Doblado mal partidas (P72, SP-707/711/810, P28)
"""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.chdir(os.path.dirname(os.path.abspath(__file__)))

JOB = "62223-1246-A01"
JOB_ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
)
LOCAL_PA = Path(r"C:\Proyectos\COTAS\Planos\JPG") / JOB / "PIEZAS_ACOTADAS"
UNC_PA = Path(JOB_ROOT) / "DOSSIER FILES" / "JPGS" / "PIEZAS_ACOTADAS"
# Y: mirror
Y_PA = Path(
    r"Y:\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223\DOSSIER FILES\JPGS\PIEZAS_ACOTADAS"
)

THK_TARGETS = (
    "62223-1248-P09",
    "62223-1248-P28_Predeterminado",
)

# Hardware / comprados que no debieron acotarse
HARDWARE_DELETE = (
    "HEX NUT 3_8-16",
    "HW-CN-01",
    "SP-792_2",
    "SP-792_3",
    "SP-855",
)


def _log(msg: str) -> None:
    print(msg, flush=True)


def _rm_tree_piece(roots: list[Path], piece: str) -> int:
    n = 0
    for root in roots:
        if not root.exists():
            continue
        for p in list(root.rglob("*")):
            if p.is_dir() and p.name == piece:
                _log(f"  DEL DIR {p}")
                shutil.rmtree(p, ignore_errors=True)
                n += 1
    return n


def _db_delete_piece(piece: str) -> int:
    try:
        import psycopg2
        from cotas_dossier_registro import _db_nesting

        conn = psycopg2.connect(**_db_nesting())
        cur = conn.cursor()
        cur.execute(
            """
            DELETE FROM public.cotas_dossier
            WHERE job = %s AND nombre_archivo ILIKE %s
            RETURNING id
            """,
            (JOB, f"%__{piece}__%"),
        )
        ids = cur.fetchall()
        conn.commit()
        conn.close()
        return len(ids)
    except Exception as exc:
        _log(f"  AVISO DB delete {piece}: {exc}")
        return 0


def limpiar_hardware() -> None:
    _log("=== 2) Limpiar hardware SIN CLASIFICACION ===")
    roots = [LOCAL_PA, UNC_PA, Y_PA]
    for piece in HARDWARE_DELETE:
        nd = _rm_tree_piece(roots, piece)
        # also loose jpg
        for root in roots:
            if not root.exists():
                continue
            for p in root.rglob(f"*__{piece}__*"):
                if p.is_file():
                    try:
                        p.unlink()
                        nd += 1
                        _log(f"  DEL FILE {p}")
                    except OSError as e:
                        _log(f"  AVISO {e}")
        ndb = _db_delete_piece(piece)
        _log(f"  {piece}: dirs/files~{nd} db_rows={ndb}")


def _move_jpgs(srcs: list[Path], dest_dir: Path) -> int:
    dest_dir.mkdir(parents=True, exist_ok=True)
    n = 0
    for src in srcs:
        if not src.is_file():
            continue
        dst = dest_dir / src.name
        try:
            if dst.exists() and os.path.normcase(str(dst)) != os.path.normcase(str(src)):
                dst.unlink()
            shutil.move(str(src), str(dst))
            n += 1
            _log(f"  MOVE {src.name} → {dest_dir}")
        except OSError as e:
            _log(f"  AVISO move {src}: {e}")
    return n


def consolidar_rutas() -> None:
    _log("=== 3) Consolidar mal ruteo ===")
    roots = [r for r in (LOCAL_PA, UNC_PA, Y_PA) if r.exists()]

    # P72: todo a Doblado/Metal (tiene HEIGHT/LEG)
    for root in roots:
        src_dir = root / "Corte" / "Plasma y Laser" / "Corte metal" / "62223-1248-P72"
        dest = root / "Doblado" / "Metal" / "62223-1248-P72"
        if src_dir.is_dir():
            jpgs = list(src_dir.glob("*.jp*"))
            _move_jpgs(jpgs, dest)
            try:
                src_dir.rmdir()
            except OSError:
                pass

    # P28: unificar en Doblado/Metal (tiene HOLE; falta THK)
    for root in roots:
        dest = root / "Doblado" / "Metal" / "62223-1248-P28_Predeterminado"
        for rel in (
            ("Corte", "Maquinado", "Maquinados metal", "62223-1248-P28_Predeterminado"),
        ):
            src_dir = root.joinpath(*rel)
            if src_dir.is_dir():
                _move_jpgs(list(src_dir.glob("*.jp*")), dest)
                try:
                    src_dir.rmdir()
                except OSError:
                    pass

    # SP-707 / SP-711 / SP-810_5: HOLE estaba en Doblado; OD/THK en Maquinado
    # → consolidar en Maquinados metal (pieza maquinada/flange, no chapa doblada)
    for piece in ("SP-707", "SP-711", "SP-810_5"):
        for root in roots:
            dest = root / "Corte" / "Maquinado" / "Maquinados metal" / piece
            src = root / "Doblado" / "Metal" / piece
            if src.is_dir():
                _move_jpgs(list(src.glob("*.jp*")), dest)
                try:
                    shutil.rmtree(src, ignore_errors=True)
                except OSError:
                    pass


def reacotar_thk() -> bool:
    _log("=== 1) Reacotar THK P09 + P28 ===")
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["COTAS_DOSSIER"] = "1"
    os.environ["PIEZAS_FILTRO"] = ",".join(THK_TARGETS)
    os.environ["PIEZAS_INCREMENTAL"] = "1"
    os.environ["PIEZAS_TRAZABLE"] = "0"  # no wipe
    os.environ["PIEZAS_WIPE_JOB"] = "0"
    os.environ["ENSAMBLES_INDEPENDIENTES"] = "0"
    os.environ["SKIP_ENSAMBLES_IND"] = "1"
    # no flat-only; need LADO THK on bent parts
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
        publicar_y_sincronizar_dossier,
    )
    from nomenclatura_capturas import nombre_job_desde_ensamble
    import creador_vistas as cv

    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            _log("ERROR: sin plano/ensamble")
            return False
        try:
            plano.Activate()
        except Exception:
            pass

        job_tok = nombre_job_desde_ensamble(ensamble)
        iniciar_sesion_dossier(job_tok, JOB_ROOT)

        from producto_tipo import clasificar_producto, aplicar_unidad_producto

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        cv.configurar_producto_flujo("TANQUE")
        # Forzar doblado para que THK/LADO corra (ambas son chapas con huecos)
        cv.configurar_piezas_doblado(set(THK_TARGETS))
        cv.configurar_piezas_corte(set())

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)
        _log(f"  filtro={os.environ['PIEZAS_FILTRO']}")
        _log(f"  salida={carpeta_piezas}")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                # False: no saltar aunque ya tengan L/W/HOLE (falta THK).
                incremental=False,
                catalogo_piezas=set(THK_TARGETS),
            )
        )
        _log(f"  flujo ok={ok}")

        # Mover THK nuevos a Doblado/Metal/<pieza>
        root = Path(carpeta_piezas)
        for piece in THK_TARGETS:
            dest = root / "Doblado" / "Metal" / piece
            dest.mkdir(parents=True, exist_ok=True)
            # sueltos + staging
            for p in list(root.glob(f"*__{piece}__THK_*.jp*")):
                _move_jpgs([p], dest)
            for st in ("_STAGING_DESPLIEGUE", "_STAGING_ESTANIADO"):
                sd = root / st
                if sd.is_dir():
                    for p in sd.glob(f"*__{piece}__*.jp*"):
                        _move_jpgs([p], dest)
            # también HEIGHT/LEG si salieron
            for p in list(root.glob(f"*__{piece}__*.jp*")):
                _move_jpgs([p], dest)

        try:
            publicar_y_sincronizar_dossier(str(carpeta_piezas), job=JOB)
        except Exception as exc:
            _log(f"  AVISO publish: {exc}")

        # sync Y: from UNC if needed — publicar already targets dossier
        return ok
    finally:
        try:
            cv.configurar_producto_flujo("")
            cv.configurar_piezas_doblado([])
            cv.configurar_piezas_corte([])
        except Exception:
            pass
        pythoncom.CoUninitialize()


def verificar() -> None:
    _log("=== VERIFY ===")
    root = Y_PA if Y_PA.exists() else UNC_PA
    for piece in THK_TARGETS:
        files = list(root.rglob(f"*__{piece}__*"))
        meds = sorted(
            {
                f.name.split("__")[2].split("_")[0].upper()
                for f in files
                if "__" in f.name and len(f.name.split("__")) >= 3
            }
        )
        has_thk = any("__THK_" in f.name.upper() for f in files)
        _log(f"  {piece}: thk={has_thk} meds={meds} n={len(files)}")
    for piece in HARDWARE_DELETE:
        left = list(root.rglob(f"*__{piece}__*"))
        _log(f"  hardware {piece}: quedan={len(left)}")


def main() -> int:
    # Orden: limpiar hardware → consolidar → reacotar THK → verify
    limpiar_hardware()
    consolidar_rutas()
    ok = reacotar_thk()
    verificar()
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
