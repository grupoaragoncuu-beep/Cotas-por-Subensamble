# -*- coding: utf-8 -*-
"""
REGLA RAPIDA flat — SOLO lo pedido:

  1) Barrenos TYP en X/Y  →  XCENTRO / YCENTRO (+ _TYP)  [centro circulo/ovalo]
  2) Ø por tipo de tamaño →  HOLE## (escanea TODOS; 1 cota por Ø distinto;
                               circulo y ovalo NUNCA se mezclan aunque midan parecido)
  3) Espesor              →  THK

Nada de LENGTH, WIDTH, Doblado ni Estañado.

Piezas = carpetas en JPGS\\Corte\\Corte (sin ``Copia de``).
Reemplaza esos JPG en el servidor + filas DB.

Inventor: machote activo + ensamble Board abierto.
  python _reacotar_barrenos_flat_corte.py
"""
from __future__ import annotations

import os
import shutil
import sys

import pythoncom

ROOT_JPGS_CORTE = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS\Corte\Corte"
)
ALT_CORTE = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\9919-BOARD2_2\Corte\Corte"
)
JOB = "9919-Board 2"

# Solo estos tokens se borran/reemplazan (no LENGTH/WIDTH).
_TOKENS_REEMPLAZO = ("__XCENTRO", "__YCENTRO", "__THK_", "__HOLE")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def _raiz() -> str:
    env = os.environ.get("CORTE_CORTE_ROOT", "").strip()
    if env and os.path.isdir(env):
        return env
    if os.path.isdir(ROOT_JPGS_CORTE):
        return ROOT_JPGS_CORTE
    if os.path.isdir(ALT_CORTE):
        return ALT_CORTE
    return ROOT_JPGS_CORTE


def _piezas(raiz: str) -> list[str]:
    if not os.path.isdir(raiz):
        return []
    out = []
    for n in sorted(os.listdir(raiz)):
        p = os.path.join(raiz, n)
        if not os.path.isdir(p):
            continue
        if n.upper().startswith("COPIA DE") or n.startswith("_"):
            continue
        out.append(n)
    return out


def _es_captura_objetivo(fn: str) -> bool:
    up = fn.upper()
    return any(t in up for t in _TOKENS_REEMPLAZO)


def _limpiar_carpeta(carpeta: str) -> int:
    if not os.path.isdir(carpeta):
        return 0
    n = 0
    for fn in list(os.listdir(carpeta)):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        if not _es_captura_objetivo(fn):
            continue
        try:
            os.remove(os.path.join(carpeta, fn))
            n += 1
        except OSError:
            pass
    return n


def _borrar_copias(raiz: str) -> int:
    n = 0
    if not os.path.isdir(raiz):
        return 0
    for nombre in list(os.listdir(raiz)):
        if not nombre.upper().startswith("COPIA DE"):
            continue
        ruta = os.path.join(raiz, nombre)
        if os.path.isdir(ruta):
            try:
                shutil.rmtree(ruta)
                print(f"  - basura: {nombre}")
                n += 1
            except OSError as e:
                print(f"  AVISO {nombre}: {e}")
    return n


def _limpiar_db(piezas: list[str]) -> int:
    from cotas_dossier_registro import _db_nesting
    import psycopg2

    borradas = 0
    with psycopg2.connect(**_db_nesting()) as conn:
        with conn.cursor() as cur:
            for pieza in piezas:
                cur.execute(
                    """
                    DELETE FROM public.cotas_dossier
                    WHERE job = %s
                      AND ruta ILIKE %s
                      AND (
                        nombre_archivo ILIKE %s
                        OR nombre_archivo ILIKE %s
                        OR nombre_archivo ILIKE %s
                        OR nombre_archivo ILIKE %s
                      )
                    """,
                    (
                        JOB,
                        f"%Corte%Corte%{pieza}%",
                        "%__HOLE%",
                        "%__XCENTRO%",
                        "%__YCENTRO%",
                        "%__THK_%",
                    ),
                )
                borradas += cur.rowcount or 0
        conn.commit()
    return borradas


def _publicar_y_db(local_corte: str, share: str, piezas: list[str]) -> tuple[int, int]:
    from cotas_dossier_registro import (
        clasificar_type_y_spoteos,
        proceso_desde_ruta,
        producto_cliente_desde_job_root,
        _db_nesting,
    )
    import psycopg2

    dossier_jpgs = os.path.dirname(os.path.dirname(share))
    job_root = os.path.dirname(os.path.dirname(dossier_jpgs))
    prod, cli = producto_cliente_desde_job_root(job_root)
    if not cli:
        cli, prod = "GIGA", "ENCLOSURES NEMA 1"

    pubs = 0
    rows = []
    for pieza in piezas:
        src = os.path.join(local_corte, pieza) if local_corte else ""
        if not os.path.isdir(src):
            # buscar en staging
            continue
        dst = os.path.join(share, pieza)
        os.makedirs(dst, exist_ok=True)
        for fn in os.listdir(src):
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if not _es_captura_objetivo(fn) and "__THK_" not in fn.upper():
                # aceptar solo XY/THK/HOLE
                if not any(
                    t in fn.upper() for t in ("XCENTRO", "YCENTRO", "__THK_")
                ):
                    continue
            s = os.path.join(src, fn)
            d = os.path.join(dst, fn)
            try:
                shutil.copy2(s, d)
                pubs += 1
                rows.append(d)
            except OSError as e:
                print(f"  AVISO copy {fn}: {e}")

    # También recolectar desde staging local si reorg no movió
    if local_corte:
        staging = os.path.join(os.path.dirname(os.path.dirname(local_corte)), "_STAGING_DESPLIEGUE")
        # PIEZAS_ACOTADAS/_STAGING_DESPLIEGUE
        base_piezas = os.path.dirname(os.path.dirname(local_corte))  # Corte -> PIEZAS
        if os.path.basename(os.path.dirname(local_corte)).upper() == "CORTE":
            base_piezas = os.path.dirname(os.path.dirname(local_corte))
        # local_corte = .../PIEZAS_ACOTADAS/Corte/Corte
        staging = os.path.join(
            os.path.dirname(os.path.dirname(local_corte)), "_STAGING_DESPLIEGUE"
        )
        if os.path.isdir(staging):
            for fn in os.listdir(staging):
                if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue
                up = fn.upper()
                if not any(t in up for t in ("XCENTRO", "YCENTRO", "__THK_")):
                    continue
                # item desde nombre JOB__ITEM__MEDIDA
                item = ""
                if "__" in fn:
                    parts = fn.split("__")
                    if len(parts) >= 2:
                        item = parts[1]
                if not item:
                    continue
                # match pieza
                pieza_match = None
                for p in piezas:
                    if p.upper() == item.upper() or item.upper().startswith(p.upper()):
                        pieza_match = p
                        break
                    if p.upper() in item.upper():
                        pieza_match = p
                        break
                if not pieza_match:
                    continue
                dst_dir = os.path.join(share, pieza_match)
                os.makedirs(dst_dir, exist_ok=True)
                d = os.path.join(dst_dir, fn)
                try:
                    shutil.copy2(os.path.join(staging, fn), d)
                    pubs += 1
                    rows.append(d)
                except OSError:
                    pass

    ins = 0
    if rows:
        with psycopg2.connect(**_db_nesting()) as conn:
            with conn.cursor() as cur:
                for ruta in rows:
                    fn = os.path.basename(ruta)
                    tipo, n_spot = clasificar_type_y_spoteos(fn)
                    if "_TYP" in fn.upper():
                        n_spot = max(n_spot, 2)
                    clase = proceso_desde_ruta(ruta) or "Corte/Corte"
                    cur.execute(
                        """
                        INSERT INTO public.cotas_dossier
                            (cliente, producto, job, type, cantidad_spoteos,
                             nombre_archivo, ruta, clasificacion)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        """,
                        (
                            cli,
                            prod,
                            JOB,
                            tipo,
                            n_spot,
                            fn,
                            os.path.abspath(ruta),
                            clase,
                        ),
                    )
                    ins += 1
            conn.commit()
    return pubs, ins


def main() -> int:
    print("=" * 62)
    print(" RAPIDO FLAT: barrenos X/Y+TYP + Ø por tipo + THK")
    print(" (NO Length, NO Width, NO Doblado, NO Estañado)")
    print("=" * 62)

    share = _raiz()
    print("destino:", share)

    print("\n[0] Borrar Copia de…")
    print("   ", _borrar_copias(share), "carpetas")
    if os.path.isdir(ALT_CORTE):
        _borrar_copias(ALT_CORTE)

    piezas = _piezas(share)
    filtro = os.environ.get("PIEZAS_FILTRO", "").strip()
    if filtro:
        piezas = [p.strip() for p in filtro.split(",") if p.strip()]
    print(f"[1] piezas flat: {len(piezas)}")
    if not piezas:
        print("ERROR: no hay piezas en Corte/Corte")
        return 1

    print("[2] limpiar HOLE/XCENTRO/YCENTRO/THK en share")
    n = sum(_limpiar_carpeta(os.path.join(share, p)) for p in piezas)
    if os.path.isdir(ALT_CORTE):
        n += sum(_limpiar_carpeta(os.path.join(ALT_CORTE, p)) for p in piezas)
    print(f"    borrados={n}")

    print("[3] limpiar DB de esos tokens")
    print(f"    filas={_limpiar_db(piezas)}")

    # Env: modo estricto
    os.environ["PIEZAS_FILTRO"] = ",".join(piezas)
    os.environ["PIEZAS_INCREMENTAL"] = "0"
    os.environ["SOLO_FLAT_CORTE"] = "1"  # solo FRENTE_1+LADO; export solo XY/THK
    os.environ["ENSAMBLES_INDEPENDIENTES"] = "0"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    print("[4] SOLO_FLAT_CORTE=1  COTAS_JOB_OVERRIDE=", JOB)

    pythoncom.CoInitialize()
    ok = False
    carpeta_piezas = None
    try:
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
            cargar_mapa_piezas_por_clasificacion,
        )
        from generador_tanque_completo import (
            CARPETA_PIEZAS_ACOTADAS,
            _recuperar_antes_de_piezas,
            _reorganizar_piezas_por_clasificacion,
        )
        from generador_vistas import ejecutar_flujo_desde_app
        import creador_vistas as _cv

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None or ensamble is None:
            print("ERROR: abre machote + ensamble Board")
            return 2
        # mm/in lo aplica ejecutar_flujo_desde_app via producto_tipo (GIGA→mm).
        try:
            h = _encontrar_hoja_machote(plano)
            if h is not None:
                h.Activate()
        except Exception:
            pass

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)
        os.makedirs(carpeta_piezas, exist_ok=True)

        _cv.configurar_piezas_corte(set(piezas))
        print(f"    Corte configurado: {len(piezas)}")

        # limpiar staging viejo
        for st in ("_STAGING_DESPLIEGUE", "_STAGING_ESTANIADO"):
            sp = os.path.join(carpeta_piezas, st)
            if os.path.isdir(sp):
                shutil.rmtree(sp, ignore_errors=True)

        _recuperar_antes_de_piezas(inv, plano)
        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=False,
                catalogo_piezas=set(piezas),
            )
        )
        print(f"    flujo ok={ok}")

        # Auditoría COM: cotas malas (paso entre barrenos / redondeo) → reparar
        print("[4b] auditoria COM XY + reparacion")
        try:
            from auditoria_cotas_flat_com import auditar_y_reparar_plano
            from generador_vistas import exportar_hojas_jpg

            audit_ok, reparadas = auditar_y_reparar_plano(inv, plano)
            print(f"    audit_ok={audit_ok} reparadas={len(reparadas)}")
            if reparadas:
                print("    re-export JPG de hojas reparadas...")
                exportar_hojas_jpg(
                    inv,
                    plano,
                    carpeta_salida=carpeta_piezas,
                    nombres_permitidos={r.upper() for r in reparadas},
                    nombre_job=JOB,
                )
            if not audit_ok:
                ok = False
        except Exception as e_audit:
            print(f"    AVISO auditoria: {e_audit}")
            import traceback

            traceback.print_exc()

        try:
            _reorganizar_piezas_por_clasificacion(
                carpeta_piezas, {"Corte": set(piezas)}
            )
        except Exception as e:
            print(f"    AVISO reorg: {e}")
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback

        traceback.print_exc()
        ok = False
    finally:
        pythoncom.CoUninitialize()

    local_corte = (
        os.path.join(carpeta_piezas, "Corte", "Corte") if carpeta_piezas else ""
    )
    print("[5] publicar XY/THK → share + DB")
    pubs, ins = _publicar_y_db(local_corte, share, piezas)
    print(f"    publicados={pubs} db={ins}")

    # Verificación
    xy = thk = hole = 0
    for p in piezas:
        carpeta = os.path.join(share, p)
        if not os.path.isdir(carpeta):
            continue
        for fn in os.listdir(carpeta):
            up = fn.upper()
            if "XCENTRO" in up or "YCENTRO" in up:
                xy += 1
            if "__THK_" in up:
                thk += 1
            if "__HOLE" in up:
                hole += 1
    print(
        f"\n=== RESULTADO share Corte/Corte: "
        f"XCENTRO/YCENTRO={xy}  HOLE={hole}  THK={thk} ==="
    )
    if xy == 0:
        print("FALLO: no hay capturas de barrenos XY. Revisar CMD [cotas] barrenos_xy.")
        return 3
    if thk == 0:
        print("FALLO: no hay capturas THK. Revisar CMD [THK] DESPLIEGUE_LADO.")
        return 4
    print("LISTO.")
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
