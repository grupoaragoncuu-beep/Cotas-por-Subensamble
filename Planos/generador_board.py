# -*- coding: utf-8 -*-
"""
Flujo BOARD / GIGA tablero eléctrico (no tanque).

Reutiliza el motor de cotas del instructivo de ensambles (ancla SI + X/Y +
TYP + 6 vistas ViewCube), pero:

- No pide TOP/SEGM/BASE.
- Acota el IAM raíz y subkits de 1er nivel (paneles/barras), no stacks HW.
- Salida: ``Planos/JPG/<job>/BOARD/<kit>/<VISTA>/*.jpg``

Entrada: ``generador_board.py`` o desvío automático desde otros generadores
vía ``producto_tipo.redirigir_si_board``.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import traceback

import pythoncom

from inventor_com import conectar_inventor

CARPETA_BOARD = "BOARD"
TIPO_DOCUMENTO_ENSAMBLE = 12291

# Subkits de 1er nivel que valen como “panel” a exportar aparte.
_RE_PANEL_KIT = re.compile(
    r"(^9919-[FMP]-)|(^GEN1[-_])|(^GENE[-_].*-201)|(^GE3R[-_])|"
    r"(POWER\s*PANEL)|(FRAME\s*STACK)|(^9919-Board)",
    re.IGNORECASE,
)
_RE_OMITIR_KIT = re.compile(
    r"(WASHER|NUT|BOLT|SCREW|(^HW[-_])|ISOLATOR\s*STACK|"
    r"carriage\s+bolt\s+stack|(^\.?[\d.]+\s*(in\s*)?stack)|"
    r"(^\d+-?\d*\s*\.?\d*in\s*stack))",
    re.IGNORECASE,
)


def _log(msg=""):
    print(msg, flush=True)


def _parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Cotas BOARD/GIGA (tablero eléctrico)"
    )
    p.add_argument(
        "--no-limpiar",
        action="store_true",
        help="Conservar JPG previos en BOARD/ (por defecto se limpia)",
    )
    p.add_argument(
        "--max",
        type=int,
        default=0,
        help="Máximo de kits (0 = raíz + paneles filtrados)",
    )
    p.add_argument(
        "--solo",
        default="",
        help="Solo kits cuyo nombre contenga este texto",
    )
    p.add_argument(
        "--listar",
        action="store_true",
        help="Solo listar kits BOARD y salir",
    )
    return p.parse_args(argv)


def _nombre_base_doc(doc_or_occ) -> str:
    try:
        ff = str(getattr(doc_or_occ, "FullFileName", "") or "")
        if ff:
            return os.path.splitext(os.path.basename(ff))[0]
    except Exception:
        pass
    try:
        return str(doc_or_occ.DisplayName or "KIT")
    except Exception:
        return "KIT"


def _contar_hijos_utiles(asm_doc) -> int:
    from generador_ensambles_instructivo import _hijos_utiles

    try:
        return len(_hijos_utiles(asm_doc))
    except Exception:
        return 0


def _es_kit_omitir(nombre: str) -> bool:
    return bool(_RE_OMITIR_KIT.search(str(nombre or "")))


def _es_panel_candidato(nombre: str, n_hijos: int) -> bool:
    if _es_kit_omitir(nombre):
        return False
    if n_hijos < 2:
        return False
    if _RE_PANEL_KIT.search(nombre):
        return True
    # Subensamble denso no-HW: candidato genérico de board.
    return n_hijos >= 6 and not _es_kit_omitir(nombre)


def recolectar_kits_board(ensamble_raiz, solo: str = "", max_n: int = 0):
    """
    Lista ``(asm_doc, nombre, qty, n_hijos)``.

    Siempre incluye el raíz. Luego sub-IAM de 1er nivel tipo panel.
    """
    from generador_ensambles_instructivo import _nombres_iam_hijos

    raiz_nom = _nombre_base_doc(ensamble_raiz)
    hijos_raiz = _contar_hijos_utiles(ensamble_raiz)
    lista = [(ensamble_raiz, raiz_nom, 1, hijos_raiz)]

    try:
        occs = ensamble_raiz.ComponentDefinition.Occurrences
        n = int(occs.Count)
    except Exception:
        n = 0

    vistos = {raiz_nom.upper()}
    for i in range(1, n + 1):
        try:
            occ = occs.Item(i)
            if getattr(occ, "Suppressed", False):
                continue
            if int(occ.DefinitionDocumentType) != TIPO_DOCUMENTO_ENSAMBLE:
                continue
            asm = occ.Definition.Document
            nom = _nombre_base_doc(asm)
            if not nom or nom.upper() in vistos:
                continue
            if _es_kit_omitir(nom):
                continue
            nh = _contar_hijos_utiles(asm)
            if not _es_panel_candidato(nom, nh):
                continue
            vistos.add(nom.upper())
            lista.append((asm, nom, 1, nh))
        except Exception:
            continue

    # Evitar exportar kits anidados si el padre ya está en la lista.
    nombres = {str(t[1]).upper() for t in lista}
    anidados = set()
    for asm_doc, nombre, _q, _h in lista:
        if asm_doc is ensamble_raiz:
            continue
        for hijo_iam in _nombres_iam_hijos(asm_doc):
            if hijo_iam in nombres and hijo_iam != str(nombre).upper():
                anidados.add(hijo_iam)
    if anidados:
        _log(f"  Kits anidados omitidos: {sorted(anidados)[:12]}")
        lista = [t for t in lista if str(t[1]).upper() not in anidados]

    solo_u = str(solo or "").strip().upper()
    if solo_u:
        lista = [t for t in lista if solo_u in str(t[1]).upper()]

    lista.sort(key=lambda t: (0 if t[0] is ensamble_raiz else 1, str(t[1]).upper()))
    if max_n and max_n > 0:
        lista = lista[:max_n]
    return lista


def ejecutar(
    gestionar_com: bool = True,
    inv_app=None,
    plano=None,
    ensamble=None,
    limpiar: bool = True,
    max_n: int = 0,
    solo: str = "",
    listar: bool = False,
):
    _log("=" * 62)
    _log(" COTAS ABIGAIL — BOARD / GIGA (tablero eléctrico)")
    _log("=" * 62)

    if gestionar_com:
        pythoncom.CoInitialize()

    ok = False
    owns_com = gestionar_com
    try:
        from generador_caras_tanque import (
            _carpeta_salida_tanque,
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )
        from generador_ensambles_instructivo import _procesar_kit
        from generador_tanque_completo import _nombre_carpeta_pieza
        from nomenclatura_capturas import nombre_job_desde_ensamble
        from producto_tipo import clasificar_producto

        if inv_app is None:
            inv_app = conectar_inventor()
        if plano is None:
            plano = _obtener_plano_activo(inv_app)
        if ensamble is None:
            ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            _log("ERROR: abre el machote (.dwg/.idw) y el .iam del board.")
            return False

        info = clasificar_producto(ensamble)
        _log(
            f"Clasificación: {info.get('tipo')} / {info.get('familia')} "
            f"— {info.get('motivo')}"
        )
        try:
            from producto_tipo import aplicar_unidad_producto

            aplicar_unidad_producto(ensamble=ensamble, info=info)
        except Exception as exc_u:
            _log(f"AVISO unidades: {exc_u}")
        if info.get("tipo") not in ("BOARD", "DESCONOCIDO"):
            _log(
                "AVISO: el ensamble no parece BOARD; se procesa igual "
                "porque se invocó generador_board explícitamente."
            )

        try:
            job = nombre_job_desde_ensamble(ensamble)
        except Exception:
            job = os.path.splitext(
                os.path.basename(ensamble.FullFileName or "BOARD")
            )[0]

        try:
            from cotas_dossier_registro import iniciar_sesion_dossier

            iniciar_sesion_dossier(job)
        except Exception as exc_dos:
            _log(f"AVISO dossier sesión: {exc_dos}")

        lista = recolectar_kits_board(ensamble, solo=solo, max_n=max_n)
        _log(f"Kits BOARD a procesar: {len(lista)}")
        for _d, nom, qty, hijos in lista:
            _log(f"  - {nom}  qty={qty}  hijos_utiles≈{hijos}")

        if listar:
            return True
        if not lista:
            _log("Sin kits BOARD útiles.")
            return True

        carpeta_job = _carpeta_salida_tanque(plano, ensamble)
        carpeta_raiz = os.path.join(carpeta_job, CARPETA_BOARD)
        if limpiar and os.path.isdir(carpeta_raiz):
            shutil.rmtree(carpeta_raiz, ignore_errors=True)
        os.makedirs(carpeta_raiz, exist_ok=True)

        try:
            base_sheet = _encontrar_hoja_machote(plano) or plano.Sheets.Item(1)
            base_sheet.Activate()
        except Exception:
            base_sheet = plano.Sheets.Item(1)

        total = 0
        for asm_doc, nombre, qty, hijos in lista:
            _log(f"\n>>> BOARD kit: {nombre} (qty={qty}, hijos≈{hijos})")
            carpeta_kit = os.path.join(
                carpeta_raiz, _nombre_carpeta_pieza(nombre)
            )
            if os.path.isdir(carpeta_kit):
                shutil.rmtree(carpeta_kit, ignore_errors=True)
            os.makedirs(carpeta_kit, exist_ok=True)
            try:
                n = _procesar_kit(
                    inv_app,
                    plano,
                    base_sheet,
                    asm_doc,
                    nombre,
                    carpeta_kit,
                    job,
                )
                total += n
            except Exception as exc:
                _log(f"ERROR kit {nombre}: {exc}")
                _log(traceback.format_exc())

        _log(f"\nListo: {total} JPG en {carpeta_raiz}")
        try:
            from cotas_dossier_registro import publicar_y_sincronizar_dossier

            publicar_y_sincronizar_dossier(carpeta_raiz, job=job)
        except Exception as exc_dos:
            _log(f"AVISO dossier sync BOARD: {exc_dos}")
        ok = True
        return True
    except Exception as exc:
        _log(f"ERROR fatal BOARD: {exc}")
        _log(traceback.format_exc())
        return False
    finally:
        try:
            from cota_estilo import set_unidad_cota

            set_unidad_cota("in")  # reset post-board: tanques default; Board re-aplica mm al entrar
        except Exception:
            pass
        try:
            from creador_vistas import set_nombre_pieza_completo

            set_nombre_pieza_completo(False)
        except Exception:
            pass
        if inv_app is not None:
            try:
                from generador_caras_tanque import (
                    _encontrar_hoja_machote,
                    _obtener_plano_activo,
                )

                pl = _obtener_plano_activo(inv_app)
                hoja = _encontrar_hoja_machote(pl)
                if hoja is not None:
                    hoja.Activate()
            except Exception:
                pass
        if owns_com:
            pythoncom.CoUninitialize()
        if ok:
            _log("PROCESO COMPLETO: flujo BOARD exportado.")


def main(argv=None):
    args = _parse_args(argv)
    limpiar = not bool(args.no_limpiar)
    ok = ejecutar(
        gestionar_com=True,
        limpiar=limpiar,
        max_n=int(args.max or 0),
        solo=str(args.solo or ""),
        listar=bool(args.listar),
    )
    return 0 if ok else 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    raise SystemExit(main())
