"""
Flujo de PIEZAS_ACOTADAS únicamente.

Genera las cotas por pieza (largo/ancho/thk/diámetros). Con
``--seleccion seleccion_caras.json`` (Top Cover + SEGM1..4) organiza la
salida en ``PIEZAS_ACOTADAS/<SEGM*|TOP>/<CLASIFICACIÓN>/<PIEZA>/``.

Con ``--solo SEGM2`` (COTAS_POR_SEG_PIEZAS) solo acota las piezas del
catálogo de esa cara — prueba rápida sin el tanque completo.

Sin selección, conserva solo ``<CLASIFICACIÓN>/<PIEZA>/``.

Se ejecuta desde la regla iLogic ``COTAS_ILOGIC_ABIGAIL`` o
``COTAS_POR_SEG_PIEZAS``.
"""

from __future__ import annotations

import os
import shutil
import sys

import pythoncom

import generador_caras_tanque
from generador_caras_tanque import (
    _cargar_seleccion_caras,
    _carpeta_salida_tanque,
    _encontrar_hoja_machote,
    _obtener_ensamble_principal,
    _obtener_plano_activo,
    _parse_ruta_seleccion,
    _parse_solo,
    cargar_mapa_piezas_por_clasificacion,
    construir_mapa_piezas_desde_seleccion,
    detectar_mapa_piezas_por_clasificacion,
    guardar_mapa_piezas_por_cara,
    guardar_mapa_piezas_por_clasificacion,
)
from generador_tanque_completo import (
    CARPETA_PIEZAS_ACOTADAS,
    SUBCARPETAS_CARA_SELECCION,
    _limpiar_exportacion_piezas,
    _recuperar_antes_de_piezas,
    _reorganizar_piezas_por_cara,
    _reorganizar_piezas_por_clasificacion,
)
from generador_vistas import ejecutar_flujo_desde_app
from inventor_com import conectar_inventor


def _reactivar_machote(inv_app):
    try:
        plano = _obtener_plano_activo(inv_app)
        hoja = _encontrar_hoja_machote(plano)
        if hoja is not None:
            hoja.Activate()
    except Exception:
        pass


def _limpiar_solo_cara_piezas(carpeta_piezas, cara):
    """Vacía únicamente ``PIEZAS_ACOTADAS/<CARA>/`` (modo --solo)."""
    os.makedirs(carpeta_piezas, exist_ok=True)
    destino = os.path.join(carpeta_piezas, cara)
    if os.path.isdir(destino):
        try:
            shutil.rmtree(destino)
            print(f"  Limpieza modo --solo: vaciada {cara}/")
        except OSError as err:
            print(f"  AVISO: no se pudo vaciar {cara}/: {err}")
    os.makedirs(destino, exist_ok=True)


def _reorganizar_clasificacion_dentro_caras(carpeta_piezas, mapa_clasificacion):
    """
    Tras ``_reorganizar_piezas_por_cara``, anida cada JPG en
    ``<CARA>/<CLASIFICACIÓN>/<PIEZA>/``.
    """
    from generador_tanque_completo import (
        SUBCARPETA_SIN_CLASIFICAR,
        SUBCARPETAS_CLASIFICACION_PIEZAS,
        SUBCARPETA_OTROS_PIEZAS,
        _clasificacion_para_pieza,
        _extraer_pieza_de_jpg,
        _limpiar_carpetas_cara_vacias_y_legacy,
        _nombre_carpeta_pieza,
    )

    caras = list(SUBCARPETAS_CARA_SELECCION) + [SUBCARPETA_OTROS_PIEZAS]
    clases_validas = {c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS}

    for cara in caras:
        cara_dir = os.path.join(carpeta_piezas, cara)
        if not os.path.isdir(cara_dir):
            continue
        jpgs = []
        for root, _dirs, files in os.walk(cara_dir):
            for nombre in files:
                if nombre.lower().endswith(".jpg"):
                    jpgs.append(os.path.join(root, nombre))
        for ruta in jpgs:
            nombre = os.path.basename(ruta)
            clase = _clasificacion_para_pieza(nombre, mapa_clasificacion)
            if clase and clase.casefold() in clases_validas:
                destino_clase = next(
                    s
                    for s in SUBCARPETAS_CLASIFICACION_PIEZAS
                    if s.casefold() == clase.casefold()
                )
            else:
                destino_clase = SUBCARPETA_SIN_CLASIFICAR
            pieza_folder = _nombre_carpeta_pieza(_extraer_pieza_de_jpg(nombre))
            destino_dir = os.path.join(
                cara_dir, destino_clase, pieza_folder
            )
            try:
                os.makedirs(destino_dir, exist_ok=True)
                destino = os.path.join(destino_dir, nombre)
                if os.path.abspath(ruta) == os.path.abspath(destino):
                    continue
                if os.path.exists(destino):
                    os.remove(destino)
                shutil.move(ruta, destino)
            except OSError as err:
                print(
                    f"AVISO: no se pudo anidar '{nombre}' en "
                    f"{cara}/{destino_clase}/{pieza_folder}/: {err}"
                )
        for root, dirs, files in os.walk(cara_dir, topdown=False):
            if root == cara_dir:
                continue
            try:
                if not os.listdir(root):
                    os.rmdir(root)
            except OSError:
                pass

    _limpiar_carpetas_cara_vacias_y_legacy(carpeta_piezas)


def ejecutar(ruta_seleccion=None, solo_cara=None):
    print("=" * 62)
    print(" COTAS ABIGAIL - SOLO PIEZAS (PIEZAS_ACOTADAS)")
    print("=" * 62)

    pythoncom.CoInitialize()
    inv_app = None
    ok = False
    try:
        inv_app = conectar_inventor()
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            print("ERROR: no se pudo recuperar el plano o ensamble activo.")
            return False

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)

        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in (
            "1",
            "true",
            "TRUE",
            "yes",
        )

        solo = str(solo_cara or "").strip().upper() or None
        if solo and solo not in SUBCARPETAS_CARA_SELECCION:
            print(
                f"ERROR: --solo {solo} no válido. "
                f"Usa: {', '.join(SUBCARPETAS_CARA_SELECCION)}"
            )
            return False

        mapa_por_cara = {}
        if ruta_seleccion:
            if not os.path.isfile(ruta_seleccion):
                print(f"ERROR: No existe el JSON de selección: {ruta_seleccion}")
                return False
            print(f"  Selección manual (Top+SEGM): {ruta_seleccion}")
            seleccion = _cargar_seleccion_caras(ruta_seleccion)
            if not solo:
                solo_json = str(seleccion.get("solo") or "").strip().upper()
                if solo_json:
                    solo = solo_json
            mapa_por_cara = construir_mapa_piezas_desde_seleccion(
                ensamble, seleccion
            )
            ruta_mapa_cara = guardar_mapa_piezas_por_cara(
                carpeta_tanque, mapa_por_cara
            )
            if ruta_mapa_cara:
                print(f"  Mapa piezas/cara (con TOP) persistido: {ruta_mapa_cara}")
            if "TOP" in mapa_por_cara:
                print(f"  TOP: {len(mapa_por_cara['TOP'])} nombres en catálogo")
        else:
            print(
                "  AVISO: sin --seleccion; PIEZAS_ACOTADAS no se dividirá "
                "por SEGM/TOP (solo clasificación)."
            )

        catalogo_filtro = None
        mapa_reorg = mapa_por_cara
        if solo:
            if not mapa_por_cara or solo not in mapa_por_cara:
                print(
                    f"ERROR: modo --solo {solo}: no hay catálogo de piezas "
                    "para esa cara en la selección. Revisa el pick."
                )
                return False
            catalogo_filtro = set(mapa_por_cara[solo] or set())
            mapa_reorg = {solo: set(catalogo_filtro)}
            print(
                f"  Modo COTAS_POR_SEG_PIEZAS: solo {solo} "
                f"({len(catalogo_filtro)} nombres en catálogo)"
            )
            if not catalogo_filtro:
                print(f"ERROR: catálogo vacío para {solo}.")
                return False

        print(
            "  Leyendo clasificación (iProperty) de cada pieza del ensamble..."
        )
        mapa_clasificacion = detectar_mapa_piezas_por_clasificacion(
            inv_app, ensamble
        )
        if mapa_clasificacion:
            ruta_mapa = guardar_mapa_piezas_por_clasificacion(
                carpeta_tanque, mapa_clasificacion
            )
            if ruta_mapa:
                print(f"  Mapa por clasificación persistido en: {ruta_mapa}")
        else:
            print(
                "  AVISO: la detección de clasificación devolvió mapa vacío. "
                "Intentando cargar el JSON persistido anterior..."
            )
            mapa_clasificacion = cargar_mapa_piezas_por_clasificacion(
                carpeta_tanque
            )
            if not mapa_clasificacion:
                print(
                    "  AVISO: no hay mapa previo tampoco; todas las piezas "
                    "caerán en SIN CLASIFICACION/."
                )

        if solo:
            _limpiar_solo_cara_piezas(carpeta_piezas, solo)
        else:
            _limpiar_exportacion_piezas(carpeta_piezas, incremental=incremental)
        _recuperar_antes_de_piezas(inv_app, plano)

        print("Ejecutando cotas por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        if solo:
            print(f"  Destino: {carpeta_piezas}\\{solo}\\")
        if incremental:
            print("  Modo incremental ACTIVO (PIEZAS_INCREMENTAL=1).")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv_app,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=incremental,
                catalogo_piezas=catalogo_filtro,
            )
        )
        if ok:
            try:
                if mapa_reorg:
                    _reorganizar_piezas_por_cara(carpeta_piezas, mapa_reorg)
                    _reorganizar_clasificacion_dentro_caras(
                        carpeta_piezas,
                        dict(
                            generador_caras_tanque.LAST_PIEZAS_POR_CLASIFICACION
                            or mapa_clasificacion
                        ),
                    )
                else:
                    _reorganizar_piezas_por_clasificacion(
                        carpeta_piezas,
                        dict(
                            generador_caras_tanque.LAST_PIEZAS_POR_CLASIFICACION
                        ),
                    )
            except Exception as err:
                print(f"AVISO: fallo en reorganización de PIEZAS_ACOTADAS: {err}")
        return ok
    finally:
        if inv_app is not None:
            _reactivar_machote(inv_app)
        pythoncom.CoUninitialize()
        if ok:
            if solo:
                print(
                    f"PROCESO COMPLETO: piezas de {solo} acotadas "
                    f"(PIEZAS_ACOTADAS\\{solo}\\)."
                )
            else:
                print("PROCESO COMPLETO: piezas acotadas exportadas.")


if __name__ == "__main__":
    ruta = _parse_ruta_seleccion(sys.argv[1:])
    solo = _parse_solo(sys.argv[1:])
    if not ruta:
        print(
            "ERROR: Falta --seleccion <seleccion_caras.json>\n"
            "Usa COTAS_ILOGIC_ABIGAIL (todas las caras) o "
            "COTAS_POR_SEG_PIEZAS (una cara: --solo SEGM1|…|TOP|BASE).",
            flush=True,
        )
        sys.exit(1)
    sys.exit(0 if ejecutar(ruta_seleccion=ruta, solo_cara=solo) else 1)
