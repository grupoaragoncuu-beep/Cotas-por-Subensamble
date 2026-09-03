"""
Flujo de PIEZAS_ACOTADAS únicamente.

Genera las cotas por pieza (largo/ancho/thk/diámetros). Con
``--seleccion seleccion_caras.json`` (Top Cover + SEGM1..4) organiza la
salida en ``PIEZAS_ACOTADAS/<SEGM*|TOP>/<CLASIFICACIÓN>/<PIEZA>/``.

Sin selección, conserva solo ``<CLASIFICACIÓN>/<PIEZA>/``.

Se ejecuta desde la regla iLogic ``COTAS_ILOGIC_ABIGAIL``.
"""

from __future__ import annotations

import os
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
    cargar_mapa_piezas_por_clasificacion,
    construir_mapa_piezas_desde_seleccion,
    detectar_mapa_piezas_por_clasificacion,
    guardar_mapa_piezas_por_cara,
    guardar_mapa_piezas_por_clasificacion,
)
from generador_tanque_completo import (
    CARPETA_PIEZAS_ACOTADAS,
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


def _reorganizar_clasificacion_dentro_caras(carpeta_piezas, mapa_clasificacion):
    """
    Tras ``_reorganizar_piezas_por_cara``, anida cada JPG en
    ``<CARA>/<CLASIFICACIÓN>/<PIEZA>/``.
    """
    import shutil

    from generador_tanque_completo import (
        SUBCARPETA_SIN_CLASIFICAR,
        SUBCARPETAS_CARA_PIEZAS,
        SUBCARPETAS_CLASIFICACION_PIEZAS,
        SUBCARPETA_OTROS_PIEZAS,
        _clasificacion_para_pieza,
        _extraer_pieza_de_jpg,
        _nombre_carpeta_pieza,
    )

    caras = list(SUBCARPETAS_CARA_PIEZAS) + [SUBCARPETA_OTROS_PIEZAS]
    clases_validas = {c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS}

    for cara in caras:
        cara_dir = os.path.join(carpeta_piezas, cara)
        if not os.path.isdir(cara_dir):
            continue
        # Recolectar JPGs en cualquier profundidad bajo la cara (plana o
        # ya en <PIEZA>/).
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
        # Limpiar carpetas de pieza vacías bajo la cara.
        for root, dirs, files in os.walk(cara_dir, topdown=False):
            if root == cara_dir:
                continue
            try:
                if not os.listdir(root):
                    os.rmdir(root)
            except OSError:
                pass


def ejecutar(ruta_seleccion=None):
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

        mapa_por_cara = {}
        if ruta_seleccion:
            if not os.path.isfile(ruta_seleccion):
                print(f"ERROR: No existe el JSON de selección: {ruta_seleccion}")
                return False
            print(f"  Selección manual (Top+SEGM): {ruta_seleccion}")
            seleccion = _cargar_seleccion_caras(ruta_seleccion)
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

        _limpiar_exportacion_piezas(carpeta_piezas, incremental=incremental)
        _recuperar_antes_de_piezas(inv_app, plano)

        print("Ejecutando cotas por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        if incremental:
            print("  Modo incremental ACTIVO (PIEZAS_INCREMENTAL=1).")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv_app,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=incremental,
            )
        )
        if ok:
            try:
                if mapa_por_cara:
                    # Primero por cara (incluye TOP), luego clasificacion adentro.
                    _reorganizar_piezas_por_cara(carpeta_piezas, mapa_por_cara)
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
            print("PROCESO COMPLETO: piezas acotadas exportadas.")


if __name__ == "__main__":
    ruta = _parse_ruta_seleccion(sys.argv[1:])
    if not ruta:
        print(
            "ERROR: Falta --seleccion <seleccion_caras.json>\n"
            "Ejecuta la regla iLogic COTAS_ILOGIC_ABIGAIL para elegir "
            "Top Cover + SEGM1..4.",
            flush=True,
        )
        sys.exit(1)
    sys.exit(0 if ejecutar(ruta_seleccion=ruta) else 1)
