"""
Flujo de PIEZAS_ACOTADAS únicamente.

Genera SOLO las cotas por pieza (equivalente al viejo COTAS_ILOGIC_ABIGAIL),
sin correr las cotas por caras del tanque. El resultado queda en
``Planos/JPG/<tanque>/PIEZAS_ACOTADAS/<CLASIFICACIÓN>/<PIEZA>/`` según el
iProperty ``Clasificación`` escrito por el iLogic
``Colorimetria Sub Assembly + Norman.iLogicVb``. Piezas sin iProperty
caen en ``SIN CLASIFICACION/<PIEZA>/``.

Se ejecuta desde la regla iLogic ``COTAS_ILOGIC_ABIGAIL``.
"""

from __future__ import annotations

import os
import sys

import pythoncom

import generador_caras_tanque
from generador_caras_tanque import (
    _carpeta_salida_tanque,
    _encontrar_hoja_machote,
    _obtener_ensamble_principal,
    _obtener_plano_activo,
    cargar_mapa_piezas_por_clasificacion,
    detectar_mapa_piezas_por_clasificacion,
    guardar_mapa_piezas_por_clasificacion,
)
from generador_tanque_completo import (
    CARPETA_PIEZAS_ACOTADAS,
    _limpiar_exportacion_piezas,
    _recuperar_antes_de_piezas,
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


def ejecutar():
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

        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in ("1", "true", "TRUE", "yes")

        # Clasificación por PROCESO usando el iProperty ``Clasificación``
        # que escribe el iLogic ``Colorimetria Sub Assembly + Norman``.
        #
        # La detección es rápida (sólo lee un iProperty por pieza), no
        # requiere geometría. Aún así persistimos el mapa a JSON para
        # facilitar auditorías y para que el flujo pueda arrancar aunque
        # el ensamble ya no esté abierto.
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
            # Fallback muy defensivo: si el ensamble no aportó nada al mapa
            # (por ejemplo COM roto), intentamos usar el JSON persistido
            # anterior para no romper la sesión.
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
                # Segmentar los JPG por clasificación del iProperty. Piezas
                # sin match caen en SIN CLASIFICACION/<pieza>/ (por decisión
                # explícita del usuario: no perder nada).
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
    sys.exit(0 if ejecutar() else 1)
