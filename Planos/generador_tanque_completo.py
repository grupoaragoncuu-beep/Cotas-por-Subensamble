"""
Flujo integrado de COTAS ABIGAIL.

1. Cotas por referencia de las cuatro caras del tanque.
2. Flujo original COTAS_ILOGIC_ABIGAIL por pieza.

Ambas salidas se organizan bajo JPG/<tanque>/ en carpetas separadas.
"""

import os
import shutil
import sys

import pythoncom

from generador_caras_tanque import (
    _carpeta_salida_tanque,
    _encontrar_hoja_machote,
    _limpiar_machote,
    _obtener_ensamble_principal,
    _obtener_plano_activo,
    _resolver_contenedor_paredes,
    ejecutar as ejecutar_caras,
)
from generador_vistas import ejecutar_flujo_desde_app
from inventor_com import conectar_inventor


CARPETA_PIEZAS_ACOTADAS = "PIEZAS_ACOTADAS"


def _limpiar_exportacion_piezas(carpeta):
    """Evita conservar JPG de piezas que ya no existan en el tanque."""
    os.makedirs(carpeta, exist_ok=True)
    for nombre in os.listdir(carpeta):
        if not nombre.lower().endswith(".jpg"):
            continue
        try:
            os.remove(os.path.join(carpeta, nombre))
        except OSError:
            pass


def _reactivar_machote(inv_app):
    """El flujo legado termina en la última pieza; dejar visible el machote."""
    try:
        plano = _obtener_plano_activo(inv_app)
        hoja = _encontrar_hoja_machote(plano)
        if hoja is not None:
            hoja.Activate()
    except Exception:
        pass


def _prevalidar_cuatro_caras(ensamble):
    """
    Comprueba que existan cuatro paredes físicas verificables.

    Acepta segmentos nombrados en la raíz (Vantran) o un subensamble
    estructural que demuestre geométricamente sus cuatro laterales (OTC).
    Nunca habilita el flujo solo porque haya cuatro IAM arbitrarios.
    """
    try:
        return _resolver_contenedor_paredes(ensamble, registrar=False)
    except Exception:
        return {"valido": False, "motivo": "error durante prevalidación"}


def _limpiar_salida_parcial_caras(carpeta_tanque):
    """Borra únicamente la salida de caras que no es válida para este tanque."""
    carpeta = os.path.join(carpeta_tanque, "COTAS_POR_REFERENCIA")
    try:
        shutil.rmtree(carpeta)
    except FileNotFoundError:
        pass
    except OSError as error:
        print(f"AVISO: no se pudo limpiar salida parcial de caras: {error}")


def ejecutar():
    print("=" * 62)
    print(" COTAS ABIGAIL - TANQUE COMPLETO")
    print("=" * 62)

    pythoncom.CoInitialize()
    inv_app = None
    ok = False
    try:
        inv_app = conectar_inventor()
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            print(
                "ERROR: no se pudo recuperar el plano o ensamble "
                "para cotas por pieza."
            )
            return False

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        alcance = _prevalidar_cuatro_caras(ensamble)
        if alcance.get("valido"):
            contenedor = alcance.get("contenedor")
            fuente = (
                str(contenedor.Name)
                if contenedor is not None
                else "segmentos nombrados en la raíz"
            )
            print(
                "Paso 1/2: Cotando accesorios por caras del tanque "
                f"(alcance: {fuente})..."
            )
            # El orquestador ya posee el contexto COM: no hacer CoUninitialize
            # intermedio (eso dejó Inventor inestable al pasar a piezas).
            if not ejecutar_caras(gestionar_com=False):
                print("ERROR: no se completaron las cotas por cara.")
                return False
            # La rutina de caras reactiva el machote; refrescar handles.
            plano = _obtener_plano_activo(inv_app)
            ensamble = _obtener_ensamble_principal(inv_app)
            carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        else:
            print(
                "Paso 1/2: Caras omitidas; no se identificaron cuatro "
                "paredes físicas verificables "
                f"({alcance.get('motivo', 'sin detalle')})."
            )
            _limpiar_machote(plano, inv_app)
            _limpiar_salida_parcial_caras(carpeta_tanque)

        carpeta_piezas = os.path.join(
            carpeta_tanque, CARPETA_PIEZAS_ACOTADAS
        )
        _limpiar_exportacion_piezas(carpeta_piezas)
        print("Paso 2/2: Ejecutando cotas originales por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        # Equivalente funcional de COTAS_ILOGIC_ABIGAIL, pero sin llamar a
        # iLogic GenerarVistas (esa llamada queda bloqueada en este tanque).
        # Conserva creador_vistas + cotas.py + THK.py + exportación JPG.
        ok = bool(
            ejecutar_flujo_desde_app(
                inv_app,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
            )
        )
        return ok
    finally:
        if inv_app is not None:
            _reactivar_machote(inv_app)
        pythoncom.CoUninitialize()
        if ok:
            print("PROCESO COMPLETO: caras y piezas acotadas exportadas.")


if __name__ == "__main__":
    sys.exit(0 if ejecutar() else 1)
