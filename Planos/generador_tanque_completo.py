"""
Flujo integrado de COTAS ABIGAIL.

1. Cotas por referencia de las 5 caras del tanque (FRONT, BACK, LEFT, RIGHT, TOP).
2. Flujo original COTAS_ILOGIC_ABIGAIL por pieza, reorganizado en subcarpetas
   por cara.

Ambas salidas se organizan bajo JPG/<tanque>/ en carpetas separadas.
"""

import os
import re
import shutil
import sys

import pythoncom

import generador_caras_tanque
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
SUBCARPETAS_CARA_PIEZAS = ("FRONT", "BACK", "LEFT", "RIGHT", "TOP")
SUBCARPETA_OTROS_PIEZAS = "OTROS"


def _limpiar_exportacion_piezas(carpeta):
    """Vacia la carpeta de piezas y sus subcarpetas de cara antes de exportar."""
    os.makedirs(carpeta, exist_ok=True)
    for nombre in os.listdir(carpeta):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isdir(ruta):
            if nombre.upper() in SUBCARPETAS_CARA_PIEZAS or nombre.upper() == SUBCARPETA_OTROS_PIEZAS:
                for jpg in os.listdir(ruta):
                    if jpg.lower().endswith(".jpg"):
                        try:
                            os.remove(os.path.join(ruta, jpg))
                        except OSError:
                            pass
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        try:
            os.remove(ruta)
        except OSError:
            pass


def _clave_pieza(texto):
    """Normaliza un nombre para comparar por token/subcadena robustamente."""
    if not texto:
        return ""
    limpio = str(texto).upper()
    # Quitar extensión y sufijos comunes de nombre de hoja.
    limpio = re.sub(r"\.JPG$", "", limpio)
    limpio = re.sub(r"[\s\-_:()\[\]{},.]+", "", limpio)
    return limpio


def _cara_para_pieza(nombre_archivo, mapa_por_cara):
    """Devuelve la cara asignada al JPG según los catálogos de piezas."""
    if not mapa_por_cara:
        return None
    base_archivo = os.path.splitext(os.path.basename(nombre_archivo))[0]
    clave_archivo = _clave_pieza(base_archivo)
    if not clave_archivo:
        return None
    mejor_cara = None
    mejor_len = 0
    for cara, piezas in mapa_por_cara.items():
        for pieza in piezas:
            clave_pieza = _clave_pieza(pieza)
            if not clave_pieza:
                continue
            if clave_pieza in clave_archivo or clave_archivo in clave_pieza:
                # Preferir el match más largo (más específico) ante empates.
                if len(clave_pieza) > mejor_len:
                    mejor_cara = cara
                    mejor_len = len(clave_pieza)
    return mejor_cara


def _reorganizar_piezas_por_cara(carpeta_piezas, mapa_por_cara):
    """
    Mueve cada JPG de pieza individual a la subcarpeta de su cara.
    Piezas sin match caen en OTROS/. Diseñado para no fallar: si algo sale
    mal solo se registra por consola y los JPG originales se conservan.
    """
    try:
        os.makedirs(carpeta_piezas, exist_ok=True)
        for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,):
            os.makedirs(os.path.join(carpeta_piezas, sub), exist_ok=True)
    except OSError as err:
        print(f"AVISO: no se pudieron preparar subcarpetas de PIEZAS_ACOTADAS: {err}")
        return {}

    conteo = {sub: 0 for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,)}
    try:
        entradas = list(os.listdir(carpeta_piezas))
    except OSError as err:
        print(f"AVISO: no se pudo listar {carpeta_piezas}: {err}")
        return conteo

    for nombre in entradas:
        ruta = os.path.join(carpeta_piezas, nombre)
        if os.path.isdir(ruta):
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        cara = _cara_para_pieza(nombre, mapa_por_cara)
        destino_sub = cara if cara in SUBCARPETAS_CARA_PIEZAS else SUBCARPETA_OTROS_PIEZAS
        destino = os.path.join(carpeta_piezas, destino_sub, nombre)
        try:
            if os.path.exists(destino):
                os.remove(destino)
            shutil.move(ruta, destino)
            conteo[destino_sub] += 1
        except OSError as err:
            print(f"AVISO: no se pudo mover '{nombre}' a {destino_sub}/: {err}")

    print("  PIEZAS_ACOTADAS por cara:")
    for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,):
        print(f"    {sub}: {conteo[sub]} JPG")
    return conteo


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
        if ok:
            try:
                _reorganizar_piezas_por_cara(
                    carpeta_piezas,
                    dict(generador_caras_tanque.LAST_PIEZAS_POR_CARA),
                )
            except Exception as err:
                print(f"AVISO: fallo en reorganización por cara de PIEZAS_ACOTADAS: {err}")
        return ok
    finally:
        if inv_app is not None:
            _reactivar_machote(inv_app)
        pythoncom.CoUninitialize()
        if ok:
            print("PROCESO COMPLETO: caras y piezas acotadas exportadas.")


if __name__ == "__main__":
    sys.exit(0 if ejecutar() else 1)
