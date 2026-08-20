"""
Flujo integrado de COTAS ABIGAIL.

1. Cotas por referencia de las 5 caras del tanque (FRONT, BACK, LEFT, RIGHT, TOP).
2. Flujo original COTAS_ILOGIC_ABIGAIL por pieza, reorganizado en subcarpetas
   por cara.

Ambas salidas se organizan bajo Planos/JPG/<tanque>/ en carpetas separadas.
"""

import os
import re
import shutil
import sys
import time

import pythoncom

import generador_caras_tanque
from generador_caras_tanque import (
    _carpeta_salida_tanque,
    _eliminar_hojas_piezas_residuales,
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

# --- Clasificación por proceso (iProperty escrita por el iLogic Colorimetria).
# Se usa como subcarpetas del flujo PIEZAS_ACOTADAS en lugar de las caras
# geométricas del tanque. El orden fija la salida de logs.
SUBCARPETAS_CLASIFICACION_PIEZAS = (
    "Almacén",
    "Corte",
    "Maquinado",
    "Doblado",
    "Plasma",
    "Plasma Doblado",
)
SUBCARPETA_SIN_CLASIFICAR = "SIN CLASIFICACION"

# Sufijos que agrega el flujo de piezas al final del nombre del JPG. Se usan
# para separar "pieza" de "tipo de cota" y así crear una carpeta por pieza.
# ORDEN IMPORTA: los sufijos compuestos (LARGO_PATA, DIAMETRO_*) deben ir
# ANTES de sus prefijos simples (LARGO, DIAMETRO) para que la alternancia
# del regex los reconozca primero y no truncue el nombre.
_SUFIJOS_JPG_PIEZA = (
    "DIAMETRO_EXTERIOR",
    "DIAMETRO_INTERIOR",
    "LARGO_PATA",
    "ANCHO",
    "LARGO",
    "THK",
    "ALTO",
)
_RE_SUFIJO_PIEZA = re.compile(
    r"^(?P<pieza>.+?)_(?P<tipo>"
    + "|".join(_SUFIJOS_JPG_PIEZA)
    + r")_\d+$",
    re.IGNORECASE,
)


def _limpiar_exportacion_piezas(carpeta, incremental=False):
    """
    Vacía la carpeta de piezas y sus subcarpetas antes de exportar.

    Si `incremental` es True se respeta lo que ya haya en la carpeta (modo F):
    solo se crea la estructura si no existe.

    Limpia AMBOS esquemas de subcarpetas:
    - Cara geométrica (FRONT/BACK/LEFT/RIGHT/TOP/OTROS) — flujo legacy.
    - Clasificación por proceso (Almacén/Corte/.../SIN CLASIFICACION) — flujo
      nuevo basado en iProperty.

    Así una migración de esquema no deja carpetas viejas mezcladas con
    las nuevas.
    """
    os.makedirs(carpeta, exist_ok=True)
    if incremental:
        return

    # Nombres válidos a limpiar: mantenemos también comparación con acentos
    # para las nuevas categorías, así "Almacén" hace match aunque el sistema
    # de archivos lo devuelva con la misma codificación.
    caras_upper = {c.upper() for c in SUBCARPETAS_CARA_PIEZAS}
    caras_upper.add(SUBCARPETA_OTROS_PIEZAS.upper())
    clasificaciones_ci = {
        c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS
    }
    clasificaciones_ci.add(SUBCARPETA_SIN_CLASIFICAR.casefold())

    for nombre in os.listdir(carpeta):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isdir(ruta):
            if (
                nombre.upper() in caras_upper
                or nombre.casefold() in clasificaciones_ci
            ):
                try:
                    shutil.rmtree(ruta)
                except OSError as err:
                    print(f"AVISO: no se pudo vaciar {nombre}/: {err}")
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


def _extraer_pieza_de_jpg(nombre_archivo):
    """
    Devuelve el nombre de pieza contenido en un JPG del flujo por pieza.

    Los archivos exportados terminan en `_<TIPO>_<n>.jpg`, donde `<TIPO>` es
    uno de LARGO/ANCHO/THK/DIAMETRO_*. Todo lo previo es el nombre de pieza.
    Si no coincide, se devuelve el nombre base sin extensión.
    """
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    match = _RE_SUFIJO_PIEZA.match(base)
    if not match:
        return base
    return match.group("pieza")


def _nombre_carpeta_pieza(nombre_pieza):
    """Sanea el nombre de pieza para usarlo como carpeta en Windows."""
    limpio = re.sub(r'[<>:"/\\|?*]+', "_", str(nombre_pieza).strip())
    limpio = limpio.rstrip(". ")
    return limpio or "PIEZA"


def _reorganizar_piezas_por_cara(carpeta_piezas, mapa_por_cara):
    """
    Mueve cada JPG a `<CARA>/<PIEZA>/<archivo>.jpg`.

    - Piezas sin match a una cara caen en `OTROS/<PIEZA>/`.
    - Si falla el movimiento, el JPG queda en su ubicación original.
    """
    try:
        os.makedirs(carpeta_piezas, exist_ok=True)
        for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,):
            os.makedirs(os.path.join(carpeta_piezas, sub), exist_ok=True)
    except OSError as err:
        print(f"AVISO: no se pudieron preparar subcarpetas de PIEZAS_ACOTADAS: {err}")
        return {}

    conteo_cara = {sub: 0 for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,)}
    piezas_por_cara = {sub: set() for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,)}
    try:
        entradas = list(os.listdir(carpeta_piezas))
    except OSError as err:
        print(f"AVISO: no se pudo listar {carpeta_piezas}: {err}")
        return conteo_cara

    for nombre in entradas:
        ruta = os.path.join(carpeta_piezas, nombre)
        if os.path.isdir(ruta):
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        cara = _cara_para_pieza(nombre, mapa_por_cara)
        destino_sub = cara if cara in SUBCARPETAS_CARA_PIEZAS else SUBCARPETA_OTROS_PIEZAS
        pieza = _extraer_pieza_de_jpg(nombre)
        pieza_folder = _nombre_carpeta_pieza(pieza)
        destino_dir = os.path.join(carpeta_piezas, destino_sub, pieza_folder)
        try:
            os.makedirs(destino_dir, exist_ok=True)
            destino = os.path.join(destino_dir, nombre)
            if os.path.exists(destino):
                os.remove(destino)
            shutil.move(ruta, destino)
            conteo_cara[destino_sub] += 1
            piezas_por_cara[destino_sub].add(pieza_folder)
        except OSError as err:
            print(
                f"AVISO: no se pudo mover '{nombre}' a "
                f"{destino_sub}/{pieza_folder}/: {err}"
            )

    print("  PIEZAS_ACOTADAS por cara:")
    for sub in SUBCARPETAS_CARA_PIEZAS + (SUBCARPETA_OTROS_PIEZAS,):
        print(
            f"    {sub}: {conteo_cara[sub]} JPG en "
            f"{len(piezas_por_cara[sub])} piezas"
        )
    return conteo_cara


def _clasificacion_para_pieza(nombre_archivo, mapa_por_clasificacion):
    """
    Devuelve la clasificación asignada al JPG según el mapa
    ``{clasificación: set(nombres_pieza)}`` construido leyendo el iProperty.

    Usa el mismo matching por substring bidireccional que
    ``_cara_para_pieza`` (los JPGs usan el nombre corto de la pieza y el
    mapa puede tener el nombre completo con revisión/estado).
    """
    if not mapa_por_clasificacion:
        return None
    base_archivo = os.path.splitext(os.path.basename(nombre_archivo))[0]
    clave_archivo = _clave_pieza(base_archivo)
    if not clave_archivo:
        return None
    mejor_clase = None
    mejor_len = 0
    for clase, piezas in mapa_por_clasificacion.items():
        for pieza in piezas:
            clave_pieza = _clave_pieza(pieza)
            if not clave_pieza:
                continue
            if clave_pieza in clave_archivo or clave_archivo in clave_pieza:
                if len(clave_pieza) > mejor_len:
                    mejor_clase = clase
                    mejor_len = len(clave_pieza)
    return mejor_clase


def _reorganizar_piezas_por_clasificacion(carpeta_piezas, mapa_por_clasificacion):
    """
    Mueve cada JPG a ``<CLASIFICACIÓN>/<PIEZA>/<archivo>.jpg``.

    Estructura resultante en PIEZAS_ACOTADAS:
        Almacén/<PIEZA>/*.jpg
        Corte/<PIEZA>/*.jpg
        Maquinado/<PIEZA>/*.jpg
        Doblado/<PIEZA>/*.jpg
        Plasma/<PIEZA>/*.jpg
        Plasma Doblado/<PIEZA>/*.jpg
        SIN CLASIFICACION/<PIEZA>/*.jpg   (para piezas sin iProperty)

    Piezas sin match en el mapa caen en ``SIN CLASIFICACION/`` (por decisión
    del usuario: no perder nada aunque el iLogic Colorimetria no se haya
    corrido para esa pieza).
    """
    subcarpetas = SUBCARPETAS_CLASIFICACION_PIEZAS + (SUBCARPETA_SIN_CLASIFICAR,)
    try:
        os.makedirs(carpeta_piezas, exist_ok=True)
        for sub in subcarpetas:
            os.makedirs(os.path.join(carpeta_piezas, sub), exist_ok=True)
    except OSError as err:
        print(
            f"AVISO: no se pudieron preparar subcarpetas de PIEZAS_ACOTADAS "
            f"por clasificación: {err}"
        )
        return {}

    conteo = {sub: 0 for sub in subcarpetas}
    piezas_por_clase = {sub: set() for sub in subcarpetas}
    try:
        entradas = list(os.listdir(carpeta_piezas))
    except OSError as err:
        print(f"AVISO: no se pudo listar {carpeta_piezas}: {err}")
        return conteo

    # Set case-insensitive de clasificaciones válidas para validar destino.
    clases_validas_cf = {c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS}

    for nombre in entradas:
        ruta = os.path.join(carpeta_piezas, nombre)
        if os.path.isdir(ruta):
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        clase = _clasificacion_para_pieza(nombre, mapa_por_clasificacion)
        if clase and clase.casefold() in clases_validas_cf:
            # Preservar el spelling canónico (con acento/mayúsculas correctas).
            destino_sub = next(
                s for s in SUBCARPETAS_CLASIFICACION_PIEZAS
                if s.casefold() == clase.casefold()
            )
        else:
            destino_sub = SUBCARPETA_SIN_CLASIFICAR
        pieza = _extraer_pieza_de_jpg(nombre)
        pieza_folder = _nombre_carpeta_pieza(pieza)
        destino_dir = os.path.join(carpeta_piezas, destino_sub, pieza_folder)
        try:
            os.makedirs(destino_dir, exist_ok=True)
            destino = os.path.join(destino_dir, nombre)
            if os.path.exists(destino):
                os.remove(destino)
            shutil.move(ruta, destino)
            conteo[destino_sub] += 1
            piezas_por_clase[destino_sub].add(pieza_folder)
        except OSError as err:
            print(
                f"AVISO: no se pudo mover '{nombre}' a "
                f"{destino_sub}/{pieza_folder}/: {err}"
            )

    print("  PIEZAS_ACOTADAS por clasificación:")
    for sub in subcarpetas:
        print(
            f"    {sub}: {conteo[sub]} JPG en "
            f"{len(piezas_por_clase[sub])} piezas"
        )
    return conteo


def _recuperar_antes_de_piezas(inv_app, plano):
    """
    Tras COTAS_POR_REFERENCIA Inventor queda saturado; sin este respiro
    creador_vistas falla al leer AllLeafOccurrences / CopyTo.
    """
    print("  Recuperando Inventor antes de PIEZAS_ACOTADAS...")
    try:
        _eliminar_hojas_piezas_residuales(plano)
    except Exception as err:
        print(f"  AVISO limpiando hojas residuales de piezas: {err}")
    try:
        inv_app.SilentOperation = False
        inv_app.ScreenUpdating = True
    except Exception:
        pass
    try:
        inv_app.ActiveDocument.Update()
    except Exception:
        pass
    try:
        inv_app.ActiveView.Update()
    except Exception:
        pass
    try:
        inv_app.UserInterfaceManager.DoEvents()
    except Exception:
        pass
    for _ in range(20):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)
    print("  Recuperación lista.")


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
        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in ("1", "true", "TRUE", "yes")
        _limpiar_exportacion_piezas(carpeta_piezas, incremental=incremental)
        # Refrescar handles tras caras + vaciar hojas residuales + respiro COM.
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            print(
                "ERROR: se perdió el plano o ensamble al pasar a "
                "cotas por pieza."
            )
            return False
        _recuperar_antes_de_piezas(inv_app, plano)
        print("Paso 2/2: Ejecutando cotas originales por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        if incremental:
            print("  Modo incremental ACTIVO (PIEZAS_INCREMENTAL=1).")

        # Detectar clasificación por proceso (iProperty) ANTES de exportar
        # para que quede persistida y disponible al reorganizar.
        print(
            "  Leyendo clasificación (iProperty) de cada pieza del ensamble..."
        )
        try:
            mapa_clasificacion = (
                generador_caras_tanque.detectar_mapa_piezas_por_clasificacion(
                    inv_app, ensamble
                )
            )
            if mapa_clasificacion:
                ruta_mapa = (
                    generador_caras_tanque
                    .guardar_mapa_piezas_por_clasificacion(
                        carpeta_tanque, mapa_clasificacion
                    )
                )
                if ruta_mapa:
                    print(f"  Mapa por clasificación persistido en: {ruta_mapa}")
        except Exception as err:
            print(f"AVISO: fallo detectando clasificación por iProperty: {err}")

        # Equivalente funcional de COTAS_ILOGIC_ABIGAIL, pero sin llamar a
        # iLogic GenerarVistas (esa llamada queda bloqueada en este tanque).
        # Conserva creador_vistas + cotas.py + THK.py + exportación JPG.
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
                _reorganizar_piezas_por_clasificacion(
                    carpeta_piezas,
                    dict(
                        generador_caras_tanque.LAST_PIEZAS_POR_CLASIFICACION
                    ),
                )
            except Exception as err:
                print(
                    f"AVISO: fallo en reorganización por clasificación de "
                    f"PIEZAS_ACOTADAS: {err}"
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
