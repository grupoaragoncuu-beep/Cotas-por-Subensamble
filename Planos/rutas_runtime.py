"""
Rutas portables del proyecto COTAS ABIGAIL.

Ninguna ruta de máquina específica (C:\\Proyectos\\..., C:\\Users\\..., C:\\Temp\\...)
debe vivir hardcodeada en el flujo de producción. Este módulo es la fuente
única para:

- Carpeta ``Planos`` (código)
- Carpeta de runtime (antes ``C:\\Temp``)
- Lectura de ``ilogic/config_planos.txt`` generado por el instalador

En otra PC: clonar el repo y ejecutar ``instalar_boton_inventor.bat``.
"""

from __future__ import annotations

import os
import sys


ENV_PLANOS_DIR = "COTAS_ABIGAIL_DIR"


def carpeta_planos():
    """Carpeta ``Planos`` (donde viven los .py del flujo)."""
    return os.path.dirname(os.path.abspath(__file__))


def carpeta_repo():
    """Raíz del repositorio (padre de ``Planos``)."""
    return os.path.dirname(carpeta_planos())


def carpeta_ilogic():
    return os.path.join(carpeta_planos(), "ilogic")


def carpeta_runtime():
    """
    Archivos temporales de comunicación entre módulos
    (listas de hojas diámetro, piezas sólidas, etc.).

    Antes estaba en ``C:\\Temp``; ahora es local al proyecto para ser
    portable y fácil de depurar. Se crea al vuelo.
    """
    base = os.path.join(carpeta_planos(), ".runtime")
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        pass
    return base


def ruta_hojas_diametro():
    return os.path.join(carpeta_runtime(), "hojas_para_diametro.txt")


def ruta_piezas_solidas():
    return os.path.join(carpeta_runtime(), "piezas_solidas.txt")


def ruta_config_planos():
    return os.path.join(carpeta_ilogic(), "config_planos.txt")


def leer_config_planos():
    """
    Lee ``ilogic/config_planos.txt`` como dict de claves mayúsculas.

    Formato esperado (una clave por línea)::

        PLANOS_DIR=C:\\ruta\\a\\Planos
        PYTHON_EXE=C:\\...\\python.exe
    """
    cfg = {}
    ruta = ruta_config_planos()
    if not os.path.isfile(ruta):
        return cfg
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue
                clave, valor = linea.split("=", 1)
                cfg[clave.strip().upper()] = valor.strip()
    except OSError:
        pass
    return cfg


def resolver_planos_dir():
    """
    Resuelve la carpeta Planos en orden:

    1. Variable de entorno ``COTAS_ABIGAIL_DIR``
    2. ``PLANOS_DIR`` en ``config_planos.txt``
    3. Carpeta de este módulo (fallback local)
    """
    env = os.environ.get(ENV_PLANOS_DIR, "").strip()
    if env and os.path.isdir(env):
        return os.path.abspath(env)

    cfg = leer_config_planos()
    desde_cfg = cfg.get("PLANOS_DIR", "").strip()
    if desde_cfg and os.path.isdir(desde_cfg):
        return os.path.abspath(desde_cfg)

    return carpeta_planos()


def resolver_python_exe():
    """Python a usar desde iLogic / scripts externos."""
    cfg = leer_config_planos()
    desde_cfg = cfg.get("PYTHON_EXE", "").strip()
    if desde_cfg and (
        os.path.isfile(desde_cfg) or desde_cfg.lower() in ("python", "py")
    ):
        return desde_cfg
    return sys.executable or "python"


# Compatibilidad con imports antiguos (constantes evaluadas al importar).
# Los módulos que necesiten la ruta fresca deben llamar a las funciones.
RUTA_HOJAS_DIAMETRO = ruta_hojas_diametro()
RUTA_PIEZAS_SOLIDAS = ruta_piezas_solidas()
