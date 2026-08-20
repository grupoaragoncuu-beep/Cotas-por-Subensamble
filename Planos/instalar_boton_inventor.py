"""
Instala la integracion formal de COTAS ABIGAIL en Autodesk Inventor.

- Escribe config_planos.txt con rutas locales
- Registra la regla en UserApplicationOptions.xml de Inventor
- Define variable de entorno COTAS_ABIGAIL_DIR para el equipo
"""
import glob
import os
import sys
import winreg

from inventor_com import conectar_inventor, configurar_carpeta_reglas_ilogic


NOMBRE_REGLA = "COTAS_ILOGIC_ABIGAIL"
ENV_VAR = "COTAS_ABIGAIL_DIR"
CERRAR_ILOGIC_MARKER = "    </ExternalRuleFilenames>"


def _carpeta_planos():
    return os.path.dirname(os.path.abspath(__file__))


def _carpeta_ilogic(planos_dir):
    return os.path.join(planos_dir, "ilogic")


def _detectar_python():
    """Prefiere la ruta absoluta del intérprete actual (portable entre PCs)."""
    try:
        if sys.executable and os.path.isfile(sys.executable):
            return sys.executable
    except Exception:
        pass
    for candidato in ("python", "py"):
        try:
            import shutil
            ruta = shutil.which(candidato)
            if ruta:
                return ruta
        except Exception:
            pass
    return "python"


def _escribir_config(planos_dir, python_exe):
    cfg_path = os.path.join(_carpeta_ilogic(planos_dir), "config_planos.txt")
    contenido = (
        "# Generado por instalar_boton_inventor.py — NO editar a mano salvo\n"
        "# que sepas lo que haces. Este archivo es local a cada PC.\n"
        f"PLANOS_DIR={planos_dir}\n"
        f"PYTHON_EXE={python_exe}\n"
    )
    with open(cfg_path, "w", encoding="utf-8") as f:
        f.write(contenido)
    return cfg_path


def _guardar_variable_entorno_usuario(nombre, valor):
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            "Environment",
            0,
            winreg.KEY_SET_VALUE,
        ) as key:
            winreg.SetValueEx(key, nombre, 0, winreg.REG_EXPAND_SZ, valor)
        return True
    except OSError as e:
        print(f"AVISO: No se pudo guardar {nombre}: {e}")
        return False


def _imprimir_instrucciones_ribbon():
    print()
    print("=" * 60)
    print(" SIGUIENTE PASO: agregar boton en la cinta de Inventor")
    print("=" * 60)
    print()
    print("IMPORTANTE: cierra y vuelve a abrir Inventor antes de continuar.")
    print()
    print("1. En Inventor, clic derecho en la cinta ->")
    print("   'Personalizar comandos de usuario'")
    print()
    print("2. En 'Elegir comandos de', selecciona:")
    print("   'Reglas de iLogic'")
    print()
    print("3. Busca las reglas y agregalas al panel que prefieras:")
    print("   - COTAS_CARAS_TANQUE       (flujo completo: caras + piezas)")
    print("   - COTAS_POR_SUBENSAMBLE    (solo cotas por caras)")
    print(f"   - {NOMBRE_REGLA}      (solo piezas acotadas)")
    print()
    print("4. Opcional: activa 'Texto' y tamano grande en el boton.")
    print()


def _actualizar_ruta_en_regla_ilogic(planos_dir):
    regla_path = os.path.join(_carpeta_ilogic(planos_dir), f"{NOMBRE_REGLA}.iLogicVb")
    with open(regla_path, "r", encoding="utf-8") as f:
        contenido = f.read()
    if "__PLANOS_DIR__" in contenido:
        contenido = contenido.replace("__PLANOS_DIR__", planos_dir)
        with open(regla_path, "w", encoding="utf-8") as f:
            f.write(contenido)
    return regla_path


def _reglas_de_la_carpeta(ilogic_dir):
    return sorted(glob.glob(os.path.join(ilogic_dir, "*.iLogicVb")))


def _ruta_ya_registrada(contenido, regla_path):
    candidatos = {
        regla_path,
        regla_path.replace("\\", "/"),
        os.path.normcase(regla_path),
    }
    contenido_norm = os.path.normcase(contenido)
    return any(c and os.path.normcase(c) in contenido_norm for c in candidatos)


def _registrar_reglas_en_xml(xml_path, rutas_reglas):
    try:
        with open(xml_path, "r", encoding="utf-16") as f:
            contenido = f.read()
    except Exception as e:
        return False, f"no se pudo leer ({e})"

    if CERRAR_ILOGIC_MARKER not in contenido:
        return False, "no tiene bloque ExternalRuleFilenames"

    nuevas = [r for r in rutas_reglas if not _ruta_ya_registrada(contenido, r)]
    if not nuevas:
        return True, "todas ya estaban registradas"

    entradas = "".join(
        f'      <ExternalRuleFilename Path="{ruta}"/>\n' for ruta in nuevas
    )
    contenido = contenido.replace(
        CERRAR_ILOGIC_MARKER, entradas + CERRAR_ILOGIC_MARKER, 1
    )

    try:
        with open(xml_path, "w", encoding="utf-16") as f:
            f.write(contenido)
    except Exception as e:
        return False, f"no se pudo escribir ({e})"

    nombres = ", ".join(os.path.basename(r) for r in nuevas)
    return True, f"registradas: {nombres}"


def _registrar_reglas_en_inventor(rutas_reglas):
    appdata = os.environ.get("APPDATA", "")
    patron = os.path.join(appdata, "Autodesk", "Inventor *", "UserApplicationOptions.xml")
    resultados = []

    for xml_path in sorted(glob.glob(patron)):
        nombre = os.path.basename(os.path.dirname(xml_path))
        if "Read-only" in nombre or "Interoperability" in nombre:
            continue

        ok, detalle = _registrar_reglas_en_xml(xml_path, rutas_reglas)
        resultados.append((nombre, ok, detalle))

    return resultados


def instalar():
    planos_dir = _carpeta_planos()
    ilogic_dir = _carpeta_ilogic(planos_dir)
    regla_path = os.path.join(ilogic_dir, f"{NOMBRE_REGLA}.iLogicVb")

    print("=" * 60)
    print(" COTAS ABIGAIL - Instalador boton Inventor")
    print("=" * 60)
    print(f"Carpeta Planos: {planos_dir}")
    print(f"Carpeta iLogic: {ilogic_dir}")
    print()
    print("IMPORTANTE: cierra Inventor antes de instalar.")
    print()

    if not os.path.isfile(regla_path):
        print(f"ERROR: Falta la regla externa: {regla_path}")
        return 1

    python_exe = _detectar_python()
    cfg_path = _escribir_config(planos_dir, python_exe)
    regla_path = _actualizar_ruta_en_regla_ilogic(planos_dir)
    print(f"OK Config escrita: {cfg_path}")
    print(f"OK Regla actualizada: {regla_path}")

    if _guardar_variable_entorno_usuario(ENV_VAR, planos_dir):
        print(f"OK Variable de entorno {ENV_VAR} configurada.")
    else:
        print(f"AVISO: Configura manualmente {ENV_VAR}={planos_dir}")

    reglas = _reglas_de_la_carpeta(ilogic_dir)
    print()
    print(f"Registrando {len(reglas)} regla(s) en configuracion de Inventor...")
    for ruta in reglas:
        print(f"  - {os.path.basename(ruta)}")
    registros = _registrar_reglas_en_inventor(reglas)
    if not registros:
        print("AVISO: No se encontro UserApplicationOptions.xml de Inventor.")
    else:
        for nombre, ok, detalle in registros:
            estado = "OK" if ok else "AVISO"
            print(f"  {estado} {nombre}: {detalle}")

    print()
    print("Conectando con Inventor (opcional)...")
    inv_app = conectar_inventor()
    if inv_app is not None:
        ok, detalle = configurar_carpeta_reglas_ilogic(inv_app, ilogic_dir)
        if ok:
            print(f"OK Carpeta iLogic tambien registrada en sesion: {detalle}")

    _imprimir_instrucciones_ribbon()
    print("Instalacion completada.")
    return 0


if __name__ == "__main__":
    sys.exit(instalar())
