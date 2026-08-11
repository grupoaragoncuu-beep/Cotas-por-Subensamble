import os

import win32com.client


def _es_instancia_disponible(inv_app):
    """Comprueba que el proxy COM no apunte a un Inventor ya cerrado."""
    try:
        _ = inv_app.Visible
        _ = inv_app.Documents.Count
        return True
    except Exception:
        return False


def conectar_inventor():
    """
    Conecta a la instancia de Inventor que el usuario ya tiene abierta.

    Al ejecutar desde un .bat, Dispatch/EnsureDispatch puede crear una segunda
    instancia vacía. GetActiveObject usa la que ya está en pantalla con el plano.
    """
    try:
        inv_app = win32com.client.GetActiveObject("Inventor.Application")
        if _es_instancia_disponible(inv_app):
            return inv_app
    except Exception:
        pass

    try:
        inv_app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        if _es_instancia_disponible(inv_app):
            return inv_app
    except Exception:
        pass

    inv_app = win32com.client.Dispatch("Inventor.Application")
    if not _es_instancia_disponible(inv_app):
        raise RuntimeError(
            "No hay una instancia disponible de Autodesk Inventor."
        )
    return inv_app


def obtener_ilogic_automation(inv_app):
    """
    Devuelve el objeto Automation del complemento iLogic, o None si no está activo.
    """
    try:
        for addin in inv_app.ApplicationAddIns:
            try:
                if "iLogic" in addin.DisplayName:
                    return addin.Automation
            except Exception:
                pass
    except Exception:
        return None

    return None


def configurar_carpeta_reglas_ilogic(inv_app, carpeta_ilogic):
    """
    Registra la carpeta de reglas externas en iLogic (persistente por sesión/config).
    """
    ilogic = obtener_ilogic_automation(inv_app)
    if ilogic is None:
        return False, "No se encontró iLogic."

    carpeta_ilogic = os.path.normpath(carpeta_ilogic)

    try:
        opciones = ilogic.FileOptions
        actuales = []

        try:
            dirs = opciones.ExternalRuleDirectories
            if dirs is not None:
                for i in range(len(dirs)):
                    actuales.append(os.path.normpath(str(dirs[i])))
        except Exception:
            pass

        if carpeta_ilogic not in actuales:
            actuales.append(carpeta_ilogic)
            opciones.ExternalRuleDirectories = actuales

        return True, carpeta_ilogic
    except Exception as e:
        return False, str(e)

