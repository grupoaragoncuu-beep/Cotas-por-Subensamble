import os
import time

import pythoncom
import win32com.client


def _es_instancia_disponible(inv_app):
    """Comprueba que el proxy COM no apunte a un Inventor ya cerrado."""
    try:
        _ = inv_app.Visible
        _ = inv_app.Documents.Count
        return True
    except Exception:
        return False


def com_vivo(inv_app):
    """True si TransientGeometry responde (señal típica post-RPC)."""
    if inv_app is None:
        return False
    try:
        _ = inv_app.TransientGeometry
        _ = inv_app.Documents.Count
        return True
    except Exception:
        return False


def conectar_inventor():
    """
    Conecta a la instancia de Inventor que el usuario ya tiene abierta.

    Al ejecutar desde un .bat, Dispatch/EnsureDispatch puede crear una segunda
    instancia vacía. GetActiveObject usa la que ya está en pantalla con el plano.

    En Python 3.13 + gen_py de Inventor, ``win32com.client.GetActiveObject``
    a veces rompe el wrapper (``KeyError: _dispobj_``); por eso se usa
    ``pythoncom.GetActiveObject`` + ``QueryInterface(IDispatch)``.
    """
    try:
        raw = pythoncom.GetActiveObject("Inventor.Application")
        disp = raw.QueryInterface(pythoncom.IID_IDispatch)
        inv_app = win32com.client.Dispatch(disp)
        if _es_instancia_disponible(inv_app):
            return inv_app
    except Exception:
        pass

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


def reconectar_inventor(espera_s=2.0):
    """
    Reatacha a Inventor tras saturación / RPC entre lotes Abigail.

    No cierra Inventor: solo bombea mensajes COM y vuelve a GetActiveObject.
    """
    espera_s = max(0.5, float(espera_s or 2.0))
    pasos = max(5, int(espera_s * 10))
    for _ in range(pasos):
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        time.sleep(0.1)
    return conectar_inventor()


def localizar_documento(inv_app, display_name=None, full_file_name=None):
    """Busca un documento abierto por DisplayName o FullFileName."""
    if inv_app is None:
        return None
    name_u = str(display_name or "").strip().upper()
    path_u = os.path.normcase(os.path.normpath(str(full_file_name or "")))
    try:
        for i in range(1, inv_app.Documents.Count + 1):
            doc = inv_app.Documents.Item(i)
            try:
                if path_u:
                    ff = os.path.normcase(
                        os.path.normpath(str(doc.FullFileName or ""))
                    )
                    if ff and ff == path_u:
                        return doc
            except Exception:
                pass
            try:
                if name_u and str(doc.DisplayName or "").strip().upper() == name_u:
                    return doc
            except Exception:
                pass
    except Exception:
        return None
    return None


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

