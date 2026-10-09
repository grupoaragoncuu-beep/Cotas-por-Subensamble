# -*- coding: utf-8 -*-
"""Fuerza ID de SP-767_2 desde círculos del modelo (HLR no ancla el bore)."""
from __future__ import annotations

import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

JOB = "62223-1246-A01"
PIECE = "SP-767_2"
JOB_ROOT = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
)
Y_JPGS = (
    r"Y:\ARGA METALS CORPORATE SYSTEM\TANKS\OTC\62223"
    r"\DOSSIER FILES\JPGS\PIEZAS_ACOTADAS"
)


def _log(m: str) -> None:
    print(m, flush=True)


def main() -> int:
    import pythoncom
    import win32com.client
    from inventor_com import conectar_inventor
    from generador_caras_tanque import (
        _carpeta_salida_tanque,
        _como_ensamble,
        _obtener_plano_activo,
    )
    from generador_tanque_completo import CARPETA_PIEZAS_ACOTADAS
    from generador_vistas import ejecutar_flujo_desde_app
    from cotas_dossier_registro import (
        iniciar_sesion_dossier,
        publicar_pieza_carpeta,
        sincronizar_lote_carpetas,
        guardar_contexto_dossier,
    )
    from producto_tipo import clasificar_producto, aplicar_unidad_producto
    import creador_vistas as cv
    import diametro as D
    from cota_estilo import texto_cota_dibujo

    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    os.environ["COTAS_DOSSIER"] = "0"
    os.environ["PIEZAS_FILTRO"] = PIECE
    os.environ["PIEZAS_INCREMENTAL"] = "1"
    os.environ["SKIP_ENSAMBLES_IND"] = "1"

    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        ensamble = None
        for i in range(1, int(inv.Documents.Count) + 1):
            d = inv.Documents.Item(i)
            nom = str(getattr(d, "DisplayName", "") or "")
            if nom.upper().startswith("62223-1246-A01") and nom.upper().endswith(
                ".IAM"
            ):
                ensamble = _como_ensamble(d)
                break
        plano = _obtener_plano_activo(inv)
        if ensamble is None or plano is None:
            _log("ERROR: sin ensamble/plano")
            return 1
        try:
            plano.Activate()
        except Exception:
            pass

        info = clasificar_producto(ensamble)
        aplicar_unidad_producto(ensamble=ensamble, info=info)
        cv.configurar_producto_flujo("TANQUE")
        cv.configurar_piezas_doblado({PIECE})
        cv.configurar_piezas_corte(set())

        carpeta = os.path.join(
            _carpeta_salida_tanque(plano, ensamble), CARPETA_PIEZAS_ACOTADAS
        )
        # Crear vistas (sin wipe de JPG existentes)
        ok = bool(
            ejecutar_flujo_desde_app(
                inv,
                ensamble,
                plano,
                carpeta_salida=carpeta,
                incremental=False,
                catalogo_piezas={PIECE},
            )
        )
        _log(f"flujo ok={ok}")

        # Buscar hoja FRENTE/DESPLIEGUE con vista y forzar ID
        plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
        tg = inv.TransientGeometry
        hoja_src = None
        for i in range(1, plano.Sheets.Count + 1):
            h = plano.Sheets.Item(i)
            nu = str(h.Name).upper()
            if PIECE.upper() not in nu:
                continue
            if "FRENTE_1" in nu or "DIAMETRO_EXTERIOR" in nu:
                if h.DrawingViews.Count >= 1:
                    hoja_src = h
                    break
        if hoja_src is None:
            _log("ERROR: sin hoja FRENTE para anclar ID")
            return 2

        try:
            hoja_src.Activate()
        except Exception:
            pass
        vista = hoja_src.DrawingViews.Item(1)
        diams = D._diametros_circulo_modelo_hoja_cm(vista)
        od = D._od_max_circulo_modelo_hoja_cm(vista)
        _log(f"diams_hoja_cm={ [round(x,2) for x in diams] } od={od and round(od,2)}")
        id_esp = None
        if od and diams:
            for d in diams:
                if d < od * 0.95 and d >= od * 0.22:
                    id_esp = d
                    break
        # Preferir bore más interior entre candidatos grandes (no el escalón
        # exterior): el menor ≥ 35% OD y ≤ 85% OD.
        if od and diams:
            bores = [d for d in diams if od * 0.35 <= d <= od * 0.85]
            if bores:
                id_esp = min(bores)
        _log(f"id_esp_cm={id_esp and round(id_esp,2)}")
        if not id_esp:
            _log("ERROR: sin ID candidato en modelo")
            return 3

        anillo = D._anillo_por_curvas_cerca_od(vista, id_esp, tol_frac=0.20)
        if anillo is None:
            # Reintento con cada bore
            for d in diams:
                if od and od * 0.30 <= d <= od * 0.90:
                    anillo = D._anillo_por_curvas_cerca_od(vista, d, tol_frac=0.22)
                    if anillo is not None:
                        id_esp = d
                        break
        if anillo is None:
            _log("ERROR: sin curva DrawingCurve para ID")
            return 4

        # Nueva hoja solo-ID
        nombre_nueva = f"{PIECE}_DESPLIEGUE_DIAMETRO_INTERIOR"
        for j in range(plano.Sheets.Count, 0, -1):
            try:
                h = plano.Sheets.Item(j)
                if str(h.Name).upper().startswith(nombre_nueva.upper()):
                    h.Delete()
            except Exception:
                continue
        nueva = hoja_src.CopyTo(plano)
        try:
            nueva.Name = nombre_nueva
        except Exception:
            pass
        try:
            nueva.Activate()
        except Exception:
            pass
        # Limpiar cotas heredadas
        try:
            dims = nueva.DrawingDimensions.GeneralDimensions
            for di in range(int(dims.Count), 0, -1):
                try:
                    dims.Item(di).Delete()
                except Exception:
                    pass
        except Exception:
            pass
        vista_n = nueva.DrawingViews.Item(1)
        # Re-resolver curva en la vista copiada
        anillo = D._anillo_por_curvas_cerca_od(vista_n, id_esp, tol_frac=0.22)
        if anillo is None:
            _log("ERROR: curva ID perdida tras CopyTo")
            return 5
        intent = nueva.CreateGeometryIntent(anillo["curva"])
        pt = D._punto_texto_barreno(
            nueva, tg, vista_n, anillo["cx"], anillo["cy"], anillo["tamaño"]
        )
        dim = nueva.DrawingDimensions.GeneralDimensions.AddDiameter(pt, intent)
        D._aplicar_estilo_y_fuera_pieza(dim, nueva, tg, vista_n)
        try:
            mv = float(dim.ModelValue)
            txt = texto_cota_dibujo(mv, nueva)
            if txt:
                dim.HideValue = True
                from cota_estilo import armar_formatted_texto_cota, get_cota_font_size_cm, COTA_BOLD

                dim.Text.FormattedText = armar_formatted_texto_cota(
                    txt, font_cm=get_cota_font_size_cm(), bold=COTA_BOLD
                )
        except Exception as e:
            _log(f"AVISO estilo: {e}")
        _log(f"OK ID ModelValue={float(dim.ModelValue)/2.54:.3f} in")

        # Exportar solo esta hoja vía flujo parcial: usar generador export
        from generador_vistas import _exportar_hoja_jpg  # type: ignore

        dest_dir = (
            os.path.join(
                carpeta, "Corte", "Maquinado", "Maquinados metal", PIECE
            )
        )
        os.makedirs(dest_dir, exist_ok=True)
        # Nombre esperado
        val_in = float(dim.ModelValue) / 2.54
        fname = f"{JOB}__{PIECE}__ID_{val_in:.2f}in.jpg"
        # Fallback si helper no existe: ActiveSheet export
        try:
            path = _exportar_hoja_jpg(inv, nueva, dest_dir, fname)
            _log(f"export helper -> {path}")
        except Exception:
            # Export manual
            out = os.path.join(dest_dir, fname)
            try:
                nueva.Activate()
            except Exception:
                pass
            # Usar SaveAs JPG del sheet via Inventor — o captura del flujo
            from generador_vistas import exportar_captura_hoja

            try:
                exportar_captura_hoja(inv, nueva, out)
                _log(f"export captura -> {out}")
            except Exception as e2:
                _log(f"ERROR export: {e2}")
                # último: copiar vía ejecutar rename de export existente
                return 6

        # Espejo Y + Doblado
        for mirror in (
            dest_dir,
            os.path.join(carpeta, "Doblado", "Metal", PIECE),
            os.path.join(Y_JPGS, "Corte", "Maquinado", "Maquinados metal", PIECE),
            os.path.join(Y_JPGS, "Doblado", "Metal", PIECE),
        ):
            os.makedirs(mirror, exist_ok=True)
            src = os.path.join(dest_dir, fname)
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(mirror, fname))

        os.environ["COTAS_DOSSIER"] = "1"
        iniciar_sesion_dossier(JOB, JOB_ROOT)
        try:
            guardar_contexto_dossier(cliente="OTC", producto="TANQUE")
        except Exception:
            pass
        unc, n = publicar_pieza_carpeta(dest_dir)
        _log(f"publish n={n} unc={bool(unc)}")
        sincronizar_lote_carpetas([dest_dir], job=JOB)
        return 0
    finally:
        try:
            cv.configurar_piezas_doblado([])
            cv.configurar_piezas_corte([])
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
