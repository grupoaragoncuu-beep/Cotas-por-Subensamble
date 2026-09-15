"""
Flujo de PIEZAS_ACOTADAS (+ ensambles independientes en corrida completa).

Genera las cotas por pieza (largo/ancho/thk/diámetros). Con
``--seleccion seleccion_caras.json`` (Top Cover + SEGM1..4) organiza la
salida en ``PIEZAS_ACOTADAS/<SEGM*|TOP>/<CLASIFICACIÓN>/<PIEZA>/``.

Tras las piezas (solo corrida completa, sin ``--solo``), el MVP bbox de
kits en ``ENSAMBLES_INDEPENDIENTES/`` queda **apagado por defecto**
(activar con ``ENSAMBLES_INDEPENDIENTES=1``). El instructivo de armado
irá en regla iLogic propia.

Con ``--solo SEGM2`` (COTAS_POR_SEG_PIEZAS) solo acota las piezas del
catálogo de esa cara — prueba rápida sin el tanque completo.

Sin selección, conserva solo ``<CLASIFICACIÓN>/<PIEZA>/``.

Se ejecuta desde la regla iLogic ``COTAS_ILOGIC_ABIGAIL`` o
``COTAS_POR_SEG_PIEZAS``.
"""

from __future__ import annotations

import os
import shutil
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
    _parse_solo,
    cargar_mapa_piezas_por_clasificacion,
    construir_mapa_piezas_desde_seleccion,
    detectar_mapa_piezas_por_clasificacion,
    guardar_mapa_piezas_por_cara,
    guardar_mapa_piezas_por_clasificacion,
)
from generador_tanque_completo import (
    CARPETA_PIEZAS_ACOTADAS,
    SUBCARPETAS_CARA_SELECCION,
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


def _limpiar_solo_cara_piezas(carpeta_piezas, cara):
    """Vacía únicamente ``PIEZAS_ACOTADAS/<CARA>/`` (modo --solo)."""
    os.makedirs(carpeta_piezas, exist_ok=True)
    destino = os.path.join(carpeta_piezas, cara)
    if os.path.isdir(destino):
        try:
            shutil.rmtree(destino)
            print(f"  Limpieza modo --solo: vaciada {cara}/")
        except OSError as err:
            print(f"  AVISO: no se pudo vaciar {cara}/: {err}")
    os.makedirs(destino, exist_ok=True)


def _reorganizar_clasificacion_dentro_caras(carpeta_piezas, mapa_clasificacion):
    """
    Tras ``_reorganizar_piezas_por_cara``, anida cada JPG en
    ``<CARA>/<CLASIFICACIÓN>/<PIEZA>/`` (Corte anida Corte/Doblado/Estañado).
    """
    from generador_tanque_completo import (
        SUBCARPETA_SIN_CLASIFICAR,
        SUBCARPETAS_CLASIFICACION_PIEZAS,
        SUBCARPETA_OTROS_PIEZAS,
        _clasificacion_para_pieza,
        _destino_dirs_clasificacion,
        _limpiar_carpetas_cara_vacias_y_legacy,
        STAGING_DESPLIEGUE,
        STAGING_ESTANIADO,
    )

    caras = list(SUBCARPETAS_CARA_SELECCION) + [SUBCARPETA_OTROS_PIEZAS]
    clases_validas = {c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS}

    for cara in caras:
        cara_dir = os.path.join(carpeta_piezas, cara)
        if not os.path.isdir(cara_dir):
            continue
        jpgs = []
        for root, _dirs, files in os.walk(cara_dir):
            for nombre in files:
                if nombre.lower().endswith(".jpg"):
                    jpgs.append(os.path.join(root, nombre))
        # Staging bajo la cara (si exportó ahí).
        for staging in (STAGING_DESPLIEGUE, STAGING_ESTANIADO):
            staging_dir = os.path.join(cara_dir, staging)
            if os.path.isdir(staging_dir):
                for nombre in os.listdir(staging_dir):
                    if nombre.lower().endswith(".jpg"):
                        jpgs.append(os.path.join(staging_dir, nombre))

        for ruta in jpgs:
            nombre = os.path.basename(ruta)
            staging = None
            parent = os.path.basename(os.path.dirname(ruta))
            if parent in (STAGING_DESPLIEGUE, STAGING_ESTANIADO):
                staging = parent
            clase = _clasificacion_para_pieza(nombre, mapa_clasificacion)
            if clase and clase.casefold() in clases_validas:
                destino_clase = next(
                    s
                    for s in SUBCARPETAS_CLASIFICACION_PIEZAS
                    if s.casefold() == clase.casefold()
                )
            else:
                destino_clase = SUBCARPETA_SIN_CLASIFICAR
            # Destino relativo a cara_dir (no a PIEZAS_ACOTADAS raíz).
            destino_dir, _clave, _ = _destino_dirs_clasificacion(
                cara_dir, destino_clase, nombre, staging
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
                    f"AVISO: no se pudo mover '{nombre}' en {cara}/: {err}"
                )

        for staging in (STAGING_DESPLIEGUE, STAGING_ESTANIADO):
            staging_dir = os.path.join(cara_dir, staging)
            if os.path.isdir(staging_dir):
                try:
                    shutil.rmtree(staging_dir, ignore_errors=True)
                except OSError:
                    pass

    _limpiar_carpetas_cara_vacias_y_legacy(carpeta_piezas)


def ejecutar(ruta_seleccion=None, solo_cara=None):
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

        # BOARD: Abigail acota TODAS las piezas (cara mayor → L/A; SM → THK).
        # Cobre ABB/GENE/RLG: además 1× captura SIN_COTA desde FRENTE.
        # El desvío a kits ViewCube es solo para reglas de caras.
        es_board = False
        try:
            from producto_tipo import clasificar_producto, aplicar_unidad_producto

            info = clasificar_producto(ensamble)
            print(
                f"  Producto: {info.get('tipo')} / {info.get('familia')} "
                f"— {info.get('motivo')}"
            )
            aplicar_unidad_producto(ensamble=ensamble, info=info)
            if info.get("tipo") == "BOARD":
                es_board = True
                print(
                    "  BOARD: flujo PIEZAS completo; "
                    "ABB/GENE/RLG → cota + SIN_COTA."
                )
        except Exception as exc_cls:
            print(f"  AVISO clasificar producto: {exc_cls}")

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        carpeta_piezas = os.path.join(carpeta_tanque, CARPETA_PIEZAS_ACOTADAS)

        try:
            from cotas_dossier_registro import iniciar_sesion_dossier
            from nomenclatura_capturas import nombre_job_desde_ensamble

            job_tok = nombre_job_desde_ensamble(ensamble)
            iniciar_sesion_dossier(job_tok)
        except Exception as exc_dos:
            print(f"  AVISO dossier sesión: {exc_dos}")

        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in (
            "1",
            "true",
            "TRUE",
            "yes",
        )

        solo = str(solo_cara or "").strip().upper() or None
        if solo and solo not in SUBCARPETAS_CARA_SELECCION:
            print(
                f"ERROR: --solo {solo} no válido. "
                f"Usa: {', '.join(SUBCARPETAS_CARA_SELECCION)}"
            )
            return False

        mapa_por_cara = {}
        if ruta_seleccion:
            if not os.path.isfile(ruta_seleccion):
                print(f"ERROR: No existe el JSON de selección: {ruta_seleccion}")
                return False
            print(f"  Selección manual (Top+SEGM): {ruta_seleccion}")
            seleccion = _cargar_seleccion_caras(ruta_seleccion)
            if not solo:
                solo_json = str(seleccion.get("solo") or "").strip().upper()
                if solo_json:
                    solo = solo_json
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

        catalogo_filtro = None
        mapa_reorg = mapa_por_cara
        if solo:
            if not mapa_por_cara or solo not in mapa_por_cara:
                print(
                    f"ERROR: modo --solo {solo}: no hay catálogo de piezas "
                    "para esa cara en la selección. Revisa el pick."
                )
                return False
            catalogo_filtro = set(mapa_por_cara[solo] or set())
            mapa_reorg = {solo: set(catalogo_filtro)}
            print(
                f"  Modo COTAS_POR_SEG_PIEZAS: solo {solo} "
                f"({len(catalogo_filtro)} nombres en catálogo)"
            )
            if not catalogo_filtro:
                print(f"ERROR: catálogo vacío para {solo}.")
                return False

        # BOARD: acota TODAS las piezas únicas. Cobre (ABB/GENE/RLG) = solo
        # doble captura SIN_COTA al exportar; no limita qué se procesa.
        # Barrenos: cualquier chapa con huecos → DESPLIEGUE + X/Y/HOLE/THK
        # (también en tanques vía creador_vistas._debe_crear_despliegue).
        if es_board and catalogo_filtro is None:
            print(
                "  BOARD: alcance completo (todas las piezas únicas). "
                "Barrenos en chapa → DESPLIEGUE X/Y/HOLE/THK. "
                "Cobre ABB/GENE/RLG → JPG + SIN_COTA + ESTANIADO."
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
            try:
                import creador_vistas as _cv_corte

                nombres_corte = set()
                for clase, piezas_c in (mapa_clasificacion or {}).items():
                    if str(clase).casefold() == "corte":
                        nombres_corte.update(piezas_c or [])
                _cv_corte.configurar_piezas_corte(nombres_corte)
                if nombres_corte:
                    print(
                        f"  Corte iProp: {len(nombres_corte)} piezas → "
                        "DESPLIEGUE forzado. "
                        "Otras chapas con barrenos también van a flat."
                    )
            except Exception as exc_corte:
                print(f"  AVISO configurar piezas Corte: {exc_corte}")
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

        if solo:
            _limpiar_solo_cara_piezas(carpeta_piezas, solo)
        else:
            _limpiar_exportacion_piezas(carpeta_piezas, incremental=incremental)
        _recuperar_antes_de_piezas(inv_app, plano)

        print("Ejecutando cotas por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        if solo:
            print(f"  Destino: {carpeta_piezas}\\{solo}\\")
        if incremental:
            print("  Modo incremental ACTIVO (PIEZAS_INCREMENTAL=1).")

        ok = bool(
            ejecutar_flujo_desde_app(
                inv_app,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=incremental,
                catalogo_piezas=catalogo_filtro,
            )
        )
        if ok:
            try:
                if mapa_reorg:
                    _reorganizar_piezas_por_cara(carpeta_piezas, mapa_reorg)
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

            try:
                from cotas_dossier_registro import publicar_y_sincronizar_dossier

                publicar_y_sincronizar_dossier(carpeta_piezas)
            except Exception as err_dos:
                print(f"AVISO dossier sync post-reorg: {err_dos}")

        # Kits OTC: solo en tanque completo. En BOARD no aplica.
        if not solo and not es_board:
            try:
                from ensambles_independientes import (
                    ejecutar_ensambles_independientes,
                )

                ok_ens = ejecutar_ensambles_independientes(
                    inv_app,
                    ensamble,
                    plano,
                    carpeta_tanque,
                    incremental=incremental,
                )
                if not ok_ens:
                    print(
                        "AVISO: ensambles independientes terminó con errores "
                        "(las piezas pueden estar OK)."
                    )
            except Exception as err:
                print(f"AVISO: ensambles independientes no ejecutado: {err}")

        return ok
    finally:
        try:
            from cota_estilo import set_unidad_cota

            set_unidad_cota("in")
        except Exception:
            pass
        try:
            from creador_vistas import set_nombre_pieza_completo

            set_nombre_pieza_completo(False)
        except Exception:
            pass
        if inv_app is not None:
            _reactivar_machote(inv_app)
        pythoncom.CoUninitialize()
        if ok:
            if solo:
                print(
                    f"PROCESO COMPLETO: piezas de {solo} acotadas "
                    f"(PIEZAS_ACOTADAS\\{solo}\\)."
                )
            else:
                print(
                    "PROCESO COMPLETO: piezas acotadas exportadas "
                    "(+ ENSAMBLES_INDEPENDIENTES si hubo kits)."
                )


if __name__ == "__main__":
    ruta = _parse_ruta_seleccion(sys.argv[1:])
    solo = _parse_solo(sys.argv[1:])
    # BOARD / cobre: iLogic lanza sin --seleccion (sin picks TOP/SEGM).
    # Tanque: la regla pasa --seleccion <json> como siempre.
    if not ruta:
        print(
            "AVISO: sin --seleccion; se acotan piezas del ensamble abierto "
            "(sin mapa SEGM/TOP). Uso típico: BOARD / GIGA.",
            flush=True,
        )
    sys.exit(0 if ejecutar(ruta_seleccion=ruta or None, solo_cara=solo) else 1)
