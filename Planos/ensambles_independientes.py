"""
Subproceso Abigail: ensambles independientes (kits OTC / armados de taller).

Detecta ``.iam`` que NO son contenedores de cara (tanque/tapa/shell/SEGM/BASE),
crea vistas ortogonales por bbox (LARGO/ANCHO/ALTO), acota la silueta y
exporta a::

    Planos/JPG/<TANQUE>/ENSAMBLES_INDEPENDIENTES/<nombre_ensamble>/*.jpg

Sin clasificación iProperty (las piezas hijas ya se clasifican en PIEZAS_ACOTADAS).
"""

from __future__ import annotations

import os
import re
import shutil
import time
import traceback

from inventor_com import conectar_inventor

CARPETA_ENSAMBLES_INDEPENDIENTES = "ENSAMBLES_INDEPENDIENTES"
TIPO_DOCUMENTO_ENSAMBLE = 12291
MIN_HIJOS_ENSAMBLE = 2

# Nombres Vantran / genéricos que nunca son "kit independiente".
_EXCLUIR_NOMBRE_RE = re.compile(
    r"(SEGMENTO|TOP[_\s-]?COVER|TOP COVER|ASSEMBLY SEGMENTO|"
    r"CASCO|SHELL|BASE DE SOLERA|PLACA BASE|"
    r"MACHOTE|MODELO VANTRAN)",
    re.IGNORECASE,
)

# Sufijos de hoja (pipeline Abigail: LARGO/ANCHO/ALTO → LENGTH/WIDTH/HEIGHT).
_SUFIJOS_VISTA = ("LARGO", "ANCHO", "ALTO")


def _log(msg):
    print(msg)


def _nombre_doc(doc_or_occ):
    try:
        if hasattr(doc_or_occ, "FullFileName"):
            base = os.path.splitext(os.path.basename(doc_or_occ.FullFileName))[0]
            if base:
                return base
    except Exception:
        pass
    try:
        return str(doc_or_occ.DisplayName).replace(".iam", "").replace(".IAM", "")
    except Exception:
        pass
    try:
        return str(doc_or_occ.Name).split(":")[0]
    except Exception:
        return "ENSAMBLE"


def _es_ensamble_occ(occ):
    try:
        return int(occ.DefinitionDocumentType) == TIPO_DOCUMENTO_ENSAMBLE
    except Exception:
        try:
            return int(occ.Definition.Document.DocumentType) == TIPO_DOCUMENTO_ENSAMBLE
        except Exception:
            return False


def _contar_hijos(asm_doc):
    try:
        return int(asm_doc.ComponentDefinition.Occurrences.Count)
    except Exception:
        return 0


def _es_excluido_candidato(nombre):
    """
    True = no acotar como kit (pero sí recorrer hijos).

    Usa reglas OTC (46/47/48 caras) + heurística de nombre Vantran.
    """
    from generador_caras_tanque import (
        _es_contenedor_cara_otc,
        _es_shell_o_tanque_otc,
        _parse_codigo_otc,
        _familia_otc_rol,
    )

    nom = str(nombre or "")
    if _es_shell_o_tanque_otc(nom):
        return True
    if _es_contenedor_cara_otc(nom):
        return True
    info = _parse_codigo_otc(nom)
    if info and info["tipo"] == "A":
        rol = _familia_otc_rol(info["familia"])
        # 48-A01 ya cubierto; 47 cualquiera y 48 A02-A06 cubiertos.
        if rol == "47":
            return True
        if rol == "48" and info["numero"] in (1, 2, 3, 4, 5, 6):
            return True
    if _EXCLUIR_NOMBRE_RE.search(nom):
        return True
    return False


def _iter_subocurrencias(occ):
    try:
        subs = occ.SubOccurrences
        n = int(subs.Count)
        for i in range(1, n + 1):
            yield subs.Item(i)
        return
    except Exception:
        pass
    try:
        doc = occ.Definition.Document
        occs = doc.ComponentDefinition.Occurrences
        n = int(occs.Count)
        for i in range(1, n + 1):
            yield occs.Item(i)
    except Exception:
        return


def recolectar_ensambles_independientes(ensamble_raiz, exclusiones_extra=None):
    """
    Lista única de ``(asm_doc, nombre_base, qty_instancias, hijos)``.

    Recorre el árbol; los contenedores de cara se omiten como candidato
    pero se exploran por dentro (kits colgados bajo SEGM).

    ``exclusiones_extra``: nombres/rutas (upper) de TOP/SEGM/BASE resueltos
    por selección manual — refuerzan el filtro OTC.
    """
    resultados = {}  # ruta -> {doc, nombre, qty, hijos}
    excl = {str(x).upper() for x in (exclusiones_extra or []) if x}

    def _nombre_excluido(nombre, asm_doc=None, occ=None):
        candidatos = []
        for raw in (
            nombre,
            getattr(occ, "Name", None) if occ is not None else None,
        ):
            if not raw:
                continue
            candidatos.append(str(raw).upper().split(":")[0].strip())
        if asm_doc is not None:
            try:
                ff = str(asm_doc.FullFileName or "")
                if ff:
                    candidatos.append(ff.upper())
                    candidatos.append(
                        os.path.splitext(os.path.basename(ff))[0].upper()
                    )
            except Exception:
                pass
        for c in candidatos:
            if not c:
                continue
            if _es_excluido_candidato(c):
                return True
            if c in excl:
                return True
        return False

    def _visitar_occ(occ):
        try:
            if occ.Suppressed:
                return
        except Exception:
            return
        if not _es_ensamble_occ(occ):
            return

        try:
            asm_doc = occ.Definition.Document
        except Exception:
            return
        try:
            nombre = _nombre_doc(asm_doc)
        except Exception:
            nombre = str(getattr(occ, "Name", "ENSAMBLE")).split(":")[0]

        if _nombre_excluido(nombre, asm_doc=asm_doc, occ=occ):
            for hijo in _iter_subocurrencias(occ):
                _visitar_occ(hijo)
            return

        try:
            ruta = str(asm_doc.FullFileName or "")
        except Exception:
            ruta = ""
        clave = ruta.upper() if ruta else nombre.upper()

        if clave in resultados:
            resultados[clave]["qty"] += 1
        else:
            n_hijos = _contar_hijos(asm_doc)
            if n_hijos >= MIN_HIJOS_ENSAMBLE:
                resultados[clave] = {
                    "doc": asm_doc,
                    "nombre": nombre,
                    "qty": 1,
                    "hijos": n_hijos,
                }
            else:
                _log(
                    f"  Omitido (wrapper <{MIN_HIJOS_ENSAMBLE} hijos): "
                    f"{nombre} ({n_hijos})"
                )

        for hijo in _iter_subocurrencias(occ):
            _visitar_occ(hijo)

    try:
        occs = ensamble_raiz.ComponentDefinition.Occurrences
        total = int(occs.Count)
    except Exception as exc:
        _log(f"ERROR: no se pudieron leer ocurrencias raíz: {exc}")
        return []

    for i in range(1, total + 1):
        try:
            _visitar_occ(occs.Item(i))
        except Exception as exc:
            _log(f"AVISO: occ raíz #{i}: {exc}")

    lista = [
        (v["doc"], v["nombre"], v["qty"], v["hijos"])
        for v in resultados.values()
    ]
    lista.sort(key=lambda t: t[1].upper())
    _log(
        f"Ensambles independientes detectados: {len(lista)} "
        f"(únicos; instancias sumadas en qty)"
    )
    for _doc, nom, qty, hijos in lista:
        _log(f"  - {nom}  qty={qty}  hijos={hijos}")
    return lista


def _bbox_centro_ejes(asm_doc, tg):
    rb = asm_doc.ComponentDefinition.RangeBox
    minx, maxx = float(rb.MinPoint.X), float(rb.MaxPoint.X)
    miny, maxy = float(rb.MinPoint.Y), float(rb.MaxPoint.Y)
    minz, maxz = float(rb.MinPoint.Z), float(rb.MaxPoint.Z)
    cx = (minx + maxx) / 2.0
    cy = (miny + maxy) / 2.0
    cz = (minz + maxz) / 2.0
    dims = [
        (abs(maxx - minx), "X", tg.CreateVector(1, 0, 0)),
        (abs(maxy - miny), "Y", tg.CreateVector(0, 1, 0)),
        (abs(maxz - minz), "Z", tg.CreateVector(0, 0, 1)),
    ]
    dims.sort(key=lambda t: t[0])
    # corto, medio, largo
    return cx, cy, cz, dims[0], dims[1], dims[2]


def _camaras_bbox(asm_doc, tg, to):
    """
    Tres cámaras: LARGO (plano largo×medio), ANCHO (medio×corto),
    ALTO (largo×corto dimensionando el corto).
    """
    import creador_vistas

    cx, cy, cz, corto, medio, largo = _bbox_centro_ejes(asm_doc, tg)
    # (sufijo, eye_dir, up_hint)
    specs = [
        ("LARGO", corto[2], medio[2]),   # mira al corto → ve largo+medio
        ("ANCHO", largo[2], corto[2]),   # mira al largo → ve medio+corto
        ("ALTO", medio[2], largo[2]),    # mira al medio → ve largo+corto
    ]
    cams = []
    for sufijo, eye, up in specs:
        eye_v = tg.CreateVector(float(eye.X), float(eye.Y), float(eye.Z))
        up_v = tg.CreateVector(float(up.X), float(up.Y), float(up.Z))
        cam = creador_vistas.crear_camara(
            asm_doc, tg, to, cx, cy, cz, eye_v, up_v
        )
        cams.append((sufijo, cam))
    return cams


def _acotar_vista_ensamble(hoja, vista, tg, nombre_hoja, sufijo):
    """Cota el eje mayor (LARGO/ANCHO) o menor (ALTO) de la silueta."""
    import THK

    try:
        datos = THK._obtener_curvas_validas(vista)
    except Exception:
        datos = None
    if not datos:
        _log(f"  ⚠️ {nombre_hoja}: sin curvas HLR")
        return False

    minx, maxx, miny, maxy = THK._bbox_global(datos)
    w = maxx - minx
    h = maxy - miny
    if max(w, h) <= THK.EPS:
        return False

    if sufijo == "ALTO":
        orient = "V" if h <= w else "H"
    else:
        orient = "H" if w >= h else "V"

    return bool(
        THK._acotar_lado_bbox(
            hoja, vista, tg, datos, nombre_hoja, orient, sufijo, ratio_min=0.90
        )
    )


def _limpiar_carpeta_ensambles(carpeta, incremental=False):
    os.makedirs(carpeta, exist_ok=True)
    if incremental:
        return
    try:
        for nombre in os.listdir(carpeta):
            ruta = os.path.join(carpeta, nombre)
            if os.path.isdir(ruta):
                shutil.rmtree(ruta, ignore_errors=True)
            elif nombre.lower().endswith(".jpg"):
                try:
                    os.remove(ruta)
                except OSError:
                    pass
    except OSError:
        pass


def _reorganizar_jpgs_por_ensamble(carpeta_raiz):
    """Mueve JPG sueltos a ``<carpeta_raiz>/<ITEM>/``."""
    from nomenclatura_capturas import extraer_item_de_captura_pieza
    from generador_tanque_completo import _nombre_carpeta_pieza

    try:
        entradas = list(os.listdir(carpeta_raiz))
    except OSError:
        return 0
    movidos = 0
    for nombre in entradas:
        if not nombre.lower().endswith(".jpg"):
            continue
        if nombre.startswith("_tmp"):
            continue
        ruta = os.path.join(carpeta_raiz, nombre)
        if not os.path.isfile(ruta):
            continue
        item = extraer_item_de_captura_pieza(nombre)
        dest_dir = os.path.join(carpeta_raiz, _nombre_carpeta_pieza(item))
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, nombre)
        try:
            if os.path.abspath(ruta) != os.path.abspath(dest):
                if os.path.exists(dest):
                    os.remove(dest)
                shutil.move(ruta, dest)
                movidos += 1
        except OSError as err:
            _log(f"  AVISO mover {nombre}: {err}")
    return movidos


def _procesar_ensamble(
    inv_app, plano, asm_doc, nombre, tg, to, base_sheet
):
    """Crea 3 hojas + cotas. Devuelve set de nombres de hoja."""
    import creador_vistas

    nombres = set()
    try:
        cams = _camaras_bbox(asm_doc, tg, to)
    except Exception as exc:
        _log(f"  ERROR cámaras {nombre}: {exc}")
        return nombres

    for sufijo, cam in cams:
        nombre_hoja = creador_vistas.construir_nombre_hoja(
            plano, nombre, sufijo
        )
        try:
            hoja = creador_vistas._crear_hoja_vista(plano, base_sheet, nombre_hoja)
        except Exception as exc:
            _log(f"  ⚠️ no se pudo crear hoja {nombre_hoja}: {exc}")
            continue
        nombres.add(str(hoja.Name).split(":")[0])
        try:
            creador_vistas._limpiar_border_y_titleblock(hoja)
        except Exception:
            pass
        px, py, ancho_util, alto_util = creador_vistas._area_util_hoja(hoja)
        vista = None
        try:
            vista = creador_vistas._crear_vista_base(
                hoja, asm_doc, tg, to, px, py, cam, False
            )
        except Exception as exc:
            _log(f"  ⚠️ vista {nombre_hoja}: {exc}")
            try:
                hoja.Delete()
            except Exception:
                pass
            nombres.discard(str(hoja.Name).split(":")[0])
            continue
        if vista is None:
            continue
        try:
            creador_vistas.escalar_vista(
                plano, vista, tg, px, py, ancho_util, alto_util
            )
        except Exception:
            pass
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        time.sleep(0.35)
        ok = _acotar_vista_ensamble(hoja, vista, tg, nombre_hoja, sufijo)
        if not ok:
            _log(f"  ⚠️ sin cota en {nombre_hoja}")
    return nombres


def ejecutar_ensambles_independientes(
    inv_app,
    ensamble_doc,
    plano,
    carpeta_tanque,
    incremental=False,
):
    """
    Punto de entrada tras PIEZAS_ACOTADAS.

    Por defecto **desactivado** tras la corrida OTC 62201: el MVP de solo
    bbox (LENGTH/WIDTH/HEIGHT) no sirve como instructivo de armado.
    El flujo real irá en regla iLogic propia (3 vistas + cotas entre piezas).

    Activar el MVP bbox solo con ``ENSAMBLES_INDEPENDIENTES=1``.
    """
    flag = os.environ.get("ENSAMBLES_INDEPENDIENTES", "0").strip().lower()
    if flag in ("0", "false", "no", "off", ""):
        _log(
            "Ensambles independientes (bbox MVP): omitido. "
            "Para forzar: ENSAMBLES_INDEPENDIENTES=1. "
            "Instructivo real → regla iLogic aparte (pendiente)."
        )
        return True
    if flag not in ("1", "true", "yes", "on"):
        _log(
            f"Ensambles independientes: flag desconocido "
            f"'{flag}', se omite."
        )
        return True

    _log("=" * 62)
    _log(" SUBPROCESO: ENSAMBLES INDEPENDIENTES")
    _log("=" * 62)

    carpeta = os.path.join(carpeta_tanque, CARPETA_ENSAMBLES_INDEPENDIENTES)
    _limpiar_carpeta_ensambles(carpeta, incremental=incremental)

    try:
        lista = recolectar_ensambles_independientes(ensamble_doc)
    except Exception as exc:
        _log(f"ERROR recolectando ensambles: {exc}")
        _log(traceback.format_exc())
        return False

    if not lista:
        _log("Sin ensambles independientes (OK en Vantran / tanques sin kits).")
        os.makedirs(carpeta, exist_ok=True)
        return True

    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    try:
        base_sheet = plano.ActiveSheet
    except Exception:
        base_sheet = plano.Sheets.Item(1)

    from generador_vistas import (
        borrar_hojas_por_nombres,
        exportar_hojas_jpg,
        _nombre_hoja_machote,
        renombrar_hojas_finales,
    )

    nombre_machote = _nombre_hoja_machote(plano)
    job = None
    try:
        from nomenclatura_capturas import nombre_job_desde_ensamble

        job = nombre_job_desde_ensamble(ensamble_doc)
    except Exception:
        job = _nombre_doc(ensamble_doc)

    # Lotes pequeños: cada ensamble = 3 hojas.
    ok_total = True
    for asm_doc, nombre, qty, hijos in lista:
        _log(f"Acotando ensamble: {nombre} (qty={qty}, hijos={hijos})")
        try:
            nombres = _procesar_ensamble(
                inv_app, plano, asm_doc, nombre, tg, to, base_sheet
            )
        except Exception as exc:
            _log(f"  ERROR procesando {nombre}: {exc}")
            _log(traceback.format_exc())
            ok_total = False
            continue
        if not nombres:
            _log(f"  AVISO: {nombre} sin hojas exportables")
            continue
        try:
            renombrar_hojas_finales(plano, nombres_permitidos=nombres)
        except Exception:
            pass
        try:
            exportar_hojas_jpg(
                inv_app,
                plano,
                carpeta_salida=carpeta,
                nombres_permitidos=nombres,
                nombre_job=job,
            )
        except Exception as exc:
            _log(f"  ERROR export {nombre}: {exc}")
            ok_total = False
        try:
            borrar_hojas_por_nombres(
                plano, nombres, nombre_machote_protegido=nombre_machote
            )
        except Exception as exc:
            _log(f"  AVISO borrando hojas {nombre}: {exc}")
        time.sleep(0.4)

    movidos = _reorganizar_jpgs_por_ensamble(carpeta)
    _log(
        f"ENSAMBLES_INDEPENDIENTES listo: {len(lista)} tipos, "
        f"{movidos} JPG en subcarpetas → {carpeta}"
    )
    return ok_total


if __name__ == "__main__":
    # Diagnóstico rápido: solo lista candidatos del ensamble activo.
    pythoncom = __import__("pythoncom")
    pythoncom.CoInitialize()
    try:
        app = conectar_inventor()
        doc = app.ActiveDocument
        if int(doc.DocumentType) != TIPO_DOCUMENTO_ENSAMBLE:
            print("Activa el .iam del tanque y vuelve a ejecutar.")
        else:
            recolectar_ensambles_independientes(doc)
    finally:
        pythoncom.CoUninitialize()
