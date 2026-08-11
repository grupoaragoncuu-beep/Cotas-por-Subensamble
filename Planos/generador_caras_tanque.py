"""
Flujo de caras del tanque (COTAS_CARAS_TANQUE).

Igual que COTAS_ILOGIC_ABIGAIL: abre CMD, trabaja por API COM y exporta JPG.

Orientacion (igual criterio que Precesador STEP PQart 2.0/3.0):
  tapa / altura  -> +Y (arriba en la foto)
  cara principal -> +Z (FRONT)
  lateral        -> +X (RIGHT)
Si el ensamble llega chueco, se leen normales de caras planas y se
corrige la postura de la camara (el tanque se ve derecho en el JPG).

Cotas:
  (0,0) en la esquina inferior-izquierda del CUERPO de la cara.
  Cotas lineales H/V desde ese origen hasta cada accesorio del
  segmento/subensamble mapeado a esa pared (soldadura).
  Ver DEBER_SER_COTAS_CARAS.md en la raiz del proyecto.
"""

import math
import os
import re
import sys
import time
import traceback

import pythoncom
import win32com.client

from inventor_com import conectar_inventor
from cota_estilo import (
    aplicar_estilo_cota,
    aplicar_estilo_texto_cota,
    texto_cota_limpio,
)
from generador_vistas import (
    ALTO_EXPORTACION,
    ANCHO_EXPORTACION,
    CARPETA_EXPORTACION,
    _actualizar_inventor,
    _bbox_hoja_a_pixeles,
    _limpiar_nombre_archivo,
    _obtener_bbox_hoja,
)
from orientacion_pqart import marco_como_pqart

try:
    from PIL import Image
except ImportError:
    Image = None

# Margen al encuadrar JPG: debe incluir TODAS las cotas fuera de la vista.
MARGEN_RECORTE_CARAS = 0.18


TIPO_DOCUMENTO_DIBUJO = 12292
TIPO_DOCUMENTO_ENSAMBLE = 12291
PREFIJO_HOJA = "TANQUE_DATUM_"
PREFIJO_SKETCH_COTAS = "TANQUE_COTAS_"

# ViewOrientationTypeEnum / DrawingViewStyleEnum / DimensionTypeEnum
K_ARBITRARY = 10763
ESTILO_HLR = 32258
COTA_HORIZONTAL = 60162
COTA_VERTICAL = 60163

OFFSET_COTA = 2.6
PASO_NIVEL_COTA = 1.15
TAM_FLECHA_COTA = 0.23
# Una foto individual conserva un respiro alrededor de toda la cara, en vez
# de ampliar la vista hasta cortar el perímetro contrario.
FACTOR_ENCUADRE_FOTO = 0.88
MARGEN_DERECHO_FOTO_CM = 0.25
EPS = 0.0001
TOL_EXTREMO_RATIO = 0.012
DOMINANCIA_RECTA = 2.0
# Tolerancia física: se convierte a hoja con la escala de la vista. Nunca
# comparar 0.30 cm del modelo contra coordenadas de hoja directamente.
TOLERANCIA_COTA_CM = 0.03
TOL_COINCIDENCIA_HOJA = 0.001
PASO_COTA_LEGIBLE = 0.52
# Ignora basura tipo 0.19 cerca del origen (bridas / biseles).
MIN_DIST_ORIGEN_CM = 0.50
# Nunca descartar extremos por cantidad: cada pieza debe conservar Xmin/Xmax
# e Ymin/Ymax. El encuadre adaptativo se ocupa de reservar el espacio.
MAX_COTAS_POR_EJE = None
COTA_ALINEADA = 60161
# Muros del cuerpo: fraccion minima del tamaño de vista (evita lugs como "pared").
FRAC_MURO_VISTA = 0.32
# Expansion del cuerpo para aceptar lugs pegados a la cara.
FRAC_EXPAND_LUGS = 0.22
NOMBRES_LUG = ("LUG", "LIFTING", "OREJA", "CANCAMO", "CANCAM")
# Piezas grandes del casco: NO se acotan (solo accesorios del segmento).
EXCLUIR_CASCO = (
    "SHELL", "PLACA", "PLATE", "WALL", "CASCO", "BODY", "TANK WALL",
    "SEGMENTO", "SEGMENT", "COVER", "TAPA", "BASE DE SOLERA", "SOLERA",
    "HEADIRON", "COMPARTMENT", "INSPECTION_PLATE", "INSPECTION PLATE",
    "MARCO SOLERA", "MARCO DE SOLERA", "TOP_COVER", "TOP COVER",
    "BASE VANTRAN", "PLACA BASE", "FONDO", "AISC", "CUADRO BASE",
    "PLAQUITA", "CBOX", "FRAME", "STRUCTURAL",
)


def _es_nombre_lug(nombre):
    upper = str(nombre or "").upper()
    return any(tag in upper for tag in NOMBRES_LUG)

RUTA_LOG = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "error_log_caras.txt",
)


def log(mensaje):
    print(mensaje)
    try:
        with open(RUTA_LOG, "a", encoding="utf-8") as archivo:
            archivo.write(f"{mensaje}\n")
    except OSError:
        pass


def _como_plano(documento):
    try:
        return win32com.client.CastTo(documento, "DrawingDocument")
    except Exception:
        return documento


def _como_ensamble(documento):
    try:
        return win32com.client.CastTo(documento, "AssemblyDocument")
    except Exception:
        return documento


def _obtener_plano_activo(inv_app):
    documento = inv_app.ActiveDocument
    if documento is not None and documento.DocumentType == TIPO_DOCUMENTO_DIBUJO:
        return _como_plano(documento)

    # El operario puede tener abierta una pieza para revisarla (por ejemplo un
    # flange circular). Localizar el machote abierto en vez de exigir foco.
    candidatas = []
    for doc in inv_app.Documents:
        try:
            if doc.DocumentType != TIPO_DOCUMENTO_DIBUJO:
                continue
            candidatas.append(_como_plano(doc))
        except Exception:
            continue
    for plano in candidatas:
        if "MACHOTE" in str(plano.DisplayName).upper():
            try:
                plano.Activate()
            except Exception:
                pass
            return plano
    if candidatas:
        try:
            candidatas[0].Activate()
        except Exception:
            pass
        return candidatas[0]
    return None


def _obtener_ensamble_principal(inv_app):
    mejor = None
    mayor = -1

    for documento in inv_app.Documents:
        try:
            if documento.DocumentType != TIPO_DOCUMENTO_ENSAMBLE:
                continue
            ensamble = _como_ensamble(documento)
            cantidad = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences.Count
            log(f"  Ensamble abierto: {documento.DisplayName} ({cantidad} piezas)")
            if cantidad > mayor:
                mayor = cantidad
                mejor = ensamble
        except Exception as error:
            log(f"  AVISO al inspeccionar documento: {error}")

    return mejor


def _piezas_desde_occurrence(occ):
    """Nombres leaf (+ display doc) colgando de una ocurrencia del tanque."""
    piezas = set()
    try:
        base = str(occ.Name).split(":")[0].strip().upper()
        if base:
            piezas.add(base)
    except Exception:
        pass

    try:
        hojas = occ.Definition.Occurrences.AllLeafOccurrences
        for i in range(1, hojas.Count + 1):
            try:
                leaf = hojas.Item(i)
                if leaf.Suppressed:
                    continue
                nb = str(leaf.Name).split(":")[0].strip().upper()
                if nb:
                    piezas.add(nb)
                try:
                    doc_pie = str(leaf.Definition.Document.DisplayName).upper()
                    doc_pie = doc_pie.replace(".IPT", "").replace(".IAM", "").strip()
                    if doc_pie:
                        piezas.add(doc_pie)
                except Exception:
                    pass
            except Exception:
                continue
    except Exception:
        try:
            doc_pie = str(occ.Definition.Document.DisplayName).upper()
            doc_pie = doc_pie.replace(".IPT", "").replace(".IAM", "").strip()
            if doc_pie:
                piezas.add(doc_pie)
        except Exception:
            pass
    return piezas


def _centroide_occurrence(occ):
    try:
        box = occ.RangeBox
        return (
            (float(box.MinPoint.X) + float(box.MaxPoint.X)) * 0.5,
            (float(box.MinPoint.Y) + float(box.MaxPoint.Y)) * 0.5,
            (float(box.MinPoint.Z) + float(box.MaxPoint.Z)) * 0.5,
        )
    except Exception:
        return None


def _es_nombre_contenedor_cara(nombre):
    u = str(nombre or "").upper()
    if "SEGMENTO" in u or "SEGMENT" in u:
        return True
    if re.search(r"SEG[\.\s_-]*\d", u):
        return True
    return False


def _es_excluir_contenedor_global(nombre):
    u = str(nombre or "").upper()
    tags = (
        "TOP_COVER", "TOP COVER", "TAPA SUPERIOR", "TAPA ", " COVER",
        "COVER_", "_COVER", "BASE DE SOLERA",
        "BASE VANTRAN", "BASE DE SOLERAS", "PARKING STAND",
        "JACKING PADS ASSEMBLY", "HEADIRON", "COMPARTMENT", "MACHOTE",
    )
    if any(t in u for t in tags):
        return True
    # "TAPA" / "COVER" como palabra completa del nombre.
    tokens = re.split(r"[\s_\-:]+", u)
    return "TAPA" in tokens or "COVER" in tokens


def _occurrence_raiz(occ):
    """Sube hasta la ocurrencia de primer nivel del tanque."""
    actual = occ
    for _ in range(16):
        try:
            padre = actual.ParentOccurrence
        except Exception:
            break
        if padre is None:
            break
        actual = padre
    return actual


def _es_rama_de_tapa(occ, cover, bbox):
    """
    Descarta subensambles superiores (top cover) aunque no digan COVER.
    Se basan en que su ocurrencia raíz vive en el tercio superior del tanque.
    """
    if cover is None or bbox is None:
        return False
    raiz = _occurrence_raiz(occ)
    centro = _centroide_occurrence(raiz)
    if centro is None:
        return False
    min_up, max_up = _rango_proyectado_bbox(bbox, cover)
    alto = max(EPS, max_up - min_up)
    proy = _dot(centro, cover)
    return proy >= (min_up + 0.70 * alto)


def _lista_ocurrencias(coleccion):
    """Convierte una colección COM de ocurrencias en lista tolerante a fallos."""
    ocurrencias = []
    try:
        for i in range(1, coleccion.Count + 1):
            ocurrencias.append(coleccion.Item(i))
    except Exception:
        pass
    return ocurrencias


def _es_ensamble_occurrence(occ):
    try:
        return int(occ.DefinitionDocumentType) == TIPO_DOCUMENTO_ENSAMBLE
    except Exception:
        try:
            return int(occ.Definition.Document.DocumentType) == TIPO_DOCUMENTO_ENSAMBLE
        except Exception:
            return False


def _cantidad_hojas_occurrence(occ):
    """Número de piezas hoja bajo una ocurrencia, o cero si no es un IAM."""
    try:
        return int(occ.Definition.Occurrences.AllLeafOccurrences.Count)
    except Exception:
        return 0


def _subocurrencias_contextuales(occ):
    """
    Hijos inmediatos dentro del contexto del tanque raíz.

    SubOccurrences conserva ParentOccurrence y RangeBox globales; usar la
    colección del Definition perdería esa ruta al trabajar con IAM anidados.
    """
    try:
        return _lista_ocurrencias(occ.SubOccurrences)
    except Exception:
        return []


def _bbox_desde_rangebox(caja):
    try:
        xmin, ymin, zmin = (
            float(caja.MinPoint.X),
            float(caja.MinPoint.Y),
            float(caja.MinPoint.Z),
        )
        xmax, ymax, zmax = (
            float(caja.MaxPoint.X),
            float(caja.MaxPoint.Y),
            float(caja.MaxPoint.Z),
        )
    except Exception:
        return None
    return {
        "min": (xmin, ymin, zmin),
        "max": (xmax, ymax, zmax),
        "cx": (xmin + xmax) * 0.5,
        "cy": (ymin + ymax) * 0.5,
        "cz": (zmin + zmax) * 0.5,
        "dx": abs(xmax - xmin),
        "dy": abs(ymax - ymin),
        "dz": abs(zmax - zmin),
        "corners": [
            (xmin, ymin, zmin),
            (xmax, ymin, zmin),
            (xmin, ymax, zmin),
            (xmax, ymax, zmin),
            (xmin, ymin, zmax),
            (xmax, ymin, zmax),
            (xmin, ymax, zmax),
            (xmax, ymax, zmax),
        ],
    }


def _bbox_occurrence(occ):
    try:
        return _bbox_desde_rangebox(occ.RangeBox)
    except Exception:
        return None


def _rango_proyectado_bbox(bbox, vector):
    """Extremos de una caja 3D sobre una dirección de cámara unitaria."""
    try:
        valores = [_dot(punto, vector) for punto in bbox["corners"]]
        return min(valores), max(valores)
    except Exception:
        return 0.0, 0.0


def _mapear_ocurrencias_por_direccion(ocurrencias, face, right, bbox):
    """Mapea ocurrencias a paredes sin escribir en el log de producción."""
    direcciones = {
        "FRONT": face,
        "BACK": _scale(face, -1.0),
        "RIGHT": right,
        "LEFT": _scale(right, -1.0),
    }
    centro = (bbox["cx"], bbox["cy"], bbox["cz"])
    mapeo = {}
    usados = set()
    for cara, normal in direcciones.items():
        mejor = None
        for indice, occ in enumerate(ocurrencias):
            if indice in usados:
                continue
            punto = _centroide_occurrence(occ)
            if punto is None:
                continue
            vec = (
                punto[0] - centro[0],
                punto[1] - centro[1],
                punto[2] - centro[2],
            )
            score = _dot(vec, normal)
            if mejor is None or score > mejor["score"]:
                mejor = {"indice": indice, "occurrence": occ, "score": score}
        if mejor is not None:
            usados.add(mejor["indice"])
            mapeo[cara] = mejor
    return mapeo


def _calificar_contenedor_paredes(occ, face, right, cover):
    """
    Valida un IAM estructural: debe tener cuatro hijos IAM ubicados en las
    cuatro paredes laterales y que recorran una porción significativa de alto.
    """
    bbox = _bbox_occurrence(occ)
    if bbox is None:
        return None

    hijos = [
        hijo
        for hijo in _subocurrencias_contextuales(occ)
        if not getattr(hijo, "Suppressed", False)
        and _es_ensamble_occurrence(hijo)
        and _cantidad_hojas_occurrence(hijo) >= 2
        and _centroide_occurrence(hijo) is not None
    ]
    if len(hijos) < 4:
        return None

    mapeo = _mapear_ocurrencias_por_direccion(hijos, face, right, bbox)
    if len(mapeo) != 4:
        return None

    direcciones = {
        "FRONT": face,
        "BACK": _scale(face, -1.0),
        "RIGHT": right,
        "LEFT": _scale(right, -1.0),
    }
    min_cover, max_cover = _rango_proyectado_bbox(bbox, cover)
    alto = max(EPS, max_cover - min_cover)
    score_total = 0.0
    for cara, datos in mapeo.items():
        normal = direcciones[cara]
        min_n, max_n = _rango_proyectado_bbox(bbox, normal)
        mitad = max(EPS, (max_n - min_n) * 0.5)
        # score ya se calculó contra el centro de la caja del candidato.
        # Restar aquí la coordenada absoluta fallaría para tanques trasladados.
        exterior = datos["score"] / mitad
        hijo_bbox = _bbox_occurrence(datos["occurrence"])
        if hijo_bbox is None:
            return None
        min_h, max_h = _rango_proyectado_bbox(hijo_bbox, cover)
        proporcion_alto = (max_h - min_h) / alto
        # Una pared real queda cerca de la envolvente exterior y cubre una
        # fracción relevante de la altura; accesorios sueltos no cumplen ambas.
        if exterior < 0.42 or proporcion_alto < 0.35:
            return None
        score_total += exterior + proporcion_alto

    return {
        "contenedor": occ,
        "bbox": bbox,
        "mapeo": mapeo,
        "ocurrencias": [mapeo[cara]["occurrence"] for cara in ("FRONT", "BACK", "RIGHT", "LEFT")],
        "score": score_total,
    }


def _iterar_contenedores_ensamble(ensamble, profundidad_maxima=2):
    """IAM contextuales bajo el tanque, priorizando los más cercanos a raíz."""
    pendientes = [
        (occ, 1)
        for occ in _lista_ocurrencias(ensamble.ComponentDefinition.Occurrences)
        if _es_ensamble_occurrence(occ)
    ]
    while pendientes:
        occ, profundidad = pendientes.pop(0)
        yield occ, profundidad
        if profundidad >= profundidad_maxima:
            continue
        for hijo in _subocurrencias_contextuales(occ):
            if _es_ensamble_occurrence(hijo):
                pendientes.append((hijo, profundidad + 1))


def _prevalidar_contenedor_paredes_rapido(ensamble):
    """
    Prevalidación sin recorrer caras B-Rep del modelo.

    El orquestador solo necesita saber si vale la pena ejecutar el flujo de
    caras; la validación geométrica completa se hace justo antes de crear las
    hojas, ya con el marco PQart calculado una sola vez.
    """
    mejor = None
    for candidato, profundidad in _iterar_contenedores_ensamble(ensamble):
        if _es_excluir_contenedor_global(str(candidato.Name)):
            continue
        hijos = [
            hijo
            for hijo in _subocurrencias_contextuales(candidato)
            if not getattr(hijo, "Suppressed", False)
            and _es_ensamble_occurrence(hijo)
            and _cantidad_hojas_occurrence(hijo) >= 2
        ]
        # Cuatro IAM complejos inmediatos es la huella estructural mínima de
        # un cuerpo de tanque; no se acepta cualquier colección de piezas.
        if len(hijos) != 4:
            continue
        hojas = sum(_cantidad_hojas_occurrence(hijo) for hijo in hijos)
        clave = (hojas, -profundidad)
        if mejor is None or clave > (mejor["hojas"], -mejor["profundidad"]):
            mejor = {
                "contenedor": candidato,
                "ocurrencias": hijos,
                "hojas": hojas,
                "profundidad": profundidad,
            }
    if mejor is None:
        return {
            "valido": False,
            "motivo": "no se encontró un IAM estructural con cuatro paredes",
            "ocurrencias": [],
        }
    return {
        "valido": True,
        "origen": "anidado_prevalidado",
        "contenedor": mejor["contenedor"],
        "ocurrencias": mejor["ocurrencias"],
    }


def _resolver_contenedor_paredes(
    ensamble,
    face=None,
    right=None,
    cover=None,
    bbox=None,
    registrar=True,
):
    """
    Localiza el alcance real de las cuatro paredes.

    Vantran conserva segmentos nombrados directamente bajo raíz. Para OTC y
    familias similares se desciende hasta un IAM estructural que demuestre por
    geometría contener las cuatro paredes, nunca solo por tener cuatro hijos.
    """
    directas = _lista_ocurrencias(ensamble.ComponentDefinition.Occurrences)
    nombradas = [
        occ
        for occ in directas
        if not getattr(occ, "Suppressed", False)
        and _es_ensamble_occurrence(occ)
        and _es_nombre_contenedor_cara(str(occ.Name))
    ]
    if len(nombradas) >= 4:
        if bbox is None:
            bbox = _bbox_ensamble(ensamble)
        if registrar:
            log(
                "  Alcance de paredes: raíz con "
                f"{len(nombradas)} segmentos nombrados."
            )
        return {
            "valido": True,
            "origen": "raiz_nombrada",
            "contenedor": None,
            "bbox_caras": bbox,
            "ocurrencias": nombradas,
            "face": face,
            "right": right,
            "cover": cover,
        }

    # El orquestador usa esta vía rápida antes de lanzar el flujo. Evita
    # recorrer miles de caras B-Rep dos veces: _crear_caras hará enseguida la
    # prueba geométrica estricta con el mismo marco que se usará en los JPG.
    if face is None or right is None or cover is None:
        return _prevalidar_contenedor_paredes_rapido(ensamble)

    if bbox is None:
        bbox = _bbox_ensamble(ensamble)

    mejor = None
    for candidato, profundidad in _iterar_contenedores_ensamble(ensamble):
        if _es_excluir_contenedor_global(str(candidato.Name)):
            continue
        calificado = _calificar_contenedor_paredes(candidato, face, right, cover)
        if calificado is None:
            continue
        calificado["profundidad"] = profundidad
        # A igualdad geométrica se prefiere el contenedor más próximo a raíz:
        # es el cuerpo completo sin cover, no un detalle dentro de una pared.
        clave = (calificado["score"], -profundidad)
        if mejor is None or clave > (mejor["score"], -mejor["profundidad"]):
            mejor = calificado

    if mejor is None:
        if registrar:
            log("  AVISO: no se encontró un contenedor geométrico de cuatro paredes.")
        return {
            "valido": False,
            "motivo": "no hay cuatro paredes laterales verificables",
            "ocurrencias": [],
            "bbox_caras": bbox,
            "face": face,
            "right": right,
            "cover": cover,
        }

    if registrar:
        hijos = ", ".join(str(occ.Name) for occ in mejor["ocurrencias"])
        log(
            "  Alcance de paredes: "
            f"{mejor['contenedor'].Name} (anidado, {hijos})"
        )
    return {
        "valido": True,
        "origen": "anidado",
        "contenedor": mejor["contenedor"],
        "bbox_caras": mejor["bbox"],
        "ocurrencias": mejor["ocurrencias"],
        "face": face,
        "right": right,
        "cover": cover,
    }


def _piezas_de_ensamble_segmento(ensamble):
    """Nombres hoja (+ DisplayName) de un ensamble Segmento* abierto aparte."""
    piezas = set()
    try:
        hojas = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
        for i in range(1, hojas.Count + 1):
            try:
                occ = hojas.Item(i)
                if occ.Suppressed:
                    continue
                base = str(occ.Name).split(":")[0].strip().upper()
                if base:
                    piezas.add(base)
                try:
                    doc_pie = str(occ.Definition.Document.DisplayName).upper()
                    doc_pie = doc_pie.replace(".IPT", "").replace(".IAM", "").strip()
                    if doc_pie:
                        piezas.add(doc_pie)
                except Exception:
                    pass
            except Exception:
                continue
    except Exception:
        pass
    return piezas


def _listar_segmentos_abiertos(inv_app):
    """Refuerzo opcional: Segmento*.iam abiertos aparte del tanque."""
    segmentos = []
    for documento in inv_app.Documents:
        try:
            if documento.DocumentType != TIPO_DOCUMENTO_ENSAMBLE:
                continue
            nombre = str(documento.DisplayName)
            up = nombre.upper()
            if "SEGMENTO" not in up and "SEGMENT" not in up:
                continue
            ensamble = _como_ensamble(documento)
            piezas = _piezas_de_ensamble_segmento(ensamble)
            if not piezas:
                continue
            segmentos.append({"nombre": nombre, "piezas": piezas, "centroide": None})
            log(f"  Segmento abierto (refuerzo): {nombre} ({len(piezas)} nombres)")
        except Exception as error:
            log(f"  AVISO catalogo segmento abierto: {error}")
    return segmentos


def _detectar_segmentos_en_ensamble(ensamble, inv_app=None, alcance=None):
    """
    Contenedores de cara dentro del alcance estructural resuelto.

    El alcance puede ser la raíz (Vantran) o hijos contextuales de un IAM
    estructural anidado (OTC). En ambos casos las ocurrencias conservan ruta
    y posición en el tanque completo.
    """
    nombrados = []
    geo = []
    try:
        if alcance is not None and alcance.get("ocurrencias"):
            ocurrencias = list(alcance["ocurrencias"])
        else:
            ocurrencias = _lista_ocurrencias(
                ensamble.ComponentDefinition.Occurrences
            )
        total = len(ocurrencias)
    except Exception as error:
        log(f"  ERROR leyendo ocurrencias del tanque: {error}")
        return []

    for occ in ocurrencias:
        try:
            if occ.Suppressed:
                continue
            nombre = str(occ.Name)
            if _es_excluir_contenedor_global(nombre):
                continue
            centro = _centroide_occurrence(occ)
            if centro is None:
                continue
            piezas = _piezas_desde_occurrence(occ)
            try:
                doc_segmento = _como_ensamble(occ.Definition.Document)
                es_segmento_ensamble = (
                    int(doc_segmento.DocumentType) == TIPO_DOCUMENTO_ENSAMBLE
                )
            except Exception:
                doc_segmento = None
                es_segmento_ensamble = False
            entry = {
                "nombre": nombre.split(":")[0].strip(),
                "piezas": piezas,
                "centroide": centro,
                # Fuente de verdad para fotos/cotas: el IAM propio del
                # segmento, localizado desde el árbol del tanque principal.
                "occurrence": occ,
                "ensamble_segmento": doc_segmento if es_segmento_ensamble else None,
                "ruta": _ruta_occurrence(occ),
            }
            if _es_nombre_contenedor_cara(nombre):
                nombrados.append(entry)
            elif len(piezas) >= 2:
                try:
                    es_asm = int(occ.DefinitionDocumentType) == TIPO_DOCUMENTO_ENSAMBLE
                except Exception:
                    es_asm = len(piezas) > 3
                if es_asm or len(piezas) >= 4:
                    geo.append(entry)
        except Exception:
            continue

    if len(nombrados) >= 2:
        segmentos = nombrados
        log(f"  Segmentos por nombre en arbol: {len(segmentos)}")
    else:
        segmentos = nombrados + geo
        log(
            f"  Contenedores de cara en arbol: nombrados={len(nombrados)} "
            f"+ geo={len(geo)} = {len(segmentos)}"
        )

    if inv_app is not None:
        for ab in _listar_segmentos_abiertos(inv_app):
            fused = False
            abu = ab["nombre"].upper()
            for seg in segmentos:
                su = seg["nombre"].upper()
                if su in abu or abu in su:
                    seg["piezas"] |= ab["piezas"]
                    fused = True
                    break
            if not fused and ab["piezas"]:
                segmentos.append(ab)

    for seg in segmentos:
        ruta = seg.get("ruta", "")
        log(
            f"    - {seg['nombre']}: {len(seg['piezas'])} piezas"
            + (f" | ruta={ruta}" if ruta else "")
        )
    return segmentos


def _mapear_segmentos_a_caras(segmentos, face, right, bbox):
    """Asigna cada contenedor a FRONT/BACK/LEFT/RIGHT por centroide."""
    if not segmentos:
        return {}

    dirs = {
        "FRONT": face,
        "BACK": _scale(face, -1.0),
        "RIGHT": right,
        "LEFT": _scale(right, -1.0),
    }
    center = (bbox["cx"], bbox["cy"], bbox["cz"])
    mapping = {}
    usados = set()
    con_centro = [s for s in segmentos if s.get("centroide")]

    for cara, normal in dirs.items():
        mejor_i = None
        mejor_score = None
        for i, seg in enumerate(con_centro):
            if i in usados:
                continue
            c = seg["centroide"]
            vec = (c[0] - center[0], c[1] - center[1], c[2] - center[2])
            score = _dot(vec, normal)
            if mejor_score is None or score > mejor_score:
                mejor_score = score
                mejor_i = i
        if mejor_i is not None:
            usados.add(mejor_i)
            seg = con_centro[mejor_i]
            mapping[cara] = seg
            log(
                f"  {cara} <- {seg['nombre']} "
                f"(outward={mejor_score:.2f}, piezas={len(seg['piezas'])})"
            )

    if len(mapping) < 4:
        log(f"  AVISO: solo {len(mapping)}/4 caras con contenedor mapeado")
    return mapping


def _ruta_es_descendiente(ruta, ancestro):
    ruta = str(ruta or "")
    ancestro = str(ancestro or "")
    return bool(ancestro) and (ruta == ancestro or ruta.startswith(ancestro + "|"))


def _ruta_tiene_contenedor_excluido(ruta):
    return any(
        _es_excluir_contenedor_global(parte)
        for parte in str(ruta or "").split("|")
    )


def _ocurrencia_es_lamina_horizontal(occ, face, right, cover, bbox):
    """
    Descarta tapa/fondo extendidos. Una solera o bracket pequeño no cumple las
    dos extensiones horizontales, por lo que sigue siendo candidato a cota.
    """
    caja = _bbox_occurrence(occ)
    if caja is None:
        return False
    min_up, max_up = _rango_proyectado_bbox(caja, cover)
    min_face, max_face = _rango_proyectado_bbox(caja, face)
    min_right, max_right = _rango_proyectado_bbox(caja, right)
    pmin_up, pmax_up = _rango_proyectado_bbox(bbox, cover)
    pmin_face, pmax_face = _rango_proyectado_bbox(bbox, face)
    pmin_right, pmax_right = _rango_proyectado_bbox(bbox, right)
    alto_rel = (max_up - min_up) / max(EPS, pmax_up - pmin_up)
    face_rel = (max_face - min_face) / max(EPS, pmax_face - pmin_face)
    right_rel = (max_right - min_right) / max(EPS, pmax_right - pmin_right)
    return alto_rel < 0.10 and face_rel > 0.55 and right_rel > 0.55


def _ocurrencia_llega_a_pared(occ, cara, face, right, bbox):
    """Exige que una pieza externa realmente esté cerca de la pared elegida."""
    caja = _bbox_occurrence(occ)
    if caja is None:
        return False
    normal = {
        "FRONT": face,
        "BACK": _scale(face, -1.0),
        "RIGHT": right,
        "LEFT": _scale(right, -1.0),
    }.get(cara)
    if normal is None:
        return False
    min_pieza, max_pieza = _rango_proyectado_bbox(caja, normal)
    min_tanque, max_tanque = _rango_proyectado_bbox(bbox, normal)
    separacion = max(0.0, min_tanque - max_pieza)
    profundidad = max(EPS, max_tanque - min_tanque)
    # Admite accesorios que sobresalen del casco y piezas soldadas ligeramente
    # hacia dentro, sin traer componentes situados en el centro del tanque.
    return separacion <= profundidad * FRAC_EXPAND_LUGS


def _es_hardware_puro(nombre):
    """
    Tornillería/fijación genérica: no aporta cota de piso y satura HLR.
    Conserva accesorios como GUN STUD / parked studs de fabricación.
    """
    upper = str(nombre or "").upper()
    if not upper:
        return False
    if "GUN STUD" in upper or "GUNSTUD" in upper:
        return False
    if upper.startswith("HW-") or "|HW-" in upper:
        return True
    tags = (
        "WASHER", "HEX NUT", "HEX BOLT", "LOCK WASHER", "FLAT WASHER",
        "MACHINE SCREW", "CAP SCREW",
    )
    return any(tag in upper for tag in tags)


def _ocurrencias_raiz_en_cara(
    ensamble,
    ocurrencia_segmento,
    cara,
    face,
    right,
    bbox,
    rutas_segmentos=None,
    cover=None,
):
    """
    Accesorios fuera del IAM de pared, incluidos los anidados bajo el
    contenedor estructural y los colgados de la raíz del tanque.

    El filtrado se hace sobre hojas con ruta completa: así una pared OTC
    anidada puede aislarse sin mezclar las otras tres, y un CBOX raíz conserva
    su posición real.
    """
    extras = []
    try:
        ruta_segmento = _ruta_occurrence(ocurrencia_segmento)
        rutas_segmentos = set(rutas_segmentos or [])
        rutas_segmentos.add(ruta_segmento)
        hojas = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
        for i in range(1, hojas.Count + 1):
            occ = hojas.Item(i)
            try:
                if occ.Suppressed:
                    continue
                ruta = _ruta_occurrence(occ)
                if _ruta_es_descendiente(ruta, ruta_segmento):
                    continue
                if any(
                    _ruta_es_descendiente(ruta, ruta_muro)
                    for ruta_muro in rutas_segmentos
                ):
                    continue
                if _ruta_tiene_contenedor_excluido(ruta):
                    continue
                if _es_hardware_puro(ruta) or _es_hardware_puro(occ.Name):
                    continue
                ruta_u = ruta.upper()
                if "FLAT-PATTERN" in ruta_u or "FLAT_PATTERN" in ruta_u:
                    continue
                # Nunca meter top cover / tapa en fotos de pared.
                if cover is not None and _es_rama_de_tapa(occ, cover, bbox):
                    continue
                centro = _centroide_occurrence(occ)
                if centro is None:
                    continue
                if _cara_fisica_de_punto(centro, face, right, bbox) != cara:
                    continue
                if not _ocurrencia_llega_a_pared(occ, cara, face, right, bbox):
                    continue
                if cover is not None and _ocurrencia_es_lamina_horizontal(
                    occ, face, right, cover, bbox
                ):
                    continue
                extras.append(occ)
            except Exception:
                continue
    except Exception as error:
        log(f"    AVISO buscando accesorios raíz de {cara}: {error}")
    return extras


def _aplicar_visibilidad_vista_cara(
    vista, ensamble, ocurrencia_segmento, ocurrencias_extra
):
    """
    Aísla hojas del segmento y accesorios externos mediante rutas completas.

    Se vuelve al aislamiento hoja-a-hoja: la variante por rama dejó visible el
    top cover al activar cadenas padre. Calidad > velocidad aquí.
    """
    ruta_segmento = _ruta_occurrence(ocurrencia_segmento)
    extras = {_ruta_occurrence(occ) for occ in ocurrencias_extra}
    visibles = 0
    fallos = 0
    try:
        hojas = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
        for i in range(1, hojas.Count + 1):
            occ = hojas.Item(i)
            try:
                ruta = _ruta_occurrence(occ)
                visible = (
                    _ruta_es_descendiente(ruta, ruta_segmento) or ruta in extras
                )
                vista.SetVisibility(occ, visible)
                if visible:
                    visibles += 1
            except Exception:
                fallos += 1
    except Exception as error:
        log(f"    AVISO aplicando visibilidad de cara: {error}")
    if fallos:
        log(f"    AVISO visibilidad: {fallos} ocurrencias no se pudieron ocultar.")
    return visibles


def _nombre_base_pieza(nombre):
    return str(nombre or "").split(":")[0].strip().upper()


def _es_accesorio_por_rol(nombre):
    """
    Acepta toda pieza que no sea casco/estructura.
    Es intencionalmente independiente del nombre de cliente: OTC/PTT pueden
    traer códigos sin palabras como FLANGE o LUG.
    """
    base = _nombre_base_pieza(nombre)
    if not base or base == "SIN_COMPONENTE":
        return False
    return not any(excluir in base for excluir in EXCLUIR_CASCO)


def _pieza_en_catalogo(nombre, catalogo):
    """True si el componente de la vista pertenece al catalogo del segmento."""
    base = _nombre_base_pieza(nombre)
    if not base or base == "SIN_COMPONENTE":
        return False

    if not _es_accesorio_por_rol(base):
        return False

    if not catalogo:
        claves = (
            "LUG", "PAD", "FLANGE", "BRACKET", "NOZZLE", "FITTING",
            "PORT", "BOSS", "CLIP", "CLAMP", "SUPPORT", "MANWAY",
            "JACK", "LIFT", "OREJA", "BOCA", "ACCES", "GROUND", "TIERRA",
            "SWITCH", "GAUGE", "NIPPLE", "PARKING", "PATCH", "PIPE",
            "BUSHING", "VALVE", "HANDHOLE", "DRAIN",
        )
        return any(k in base for k in claves)

    for pieza in catalogo:
        if not pieza:
            continue
        if pieza == base or pieza in base or base in pieza:
            if any(ex in pieza for ex in EXCLUIR_CASCO) and not _es_nombre_lug(base):
                return False
            return True
    return False


def _nombre_en_catalogo(nombre, catalogo):
    """Pertenencia al contenedor sin aplicar la exclusión casco/accesorio."""
    base = _nombre_base_pieza(nombre)
    if not base:
        return False
    for pieza in catalogo:
        if not pieza:
            continue
        if pieza == base or pieza in base or base in pieza:
            return True
    return False


def _es_accesorio_segmento(nombre, catalogo):
    """Compat: solo piezas del segmento activo."""
    return _pieza_en_catalogo(nombre, catalogo)


def _punto3d_curva(curva):
    """Punto 3D del modelo asociado a una DrawingCurve (para profundidad)."""
    try:
        mg = curva.ModelGeometry
    except Exception:
        return None

    try:
        p = mg.Point
        return (float(p.X), float(p.Y), float(p.Z))
    except Exception:
        pass

    try:
        p1 = mg.StartVertex.Point
        p2 = mg.EndVertex.Point
        return (
            (float(p1.X) + float(p2.X)) * 0.5,
            (float(p1.Y) + float(p2.Y)) * 0.5,
            (float(p1.Z) + float(p2.Z)) * 0.5,
        )
    except Exception:
        pass

    try:
        occ = mg.ContainingOccurrence
        box = occ.RangeBox
        return (
            (float(box.MinPoint.X) + float(box.MaxPoint.X)) * 0.5,
            (float(box.MinPoint.Y) + float(box.MaxPoint.Y)) * 0.5,
            (float(box.MinPoint.Z) + float(box.MaxPoint.Z)) * 0.5,
        )
    except Exception:
        return None


def _cara_fisica_de_punto(punto, face, right, bbox):
    """
    Pared física más cercana a un punto 3D del accesorio.
    No depende de que el cliente use nombres como "Segmento 1".
    """
    vec = (
        float(punto[0]) - float(bbox["cx"]),
        float(punto[1]) - float(bbox["cy"]),
        float(punto[2]) - float(bbox["cz"]),
    )
    hacia_face = _dot(vec, face)
    hacia_right = _dot(vec, right)

    if abs(hacia_face) >= abs(hacia_right):
        return "FRONT" if hacia_face >= 0.0 else "BACK"
    return "RIGHT" if hacia_right >= 0.0 else "LEFT"


def _datos_son_de_cara_fisica(datos, cara, face, right, bbox):
    """
    Una pieza se acota solo en la pared donde vive su centroide 3D.
    Esto evita que nombres repetidos o catálogos de otros segmentos se mezclen.
    """
    votos = {}
    for dato in datos:
        punto = _punto3d_curva(dato["curve"])
        if punto is None:
            continue
        encontrada = _cara_fisica_de_punto(punto, face, right, bbox)
        votos[encontrada] = votos.get(encontrada, 0) + 1

    if not votos:
        return False
    ganadora = max(votos, key=votos.get)
    return ganadora == cara


def _huella_grupo_en_vista(datos, vista):
    """Cobertura proyectada de una ocurrencia completa dentro de la vista."""
    if not datos:
        return 0.0, 0.0, 0.0
    try:
        ancho = max(d["maxx"] for d in datos) - min(d["minx"] for d in datos)
        alto = max(d["maxy"] for d in datos) - min(d["miny"] for d in datos)
        rel_ancho = ancho / max(EPS, float(vista.Width))
        rel_alto = alto / max(EPS, float(vista.Height))
        return rel_ancho, rel_alto, rel_ancho * rel_alto
    except Exception:
        return 0.0, 0.0, 0.0


def _nombre_final_componente(nombre):
    """Última ocurrencia del path de un componente, para logs y exclusiones."""
    texto = str(nombre or "")
    if "|" in texto:
        texto = texto.rsplit("|", 1)[-1]
    return texto.split(":")[0].strip()


def _vector_hacia_camara(vista):
    """Unitario target -> eye (lo mas de frente tiene mayor proyeccion)."""
    cam = vista.Camera
    ex, ey, ez = float(cam.Eye.X), float(cam.Eye.Y), float(cam.Eye.Z)
    tx, ty, tz = float(cam.Target.X), float(cam.Target.Y), float(cam.Target.Z)
    vx, vy, vz = ex - tx, ey - ty, ez - tz
    n = math.sqrt(vx * vx + vy * vy + vz * vz) or 1.0
    return (vx / n, vy / n, vz / n), (tx, ty, tz)


def _score_frente_segmento(vista, piezas, grupos):
    """
    Que tan 'de frente' esta un segmento en ESTA vista:
    promedio de (punto - target) · look hacia la camara.
    """
    look, target = _vector_hacia_camara(vista)
    vals = []
    piezas_vistas = 0
    for nombre, datos in grupos.items():
        if not _pieza_en_catalogo(nombre, piezas):
            continue
        piezas_vistas += 1
        for dato in datos[:8]:
            pt = _punto3d_curva(dato["curve"])
            if pt is None:
                continue
            vals.append(
                (pt[0] - target[0]) * look[0]
                + (pt[1] - target[1]) * look[1]
                + (pt[2] - target[2]) * look[2]
            )
    if not vals:
        return None
    return (sum(vals) / len(vals), piezas_vistas, len(vals))


def _elegir_segmento_de_vista(vista, segmentos, grupos):
    """
    Elige el Assembly Segmento* mas enfrente de la camara de esta cara.
    No depende del numero del archivo ni del nombre FRONT/BACK.
    """
    if not segmentos:
        return set(), "(sin segmento abierto: keywords)"

    mejor = None
    mejor_key = None
    for seg in segmentos:
        score = _score_frente_segmento(vista, seg["piezas"], grupos)
        if score is None:
            continue
        # Prioriza profundidad hacia camara; desempata por #piezas visibles.
        key = (score[0], score[1])
        if mejor_key is None or key > mejor_key:
            mejor_key = key
            mejor = (seg, score)

    if mejor is None:
        # Ningun segmento tiene curvas en esta vista: no acotar basura.
        return set(), "(ningun segmento visible)"

    seg, score = mejor
    log(
        f"    Segmento de esta vista: {seg['nombre']} "
        f"(frente={score[0]:.2f}, piezas={score[1]}, curvas={score[2]})"
    )
    return seg["piezas"], seg["nombre"]


def _hojas_de_caras(plano):
    hojas = []
    for indice in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(indice)
        if str(hoja.Name).upper().startswith(PREFIJO_HOJA):
            hojas.append(hoja)
    return hojas


def _es_hoja_modelo_autocad(hoja):
    nombre = str(getattr(hoja, "Name", "") or "")
    return "Model (AutoCAD)" in nombre


def _es_hoja_de_caras(hoja):
    return str(getattr(hoja, "Name", "") or "").upper().startswith(PREFIJO_HOJA)


def _es_hoja_pieza_acotada(hoja):
    """Hojas residuales del flujo por pieza (ANCHO/LARGO/THK/diámetro)."""
    nombre = str(getattr(hoja, "Name", "") or "").upper()
    if _es_hoja_de_caras(hoja) or _es_hoja_modelo_autocad(hoja):
        return False
    claves = (
        "_ANCHO", "_LARGO", "_THK",
        "DIAMETRO_EXTERIOR", "DIAMETRO_INTERIOR", "_DIAMETRO",
    )
    return any(clave in nombre for clave in claves)


def _encontrar_hoja_machote(plano):
    """Hoja de layout del machote (nunca Model AutoCAD ni TANQUE_DATUM)."""
    candidatas = []
    for indice in range(1, plano.Sheets.Count + 1):
        try:
            hoja = plano.Sheets.Item(indice)
        except Exception:
            continue
        if (
            _es_hoja_modelo_autocad(hoja)
            or _es_hoja_de_caras(hoja)
            or _es_hoja_pieza_acotada(hoja)
        ):
            continue
        candidatas.append(hoja)

    for hoja in candidatas:
        nombre = str(hoja.Name).upper()
        if "MODELO" in nombre or "COMPLETO" in nombre or "MACHOTE" in nombre:
            return hoja
    for hoja in candidatas:
        nombre = str(hoja.Name).upper()
        if "HOJA" in nombre:
            return hoja
    return candidatas[0] if candidatas else None


def _eliminar_hojas_piezas_residuales(plano):
    """
    Quita hojas basura del flujo por pieza antes de copiar la plantilla.
    Un machote con cientos de hojas hace lentísimo cada Update/Activate.
    """
    hoja_machote = _encontrar_hoja_machote(plano)
    if hoja_machote is not None:
        try:
            hoja_machote.Activate()
        except Exception:
            pass

    borradas = 0
    for indice in range(plano.Sheets.Count, 0, -1):
        try:
            hoja = plano.Sheets.Item(indice)
            if not _es_hoja_pieza_acotada(hoja):
                continue
            if plano.Sheets.Count <= 1:
                break
            hoja.Delete()
            borradas += 1
        except Exception:
            continue
    if borradas:
        log(f"  Hojas residuales de piezas eliminadas: {borradas}")


def _eliminar_hojas_anteriores(plano):
    """Borra solo hojas TANQUE_DATUM_*. Activa antes el layout del machote."""
    hoja_machote = _encontrar_hoja_machote(plano)
    if hoja_machote is not None:
        try:
            hoja_machote.Activate()
        except Exception:
            pass

    for indice in range(plano.Sheets.Count, 0, -1):
        try:
            hoja = plano.Sheets.Item(indice)
            if not _es_hoja_de_caras(hoja):
                continue
            # Nunca dejar el documento sin layout de dibujo.
            if plano.Sheets.Count <= 1:
                _limpiar_vistas(hoja)
                _borrar_cotas_hoja(hoja)
                log(f"  Ultima hoja vaciada (no borrada): {hoja.Name}")
                break
            log(f"  Eliminando hoja de caras: {hoja.Name}")
            hoja.Delete()
        except Exception as error:
            log(f"  AVISO al eliminar hoja #{indice}: {error}")


def _nombre_tanque(ensamble):
    nombre = str(getattr(ensamble, "DisplayName", None) or "TANQUE")
    nombre = nombre.replace(".iam", "").replace(".IAM", "").strip()
    return _limpiar_nombre_archivo(nombre) or "TANQUE"


def _carpeta_salida_tanque(plano, ensamble):
    """JPG/<nombre_del_tanque>/ junto al machote."""
    ruta_dibujo = str(plano.FullFileName or "")
    base = os.path.dirname(ruta_dibujo) if ruta_dibujo else os.path.dirname(os.path.abspath(__file__))
    carpeta = os.path.join(base, CARPETA_EXPORTACION, _nombre_tanque(ensamble))
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


def _limpiar_machote(plano, inv_app=None):
    """
    Solo quita los dibujos de caras (hojas TANQUE_DATUM_*).
    Debe quedar visible la hoja del machote (con marco), NO Model AutoCAD negro.
    """
    log("Limpiando dibujos del machote (sin tocar la hoja plantilla)...")

    # 1) Activar SIEMPRE la hoja del machote ANTES de borrar.
    hoja_machote = _encontrar_hoja_machote(plano)
    if hoja_machote is not None:
        try:
            hoja_machote.Activate()
            log(f"  Hoja machote activa: {hoja_machote.Name}")
        except Exception as error:
            log(f"  AVISO al activar machote: {error}")
    else:
        log("  AVISO: no se encontro hoja de plantilla; se evitara dejar Model AutoCAD activo")

    # 2) Borrar solo hojas de caras.
    try:
        _eliminar_hojas_anteriores(plano)
    except Exception as error:
        log(f"  AVISO al limpiar hojas de caras: {error}")

    # 3) Reactivar hoja de layout (nunca Model AutoCAD).
    hoja_machote = _encontrar_hoja_machote(plano)
    if hoja_machote is not None:
        try:
            hoja_machote.Activate()
            log(f"  Reactivada hoja: {hoja_machote.Name}")
        except Exception as error:
            log(f"  AVISO al reactivar: {error}")
    else:
        # Ultimo recurso: primera hoja que NO sea Model AutoCAD
        for indice in range(1, plano.Sheets.Count + 1):
            try:
                hoja = plano.Sheets.Item(indice)
                if _es_hoja_modelo_autocad(hoja):
                    continue
                hoja.Activate()
                log(f"  Reactivada hoja alternativa: {hoja.Name}")
                break
            except Exception:
                continue

    try:
        plano.Update()
    except Exception:
        pass

    if inv_app is not None:
        try:
            inv_app.ScreenUpdating = True
        except Exception:
            pass
        try:
            _actualizar_inventor(inv_app)
        except Exception:
            pass
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        try:
            inv_app.ActiveView.Fit()
        except Exception:
            pass

    log("  Machote limpio: sin dibujos de caras, hoja plantilla visible.")


def _elegir_hoja_base(plano):
    candidatas = []
    try:
        candidatas.append(plano.ActiveSheet)
    except Exception:
        pass

    for indice in range(1, plano.Sheets.Count + 1):
        try:
            hoja = plano.Sheets.Item(indice)
            if hoja not in candidatas:
                candidatas.append(hoja)
        except Exception:
            pass

    for hoja in candidatas:
        nombre = str(getattr(hoja, "Name", "?"))
        if "Model (AutoCAD)" in nombre:
            continue
        try:
            _ = hoja.DrawingViews.Count
            log(f"  Hoja base elegida: {nombre}")
            return hoja
        except Exception as error:
            log(f"  Hoja no usable '{nombre}': {error}")

    return plano.ActiveSheet


def _limpiar_vistas(hoja):
    for indice in range(hoja.DrawingViews.Count, 0, -1):
        try:
            hoja.DrawingViews.Item(indice).Delete()
        except Exception:
            pass


def _crear_hoja(plano, hoja_base, nombre):
    nueva = hoja_base.CopyTo(plano)
    nueva.Name = nombre
    _limpiar_vistas(nueva)
    # La hoja se usa exclusivamente como lienzo temporal para JPG. El marco,
    # cartucho y divisiones del machote no deben aparecer en la fotografía.
    for atributo in ("Border", "TitleBlock"):
        try:
            objeto = getattr(nueva, atributo)
            if objeto is not None:
                objeto.Delete()
        except Exception:
            pass
    return nueva


K_PLANE_SURFACE = 5890
ANGULO_CLUSTER_DEG = 12.0
AREA_MINIMA_CARA_CM2 = 50.0


def _bbox_ensamble(ensamble):
    box = ensamble.ComponentDefinition.RangeBox
    xmin, ymin, zmin = box.MinPoint.X, box.MinPoint.Y, box.MinPoint.Z
    xmax, ymax, zmax = box.MaxPoint.X, box.MaxPoint.Y, box.MaxPoint.Z
    corners = [
        (xmin, ymin, zmin),
        (xmax, ymin, zmin),
        (xmin, ymax, zmin),
        (xmax, ymax, zmin),
        (xmin, ymin, zmax),
        (xmax, ymin, zmax),
        (xmin, ymax, zmax),
        (xmax, ymax, zmax),
    ]
    return {
        "min": (xmin, ymin, zmin),
        "max": (xmax, ymax, zmax),
        "cx": (xmin + xmax) / 2.0,
        "cy": (ymin + ymax) / 2.0,
        "cz": (zmin + zmax) / 2.0,
        "dx": abs(xmax - xmin),
        "dy": abs(ymax - ymin),
        "dz": abs(zmax - zmin),
        "corners": corners,
    }


def _v3(x, y, z):
    return (float(x), float(y), float(z))


def _norm(v):
    length = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    if length < 1e-12:
        return (0.0, 0.0, 0.0)
    return (v[0] / length, v[1] / length, v[2] / length)


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a, b):
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _scale(v, s):
    return (v[0] * s, v[1] * s, v[2] * s)


def _abs_dot(a, b):
    return abs(_dot(a, b))


def _transformar_normal(n_local, matriz):
    """Aplica solo la rotacion 3x3 de la matriz de ocurrencia."""
    # Inventor Matrix: Cell(row, col) 1-based. Vectores columna.
    try:
        x = (
            matriz.Cell(1, 1) * n_local[0]
            + matriz.Cell(1, 2) * n_local[1]
            + matriz.Cell(1, 3) * n_local[2]
        )
        y = (
            matriz.Cell(2, 1) * n_local[0]
            + matriz.Cell(2, 2) * n_local[1]
            + matriz.Cell(2, 3) * n_local[2]
        )
        z = (
            matriz.Cell(3, 1) * n_local[0]
            + matriz.Cell(3, 2) * n_local[1]
            + matriz.Cell(3, 3) * n_local[2]
        )
        return _norm((x, y, z))
    except Exception:
        return _norm(n_local)


def _vector_global_a_local(vector, transformacion):
    """
    Convierte una dirección del tanque principal a coordenadas del IAM de
    segmento. Así una cámara mantiene la misma cara física aunque el segmento
    esté girado respecto a su propio origen.
    """
    try:
        inversa = transformacion.Copy()
        inversa.Invert()
        x = (
            inversa.Cell(1, 1) * vector[0]
            + inversa.Cell(1, 2) * vector[1]
            + inversa.Cell(1, 3) * vector[2]
        )
        y = (
            inversa.Cell(2, 1) * vector[0]
            + inversa.Cell(2, 2) * vector[1]
            + inversa.Cell(2, 3) * vector[2]
        )
        z = (
            inversa.Cell(3, 1) * vector[0]
            + inversa.Cell(3, 2) * vector[1]
            + inversa.Cell(3, 3) * vector[2]
        )
        local = _norm((x, y, z))
        if _dot(local, local) > 0.5:
            return local
    except Exception:
        pass
    return _norm(vector)


def _recolectar_normales_caras(ensamble):
    """
    Recorre piezas hoja y junta normales de caras planas grandes
    en coordenadas del ensamble. Asi detectamos si el tanque esta chueco.
    """
    muestras = []
    ocurrencias = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
    k_plane = K_PLANE_SURFACE

    for i in range(1, ocurrencias.Count + 1):
        try:
            occ = ocurrencias.Item(i)
            if occ.Suppressed:
                continue
            part = occ.Definition.Document
            if part is None:
                continue
            try:
                part = win32com.client.CastTo(part, "PartDocument")
            except Exception:
                pass

            matriz = occ.Transformation
            bodies = part.ComponentDefinition.SurfaceBodies
            for b in range(1, bodies.Count + 1):
                body = bodies.Item(b)
                faces = body.Faces
                for f in range(1, faces.Count + 1):
                    face = faces.Item(f)
                    try:
                        if face.SurfaceType != k_plane:
                            continue
                        area = float(face.Evaluator.Area)
                        if area < AREA_MINIMA_CARA_CM2:
                            continue
                        n_local = face.Geometry.Normal
                        n_asm = _transformar_normal(
                            (n_local.X, n_local.Y, n_local.Z), matriz
                        )
                        if _dot(n_asm, n_asm) < 0.5:
                            continue
                        muestras.append((n_asm, area))
                    except Exception:
                        continue
        except Exception:
            continue

    log(f"  Caras planas grandes analizadas: {len(muestras)}")
    return muestras


def _clusterizar_normales(muestras):
    cos_lim = math.cos(math.radians(ANGULO_CLUSTER_DEG))
    clusters = []  # cada uno: {"dir": (x,y,z), "area": float}

    for normal, area in muestras:
        n = _norm(normal)
        if _dot(n, n) < 0.5:
            continue

        colocado = False
        for cluster in clusters:
            d = cluster["dir"]
            ad = _dot(n, d)
            if abs(ad) >= cos_lim:
                # Misma direccion u opuesta: unificamos sentido con el dominante.
                if ad < 0:
                    n = _scale(n, -1.0)
                # Promedio ponderado por area
                peso = cluster["area"]
                nuevo = _norm(
                    (
                        d[0] * peso + n[0] * area,
                        d[1] * peso + n[1] * area,
                        d[2] * peso + n[2] * area,
                    )
                )
                cluster["dir"] = nuevo
                cluster["area"] += area
                colocado = True
                break

        if not colocado:
            clusters.append({"dir": n, "area": area})

    clusters.sort(key=lambda c: c["area"], reverse=True)
    return clusters


def _ortogonalizar_tres(dirs):
    """Devuelve hasta 3 ejes ortonormales a partir de direcciones dominantes."""
    if not dirs:
        return []

    x = _norm(dirs[0])
    ejes = [x]

    for candidato in dirs[1:]:
        # Quitar componentes ya usadas
        v = candidato
        for e in ejes:
            v = (
                v[0] - e[0] * _dot(v, e),
                v[1] - e[1] * _dot(v, e),
                v[2] - e[2] * _dot(v, e),
            )
        v = _norm(v)
        if _dot(v, v) > 0.5:
            ejes.append(v)
        if len(ejes) == 3:
            break

    if len(ejes) == 2:
        ejes.append(_norm(_cross(ejes[0], ejes[1])))
    elif len(ejes) == 1:
        # Completar con mundo
        tmp = (0.0, 1.0, 0.0)
        if _abs_dot(ejes[0], tmp) > 0.9:
            tmp = (0.0, 0.0, 1.0)
        y = _norm(_cross(tmp, ejes[0]))
        z = _norm(_cross(ejes[0], y))
        ejes.extend([y, z])

    return ejes


def _ejes_desde_geometria(ensamble, bbox):
    """
    Calcula ejes del tanque a partir de sus paredes reales.
    Si el STEP/ensamble esta chueco respecto al mundo, las camaras
    siguen las normales de las caras (no los ejes XYZ mundiales).
    """
    muestras = _recolectar_normales_caras(ensamble)
    clusters = _clusterizar_normales(muestras)

    if len(clusters) < 2:
        log("  AVISO: pocas caras para orientar; se usan ejes mundiales del bbox.")
        return None

    for idx, c in enumerate(clusters[:5]):
        d = c["dir"]
        log(
            f"  Cluster normal #{idx + 1}: "
            f"({d[0]:.3f},{d[1]:.3f},{d[2]:.3f}) area={c['area']:.1f}"
        )

    # Altura = direccion de cluster mas alineada con "arriba" mundial
    # entre las de mayor area. Preferimos la mas vertical.
    mundo_up_cands = [(0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (1.0, 0.0, 0.0)]
    candidatos = clusters[:6]

    mejor_altura = None
    mejor_score = -1.0
    for c in candidatos:
        for up in mundo_up_cands:
            score = _abs_dot(c["dir"], up) * math.sqrt(c["area"])
            # Premiar mas si es claramente vertical
            score += _abs_dot(c["dir"], up) * 1000.0
            if score > mejor_score:
                mejor_score = score
                mejor_altura = c["dir"] if _dot(c["dir"], up) >= 0 else _scale(c["dir"], -1.0)

    if mejor_altura is None:
        return None

    # Paredes = clusters casi horizontales (perp. a altura) de mayor area
    paredes = []
    for c in candidatos:
        if _abs_dot(c["dir"], mejor_altura) > 0.35:
            continue
        paredes.append(c["dir"])

    if len(paredes) < 1:
        # Fallback: ortogonalizar clusters grandes
        ejes = _ortogonalizar_tres([c["dir"] for c in candidatos])
        # Elegir altura = mas vertical
        altura = max(ejes, key=lambda e: max(_abs_dot(e, u) for u in mundo_up_cands))
        if _dot(altura, (0.0, 1.0, 0.0)) < 0 and _dot(altura, (0.0, 0.0, 1.0)) < 0:
            altura = _scale(altura, -1.0)
        paredes_eje = [e for e in ejes if _abs_dot(e, altura) < 0.5]
        if len(paredes_eje) < 2:
            paredes_eje = _ortogonalizar_tres(ejes)
            paredes_eje = [e for e in paredes_eje if _abs_dot(e, altura) < 0.5]
        log(
            f"  Ejes (fallback ortogonal): altura={altura} paredes={paredes_eje[:2]}"
        )
        return altura, paredes_eje[:2]

    # Segunda pared ortogonal a altura y a la primera
    p1 = _norm(paredes[0])
    p2 = None
    for p in paredes[1:]:
        v = (
            p[0] - p1[0] * _dot(p, p1) - mejor_altura[0] * _dot(p, mejor_altura),
            p[1] - p1[1] * _dot(p, p1) - mejor_altura[1] * _dot(p, mejor_altura),
            p[2] - p1[2] * _dot(p, p1) - mejor_altura[2] * _dot(p, mejor_altura),
        )
        v = _norm(v)
        if _dot(v, v) > 0.5:
            p2 = v
            break
    if p2 is None:
        p2 = _norm(_cross(mejor_altura, p1))

    # Asegurar tipode derecha
    if _dot(_cross(p1, p2), mejor_altura) < 0:
        p2 = _scale(p2, -1.0)

    log(
        f"  Ejes tanque por caras: altura=({mejor_altura[0]:.3f},{mejor_altura[1]:.3f},{mejor_altura[2]:.3f}) "
        f"FRONT=({p1[0]:.3f},{p1[1]:.3f},{p1[2]:.3f}) "
        f"RIGHT=({p2[0]:.3f},{p2[1]:.3f},{p2[2]:.3f})"
    )
    return mejor_altura, [p1, p2]


def _ejes_mundo_bbox(bbox):
    extents = [("X", bbox["dx"], (1.0, 0.0, 0.0)),
               ("Y", bbox["dy"], (0.0, 1.0, 0.0)),
               ("Z", bbox["dz"], (0.0, 0.0, 1.0))]
    extents.sort(key=lambda item: item[1], reverse=True)
    altura = extents[0][2]
    paredes = [extents[1][2], extents[2][2]]
    log(
        f"  Ejes mundiales bbox: altura={extents[0][0]} "
        f"paredes={extents[1][0]}/{extents[2][0]}"
    )
    return altura, paredes


def _tg_vec(tg, v):
    return tg.CreateVector(v[0], v[1], v[2])


def _direcciones_caras_pqart(tg, cover, face, right):
    """
    Misma triada que PQart 2.0/3.0:
      cover = +Y (arriba en la foto)
      face  = +Z (FRONT)
      right = +X (RIGHT)
    Cada hoja mira una pared exterior, derecha (no chueca).
    """
    return [
        ("FRONT", _tg_vec(tg, face), _tg_vec(tg, cover)),
        ("BACK", _tg_vec(tg, _scale(face, -1.0)), _tg_vec(tg, cover)),
        ("RIGHT", _tg_vec(tg, right), _tg_vec(tg, cover)),
        ("LEFT", _tg_vec(tg, _scale(right, -1.0)), _tg_vec(tg, cover)),
    ]


def _crear_camara(ensamble, tg, to, bbox, eye_dir, up_hint):
    eye_dir = eye_dir.Copy()
    eye_dir.Normalize()
    # Up = tapa (cover) del marco, no forzar XYZ mundo.
    try:
        up_hint = up_hint.Copy()
        up_hint.Normalize()
    except Exception:
        up_hint = tg.CreateVector(0.0, 1.0, 0.0)
    proy = eye_dir.DotProduct(up_hint)
    up_hint = tg.CreateVector(
        up_hint.X - eye_dir.X * proy,
        up_hint.Y - eye_dir.Y * proy,
        up_hint.Z - eye_dir.Z * proy,
    )
    if up_hint.Length < 0.001:
        up_hint = tg.CreateVector(0.0, 1.0, 0.0)
        proy = eye_dir.DotProduct(up_hint)
        up_hint = tg.CreateVector(
            up_hint.X - eye_dir.X * proy,
            up_hint.Y - eye_dir.Y * proy,
            up_hint.Z - eye_dir.Z * proy,
        )
    up_hint.Normalize()

    right = eye_dir.CrossProduct(up_hint)
    if right.Length < 0.001:
        temp = tg.CreateVector(1, 0, 0)
        if abs(eye_dir.DotProduct(temp)) > 0.9:
            temp = tg.CreateVector(0, 0, 1)
        right = eye_dir.CrossProduct(temp)
    right.Normalize()
    true_up = right.CrossProduct(eye_dir)
    true_up.Normalize()

    dist = max(bbox["dx"], bbox["dy"], bbox["dz"], 10.0) * 2.5
    cam = to.CreateCamera()
    cam.SceneObject = ensamble.ComponentDefinition
    cam.Perspective = False
    cam.Target = tg.CreatePoint(bbox["cx"], bbox["cy"], bbox["cz"])
    cam.Eye = tg.CreatePoint(
        bbox["cx"] + eye_dir.X * dist,
        bbox["cy"] + eye_dir.Y * dist,
        bbox["cz"] + eye_dir.Z * dist,
    )
    cam.UpVector = tg.CreateUnitVector(true_up.X, true_up.Y, true_up.Z)
    try:
        cam.ApplyWithoutTransition()
    except Exception:
        try:
            cam.Apply()
        except Exception:
            pass
    return cam


def _tipo_dim(dim):
    try:
        return int(dim.DimensionType)
    except Exception:
        try:
            return int(dim.DimensionType.value)
        except Exception:
            return -1


def _es_cota_ortogonal(dim, tipo_esperado):
    """Rechaza alineadas. No usar RangeBox (las extension lines inflan dy/dx)."""
    tipo = _tipo_dim(dim)
    if tipo < 0:
        return True  # conservar si no se puede leer
    if tipo == COTA_ALINEADA:
        return False
    if tipo_esperado is not None and tipo != int(tipo_esperado):
        return False
    return True


def _purgar_solo_alineadas(hoja):
    """Borra unicamente cotas alineadas; no tocar H/V validas."""
    try:
        dims = hoja.DrawingDimensions.GeneralDimensions
        for i in range(dims.Count, 0, -1):
            try:
                dim = dims.Item(i)
                if _tipo_dim(dim) == COTA_ALINEADA:
                    dim.Delete()
            except Exception:
                pass
    except Exception:
        pass


def _limpiar_bolitas(hoja, vista=None):
    """
    Borra centermarks/centerlines/workpoints (bolitas).
    No toca OriginIndicator: marca el (0,0) de soldadura.
    """
    if vista is not None:
        for attr in ("ShowWorkPoints", "ShowWorkAxes", "ShowWorkPlanes"):
            try:
                setattr(vista, attr, False)
            except Exception:
                pass
        try:
            vista.Include3DAnnotations = False
        except Exception:
            pass
        try:
            cms = vista.Centermarks
            for i in range(cms.Count, 0, -1):
                try:
                    cms.Item(i).Delete()
                except Exception:
                    pass
        except Exception:
            pass

    for attr in ("Centermarks", "Centerlines"):
        try:
            col = getattr(hoja, attr)
            for _ in range(5):
                n = col.Count
                if n <= 0:
                    break
                for i in range(n, 0, -1):
                    try:
                        col.Item(i).Delete()
                    except Exception:
                        pass
        except Exception:
            pass


def _quitar_flechas_tipo_bolita(dim, hoja=None):
    """
    NO tocar dim.Style (es compartido del machote y apaga cotas).
    Solo ajustes de instancia si existen.
    """
    for attr in ("ArrowheadsInside",):
        try:
            setattr(dim, attr, False)
        except Exception:
            pass


def _intent_recta_orientada(hoja, dato, lado, exigir_vertical):
    """Prefiere curvas H/V; si no califica, igual crea intent (no bloquear cotas)."""
    ok = True
    if exigir_vertical:
        if dato["dy"] < max(0.08, dato["dx"] * 1.2):
            ok = False
    else:
        if dato["dx"] < max(0.08, dato["dy"] * 1.2):
            ok = False
    intent = _intent_seguro(hoja, dato, lado)
    if intent is not None:
        return intent
    if not ok:
        return None
    return intent


def _crear_cota_desde_muro(hoja, tg, vista, intent_origen, dato_obj, lado, tipo, texto_xy, dist_hoja):
    """Cota H/V desde el muro. Solo borra si queda alineada (tipo 60161)."""
    intent_obj = _intent_seguro(hoja, dato_obj, lado)
    if intent_obj is None or intent_origen is None:
        return False

    try:
        dim = hoja.DrawingDimensions.GeneralDimensions.AddLinear(
            tg.CreatePoint2d(float(texto_xy[0]), float(texto_xy[1])),
            intent_origen,
            intent_obj,
            int(tipo),
        )
    except Exception:
        return False

    tipo_real = _tipo_dim(dim)
    if tipo_real == COTA_ALINEADA:
        try:
            dim.Delete()
        except Exception:
            pass
        return False

    # Si no es H/V pedida, borrar; si no se lee tipo, conservar.
    if tipo_real > 0 and tipo_real != int(tipo):
        try:
            dim.Delete()
        except Exception:
            pass
        return False

    _quitar_flechas_tipo_bolita(dim, hoja)
    aplicar_estilo_cota(dim, hoja=hoja)
    return True


def _enderezar_vista_en_hoja(vista):
    """Rota la vista en la hoja para que bordes largos queden H/V."""
    try:
        curvas = vista.DrawingCurves
        total = curvas.Count
    except Exception:
        return

    muestras = []
    for i in range(1, total + 1):
        try:
            curva = curvas.Item(i)
            caja = curva.Evaluator2D.RangeBox
            dx = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
            dy = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
            largo = math.hypot(dx, dy)
            if largo < 1.5:
                continue
            try:
                p1 = curva.StartPoint
                p2 = curva.EndPoint
                ang = math.degrees(
                    math.atan2(float(p2.Y) - float(p1.Y), float(p2.X) - float(p1.X))
                )
            except Exception:
                continue

            while ang > 90.0:
                ang -= 180.0
            while ang < -90.0:
                ang += 180.0

            if dx >= dy * 1.8:
                if abs(ang) <= 45.0:
                    muestras.append((ang, largo))
            elif dy >= dx * 1.8:
                desv = ang - (90.0 if ang >= 0 else -90.0)
                while desv > 90.0:
                    desv -= 180.0
                while desv < -90.0:
                    desv += 180.0
                if abs(desv) <= 45.0:
                    muestras.append((desv, largo))
        except Exception:
            continue

    if len(muestras) < 4:
        return

    pesos = []
    for ang, w in muestras:
        reps = max(1, min(20, int(w)))
        pesos.extend([ang] * reps)
    pesos.sort()
    mediana = pesos[len(pesos) // 2]

    if abs(mediana) < 0.5 or abs(mediana) > 40.0:
        return

    try:
        actual = float(vista.Rotate)
        vista.Rotate = actual - mediana
        log(f"    Vista enderezada en hoja: {mediana:.2f} deg")
    except Exception as error:
        log(f"    AVISO Rotate vista: {error}")


def _ajustar_vista(vista, hoja, tg):
    if vista.Width <= 0 or vista.Height <= 0:
        return

    # Tamaño inicial amplio. Después de contar los extremos reales, el
    # encuadre final reserva exactamente el espacio necesario para las cotas.
    ancho = hoja.Width * 0.74
    alto = hoja.Height * 0.74
    factor = min(ancho / vista.Width, alto / vista.Height) * 0.92
    factor = max(0.01, min(factor, 4.0))
    vista.Scale = vista.Scale * factor
    vista.Position = tg.CreatePoint2d(hoja.Width * 0.60, hoja.Height * 0.61)


def _info_curva(curva):
    try:
        # Solo geometria realmente visible en HLR (evita puntos fantasma).
        try:
            if curva.Segments.Count <= 0:
                return None
        except Exception:
            pass

        caja = curva.Evaluator2D.RangeBox
        minx = float(caja.MinPoint.X)
        maxx = float(caja.MaxPoint.X)
        miny = float(caja.MinPoint.Y)
        maxy = float(caja.MaxPoint.Y)
        dx = abs(maxx - minx)
        dy = abs(maxy - miny)
        if dx < EPS and dy < EPS:
            return None
        return {
            "curve": curva,
            "minx": minx,
            "maxx": maxx,
            "miny": miny,
            "maxy": maxy,
            "dx": dx,
            "dy": dy,
            "cx": (minx + maxx) / 2.0,
            "cy": (miny + maxy) / 2.0,
        }
    except Exception:
        return None


def _ruta_occurrence(ocurrencia):
    """Ruta estable de una ocurrencia hoja dentro del IAM principal."""
    partes = []
    actual = ocurrencia
    for _ in range(12):
        if actual is None:
            break
        try:
            partes.append(str(actual.Name))
            actual = actual.ParentOccurrence
        except Exception:
            break
    return "|".join(reversed(partes)) if partes else ""


def _nombre_componente(curva):
    try:
        ocurrencia = curva.ModelGeometry.ContainingOccurrence
        if ocurrencia is not None:
            # Nombre completo para no mezclar "PIPE FLANGE 0.250" de dos
            # caras distintas. El nombre simple no es único en muchos tanques.
            ruta = _ruta_occurrence(ocurrencia)
            if ruta:
                return ruta
    except Exception:
        pass
    return "SIN_COMPONENTE"


def _curvas_por_componente(vista):
    grupos = {}
    try:
        curvas = vista.DrawingCurves
        total = curvas.Count
    except Exception as error:
        log(f"    No se pudieron leer curvas: {error}")
        return {}, []

    log(f"    Curvas en la vista: {total}")
    todas = []

    for indice in range(1, total + 1):
        try:
            curva = curvas.Item(indice)
        except Exception:
            continue
        info = _info_curva(curva)
        if info is None:
            continue
        todas.append(info)
        nombre = _nombre_componente(curva)
        grupos.setdefault(nombre, []).append(info)

    return grupos, todas


def _es_recta_dominante(d, lado):
    if lado in ("izq", "der"):
        return d["dy"] >= max(EPS, d["dx"] * DOMINANCIA_RECTA)
    return d["dx"] >= max(EPS, d["dy"] * DOMINANCIA_RECTA)


def _elegir_extrema(datos, lado, tol):
    if not datos:
        return None

    if lado == "izq":
        objetivo = min(d["minx"] for d in datos)
        cands = [d for d in datos if abs(d["minx"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante(d, lado)]
        base = rectos or cands
        return max(base, key=lambda d: (d["dy"], d["dx"])) if base else None

    if lado == "der":
        objetivo = max(d["maxx"] for d in datos)
        cands = [d for d in datos if abs(d["maxx"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante(d, lado)]
        base = rectos or cands
        return max(base, key=lambda d: (d["dy"], d["dx"])) if base else None

    if lado == "inf":
        objetivo = min(d["miny"] for d in datos)
        cands = [d for d in datos if abs(d["miny"] - objetivo) <= tol]
        rectos = [d for d in cands if _es_recta_dominante(d, lado)]
        base = rectos or cands
        return max(base, key=lambda d: (d["dx"], d["dy"])) if base else None

    # sup
    objetivo = max(d["maxy"] for d in datos)
    cands = [d for d in datos if abs(d["maxy"] - objetivo) <= tol]
    rectos = [d for d in cands if _es_recta_dominante(d, lado)]
    base = rectos or cands
    return max(base, key=lambda d: (d["dx"], d["dy"])) if base else None


def _puntos_clave(curva):
    puntos = []
    usados = set()
    for attr in ("StartPoint", "MidPoint", "EndPoint"):
        try:
            p = getattr(curva, attr)
            if p is None:
                continue
            key = (round(float(p.X), 6), round(float(p.Y), 6))
            if key not in usados:
                usados.add(key)
                puntos.append(p)
        except Exception:
            pass
    return puntos


def _intent_seguro(hoja, dato, lado):
    curva = dato["curve"]
    puntos = _puntos_clave(curva)
    if puntos:
        try:
            if lado == "izq":
                p = min(puntos, key=lambda t: float(t.X))
            elif lado == "der":
                p = max(puntos, key=lambda t: float(t.X))
            elif lado == "inf":
                p = min(puntos, key=lambda t: float(t.Y))
            else:
                p = max(puntos, key=lambda t: float(t.Y))
            return hoja.CreateGeometryIntent(curva, p)
        except Exception:
            pass
    try:
        return hoja.CreateGeometryIntent(curva)
    except Exception:
        return None


def _bbox_cuerpo(todas, vista):
    """Caja del cuerpo por aristas largas (ignora piezas sueltas / fantasma)."""
    hmin = max(1.0, float(vista.Height) * FRAC_MURO_VISTA)
    wmin = max(1.0, float(vista.Width) * FRAC_MURO_VISTA)
    largas = [d for d in todas if d["dx"] >= wmin or d["dy"] >= hmin]
    if len(largas) < 2:
        largas = todas
    minx = min(d["minx"] for d in largas)
    maxx = max(d["maxx"] for d in largas)
    miny = min(d["miny"] for d in largas)
    maxy = max(d["maxy"] for d in largas)
    mx = max(0.5, (maxx - minx) * FRAC_EXPAND_LUGS)
    my = max(0.5, (maxy - miny) * FRAC_EXPAND_LUGS)
    return (minx - mx, maxx + mx, miny - my, maxy + my)


def _curva_en_cara(dato, bbox):
    return (
        bbox[0] <= dato["cx"] <= bbox[1]
        and bbox[2] <= dato["cy"] <= bbox[3]
    )


def _filtrar_curvas_cara(todas, vista):
    bbox = _bbox_cuerpo(todas, vista)
    filtradas = [d for d in todas if _curva_en_cara(d, bbox)]
    if len(filtradas) < 8:
        return todas, bbox
    return filtradas, bbox


def _elegir_muro_datum(datos, lado, vista):
    """
    (0,0) en la esquina del cuerpo de la cara:
    - izq: vertical LARGA mas a la izquierda (no pico de lug corto)
    - inf: horizontal LARGA mas abajo
    """
    if not datos:
        return None

    hmin = max(1.2, float(vista.Height) * FRAC_MURO_VISTA)
    wmin = max(1.2, float(vista.Width) * FRAC_MURO_VISTA)

    if lado == "izq":
        muros = [
            d for d in datos
            if d["dy"] >= hmin and d["dy"] >= max(EPS, d["dx"] * DOMINANCIA_RECTA)
        ]
        if not muros:
            muros = [d for d in datos if d["dy"] >= d["dx"] * 1.5]
        if not muros:
            return _elegir_extrema(datos, "izq", max(0.05, vista.Width * 0.02))
        # Entre muros largos, el de mas a la izquierda.
        xmin = min(d["minx"] for d in muros)
        tol = max(0.08, float(vista.Width) * 0.01)
        cands = [d for d in muros if abs(d["minx"] - xmin) <= tol]
        return max(cands, key=lambda d: d["dy"])

    if lado == "inf":
        muros = [
            d for d in datos
            if d["dx"] >= wmin and d["dx"] >= max(EPS, d["dy"] * DOMINANCIA_RECTA)
        ]
        if not muros:
            muros = [d for d in datos if d["dx"] >= d["dy"] * 1.5]
        if not muros:
            return _elegir_extrema(datos, "inf", max(0.05, vista.Height * 0.02))
        ymin = min(d["miny"] for d in muros)
        tol = max(0.08, float(vista.Height) * 0.01)
        cands = [d for d in muros if abs(d["miny"] - ymin) <= tol]
        return max(cands, key=lambda d: d["dx"])

    return _elegir_extrema(datos, lado, 0.1)


def _intent_en_coordenada(hoja, tg, dato, x, y):
    try:
        return hoja.CreateGeometryIntent(dato["curve"], tg.CreatePoint2d(x, y))
    except Exception:
        return None


def _agregar_posicion(
    lista, valor, curva_dato, tolerancia, lado, pieza_id, contacto=False
):
    """
    Conserva los cuatro extremos de una misma pieza, incluso si su espesor
    proyectado es pequeño. Solo fusiona extremos de piezas diferentes cuando
    son realmente la misma coordenada de hoja.
    """
    for existente in lista:
        if existente["pieza_id"] == pieza_id:
            if existente["lado"] == lado and abs(existente["valor"] - valor) <= EPS:
                return
            continue
    lista.append(
        {
            "valor": valor,
            "dato": curva_dato,
            "lado": lado,
            "pieza_id": pieza_id,
            "contacto": bool(contacto),
        }
    )


def _agrupar_referencias_typ(
    posiciones, tolerancia, eje=None, origen=None, vista=None, hoja=None
):
    """
    Consolida referencias que producen exactamente la misma cota visible.

    La dimensión se exporta una sola vez con TYP; se conservan sus miembros
    para marcar en azul todos los extremos a los que aplica la igualdad.
    """
    if not posiciones:
        return []

    grupos = []
    for posicion in sorted(posiciones, key=lambda item: float(item["valor"])):
        # La escala del dibujo puede volver diferentes dos coordenadas de
        # hoja que Inventor finalmente imprime igual (por ejemplo 50.000).
        # La clave se calcula con exactamente el mismo formato del JPG.
        if eje is not None and origen is not None and vista is not None:
            clave = _valor_real_desde_hoja(
                vista, origen, posicion["valor"], hoja
            )
        else:
            clave = round(float(posicion["valor"]) / max(tolerancia, EPS))

        grupo_destino = None
        for grupo in reversed(grupos):
            if grupo["clave"] != clave:
                continue
            # Xmin/Xmax o Ymin/Ymax de la misma pieza no son redundantes:
            # representan límites físicos distintos aunque redondeen igual.
            mismo_extremo_de_pieza = any(
                miembro["pieza_id"] == posicion["pieza_id"]
                and miembro["lado"] != posicion["lado"]
                for miembro in grupo["miembros"]
            )
            if not mismo_extremo_de_pieza:
                grupo_destino = grupo
                break

        if grupo_destino is None:
            grupos.append(
                {
                    "clave": clave,
                    "valor": float(posicion["valor"]),
                    "miembros": [posicion],
                }
            )
        else:
            grupo_destino["miembros"].append(posicion)

    consolidadas = []
    for grupo in grupos:
        miembros = grupo["miembros"]
        referencia = dict(miembros[0])
        piezas = {str(miembro["pieza_id"]) for miembro in miembros}
        referencia["miembros"] = miembros
        referencia["typ"] = len(piezas) > 1
        # Promediar solo elimina ruido de proyección; el miembro representativo
        # sigue definiendo la arista a la cual llega la línea de extensión.
        referencia["valor"] = sum(
            float(miembro["valor"]) for miembro in miembros
        ) / len(miembros)
        consolidadas.append(referencia)
    return consolidadas


def _es_componente_circular(datos, envolvente, nombre_pieza=""):
    """
    Un flange/boss circular se localiza por su centro. Las líneas rectas de un
    cuadro con barrenos impiden que se confunda con una pieza circular.
    """
    if envolvente is None:
        return False
    ancho = float(envolvente["dx"])
    alto = float(envolvente["dy"])
    if max(ancho, alto) <= EPS:
        return False
    # Los nipples se modelan como cilindros huecos de longitud variable: su
    # caja incluye esa profundidad, pero su datum de montaje siempre es el
    # eje/centro del diámetro que se ve de frente.
    if "NIPPLE" in str(nombre_pieza).upper():
        return True

    # Leer la geometría 3D de la pieza, no los trazos HLR. Una pieza como
    # OIL GAUGE BOSS tiene un barreno central y muchos bordes/toroides; exigir
    # que *todas* sus aristas sean Circle la clasificaba incorrectamente como
    # rectangular.
    try:
        ocurrencia = datos[0]["curve"].ModelGeometry.ContainingOccurrence
        documento = ocurrencia.Definition.Document
        cuerpos = documento.ComponentDefinition.SurfaceBodies
        circulos = 0
        rectas = 0
        cilindros = 0
        toroides = 0
        for i in range(1, cuerpos.Count + 1):
            cuerpo = cuerpos.Item(i)
            for j in range(1, cuerpo.Faces.Count + 1):
                try:
                    tipo_superficie = int(cuerpo.Faces.Item(j).SurfaceType)
                    if tipo_superficie == 5891:  # kCylinderSurface
                        cilindros += 1
                    elif tipo_superficie == 5893:  # kTorusSurface
                        toroides += 1
                except Exception:
                    continue
            for j in range(1, cuerpo.Edges.Count + 1):
                try:
                    geometria = cuerpo.Edges.Item(j).Geometry
                    tipo = str(type(geometria)).upper()
                    # win32com expone tipos como "...Circle'>" (sin punto
                    # posterior), por lo que la coincidencia antigua
                    # ".CIRCLE." no reconocía nipples ni tubos redondos.
                    if "CIRCLE" in tipo:
                        circulos += 1
                    elif "LINESEGMENT" in tipo or ".LINE." in tipo:
                        rectas += 1
                except Exception:
                    continue
        # Un cuerpo circular directo (flange, tierra/boss redondo) se compone
        # de contornos Circle; aplica aunque su profundidad distorsione la
        # envolvente proyectada (PIPE HALF NIPPLE, por ejemplo).
        if circulos >= 2 and rectas == 0:
            return True
        # Bosses mecanizados pueden llevar planos pequeños, barreno y filetes,
        # pero su forma dominante sigue siendo cilíndrica y el perfil frontal
        # es cuadrado. Un patch con uno o dos barrenos no alcanza este umbral.
        if (
            abs(ancho - alto) <= max(ancho, alto) * 0.12
            and cilindros >= 4
            and toroides >= 1
        ):
            return True
    except Exception:
        pass
    return False


def _agregar_inicio_y_fin(
    posiciones_x,
    posiciones_y,
    datos,
    tol_extremo,
    tol_pos,
    pieza_id,
    contacto=False,
):
    """
    Desde (0,0) se acota el INICIO y el FIN de cada componente:
    Xmin/Xmax e Ymin/Ymax. No solo una esquina de referencia.
    """
    izq = _elegir_extrema(datos, "izq", tol_extremo)
    der = _elegir_extrema(datos, "der", tol_extremo)
    inf = _elegir_extrema(datos, "inf", tol_extremo)
    sup = _elegir_extrema(datos, "sup", tol_extremo)
    if izq is not None:
        _agregar_posicion(
            posiciones_x, izq["minx"], izq, tol_pos, "izq", pieza_id, contacto
        )
    if der is not None:
        _agregar_posicion(
            posiciones_x, der["maxx"], der, tol_pos, "der", pieza_id, contacto
        )
    if inf is not None:
        _agregar_posicion(
            posiciones_y, inf["miny"], inf, tol_pos, "inf", pieza_id, contacto
        )
    if sup is not None:
        _agregar_posicion(
            posiciones_y, sup["maxy"], sup, tol_pos, "sup", pieza_id, contacto
        )


def _envolvente_occurrence_en_hoja(datos, vista, tg):
    """
    Proyecta los ocho vértices del RangeBox de la ocurrencia a la hoja.
    Las curvas HLR de una solera pueden mostrar solo una arista; la caja 3D
    garantiza que se dimensionen los cuatro extremos reales de la pieza.
    """
    if not datos:
        return None
    try:
        ocurrencia = datos[0]["curve"].ModelGeometry.ContainingOccurrence
        caja = ocurrencia.RangeBox
        xs = (float(caja.MinPoint.X), float(caja.MaxPoint.X))
        ys = (float(caja.MinPoint.Y), float(caja.MaxPoint.Y))
        zs = (float(caja.MinPoint.Z), float(caja.MaxPoint.Z))
        proyectados = []
        for x in xs:
            for y in ys:
                for z in zs:
                    punto = vista.ModelToSheetSpace(tg.CreatePoint(x, y, z))
                    proyectados.append((float(punto.X), float(punto.Y)))
        if len(proyectados) != 8:
            return None
        minx = min(p[0] for p in proyectados)
        maxx = max(p[0] for p in proyectados)
        miny = min(p[1] for p in proyectados)
        maxy = max(p[1] for p in proyectados)
        if maxx - minx < EPS and maxy - miny < EPS:
            return None
        dato = dict(datos[0])
        dato.update(
            {
                # Se conserva el grupo de curvas HLR para refrescar la
                # posición exacta cuando la vista cambie de escala/posición
                # durante la exportación individual.
                "curvas": list(datos),
                "minx": minx,
                "maxx": maxx,
                "miny": miny,
                "maxy": maxy,
                "dx": maxx - minx,
                "dy": maxy - miny,
                "cx": (minx + maxx) * 0.5,
                "cy": (miny + maxy) * 0.5,
            }
        )
        return dato
    except Exception:
        return None


def _envolvente_visible_refrescada(dato, vista=None, tg=None):
    """
    Devuelve la envolvente HLR vigente en hoja para una pieza.

    No se transforma la coordenada de forma matemática tras mover una vista:
    Inventor puede aplicar una rotación/traslación que no coincide con esa
    aproximación. La cota debe usar las curvas visibles actuales para que su
    línea de extensión termine sobre la arista que el operario ve.
    """
    puntos_modelo = dato.get("puntos_modelo")
    if puntos_modelo and vista is not None and tg is not None:
        try:
            proyectados = []
            for x, y, z in puntos_modelo:
                punto = vista.ModelToSheetSpace(tg.CreatePoint(x, y, z))
                proyectados.append((float(punto.X), float(punto.Y)))
            if len(proyectados) >= 2:
                resultado = dict(dato)
                resultado["minx"] = min(p[0] for p in proyectados)
                resultado["maxx"] = max(p[0] for p in proyectados)
                resultado["miny"] = min(p[1] for p in proyectados)
                resultado["maxy"] = max(p[1] for p in proyectados)
                resultado["dx"] = resultado["maxx"] - resultado["minx"]
                resultado["dy"] = resultado["maxy"] - resultado["miny"]
                resultado["cx"] = (resultado["minx"] + resultado["maxx"]) * 0.5
                resultado["cy"] = (resultado["miny"] + resultado["maxy"]) * 0.5
                return resultado
        except Exception:
            return None

    originales = dato.get("curvas") or [dato]
    curvas = []
    for anterior in originales:
        try:
            actual = _info_curva(anterior["curve"])
            if actual is not None:
                curvas.append(actual)
        except Exception:
            continue
    if not curvas:
        return None

    resultado = dict(curvas[0])
    resultado["curvas"] = curvas
    resultado["minx"] = min(curva["minx"] for curva in curvas)
    resultado["maxx"] = max(curva["maxx"] for curva in curvas)
    resultado["miny"] = min(curva["miny"] for curva in curvas)
    resultado["maxy"] = max(curva["maxy"] for curva in curvas)
    resultado["dx"] = resultado["maxx"] - resultado["minx"]
    resultado["dy"] = resultado["maxy"] - resultado["miny"]
    resultado["cx"] = (resultado["minx"] + resultado["maxx"]) * 0.5
    resultado["cy"] = (resultado["miny"] + resultado["maxy"]) * 0.5
    return resultado


def _valor_referencia_desde_envolvente(eje, lado, envolvente):
    if eje == "X":
        campos = {"izq": "minx", "der": "maxx", "centro": "cx"}
    else:
        campos = {"inf": "miny", "sup": "maxy", "centro": "cy"}
    campo = campos.get(lado)
    if campo is None:
        raise ValueError(f"Lado de referencia no reconocido: {lado}")
    return float(envolvente[campo])


def _refrescar_miembros_referencia(eje, referencia, vista, tg):
    """Actualiza cada extremo equivalente tras recolocar la vista."""
    miembros_actuales = []
    for miembro in referencia.get("miembros", [referencia]):
        envolvente = _envolvente_visible_refrescada(miembro["dato"], vista, tg)
        if envolvente is None:
            continue
        actual = dict(miembro)
        actual["dato"] = envolvente
        actual["valor"] = _valor_referencia_desde_envolvente(
            eje, actual["lado"], envolvente
        )
        miembros_actuales.append(actual)
    return miembros_actuales


def _refrescar_posiciones_para_typ(eje, posiciones, vista, tg):
    """
    Relee la geometría HLR antes de consolidar TYP.

    Las envolventes 3D sirven para no perder espesores, pero una fotografía
    individual se acota contra sus curvas HLR visibles. Comparar estas mismas
    curvas evita que dos JPG muestren 50.000 por separado.
    """
    refrescadas = []
    for posicion in posiciones:
        envolvente = _envolvente_visible_refrescada(
            posicion["dato"], vista, tg
        )
        if envolvente is None:
            refrescadas.append(posicion)
            continue
        actual = dict(posicion)
        actual["dato"] = envolvente
        actual["valor"] = _valor_referencia_desde_envolvente(
            eje, actual["lado"], envolvente
        )
        refrescadas.append(actual)
    return refrescadas


def _occurrence_hoja_en_vista(vista, ruta):
    """Obtiene la ocurrencia leaf real del modelo de una DrawingView."""
    try:
        documento = vista.ReferencedDocumentDescriptor.ReferencedDocument
        ensamble = _como_ensamble(documento)
        hojas = ensamble.ComponentDefinition.Occurrences.AllLeafOccurrences
        for i in range(1, hojas.Count + 1):
            ocurrencia = hojas.Item(i)
            if _ruta_occurrence(ocurrencia) == ruta:
                return ocurrencia
    except Exception:
        pass
    return None


def _caras_planas_alineadas(ocurrencia, direccion_camara):
    """Caras planas de una pieza paralelas al plano de la vista."""
    candidatas = []
    try:
        cuerpos = ocurrencia.Definition.Document.ComponentDefinition.SurfaceBodies
        for i in range(1, cuerpos.Count + 1):
            cuerpo = cuerpos.Item(i)
            for j in range(1, cuerpo.Faces.Count + 1):
                cara = cuerpo.Faces.Item(j)
                try:
                    if int(cara.SurfaceType) != 5890 or not cara.Vertices.Count:
                        continue
                    proxy = ocurrencia.CreateGeometryProxy(cara)
                    normal = proxy.Geometry.Normal
                    normal_tupla = _norm(
                        (float(normal.X), float(normal.Y), float(normal.Z))
                    )
                    if abs(_dot(normal_tupla, direccion_camara)) < 0.92:
                        continue
                    puntos = [
                        proxy.Vertices.Item(k).Point
                        for k in range(1, proxy.Vertices.Count + 1)
                    ]
                    modelo = [
                        (float(p.X), float(p.Y), float(p.Z)) for p in puntos
                    ]
                    if len(modelo) < 3:
                        continue
                    profundidad = sum(
                        _dot(punto, direccion_camara) for punto in modelo
                    ) / len(modelo)
                    candidatas.append(
                        {
                            "puntos_modelo": modelo,
                            "profundidad": profundidad,
                            "area": float(cara.Evaluator.Area),
                        }
                    )
                except Exception:
                    continue
    except Exception:
        pass
    return candidatas


def _cara_contacto_lug_en_hoja(
    ruta_lug, ruta_panel, datos_lug, vista, tg
):
    """
    Proyecta la cara física del lug que apoya sobre la placa madre.

    No infiere el contacto desde contornos HLR: compara los planos reales del
    lug y la placa, por lo que elimina aristas de soleras o del perfil del lug
    que no forman parte de la zona de anclaje.
    """
    if not ruta_lug or not ruta_panel or not datos_lug:
        return None
    try:
        occ_lug = _occurrence_hoja_en_vista(vista, ruta_lug)
        occ_panel = _occurrence_hoja_en_vista(vista, ruta_panel)
        if occ_lug is None or occ_panel is None:
            return None

        direccion_camara, _ = _vector_hacia_camara(vista)
        paneles = _caras_planas_alineadas(occ_panel, direccion_camara)
        lugs = _caras_planas_alineadas(occ_lug, direccion_camara)
        if not paneles or not lugs:
            return None

        # Solo las caras grandes de la placa representan sus dos pieles; los
        # barrenos y pequeños rasgos no deben convertirse en planos de apoyo.
        area_max_panel = max(cara["area"] for cara in paneles)
        planos_placa = [
            cara["profundidad"]
            for cara in paneles
            if cara["area"] >= area_max_panel * 0.60
        ]
        if not planos_placa:
            return None

        # La cara de apoyo es la del lug coplanar con una de esas pieles. Si
        # dos caras están en el mismo plano, elegir la de mayor área: para este
        # modelo es el rectángulo de 8 in x 1 in mostrado en verde.
        for cara in lugs:
            cara["distancia_placa"] = min(
                abs(cara["profundidad"] - plano) for plano in planos_placa
            )
        distancia_minima = min(cara["distancia_placa"] for cara in lugs)
        tolerancia_contacto = max(0.08, distancia_minima + 0.04)
        candidatas = [
            cara
            for cara in lugs
            if cara["distancia_placa"] <= tolerancia_contacto
        ]
        contacto = max(candidatas, key=lambda cara: cara["area"])

        proyectados = []
        for x, y, z in contacto["puntos_modelo"]:
            punto = vista.ModelToSheetSpace(tg.CreatePoint(x, y, z))
            proyectados.append((float(punto.X), float(punto.Y)))
        minx = min(p[0] for p in proyectados)
        maxx = max(p[0] for p in proyectados)
        miny = min(p[1] for p in proyectados)
        maxy = max(p[1] for p in proyectados)
        dato = dict(datos_lug[0])
        dato.update(
            {
                "puntos_modelo": contacto["puntos_modelo"],
                "minx": minx,
                "maxx": maxx,
                "miny": miny,
                "maxy": maxy,
                "dx": maxx - minx,
                "dy": maxy - miny,
                "cx": (minx + maxx) * 0.5,
                "cy": (miny + maxy) * 0.5,
                "contacto": True,
                "eje_contacto": "Y" if maxy - miny >= maxx - minx else "X",
                "area_contacto": contacto["area"],
            }
        )
        return dato
    except Exception as error:
        log(f"    AVISO contacto lug: {error}")
        return None


def _reservar_stack_cotas(cantidad, limite_hoja):
    """Espacio de hoja para un stack H/V legible de la cantidad indicada."""
    requerido = OFFSET_COTA + 1.15 + max(0, cantidad - 1) * PASO_COTA_LEGIBLE
    return min(float(limite_hoja) * 0.62, max(3.80, requerido))


def _transformar_coordenada_vista(valor, centro_anterior, centro_nuevo, factor):
    return float(centro_nuevo) + (float(valor) - float(centro_anterior)) * factor


def _optimizar_espacio_cotas(
    hoja, vista, tg, todas, posiciones_x, posiciones_y, origen_x, origen_y
):
    """
    Amplía y desplaza la vista después de conocer cuántos extremos hay.
    La reserva izquierda/inferior crece con las cotas; lo restante se usa para
    la pieza. Así el JPG aprovecha la hoja incluso en tanques densos.
    """
    try:
        ancho_actual = float(vista.Width)
        alto_actual = float(vista.Height)
        escala_anterior = float(vista.Scale)
        cx_anterior = float(vista.Position.X)
        cy_anterior = float(vista.Position.Y)
        if ancho_actual <= EPS or alto_actual <= EPS or escala_anterior <= EPS:
            return origen_x, origen_y

        reserva_izq = _reservar_stack_cotas(len(posiciones_y), hoja.Width)
        reserva_inf = _reservar_stack_cotas(len(posiciones_x), hoja.Height)
        margen_der = max(0.70, hoja.Width * 0.025)
        margen_sup = max(0.70, hoja.Height * 0.025)
        ancho_disponible = max(1.0, float(hoja.Width) - reserva_izq - margen_der)
        alto_disponible = max(1.0, float(hoja.Height) - reserva_inf - margen_sup)
        factor = min(ancho_disponible / ancho_actual, alto_disponible / alto_actual) * 0.96
        factor = max(0.18, min(factor, 4.0))
        cx_nuevo = reserva_izq + ancho_disponible * 0.5
        cy_nuevo = reserva_inf + alto_disponible * 0.5

        vista.Scale = escala_anterior * factor
        vista.Position = tg.CreatePoint2d(cx_nuevo, cy_nuevo)

        vistos = set()
        for dato in todas:
            ident = id(dato)
            if ident in vistos:
                continue
            vistos.add(ident)
            dato["minx"] = _transformar_coordenada_vista(
                dato["minx"], cx_anterior, cx_nuevo, factor
            )
            dato["maxx"] = _transformar_coordenada_vista(
                dato["maxx"], cx_anterior, cx_nuevo, factor
            )
            dato["miny"] = _transformar_coordenada_vista(
                dato["miny"], cy_anterior, cy_nuevo, factor
            )
            dato["maxy"] = _transformar_coordenada_vista(
                dato["maxy"], cy_anterior, cy_nuevo, factor
            )
            dato["dx"] = abs(dato["maxx"] - dato["minx"])
            dato["dy"] = abs(dato["maxy"] - dato["miny"])
            dato["cx"] = (dato["minx"] + dato["maxx"]) * 0.5
            dato["cy"] = (dato["miny"] + dato["maxy"]) * 0.5

        for pos in posiciones_x:
            dato = pos["dato"]
            ident = id(dato)
            if ident not in vistos:
                vistos.add(ident)
                dato["minx"] = _transformar_coordenada_vista(
                    dato["minx"], cx_anterior, cx_nuevo, factor
                )
                dato["maxx"] = _transformar_coordenada_vista(
                    dato["maxx"], cx_anterior, cx_nuevo, factor
                )
                dato["miny"] = _transformar_coordenada_vista(
                    dato["miny"], cy_anterior, cy_nuevo, factor
                )
                dato["maxy"] = _transformar_coordenada_vista(
                    dato["maxy"], cy_anterior, cy_nuevo, factor
                )
                dato["dx"] = abs(dato["maxx"] - dato["minx"])
                dato["dy"] = abs(dato["maxy"] - dato["miny"])
                dato["cx"] = (dato["minx"] + dato["maxx"]) * 0.5
                dato["cy"] = (dato["miny"] + dato["maxy"]) * 0.5
            pos["valor"] = _transformar_coordenada_vista(
                pos["valor"], cx_anterior, cx_nuevo, factor
            )
        for pos in posiciones_y:
            dato = pos["dato"]
            ident = id(dato)
            if ident not in vistos:
                vistos.add(ident)
                dato["minx"] = _transformar_coordenada_vista(
                    dato["minx"], cx_anterior, cx_nuevo, factor
                )
                dato["maxx"] = _transformar_coordenada_vista(
                    dato["maxx"], cx_anterior, cx_nuevo, factor
                )
                dato["miny"] = _transformar_coordenada_vista(
                    dato["miny"], cy_anterior, cy_nuevo, factor
                )
                dato["maxy"] = _transformar_coordenada_vista(
                    dato["maxy"], cy_anterior, cy_nuevo, factor
                )
                dato["dx"] = abs(dato["maxx"] - dato["minx"])
                dato["dy"] = abs(dato["maxy"] - dato["miny"])
                dato["cx"] = (dato["minx"] + dato["maxx"]) * 0.5
                dato["cy"] = (dato["miny"] + dato["maxy"]) * 0.5
            pos["valor"] = _transformar_coordenada_vista(
                pos["valor"], cy_anterior, cy_nuevo, factor
            )

        origen_x = _transformar_coordenada_vista(
            origen_x, cx_anterior, cx_nuevo, factor
        )
        origen_y = _transformar_coordenada_vista(
            origen_y, cy_anterior, cy_nuevo, factor
        )
        log(
            f"    Encuadre adaptativo: X={len(posiciones_x)} Y={len(posiciones_y)} "
            f"reserva inf={reserva_inf:.2f} izq={reserva_izq:.2f}"
        )
    except Exception as error:
        log(f"    AVISO: no se pudo optimizar encuadre de cotas: {error}")
    return origen_x, origen_y


def _estado_base_vista(vista):
    return {
        "scale": float(vista.Scale),
        "x": float(vista.Position.X),
        "y": float(vista.Position.Y),
        "width": float(vista.Width),
        "height": float(vista.Height),
    }


def _dato_transformado_para_foto(dato, estado, nuevo_x, nuevo_y, factor):
    nuevo = dict(dato)
    nuevo["minx"] = _transformar_coordenada_vista(
        dato["minx"], estado["x"], nuevo_x, factor
    )
    nuevo["maxx"] = _transformar_coordenada_vista(
        dato["maxx"], estado["x"], nuevo_x, factor
    )
    nuevo["miny"] = _transformar_coordenada_vista(
        dato["miny"], estado["y"], nuevo_y, factor
    )
    nuevo["maxy"] = _transformar_coordenada_vista(
        dato["maxy"], estado["y"], nuevo_y, factor
    )
    nuevo["dx"] = abs(nuevo["maxx"] - nuevo["minx"])
    nuevo["dy"] = abs(nuevo["maxy"] - nuevo["miny"])
    nuevo["cx"] = (nuevo["minx"] + nuevo["maxx"]) * 0.5
    nuevo["cy"] = (nuevo["miny"] + nuevo["maxy"]) * 0.5
    return nuevo


def _preparar_encuadre_estable_cara(plano, hoja, vista, tg, estado):
    """
    Un solo reframe por cara, con margen para cotas X e Y.

    Evita regenerar HLR al cambiar Scale/Position en cada foto (~55 s/JPG).
    """
    reserva_izq = 4.80
    reserva_inf = 4.80
    margen_der = max(0.70, hoja.Width * 0.025)
    margen_sup = max(0.70, hoja.Height * 0.025)
    ancho_disponible = max(1.0, float(hoja.Width) - reserva_izq - margen_der)
    alto_disponible = max(1.0, float(hoja.Height) - reserva_inf - margen_sup)
    factor = min(
        ancho_disponible / max(EPS, estado["width"]),
        alto_disponible / max(EPS, estado["height"]),
    ) * FACTOR_ENCUADRE_FOTO
    factor = max(0.18, min(factor, 5.0))
    nuevo_x = reserva_izq + ancho_disponible * 0.5
    nuevo_y = reserva_inf + alto_disponible * 0.5
    vista.Scale = estado["scale"] * factor
    vista.Position = tg.CreatePoint2d(nuevo_x, nuevo_y)
    try:
        plano.Update()
    except Exception:
        pass
    log(
        f"    Encuadre estable de cara: factor={factor:.3f} "
        f"pos=({nuevo_x:.2f},{nuevo_y:.2f})"
    )


def _refrescar_referencias_exportacion(referencias, dato_panel, vista, tg):
    """
    Relee HLR una sola vez tras el encuadre estable de la cara.
    Las fotos posteriores reutilizan estas coordenadas de hoja.
    """
    envolvente_panel = _envolvente_visible_refrescada(dato_panel, vista, tg)
    if envolvente_panel is None:
        raise RuntimeError(
            "No se pudo refrescar la placa madre tras el encuadre estable."
        )
    origen_x = float(envolvente_panel["minx"])
    origen_y = float(envolvente_panel["miny"])

    refresadas = []
    omitidas = 0
    for eje, referencia in referencias:
        miembros_actuales = _refrescar_miembros_referencia(
            eje, referencia, vista, tg
        )
        if not miembros_actuales:
            omitidas += 1
            continue
        representativa = next(
            (
                miembro
                for miembro in miembros_actuales
                if miembro["pieza_id"] == referencia["pieza_id"]
                and miembro["lado"] == referencia["lado"]
            ),
            miembros_actuales[0],
        )
        referencia_nueva = dict(referencia)
        referencia_nueva["dato"] = representativa["dato"]
        referencia_nueva["valor"] = representativa["valor"]
        referencia_nueva["miembros"] = miembros_actuales
        refresadas.append((eje, referencia_nueva))
    if omitidas:
        log(f"    AVISO: {omitidas} referencias sin HLR tras el encuadre.")
    return refresadas, origen_x, origen_y


def _preparar_foto_de_referencia(
    plano,
    hoja,
    vista,
    tg,
    estado,
    eje,
    referencia,
    dato_panel,
):
    """
    Compat: reframe puntual. El export de producción usa encuadre estable.
    """
    reserva_izq = 1.20 if eje == "X" else 4.80
    reserva_inf = 4.80 if eje == "X" else 1.20
    margen_der = max(0.70, hoja.Width * 0.025)
    margen_sup = max(0.70, hoja.Height * 0.025)
    ancho_disponible = max(1.0, float(hoja.Width) - reserva_izq - margen_der)
    alto_disponible = max(1.0, float(hoja.Height) - reserva_inf - margen_sup)
    factor = min(
        ancho_disponible / estado["width"],
        alto_disponible / estado["height"],
    ) * FACTOR_ENCUADRE_FOTO
    factor = max(0.18, min(factor, 5.0))
    nuevo_x = reserva_izq + ancho_disponible * 0.5
    nuevo_y = reserva_inf + alto_disponible * 0.5

    vista.Scale = estado["scale"] * factor
    vista.Position = tg.CreatePoint2d(nuevo_x, nuevo_y)

    try:
        plano.Update()
    except Exception:
        pass

    envolvente_panel = _envolvente_visible_refrescada(dato_panel, vista, tg)
    miembros_actuales = _refrescar_miembros_referencia(
        eje, referencia, vista, tg
    )
    if envolvente_panel is None or not miembros_actuales:
        raise RuntimeError(
            "No se pudo refrescar la placa o el accesorio en la vista individual."
        )

    representativa = next(
        (
            miembro
            for miembro in miembros_actuales
            if miembro["pieza_id"] == referencia["pieza_id"]
            and miembro["lado"] == referencia["lado"]
        ),
        miembros_actuales[0],
    )
    referencia_nueva = dict(referencia)
    referencia_nueva["dato"] = representativa["dato"]
    referencia_nueva["valor"] = representativa["valor"]
    referencia_nueva["miembros"] = miembros_actuales
    origen_x = float(envolvente_panel["minx"])
    origen_y = float(envolvente_panel["miny"])
    return referencia_nueva, origen_x, origen_y


def _borrar_cotas_hoja(hoja):
    try:
        dims = hoja.DrawingDimensions
        for indice in range(dims.Count, 0, -1):
            try:
                dims.Item(indice).Delete()
            except Exception:
                pass
    except Exception:
        pass


def _borrar_sketches_cotas(hoja):
    """Borra solo la geometría gráfica de cotas creada por esta regla."""
    try:
        sketches = hoja.Sketches
        for indice in range(sketches.Count, 0, -1):
            try:
                sketch = sketches.Item(indice)
                if str(sketch.Name).upper().startswith(PREFIJO_SKETCH_COTAS):
                    sketch.Delete()
            except Exception:
                pass
    except Exception:
        pass


def _punto_sketch_desde_hoja(sketch, tg, x, y):
    """Convierte coordenadas de hoja a la geometría del DrawingSketch."""
    punto_hoja = tg.CreatePoint2d(float(x), float(y))
    try:
        return sketch.SheetToSketchSpace(punto_hoja)
    except Exception:
        # En versiones donde el sketch comparte coordenadas de hoja.
        return punto_hoja


def _linea_cota_sketch(sketch, tg, x1, y1, x2, y2, color=None):
    try:
        linea = sketch.SketchLines.AddByTwoPoints(
            _punto_sketch_desde_hoja(sketch, tg, x1, y1),
            _punto_sketch_desde_hoja(sketch, tg, x2, y2),
        )
        if color is not None:
            try:
                linea.Color = color
            except Exception:
                pass
        return linea
    except Exception:
        return None


def _texto_cota_sketch(sketch, tg, x, y, texto, inv_app):
    try:
        caja = sketch.TextBoxes.AddFitted(
            _punto_sketch_desde_hoja(sketch, tg, x, y), str(texto)
        )
        aplicar_estilo_texto_cota(caja, texto, inv_app)
        return caja
    except Exception:
        return None


def _circulo_cota_sketch(sketch, tg, x, y, radio, color=None):
    try:
        circulo = sketch.SketchCircles.AddByCenterRadius(
            _punto_sketch_desde_hoja(sketch, tg, x, y),
            float(radio),
        )
        if color is not None:
            try:
                circulo.Color = color
            except Exception:
                pass
        return circulo
    except Exception:
        return None


def _marcas_typ_en_accesorios(sketch, tg, eje, referencia, color):
    """
    Marca con puntos azules las aristas/centros de todos los accesorios a los
    que aplica la única cota TYP.
    """
    if not referencia.get("typ"):
        return
    vistos = set()
    for miembro in referencia.get("miembros", []):
        dato = miembro.get("dato")
        if dato is None:
            continue
        if eje == "X":
            x = float(miembro["valor"])
            y = float(dato["cy"])
        else:
            x = float(dato["cx"])
            y = float(miembro["valor"])
        clave = (round(x, 4), round(y, 4))
        if clave in vistos:
            continue
        vistos.add(clave)
        _circulo_cota_sketch(sketch, tg, x, y, 0.16, color)


def _flechas_horizontales(sketch, tg, x1, x2, y, color):
    """Dos flechas V simples sobre una línea de cota horizontal."""
    signo = 1.0 if x2 >= x1 else -1.0
    a = TAM_FLECHA_COTA
    _linea_cota_sketch(sketch, tg, x1, y, x1 + signo * a, y + a * 0.55, color)
    _linea_cota_sketch(sketch, tg, x1, y, x1 + signo * a, y - a * 0.55, color)
    _linea_cota_sketch(sketch, tg, x2, y, x2 - signo * a, y + a * 0.55, color)
    _linea_cota_sketch(sketch, tg, x2, y, x2 - signo * a, y - a * 0.55, color)


def _flechas_verticales(sketch, tg, x, y1, y2, color):
    """Dos flechas V simples sobre una línea de cota vertical."""
    signo = 1.0 if y2 >= y1 else -1.0
    a = TAM_FLECHA_COTA
    _linea_cota_sketch(sketch, tg, x, y1, x + a * 0.55, y1 + signo * a, color)
    _linea_cota_sketch(sketch, tg, x, y1, x - a * 0.55, y1 + signo * a, color)
    _linea_cota_sketch(sketch, tg, x, y2, x + a * 0.55, y2 - signo * a, color)
    _linea_cota_sketch(sketch, tg, x, y2, x - a * 0.55, y2 - signo * a, color)


def _valor_real_desde_hoja(vista, inicio, final, hoja):
    """
    Vista ortográfica: distancia de hoja / escala = distancia real del modelo.
    Inventor guarda longitudes internas en cm; texto_cota_limpio aplica las
    unidades activas del machote sin poner el sufijo.
    """
    try:
        escala = abs(float(vista.Scale))
        if escala <= EPS:
            return ""
        return texto_cota_limpio(abs(float(final) - float(inicio)) / escala, hoja)
    except Exception:
        return ""


def _dibujar_cotas_hv_desde_origen(
    hoja, vista, tg, inv_app, posiciones_x, posiciones_y, origen_x, origen_y
):
    """
    Cotas gráficas H/V con líneas, extensiones, flechas y valor real.

    Se usan porque Dimension API de Inventor produjo alternativas defectuosas:
    AddLinear -> alineadas diagonales; Ordinate -> números sin líneas.
    Las vistas son ortográficas y se regeneran en cada corrida, por lo que las
    posiciones y valores se obtienen de su escala real en ese instante.
    """
    _borrar_sketches_cotas(hoja)
    try:
        sketch = hoja.Sketches.Add()
        sketch.Name = PREFIJO_SKETCH_COTAS + "HV"
        sketch.Edit()
    except Exception as error:
        log(f"    ERROR creando sketch de cotas H/V: {error}")
        return 0, len(posiciones_x) + len(posiciones_y)

    color = None
    try:
        color = inv_app.TransientObjects.CreateColor(0, 0, 128)
    except Exception:
        pass

    creadas = 0
    fallos = 0
    try:
        base_abajo = float(vista.Top) - float(vista.Height) - OFFSET_COTA
        base_izq = float(vista.Left) - OFFSET_COTA
        # Encoger el paso si una cara tiene muchas posiciones, para que la
        # última cota siempre quede dentro de la hoja.
        paso_x = PASO_NIVEL_COTA
        paso_y = PASO_NIVEL_COTA
        if len(posiciones_x) > 1:
            paso_x = max(0.36, min(
                PASO_NIVEL_COTA,
                max(0.36, (base_abajo - 1.10) / (len(posiciones_x) - 1)),
            ))
        if len(posiciones_y) > 1:
            paso_y = max(0.36, min(
                PASO_NIVEL_COTA,
                max(0.36, (base_izq - 1.10) / (len(posiciones_y) - 1)),
            ))

        for indice, pos in enumerate(posiciones_x):
            nivel = indice
            x = float(pos["valor"])
            y_dim = base_abajo - nivel * paso_x
            dato = pos["dato"]
            y_obj = (float(dato["miny"]) + float(dato["maxy"])) * 0.5
            valor = _valor_real_desde_hoja(vista, origen_x, x, hoja)
            if not valor:
                fallos += 1
                continue

            ok = _linea_cota_sketch(sketch, tg, origen_x, y_dim, x, y_dim, color)
            ok = _linea_cota_sketch(sketch, tg, origen_x, origen_y, origen_x, y_dim, color) and ok
            ok = _linea_cota_sketch(sketch, tg, x, y_obj, x, y_dim, color) and ok
            _flechas_horizontales(sketch, tg, origen_x, x, y_dim, color)
            _marcas_typ_en_accesorios(sketch, tg, "X", pos, color)
            texto_valor = f"{valor} TYP" if pos.get("typ") else valor
            texto = _texto_cota_sketch(
                sketch,
                tg,
                (origen_x + x) * 0.5,
                y_dim + 0.12,
                texto_valor,
                inv_app,
            )
            if ok is not None and texto is not None:
                creadas += 1
            else:
                fallos += 1

        for indice, pos in enumerate(posiciones_y):
            nivel = indice
            y = float(pos["valor"])
            x_dim = base_izq - nivel * paso_y
            dato = pos["dato"]
            x_obj = (float(dato["minx"]) + float(dato["maxx"])) * 0.5
            valor = _valor_real_desde_hoja(vista, origen_y, y, hoja)
            if not valor:
                fallos += 1
                continue

            ok = _linea_cota_sketch(sketch, tg, x_dim, origen_y, x_dim, y, color)
            ok = _linea_cota_sketch(sketch, tg, origen_x, origen_y, x_dim, origen_y, color) and ok
            ok = _linea_cota_sketch(sketch, tg, x_obj, y, x_dim, y, color) and ok
            _flechas_verticales(sketch, tg, x_dim, origen_y, y, color)
            _marcas_typ_en_accesorios(sketch, tg, "Y", pos, color)
            texto_valor = f"{valor} TYP" if pos.get("typ") else valor
            texto = _texto_cota_sketch(
                sketch,
                tg,
                x_dim + 0.12,
                (origen_y + y) * 0.5,
                texto_valor,
                inv_app,
            )
            if ok is not None and texto is not None:
                creadas += 1
            else:
                fallos += 1
    finally:
        try:
            sketch.ExitEdit()
        except Exception:
            pass

    return creadas, fallos


def _punto_curva_mas_cercano(todas, x, y, tg):
    """Punto visible mas cercano a (x,y) en hoja — para OriginIndicator."""
    mejor = None
    mejor_d2 = None
    for dato in todas:
        for p in _puntos_clave(dato["curve"]):
            try:
                dx = float(p.X) - x
                dy = float(p.Y) - y
                d2 = dx * dx + dy * dy
            except Exception:
                continue
            if mejor_d2 is None or d2 < mejor_d2:
                mejor_d2 = d2
                mejor = (dato["curve"], p)
    if mejor is None:
        return None
    return mejor


def _fijar_origen_indicador(vista, hoja, tg, todas, origen_x, origen_y, ref_izq, ref_inf):
    """
    Origen (0,0) en esquina inferior-izquierda del cuerpo.
    Requiere punto REAL de curva (CreateGeometryIntent con Point2d suelto falla).
    """
    hallado = _punto_curva_mas_cercano(todas, origen_x, origen_y, tg)
    intent = None
    if hallado is not None:
        try:
            intent = hoja.CreateGeometryIntent(hallado[0], hallado[1])
        except Exception:
            intent = None

    if intent is None:
        intent = _intent_en_coordenada(hoja, tg, ref_inf, origen_x, origen_y)
    if intent is None:
        intent = _intent_seguro(hoja, ref_inf, "inf")
    if intent is None:
        intent = _intent_seguro(hoja, ref_izq, "izq")
    if intent is None:
        return False

    try:
        if vista.HasOriginIndicator:
            vista.OriginIndicator.Intent = intent
        else:
            vista.CreateOriginIndicator(intent)
    except Exception as error:
        log(f"    AVISO OriginIndicator: {error}")
        try:
            intent2 = _intent_seguro(hoja, ref_inf, "inf")
            if intent2 is None:
                return False
            if vista.HasOriginIndicator:
                vista.OriginIndicator.Intent = intent2
            else:
                vista.CreateOriginIndicator(intent2)
        except Exception as error2:
            log(f"    ERROR OriginIndicator: {error2}")
            return False

    # Dejarlo visible: marca el (0,0) de soldadura.
    try:
        vista.OriginIndicator.Visible = True
    except Exception:
        pass
    return True


def _crear_cota_ordenada_desde_origen(hoja, tg, dato_obj, lado, tipo, texto_xy):
    """
    Cota baseline H/V desde OriginIndicator.

    OrdinateDimensions no admite tipo alineado: elimina de raíz el abanico
    diagonal que AddLinear estaba creando en algunas caras.
    """
    intent_obj = _intent_seguro(hoja, dato_obj, lado)
    if intent_obj is None:
        return False

    try:
        dim = hoja.DrawingDimensions.OrdinateDimensions.Add(
            intent_obj,
            tg.CreatePoint2d(float(texto_xy[0]), float(texto_xy[1])),
            int(tipo),
        )
    except Exception as error:
        log(f"    AVISO cota H/V no creada: {error}")
        return False

    # No usar HideValue/FormatedText: eso dejó los números sueltos.
    aplicar_estilo_cota(dim, hoja=hoja, solo_color=True)
    return True


def _acotar_vista(
    hoja,
    vista,
    tg,
    inv_app,
    cara,
    face,
    right,
    bbox,
    catalogo=None,
    nombre_seg="",
    analizar_solo_segmento=False,
):
    """
    Cotas H/V desde (0,0) a cada pieza independiente sobre la pared física.
    El segmento mapeado es contexto; la selección definitiva es geométrica.
    """
    if catalogo is None:
        catalogo = set()

    grupos, todas = _curvas_por_componente(vista)
    if not todas:
        log("    Sin curvas visibles para acotar.")
        return None

    todas, _bbox = _filtrar_curvas_cara(todas, vista)
    grupos = {}
    for info in todas:
        nombre = _nombre_componente(info["curve"])
        grupos.setdefault(nombre, []).append(info)

    if nombre_seg:
        log(f"    Contenedor mapeado: {nombre_seg} ({len(catalogo)} nombres)")
    else:
        log("    AVISO: sin contenedor mapeado; filtro físico de cara")

    _borrar_cotas_hoja(hoja)
    _borrar_sketches_cotas(hoja)
    _limpiar_bolitas(hoja, vista)

    posiciones_x = []
    posiciones_y = []
    tol = max(0.05, max(vista.Width, vista.Height) * TOL_EXTREMO_RATIO)
    escala_hoja = max(EPS, abs(float(vista.Scale)))
    min_dist_origen_hoja = MIN_DIST_ORIGEN_CM * escala_hoja
    piezas_superficie = 0
    sin_componente = 0
    fuera_de_cara = 0
    nombres_ok = []
    envolventes_3d = 0
    grupos_en_cara = []
    for nombre, datos in grupos.items():
        if not datos:
            continue
        nombre_pieza = _nombre_final_componente(nombre)
        # Un barreno es una curva de la placa, no una ocurrencia/pieza. Las
        # curvas sin ocurrencia nunca son candidatas de fabricación.
        if _nombre_base_pieza(nombre_pieza) == "SIN_COMPONENTE":
            sin_componente += 1
            continue
        pertenece_al_segmento = _nombre_en_catalogo(nombre_pieza, catalogo)
        if (
            not analizar_solo_segmento
            and not pertenece_al_segmento
            and not _datos_son_de_cara_fisica(datos, cara, face, right, bbox)
        ):
            fuera_de_cara += 1
            continue
        grupos_en_cara.append((nombre, nombre_pieza, datos))

    # Se excluye solo la placa madre. Para identificarla se priorizan las
    # ocurrencias del segmento mapeado y, entre ellas, la mayor huella. Esto
    # evita confundir un CBOX/bracket grande con la pared; las piezas fuera
    # del subensamble también se mantienen y se acotan si viven en la cara.
    paneles_segmento = [
        (nombre, datos, _huella_grupo_en_vista(datos, vista)[2])
        for nombre, nombre_pieza, datos in grupos_en_cara
        if _nombre_en_catalogo(nombre_pieza, catalogo)
    ]
    paneles = paneles_segmento or [
        (nombre, datos, _huella_grupo_en_vista(datos, vista)[2])
        for nombre, _nombre_pieza, datos in grupos_en_cara
    ]
    nombre_panel = max(paneles, key=lambda item: item[2])[0] if paneles else None
    if nombre_panel:
        log(f"    Placa madre excluida: {_nombre_final_componente(nombre_panel)}")
    else:
        log("    AVISO: no se identificó placa madre; se incluirán todas las piezas.")

    datos_panel = next(
        (datos for nombre, _nombre_pieza, datos in grupos_en_cara if nombre == nombre_panel),
        [],
    )
    envolvente_panel = _envolvente_occurrence_en_hoja(datos_panel, vista, tg)
    if envolvente_panel is not None:
        # Datum de piso: esquina inferior-izquierda de la placa madre del
        # segmento, no de la base exterior ni de las soleras.
        origen_x = float(envolvente_panel["minx"])
        origen_y = float(envolvente_panel["miny"])
        log(
            f"    Origen (0,0) placa segmento: X={origen_x:.3f} "
            f"Y={origen_y:.3f}"
        )
    else:
        ref_izq = _elegir_muro_datum(todas, "izq", vista)
        ref_inf = _elegir_muro_datum(todas, "inf", vista)
        if ref_izq is None or ref_inf is None:
            log("    No se pudo fijar el origen (0,0).")
            return None
        origen_x = float(ref_izq["minx"])
        origen_y = float(ref_inf["miny"])
        log("    AVISO: datum de respaldo por muro; placa madre no proyectable.")

    for nombre, nombre_pieza, datos in grupos_en_cara:
        if nombre == nombre_panel:
            continue
        piezas_superficie += 1
        nombres_ok.append(nombre_pieza[:40])
        es_lug = _es_nombre_lug(nombre_pieza)
        tol_pos = (0.01 if es_lug else TOLERANCIA_COTA_CM) * escala_hoja
        envolvente = _envolvente_occurrence_en_hoja(datos, vista, tg)
        datos_extremos = [envolvente] if envolvente is not None else datos
        if envolvente is not None:
            envolventes_3d += 1
        if _es_componente_circular(datos, envolvente, nombre_pieza):
            _agregar_posicion(
                posiciones_x,
                envolvente["cx"],
                envolvente,
                tol_pos,
                "centro",
                nombre,
            )
            _agregar_posicion(
                posiciones_y,
                envolvente["cy"],
                envolvente,
                tol_pos,
                "centro",
                nombre,
            )
        else:
            _agregar_inicio_y_fin(
                posiciones_x, posiciones_y, datos_extremos, tol, tol_pos, nombre
            )
            # Además de la envolvente exterior, un lifting lug necesita las
            # referencias del tramo que realmente toca/suelda sobre la placa.
            contacto_lug = (
                _cara_contacto_lug_en_hoja(
                    nombre, nombre_panel, datos, vista, tg
                )
                if es_lug
                else None
            )
            if contacto_lug is not None:
                if contacto_lug["eje_contacto"] == "Y":
                    _agregar_posicion(
                        posiciones_y,
                        contacto_lug["miny"],
                        contacto_lug,
                        tol_pos,
                        "inf",
                        nombre,
                        contacto=True,
                    )
                    _agregar_posicion(
                        posiciones_y,
                        contacto_lug["maxy"],
                        contacto_lug,
                        tol_pos,
                        "sup",
                        nombre,
                        contacto=True,
                    )
                else:
                    _agregar_posicion(
                        posiciones_x,
                        contacto_lug["minx"],
                        contacto_lug,
                        tol_pos,
                        "izq",
                        nombre,
                        contacto=True,
                    )
                    _agregar_posicion(
                        posiciones_x,
                        contacto_lug["maxx"],
                        contacto_lug,
                        tol_pos,
                        "der",
                        nombre,
                        contacto=True,
                    )
                log(
                    "    Lug con cara de contacto a placa: "
                    f"{contacto_lug['eje_contacto']} "
                    f"{contacto_lug['dx']:.2f} x {contacto_lug['dy']:.2f} "
                    f"cm hoja, área={contacto_lug['area_contacto']:.2f} cm²"
                )
            elif es_lug:
                log("    AVISO: no se detectó línea de contacto del lifting lug.")

    posiciones_x = [
        p for p in posiciones_x
        if abs(p["valor"] - origen_x) >= min_dist_origen_hoja
    ]
    posiciones_y = [
        p for p in posiciones_y
        if abs(p["valor"] - origen_y) >= min_dist_origen_hoja
    ]
    # La referencia TYP se decide contra la misma silueta HLR que se usará
    # para dibujar cada foto, no contra su RangeBox tridimensional.
    envolvente_panel_typ = _envolvente_visible_refrescada(
        envolvente_panel, vista, tg
    )
    origen_x_typ = (
        float(envolvente_panel_typ["minx"])
        if envolvente_panel_typ is not None
        else origen_x
    )
    origen_y_typ = (
        float(envolvente_panel_typ["miny"])
        if envolvente_panel_typ is not None
        else origen_y
    )
    posiciones_x = _refrescar_posiciones_para_typ(
        "X", posiciones_x, vista, tg
    )
    posiciones_y = _refrescar_posiciones_para_typ(
        "Y", posiciones_y, vista, tg
    )
    posiciones_x.sort(key=lambda p: p["valor"])
    posiciones_y.sort(key=lambda p: p["valor"])
    tolerancia_typ = max(
        TOL_COINCIDENCIA_HOJA,
        TOLERANCIA_COTA_CM * escala_hoja,
    )
    posiciones_x = _agrupar_referencias_typ(
        posiciones_x,
        tolerancia_typ,
        eje="X",
        origen=origen_x_typ,
        vista=vista,
        hoja=hoja,
    )
    posiciones_y = _agrupar_referencias_typ(
        posiciones_y,
        tolerancia_typ,
        eje="Y",
        origen=origen_y_typ,
        vista=vista,
        hoja=hoja,
    )
    referencias_typ = sum(
        1 for posicion in posiciones_x + posiciones_y if posicion["typ"]
    )

    log(
        f"    Piezas de superficie={piezas_superficie} | "
        f"envolventes 3D={envolventes_3d} sin ocurrencia={sin_componente} "
        f"otra cara={fuera_de_cara} | "
        f"cotas X={len(posiciones_x)} Y={len(posiciones_y)} TYP={referencias_typ}"
    )
    if nombres_ok:
        log(f"    Piezas: {', '.join(nombres_ok[:16])}")

    _limpiar_bolitas(hoja, vista)

    return {
        "origen_x": origen_x,
        "origen_y": origen_y,
        "dato_panel": envolvente_panel or {"curvas": list(datos_panel)},
        "posiciones_x": posiciones_x,
        "posiciones_y": posiciones_y,
        "piezas_superficie": piezas_superficie,
    }


def _crear_caras(inv_app, plano, ensamble):
    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    _eliminar_hojas_piezas_residuales(plano)
    hoja_base = _elegir_hoja_base(plano)
    _eliminar_hojas_anteriores(plano)

    bbox_tanque = _bbox_ensamble(ensamble)
    log("Orientacion tipo PQart (tapa=+Y, cara=+Z, lateral=+X)...")
    cover, face, right = marco_como_pqart(ensamble, bbox_tanque, log)
    log(
        f"  Camara foto (cara real): face=({face[0]:.3f},{face[1]:.3f},{face[2]:.3f}) "
        f"right=({right[0]:.3f},{right[1]:.3f},{right[2]:.3f}) "
        f"cover=({cover[0]:.3f},{cover[1]:.3f},{cover[2]:.3f})"
    )

    log("Detectando el contenedor estructural de cuatro paredes...")
    alcance = _resolver_contenedor_paredes(
        ensamble,
        face=face,
        right=right,
        cover=cover,
        bbox=bbox_tanque,
    )
    if not alcance.get("valido"):
        raise RuntimeError(
            "No se detectaron cuatro paredes reales: "
            + str(alcance.get("motivo", "sin detalle"))
        )
    bbox_caras = alcance.get("bbox_caras") or bbox_tanque
    log("Detectando segmentos/contenedores de cara en el alcance resuelto...")
    segmentos = _detectar_segmentos_en_ensamble(ensamble, inv_app, alcance)

    log("Mapeo segmento -> cara (por centroide)...")
    mapa_caras = _mapear_segmentos_a_caras(
        segmentos, face, right, bbox_caras
    )
    caras = _direcciones_caras_pqart(tg, cover, face, right)
    rutas_segmentos = {
        _ruta_occurrence(seg["occurrence"])
        for seg in mapa_caras.values()
        if seg.get("occurrence") is not None
    }

    creadas = []
    planes_cotas = {}
    for nombre, eye_dir, up_hint in caras:
        nombre_hoja = PREFIJO_HOJA + nombre
        log(f"Paso cara {nombre} (camara PQart de pared)...")

        seg = mapa_caras.get(nombre)
        catalogo = set(seg["piezas"]) if seg else set()
        nombre_seg = seg["nombre"] if seg else ""
        ensamble_segmento = seg.get("ensamble_segmento") if seg else None
        ocurrencia_segmento = seg.get("occurrence") if seg else None

        try:
            if ensamble_segmento is None or ocurrencia_segmento is None:
                raise RuntimeError(
                    "No se encontró el IAM del segmento desde el tanque principal."
                )
            hoja = _crear_hoja(plano, hoja_base, nombre_hoja)
            hoja.Activate()

            centro = tg.CreatePoint2d(hoja.Width * 0.58, hoja.Height * 0.55)
            extras_raiz = _ocurrencias_raiz_en_cara(
                ensamble,
                ocurrencia_segmento,
                nombre,
                face,
                right,
                bbox_caras,
                rutas_segmentos=rutas_segmentos,
                cover=cover,
            )
            if extras_raiz:
                log(
                    "    Accesorios externos incluidos: "
                    + ", ".join(str(occ.Name) for occ in extras_raiz)
                )
            # El tanque principal aporta la geometría completa. La visibilidad
            # se limita enseguida al segmento y sus accesorios raíz de esta
            # cara, manteniendo el (0,0) de la placa del segmento.
            cam = _crear_camara(
                ensamble,
                tg,
                to,
                bbox_tanque,
                eye_dir,
                up_hint,
            )
            log(
                f"    Fuente de vista: tanque + {ensamble_segmento.DisplayName}"
            )

            vista = hoja.DrawingViews.AddBaseView(
                ensamble,
                centro,
                1.0,
                K_ARBITRARY,
                ESTILO_HLR,
                "",
                cam,
            )
            vista.Name = "TANQUE_" + nombre
            plano.Update()
            visibles = _aplicar_visibilidad_vista_cara(
                vista,
                ensamble,
                ocurrencia_segmento,
                extras_raiz,
            )
            log(f"    Vista aislada: {visibles} ocurrencias visibles")
            plano.Update()
            _enderezar_vista_en_hoja(vista)
            plano.Update()
            _ajustar_vista(vista, hoja, tg)
            plano.Update()
            _limpiar_bolitas(hoja, vista)

            plan_cotas = _acotar_vista(
                hoja,
                vista,
                tg,
                inv_app,
                nombre,
                face,
                right,
                bbox_caras,
                catalogo,
                nombre_seg,
                analizar_solo_segmento=False,
            )
            if plan_cotas is None:
                raise RuntimeError("No se pudieron preparar las referencias de la cara.")
            componentes = plan_cotas["piezas_superficie"]
            cotas = len(plan_cotas["posiciones_x"]) + len(plan_cotas["posiciones_y"])
            _limpiar_bolitas(hoja, vista)
            try:
                if vista.HasOriginIndicator:
                    vista.OriginIndicator.Visible = False
            except Exception:
                pass
            plano.Update()
            _limpiar_bolitas(hoja, vista)
            try:
                if vista.HasOriginIndicator:
                    vista.OriginIndicator.Visible = False
            except Exception:
                pass
            log(
                f"  OK {nombre_hoja}: vista lista "
                f"({componentes} piezas, {cotas} referencias)"
            )
            creadas.append(hoja)
            planes_cotas[nombre] = {
                "hoja": hoja,
                "vista": vista,
                **plan_cotas,
            }
            _actualizar_inventor(inv_app)
            time.sleep(0.25)
        except Exception as error:
            log(f"  ERROR en {nombre}: {error}")
            log(traceback.format_exc())

    return creadas, planes_cotas


def _bbox_contenido_hoja(hoja, vista_objetivo=None):
    """
    Bbox de vista + cotas nativas + sketch de cotas H/V.
    Sin esto el recorte deja solo geometria y las cotas desaparecen del JPG.
    """
    # El marco/título del machote no forma parte de la fotografía. La caja de
    # DrawingView incluye el área de hoja alrededor del dibujo y puede abarcar
    # el marco; partir de las curvas HLR deja únicamente el tanque real.
    bbox = None
    if vista_objetivo is not None:
        try:
            minx = float(vista_objetivo.Left)
            maxx = minx + float(vista_objetivo.Width)
            maxy = float(vista_objetivo.Top)
            miny = maxy - float(vista_objetivo.Height)
            bbox = [minx, maxx, miny, maxy]
        except Exception:
            bbox = None
    elif bbox is None:
        try:
            vistas = []
            for i in range(1, hoja.DrawingViews.Count + 1):
                vista = hoja.DrawingViews.Item(i)
                try:
                    if str(vista.Name).upper().startswith("TANQUE_"):
                        vistas.append(vista)
                except Exception:
                    continue
            # En una hoja copiada puede existir una vista heredada del machote
            # (marco/cartucho). Para el JPG solo cuenta la vista creada aquí.
            if not vistas:
                vistas = [
                    hoja.DrawingViews.Item(i)
                    for i in range(1, hoja.DrawingViews.Count + 1)
                ]
            for vista in vistas:
                try:
                    curvas = vista.DrawingCurves
                    for j in range(1, curvas.Count + 1):
                        curva = curvas.Item(j)
                        # Curvas del formato DWG/marco no pertenecen a una
                        # ocurrencia del IAM. Excluirlas antes de calcular el
                        # encuadre de la fotografía.
                        if _nombre_componente(curva) == "SIN_COMPONENTE":
                            continue
                        dato = _info_curva(curva)
                        if dato is None:
                            continue
                        minx = float(dato["minx"])
                        maxx = float(dato["maxx"])
                        miny = float(dato["miny"])
                        maxy = float(dato["maxy"])
                        if bbox is None:
                            bbox = [minx, maxx, miny, maxy]
                        else:
                            bbox[0] = min(bbox[0], minx)
                            bbox[1] = max(bbox[1], maxx)
                            bbox[2] = min(bbox[2], miny)
                            bbox[3] = max(bbox[3], maxy)
                except Exception:
                        continue
        except Exception:
            pass
    if bbox is None:
        bbox = _obtener_bbox_hoja(hoja)
    n_ord = 0
    n_sketch = 0

    # Sketch explícito de cotas H/V (líneas, flechas, textos).
    try:
        sketches = hoja.Sketches
        for i in range(1, sketches.Count + 1):
            try:
                sketch = sketches.Item(i)
                if not str(sketch.Name).upper().startswith(PREFIJO_SKETCH_COTAS):
                    continue
                puntos = []
                # RangeBox de DrawingSketch no es confiable por COM; leer
                # las puntas reales de cada línea es estable.
                try:
                    lines = sketch.SketchLines
                    for j in range(1, lines.Count + 1):
                        linea = lines.Item(j)
                        geom = linea.Geometry
                        puntos.extend([geom.StartPoint, geom.EndPoint])
                except Exception:
                    pass
                if puntos:
                    xs = []
                    ys = []
                    for punto in puntos:
                        try:
                            p = sketch.SketchToSheetSpace(punto)
                        except Exception:
                            p = punto
                        xs.append(float(p.X))
                        ys.append(float(p.Y))
                    minx, maxx = min(xs), max(xs)
                    miny, maxy = min(ys), max(ys)
                    if bbox is None:
                        bbox = [minx, maxx, miny, maxy]
                    else:
                        bbox[0] = min(bbox[0], minx)
                        bbox[1] = max(bbox[1], maxx)
                        bbox[2] = min(bbox[2], miny)
                        bbox[3] = max(bbox[3], maxy)
                try:
                    n_sketch += max(1, int(sketch.TextBoxes.Count) - 1)
                except Exception:
                    n_sketch += 1
            except Exception:
                pass
    except Exception:
        pass

    # El machote puede contener cotas nativas del formato; no son parte de la
    # imagen de fabricación y nunca deben ampliar el recorte.
    return bbox, n_sketch


def _bbox_foto_referencia(vista, eje, referencia, origen_x, origen_y):
    """
    Caja de una foto individual calculada desde la vista y la única cota
    dibujada. No se lee el RangeBox del DrawingSketch: en DWG ese RangeBox
    puede incluir el marco completo del machote.
    """
    minx = float(vista.Left)
    maxx = minx + float(vista.Width)
    maxy = float(vista.Top)
    miny = maxy - float(vista.Height)
    dato = referencia["dato"]

    if eje == "X":
        x = float(referencia["valor"])
        y_dim = miny - OFFSET_COTA
        puntos = [
            (origen_x, y_dim),
            (x, y_dim),
            (origen_x, origen_y),
            (x, float(dato["cy"])),
        ]
    else:
        y = float(referencia["valor"])
        x_dim = minx - OFFSET_COTA
        puntos = [
            (x_dim, origen_y),
            (x_dim, y),
            (origen_x, origen_y),
            (float(dato["cx"]), y),
        ]

    for x, y in puntos:
        minx = min(minx, float(x))
        maxx = max(maxx, float(x))
        miny = min(miny, float(y))
        maxy = max(maxy, float(y))

    # Las cotas solo salen por izquierda/abajo. No añadir margen superior ni
    # derecho: ahí reaparecería el marco del machote.
    # Las hojas DWG conservan una regla horizontal apenas arriba de la vista.
    # Retirar solo esa guarda superior evita que aparezca como marco sin
    # recortar el contorno real de la placa.
    return [
        minx - 0.45,
        maxx + MARGEN_DERECHO_FOTO_CM,
        miny - 0.45,
        maxy - 0.28,
    ]


def _recortar_jpg_caras(
    hoja, ruta_temporal, ruta_final, vista_objetivo=None, bbox_forzada=None
):
    """
    Encuadra vista + cotas. Nunca recortar solo la geometria si hay cotas.
    """
    if Image is None:
        os.replace(ruta_temporal, ruta_final)
        return

    if bbox_forzada is not None:
        bbox_hoja, n_cotas = bbox_forzada, 1
    else:
        bbox_hoja, n_cotas = _bbox_contenido_hoja(hoja, vista_objetivo)
    if not bbox_hoja:
        os.replace(ruta_temporal, ruta_final)
        return

    try:
        img = Image.open(ruta_temporal)
        img_w, img_h = img.size
        margen = MARGEN_RECORTE_CARAS
        # Las cotas H/V del DrawingSketch ya tienen su bbox completo:
        # margen pequeño para encuadrar, no volver a exportar toda la hoja.
        try:
            for i in range(1, hoja.Sketches.Count + 1):
                if str(hoja.Sketches.Item(i).Name).upper().startswith(
                    PREFIJO_SKETCH_COTAS
                ):
                    # La cota y el modelo ya están incluidos en su bbox.
                    # Cualquier margen de hoja vuelve a introducir el marco
                    # impreso del machote.
                    margen = 0.0
                    break
        except Exception:
            pass
        if n_cotas > 0:
            margen = max(margen, 0.0)

        left_px, upper_px, right_px, lower_px = _bbox_hoja_a_pixeles(
            hoja, bbox_hoja, img_w, img_h, margen_ratio=margen
        )
        if (right_px - left_px) < 80 or (lower_px - upper_px) < 80:
            img.close()
            os.replace(ruta_temporal, ruta_final)
            return

        recorte = img.crop((left_px, upper_px, right_px, lower_px)).convert("RGB")
        recorte.save(ruta_final, quality=95, subsampling=0)
        img.close()
        try:
            os.remove(ruta_temporal)
        except OSError:
            pass
        log(f"    Encuadre JPG: cotas_detectadas={n_cotas}")
    except Exception as error:
        log(f"    AVISO recorte JPG: {error}")
        try:
            os.replace(ruta_temporal, ruta_final)
        except OSError:
            pass


def _etiqueta_referencia(eje, referencia):
    lado = str(referencia.get("lado", "")).lower()
    extremos = {
        ("X", "izq"): "XMIN",
        ("X", "der"): "XMAX",
        ("Y", "inf"): "YMIN",
        ("Y", "sup"): "YMAX",
        ("X", "centro"): "XCENTRO",
        ("Y", "centro"): "YCENTRO",
    }
    etiqueta = extremos.get((eje, lado), eje + "_" + lado.upper())
    return f"{etiqueta}_CONTACTO" if referencia.get("contacto") else etiqueta


def _nombre_archivo_referencia(indice, eje, referencia):
    pieza = _nombre_final_componente(referencia.get("pieza_id", "PIEZA"))
    pieza = _limpiar_nombre_archivo(pieza)[:70] or "PIEZA"
    sufijo_typ = "_TYP" if referencia.get("typ") else ""
    return (
        f"{indice:03d}_{_etiqueta_referencia(eje, referencia)}"
        f"{sufijo_typ}_{pieza}.jpg"
    )


def _limpiar_exportaciones_cara(carpeta):
    """La carpeta es exclusiva de fotos por referencia generadas por esta regla."""
    os.makedirs(carpeta, exist_ok=True)
    for nombre in os.listdir(carpeta):
        if not nombre.lower().endswith(".jpg"):
            continue
        try:
            os.remove(os.path.join(carpeta, nombre))
        except OSError:
            pass


def _exportar_caras_jpg(inv_app, plano, ensamble, planes_cotas):
    if not planes_cotas:
        log("ERROR: No hay referencias de cotas para exportar.")
        return 0, 0, ""
    if not str(plano.FullFileName or ""):
        log("ERROR: Guarda el machote antes de exportar.")
        return 0, 0, ""

    carpeta_raiz = os.path.join(
        _carpeta_salida_tanque(plano, ensamble), "COTAS_POR_REFERENCIA"
    )
    white = inv_app.TransientObjects.CreateColor(255, 255, 255)
    exportadas = 0
    esperadas = 0
    tg = inv_app.TransientGeometry

    log("Exportando una fotografía JPG por referencia dimensional...")
    log("  Modo: encuadre estable por cara (sin reframe HLR por foto)")
    log(f"  Carpeta: {carpeta_raiz}")

    for cara, plan in planes_cotas.items():
        hoja = plan["hoja"]
        vista = plan["vista"]
        referencias = [
            ("X", referencia) for referencia in plan["posiciones_x"]
        ] + [
            ("Y", referencia) for referencia in plan["posiciones_y"]
        ]
        referencias.sort(key=lambda item: (item[0], item[1]["valor"]))
        esperadas += len(referencias)
        carpeta_cara = os.path.join(carpeta_raiz, cara)
        _limpiar_exportaciones_cara(carpeta_cara)
        estado = _estado_base_vista(vista)
        log(f"  {cara}: {len(referencias)} fotos individuales")
        inicio_cara = time.perf_counter()

        try:
            hoja.Activate()
        except Exception:
            pass
        _borrar_cotas_hoja(hoja)
        _borrar_sketches_cotas(hoja)
        _preparar_encuadre_estable_cara(plano, hoja, vista, tg, estado)
        try:
            referencias_foto, origen_x, origen_y = _refrescar_referencias_exportacion(
                referencias, plan["dato_panel"], vista, tg
            )
        except Exception as error:
            log(f"  ERROR encuadre/releída {cara}: {error}")
            try:
                vista.Scale = estado["scale"]
                vista.Position = tg.CreatePoint2d(estado["x"], estado["y"])
                plano.Update()
            except Exception:
                pass
            continue

        for indice, (eje, referencia_foto) in enumerate(referencias_foto, start=1):
            temporal = os.path.join(carpeta_cara, f"_tmp_{indice:03d}.jpg")
            salida = os.path.join(
                carpeta_cara,
                _nombre_archivo_referencia(indice, eje, referencia_foto),
            )
            try:
                _borrar_sketches_cotas(hoja)
                if eje == "X":
                    creadas, fallos = _dibujar_cotas_hv_desde_origen(
                        hoja,
                        vista,
                        tg,
                        inv_app,
                        [referencia_foto],
                        [],
                        origen_x,
                        origen_y,
                    )
                else:
                    creadas, fallos = _dibujar_cotas_hv_desde_origen(
                        hoja,
                        vista,
                        tg,
                        inv_app,
                        [],
                        [referencia_foto],
                        origen_x,
                        origen_y,
                    )
                if creadas != 1 or fallos:
                    raise RuntimeError("No se pudo dibujar la única cota de la foto.")

                bbox_foto = _bbox_foto_referencia(
                    vista,
                    eje,
                    referencia_foto,
                    origen_x,
                    origen_y,
                )
                inv_app.ActiveView.Camera.SaveAsBitmap(
                    temporal,
                    ANCHO_EXPORTACION,
                    ALTO_EXPORTACION,
                    white,
                )
                _recortar_jpg_caras(
                    hoja,
                    temporal,
                    salida,
                    vista,
                    bbox_forzada=bbox_foto,
                )
                exportadas += 1
                if indice == 1 or indice % 25 == 0:
                    elapsed = max(0.001, time.perf_counter() - inicio_cara)
                    log(
                        f"    {cara} progreso {indice}/{len(referencias_foto)} "
                        f"({exportadas} ok, {elapsed / indice:.2f} s/foto)"
                    )
            except Exception as error:
                log(
                    f"  ERROR {cara} #{indice} "
                    f"{_etiqueta_referencia(eje, referencia_foto)}: {error}"
                )
                try:
                    if os.path.isfile(temporal):
                        os.remove(temporal)
                except OSError:
                    pass
            finally:
                _borrar_sketches_cotas(hoja)

        try:
            vista.Scale = estado["scale"]
            vista.Position = tg.CreatePoint2d(estado["x"], estado["y"])
            plano.Update()
        except Exception:
            pass
        elapsed_cara = max(0.001, time.perf_counter() - inicio_cara)
        log(
            f"  {cara} export terminado en {elapsed_cara:.1f}s "
            f"({elapsed_cara / max(1, len(referencias_foto)):.2f} s/foto)"
        )

    log(f"JPG exportados: {exportadas}/{esperadas}")
    return exportadas, esperadas, carpeta_raiz


def ejecutar(gestionar_com=True):
    if gestionar_com:
        pythoncom.CoInitialize()

    try:
        with open(RUTA_LOG, "w", encoding="utf-8") as archivo:
            archivo.write("")
    except OSError:
        pass

    silent_prev = False
    screen_prev = True
    inv_app = None
    plano = None

    try:
        log("=" * 62)
        log(" COTAS ABIGAIL - CARAS DEL TANQUE")
        log("=" * 62)
        log("Conectando con Inventor...")

        inv_app = conectar_inventor()
        plano = _obtener_plano_activo(inv_app)
        if plano is None:
            log("ERROR: El documento activo debe ser un plano .dwg/.idw.")
            return False

        if not str(plano.FullFileName or ""):
            log("ERROR: Guarda el plano antes de ejecutar.")
            return False

        log(f"Plano activo: {plano.DisplayName}")
        log("Buscando ensamble abierto...")
        ensamble = _obtener_ensamble_principal(inv_app)
        if ensamble is None:
            log("ERROR: No hay ningun ensamble (.iam) abierto.")
            return False

        log(f"Ensamble seleccionado: {ensamble.DisplayName}")
        log(f"Carpeta de salida: {_carpeta_salida_tanque(plano, ensamble)}")

        try:
            silent_prev = inv_app.SilentOperation
            screen_prev = inv_app.ScreenUpdating
            inv_app.SilentOperation = True
            inv_app.ScreenUpdating = False
        except Exception:
            pass

        log("Paso 1/2: Vistas de paredes + referencias reales desde (0,0)...")
        creadas, planes_cotas = _crear_caras(inv_app, plano, ensamble)

        if len(creadas) != 4:
            log(f"ERROR: Se crearon {len(creadas)}/4 hojas.")
            for i in range(1, plano.Sheets.Count + 1):
                try:
                    log(f"  Hoja actual #{i}: {plano.Sheets.Item(i).Name}")
                except Exception:
                    pass
            return False

        log("Paso 2/2: Exportando una fotografía por referencia...")
        exportadas, esperadas, carpeta = _exportar_caras_jpg(
            inv_app, plano, ensamble, planes_cotas
        )
        if exportadas != esperadas:
            return False

        log("")
        log("PROCESO COMPLETO:")
        log("  1) Cuatro paredes del tanque")
        log("  2) Una cota desde (0,0) por JPG")
        log(f"  3) {exportadas} JPG en: {carpeta}")
        log("  4) Machote limpiado para reutilizar")
        return True

    except Exception as error:
        log(f"ERROR NO CONTROLADO: {error}")
        log(traceback.format_exc())
        return False
    finally:
        # Siempre dejar el machote limpio (exito o error a media corrida).
        try:
            if plano is not None:
                _limpiar_machote(plano, inv_app)
        except Exception as error:
            log(f"AVISO: no se pudo limpiar el machote: {error}")
        try:
            if inv_app is not None:
                inv_app.SilentOperation = silent_prev
                inv_app.ScreenUpdating = screen_prev
        except Exception:
            pass
        if gestionar_com:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(0 if ejecutar() else 1)
