import contextlib
import io
import os
import re
import sys
import time
import importlib
import win32com.client

from inventor_com import conectar_inventor, obtener_ilogic_automation

try:
    from PIL import Image
except ImportError:
    print("❌ Falta Pillow. Instálalo con:")
    print(f'"{sys.executable}" -m pip install Pillow')
    raise

import creador_vistas
from rutas_runtime import ruta_hojas_diametro, ruta_piezas_solidas


# ============================================================
# CONFIGURACIÓN
# ============================================================
NOMBRE_REGLA_ILOGIC = "GenerarVistas"
CARPETA_EXPORTACION = "JPG"
# Portable: Planos/.runtime/ (antes C:\Temp\...)
RUTA_HOJAS_DIAMETRO = ruta_hojas_diametro()
# Sincronizado con diametro.py::RUTA_PIEZAS_SOLIDAS.
RUTA_PIEZAS_SOLIDAS = ruta_piezas_solidas()

ANCHO_EXPORTACION = 2400
ALTO_EXPORTACION = 1760
MARGEN_RECORTE = 0.06


# ============================================================
# UTILIDADES GENERALES
# ============================================================
def _agregar_carpeta_actual_al_path():
    carpeta_actual = os.path.dirname(os.path.abspath(__file__))
    if carpeta_actual not in sys.path:
        sys.path.insert(0, carpeta_actual)
    return carpeta_actual


def _importar_modulo(nombre_modulo: str):
    """
    Importa o recarga un módulo que está en la misma carpeta de este script.
    """
    if nombre_modulo in sys.modules:
        return importlib.reload(sys.modules[nombre_modulo])
    return importlib.import_module(nombre_modulo)


def _actualizar_inventor(inv_app):
    """
    Fuerza actualización del documento y de la vista activa para reducir
    problemas de sincronización entre pasos.
    """
    try:
        inv_app.ActiveDocument.Update()
    except:
        pass

    try:
        inv_app.ActiveView.Update()
    except:
        pass

    try:
        inv_app.UserInterfaceManager.DoEvents()
    except:
        pass


def _obtener_ilogic_automation(inv_app):
    return obtener_ilogic_automation(inv_app)


def _obtener_o_activar_plano(inv_app):
    """
    Devuelve el DrawingDocument activo. Si el foco está en un .iam/.ipt,
    busca un plano abierto y lo activa automáticamente.
    """
    kDrawingDocument = 12292

    try:
        doc = inv_app.ActiveDocument
    except Exception:
        doc = None

    if doc is not None and doc.DocumentType == kDrawingDocument:
        return doc

    candidatos = []
    try:
        total = inv_app.Documents.Count
    except Exception:
        total = 0

    for i in range(1, total + 1):
        try:
            candidato = inv_app.Documents.Item(i)
            if candidato.DocumentType == kDrawingDocument:
                candidatos.append(candidato)
        except Exception:
            pass

    if not candidatos:
        return None

    preferido = None
    for candidato in candidatos:
        try:
            nombre = str(candidato.DisplayName).upper()
        except Exception:
            nombre = ""
        if "MACHOTE" in nombre or "PLANO" in nombre:
            preferido = candidato
            break

    plano = preferido if preferido is not None else candidatos[0]

    try:
        plano.Activate()
        time.sleep(0.3)
        _actualizar_inventor(inv_app)
        print(f"Plano activado automaticamente: {plano.DisplayName}")
    except Exception as e:
        print(f"AVISO: No se pudo activar el plano automaticamente: {e}")

    return plano

def _limpiar_nombre_archivo(nombre: str) -> str:
    """
    Limpia caracteres inválidos para nombre de archivo en Windows.
    """
    nombre = re.sub(r'[<>:"/\\\\|?*]', "_", nombre)
    nombre = nombre.strip().rstrip(".")
    return nombre


def _separar_nombre_hoja(nombre: str):
    """
    Separa:
    'TAPA COVER_FRENTE_1:122' -> ('TAPA COVER_FRENTE_1', ':122')
    """
    partes = nombre.rsplit(":", 1)

    if len(partes) == 2 and partes[1].isdigit():
        return partes[0], ":" + partes[1]

    return nombre, ""


def _cargar_hojas_para_diametro():
    """
    Lee Planos/.runtime/hojas_para_diametro.txt y devuelve:
    - set de nombres completos en mayúsculas
    - set de nombres base en mayúsculas (sin :n)
    """
    visibles = set()
    bases = set()

    try:
        if not os.path.exists(RUTA_HOJAS_DIAMETRO):
            return visibles, bases

        with open(RUTA_HOJAS_DIAMETRO, "r", encoding="utf-8") as f:
            for linea in f:
                nombre = linea.strip()
                if not nombre:
                    continue

                nombre_up = nombre.upper()
                visibles.add(nombre_up)

                base, _ = _separar_nombre_hoja(nombre_up)
                bases.add(base)

        print(f"📘 Hojas de diámetro cargadas desde: {RUTA_HOJAS_DIAMETRO}")
        return visibles, bases

    except Exception as e:
        print(f"⚠️ No se pudieron cargar las hojas de diámetro: {e}")
        return set(), set()


def _convertir_nombre_tecnico_hoja(
    nombre: str,
    hojas_diametro_visibles=None,
    hojas_diametro_bases=None,
    hojas_lado_sin_thk_bases=None,
) -> str:
    """
    Convierte nombres finales de hoja.

    Regla:
    - Si la hoja fue mandada a diámetro:
        FRENTE_1 -> DIAMETRO_EXTERIOR
        FRENTE_2 -> DIAMETRO_INTERIOR
    - Si no fue mandada a diámetro:
        FRENTE_1 -> LARGO
        FRENTE_2 -> ANCHO
    - LADO -> THK **solo si THK.py logró cotar la hoja**. Si la hoja aparece
      en ``hojas_lado_sin_thk_bases`` (lista de pendientes de THK), su nombre
      se conserva como ``_LADO`` para que el JPG resultante no mienta al
      llamarse ``_THK`` sin cota real.
    - ALTO se mantiene (4ª captura de perfiles U/L)
    """
    if hojas_diametro_visibles is None:
        hojas_diametro_visibles = set()

    if hojas_diametro_bases is None:
        hojas_diametro_bases = set()

    if hojas_lado_sin_thk_bases is None:
        hojas_lado_sin_thk_bases = set()

    base, sufijo = _separar_nombre_hoja(nombre)

    base_up = base.upper()
    visible_up = nombre.upper()

    es_diametro = (
        visible_up in hojas_diametro_visibles
        or base_up in hojas_diametro_bases
    )

    if "_ALTO" in base_up:
        pass  # ya es nombre final

    elif "_LARGO_PATA" in base_up:
        pass  # ya es nombre final (nueva hoja del canal U/L/C)

    elif "_LADO" in base:
        # No renombramos a _THK cuando el resolver de THK falló para esta hoja.
        if base_up in hojas_lado_sin_thk_bases:
            pass
        else:
            base = base.replace("_LADO", "_THK")

    elif es_diametro and "_FRENTE_1" in base:
        base = base.replace("_FRENTE_1", "_DIAMETRO_EXTERIOR")

    elif es_diametro and "_FRENTE_2" in base:
        base = base.replace("_FRENTE_2", "_DIAMETRO_INTERIOR")

    elif "_FRENTE_1" in base:
        base = base.replace("_FRENTE_1", "_LARGO")

    elif "_FRENTE_2" in base:
        base = base.replace("_FRENTE_2", "_ANCHO")

    return base + sufijo


def renombrar_hojas_finales(doc, nombres_permitidos=None, hojas_lado_sin_thk=None):
    """
    Renombra hojas al final del flujo:
    - Circulares:
        FRENTE_1 -> DIAMETRO_EXTERIOR
        FRENTE_2 -> DIAMETRO_INTERIOR
    - No circulares:
        FRENTE_1 -> LARGO
        FRENTE_2 -> ANCHO
    - LADO -> THK (**solo si THK cotó la hoja**; si está en
      ``hojas_lado_sin_thk`` mantenemos ``_LADO`` para no engañar con un
      nombre ``_THK`` cuando en realidad no hay cota).

    Devuelve un dict `{nombre_original_upper: nombre_nuevo_str}` con los
    renombrados aplicados (útil para tracking en flujos por lotes).
    """
    try:
        draw_doc = win32com.client.CastTo(doc, "DrawingDocument")
    except:
        draw_doc = doc

    hojas_diametro_visibles, hojas_diametro_bases = _cargar_hojas_para_diametro()

    hojas_lado_sin_thk_bases = set()
    if hojas_lado_sin_thk:
        for nombre in hojas_lado_sin_thk:
            base, _ = _separar_nombre_hoja(str(nombre))
            hojas_lado_sin_thk_bases.add(base.upper())

    cambios = 0
    mapeo = {}

    permitidos_up = None
    if nombres_permitidos is not None:
        permitidos_up = {str(x).upper() for x in nombres_permitidos}

    try:
        total_hojas = draw_doc.Sheets.Count
    except Exception as e:
        print(f"⚠️ No se pudieron leer las hojas para renombrar: {e}")
        return mapeo

    for i in range(1, total_hojas + 1):
        try:
            hoja = draw_doc.Sheets.Item(i)

            nombre_actual_visible = str(hoja.Name)
            nombre_actual_base, _ = _separar_nombre_hoja(nombre_actual_visible)

            if permitidos_up is not None and nombre_actual_base.upper() not in permitidos_up:
                continue

            nombre_nuevo_visible = _convertir_nombre_tecnico_hoja(
                nombre_actual_visible,
                hojas_diametro_visibles,
                hojas_diametro_bases,
                hojas_lado_sin_thk_bases,
            )

            nombre_nuevo_base, _ = _separar_nombre_hoja(nombre_nuevo_visible)

            if nombre_nuevo_base != nombre_actual_base:
                hoja.Name = nombre_nuevo_base
                time.sleep(0.05)

                try:
                    nombre_resultado = str(hoja.Name)
                except:
                    nombre_resultado = nombre_nuevo_visible

                print(f"📝 Renombrada hoja: {nombre_actual_visible}  ->  {nombre_resultado}")
                mapeo[nombre_actual_base.upper()] = nombre_resultado
                cambios += 1
            else:
                mapeo[nombre_actual_base.upper()] = nombre_actual_visible

        except Exception as e:
            print(f"⚠️ No se pudo renombrar una hoja: {e}")

    print(f"✅ Renombrado final terminado. Hojas cambiadas: {cambios}")
    return mapeo


# ============================================================
# UTILIDADES POR LOTES (modo D)
# ============================================================
# Sufijos técnicos que pueden aparecer en JPGs ya exportados. Usado para el
# modo incremental (F) que detecta piezas ya listas.
_SUFIJOS_JPG_PIEZA_EXPORTADA = (
    "DIAMETRO_EXTERIOR",
    "DIAMETRO_INTERIOR",
    "LARGO_PATA",
    "ANCHO",
    "LARGO",
    "THK",
    "ALTO",
    "FRENTE_1",
    "FRENTE_2",
    "LADO",
)
_RE_JPG_PIEZA = re.compile(
    r"^(?P<pieza>.+?)_(?:" + "|".join(_SUFIJOS_JPG_PIEZA_EXPORTADA) + r")(?:_\d+)?$",
    re.IGNORECASE,
)


def _acumular_pieza_desde_archivo(nombre_archivo, out):
    if not nombre_archivo.lower().endswith(".jpg"):
        return
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    m = _RE_JPG_PIEZA.match(base)
    if m:
        out.add(m.group("pieza"))


def _piezas_ya_exportadas(carpeta):
    """
    Devuelve set con nombres_base de piezas que ya tienen ≥1 JPG en `carpeta`.

    Cubre dos estructuras:
    - Plana: `<carpeta>/<PIEZA>_<TIPO>_<N>.jpg`
    - Reorganizada: `<carpeta>/<CARA>/<PIEZA_FOLDER>/<PIEZA>_<TIPO>_<N>.jpg`
    """
    piezas = set()
    if not carpeta or not os.path.isdir(carpeta):
        return piezas
    try:
        entradas = os.listdir(carpeta)
    except OSError:
        return piezas

    for nombre in entradas:
        ruta = os.path.join(carpeta, nombre)
        if os.path.isfile(ruta):
            _acumular_pieza_desde_archivo(nombre, piezas)
            continue
        # Subcarpetas por cara (FRONT/BACK/LEFT/RIGHT/TOP/OTROS o similar).
        try:
            sub_entradas = os.listdir(ruta)
        except OSError:
            continue
        for sub in sub_entradas:
            sub_ruta = os.path.join(ruta, sub)
            if os.path.isfile(sub_ruta):
                _acumular_pieza_desde_archivo(sub, piezas)
                continue
            try:
                archivos_pieza = os.listdir(sub_ruta)
            except OSError:
                continue
            for arch in archivos_pieza:
                _acumular_pieza_desde_archivo(arch, piezas)
    return piezas


def borrar_hojas_por_nombres(doc, nombres_a_borrar, nombre_machote_protegido=None):
    """
    Borra del machote las hojas cuyo nombre (base, upper) esté en
    `nombres_a_borrar`. Se aplica un candado extra sobre la hoja machote
    para no borrarla nunca.
    """
    if not nombres_a_borrar:
        return 0
    try:
        draw_doc = win32com.client.CastTo(doc, "DrawingDocument")
    except:
        draw_doc = doc
    objetivo = {str(x).upper() for x in nombres_a_borrar}
    if nombre_machote_protegido:
        objetivo.discard(str(nombre_machote_protegido).upper())

    borradas = 0
    try:
        total = draw_doc.Sheets.Count
    except Exception as e:
        print(f"⚠️ No se pudieron leer hojas para borrar: {e}")
        return 0

    for i in range(total, 0, -1):
        try:
            hoja = draw_doc.Sheets.Item(i)
            nombre = str(hoja.Name)
            base, _ = _separar_nombre_hoja(nombre)
            if base.upper() in objetivo:
                hoja.Delete()
                borradas += 1
        except Exception as e:
            print(f"  AVISO borrando hoja #{i}: {e}")
    if borradas:
        print(f"  🧹 Borradas {borradas} hojas del lote tras exportar.")
    return borradas

# ============================================================
# RECORTE AUTOMÁTICO DEL JPG
# ============================================================
def _expandir_bbox(bbox_actual, minx, maxx, miny, maxy):
    if bbox_actual is None:
        return [minx, maxx, miny, maxy]

    bbox_actual[0] = min(bbox_actual[0], minx)
    bbox_actual[1] = max(bbox_actual[1], maxx)
    bbox_actual[2] = min(bbox_actual[2], miny)
    bbox_actual[3] = max(bbox_actual[3], maxy)
    return bbox_actual


def _obtener_bbox_pieza(hoja):
    """Bbox SOLO de la geometría de la pieza (DrawingCurves de las vistas).

    No incluye cotas, no incluye viewport. Es lo que se usa para calcular el
    centro donde debe quedar la pieza en el JPG final.
    """
    bbox = None

    try:
        total_views = hoja.DrawingViews.Count
    except Exception:
        total_views = 0

    for i in range(1, total_views + 1):
        try:
            view = hoja.DrawingViews.Item(i)
        except Exception:
            continue

        view_bbox = None
        try:
            total_curvas = view.DrawingCurves.Count
        except Exception:
            total_curvas = 0

        for j in range(1, total_curvas + 1):
            try:
                curva = view.DrawingCurves.Item(j)
                caja = curva.Evaluator2D.RangeBox
                minx = float(caja.MinPoint.X)
                maxx = float(caja.MaxPoint.X)
                miny = float(caja.MinPoint.Y)
                maxy = float(caja.MaxPoint.Y)
                view_bbox = _expandir_bbox(view_bbox, minx, maxx, miny, maxy)
            except Exception:
                continue

        if view_bbox is None:
            # Fallback al viewport si aún no hay curvas materializadas.
            try:
                left = float(view.Left)
                top = float(view.Top)
                width = float(view.Width)
                height = float(view.Height)
                view_bbox = (left, left + width, top - height, top)
            except Exception:
                view_bbox = None

        if view_bbox is not None:
            bbox = _expandir_bbox(bbox, *view_bbox)

    return bbox


def _obtener_bbox_cotas(hoja):
    """
    Bbox de las cotas dibujadas en la hoja. Incluye tanto ``RangeBox`` como
    la ``Text.Origin`` cuando esté disponible, porque en algunos casos el
    ``RangeBox`` no cubre el texto del número (queda parcialmente fuera del
    bbox reportado y la cota se pierde al recortar el JPG).
    """
    bbox = None
    try:
        total_dims = hoja.DrawingDimensions.GeneralDimensions.Count
    except Exception:
        total_dims = 0

    for i in range(1, total_dims + 1):
        try:
            dim = hoja.DrawingDimensions.GeneralDimensions.Item(i)
        except Exception:
            continue

        try:
            box = dim.RangeBox
            minx = float(box.MinPoint.X)
            maxx = float(box.MaxPoint.X)
            miny = float(box.MinPoint.Y)
            maxy = float(box.MaxPoint.Y)
            bbox = _expandir_bbox(bbox, minx, maxx, miny, maxy)
        except Exception:
            pass

        # Incluir la posición del texto por seguridad: RangeBox a veces
        # reporta sólo la línea de cota sin el número.
        try:
            texto = dim.Text
            origen = texto.Origin
            tx = float(origen.X)
            ty = float(origen.Y)
            # Añadimos un padding pequeño alrededor del punto del texto para
            # cubrir el ancho aproximado del número (~1.2 cm).
            bbox = _expandir_bbox(bbox, tx - 0.6, tx + 0.6, ty - 0.4, ty + 0.4)
        except Exception:
            pass

    return bbox


def _obtener_bbox_hoja(hoja):
    """
    Bbox de la unión de pieza + cotas. Mantenida por compatibilidad.
    Para el recorte centrado usar ``_bbox_recorte_centrado``.
    """
    bbox_pieza = _obtener_bbox_pieza(hoja)
    bbox_cotas = _obtener_bbox_cotas(hoja)

    bbox = bbox_pieza
    if bbox_cotas is not None:
        if bbox is None:
            bbox = bbox_cotas
        else:
            bbox = _expandir_bbox(bbox, *bbox_cotas)
    return bbox


def _bbox_recorte_ideal(hoja):
    """
    Calcula el bbox IDEAL para el JPG final, SIN clampear al borde físico
    del sheet. Si el bbox se sale del sheet, luego se compensa con padding
    blanco en ``_recortar_exportacion_jpg``.

    Reglas del bbox ideal:
    - Centrado en el centro geométrico de la pieza.
    - Incluye todas las cotas dibujadas (con margen extra por seguridad).
    - Padding generoso alrededor: 25 % del tamaño de la pieza, mínimo 2.5 cm.
    - Padding EXTRA cuando NO se detectan cotas en la hoja (la cota puede
      no haberse creado, o `_obtener_bbox_cotas` no la detectó): en ese
      caso se añaden 4 cm adicionales por lado para no perder una cota que
      Inventor sí dibujó pero que este código no llegó a ver.
    - NO se fuerza aspect ratio: antes se comprimía a [0.55, 1.80] y eso
      era el origen del "corta las cotas"; ahora respetamos el ratio real
      del contenido y compensamos con padding blanco en el paso siguiente.
    """
    bbox_pieza = _obtener_bbox_pieza(hoja)
    if bbox_pieza is None:
        return None
    minx_p, maxx_p, miny_p, maxy_p = bbox_pieza
    cx = (minx_p + maxx_p) / 2.0
    cy = (miny_p + maxy_p) / 2.0
    ancho_pieza = max(1e-6, maxx_p - minx_p)
    alto_pieza = max(1e-6, maxy_p - miny_p)

    bbox_cotas = _obtener_bbox_cotas(hoja)

    # Radios desde el centro de la pieza hasta cubrir todo el contenido
    # visible (pieza + cotas).
    dx_max = max(cx - minx_p, maxx_p - cx)
    dy_max = max(cy - miny_p, maxy_p - cy)

    if bbox_cotas is not None:
        minx_c, maxx_c, miny_c, maxy_c = bbox_cotas
        dx_max = max(dx_max, cx - minx_c, maxx_c - cx)
        dy_max = max(dy_max, cy - miny_c, maxy_c - cy)

    # Padding generoso: nunca menos de 2.5 cm por lado.
    pad_x = max(2.5, ancho_pieza * 0.25)
    pad_y = max(2.5, alto_pieza * 0.25)

    # Si no detectamos cotas en la hoja, damos padding EXTRA por si Inventor
    # dibujó algo que _obtener_bbox_cotas no llegó a ver (cotas de tipos
    # exóticos, notas, etc.). Mejor sobre-incluir que perder cotas.
    if bbox_cotas is None:
        pad_x += 4.0
        pad_y += 4.0

    dx_max += pad_x
    dy_max += pad_y

    minx = cx - dx_max
    maxx = cx + dx_max
    miny = cy - dy_max
    maxy = cy + dy_max

    return (minx, maxx, miny, maxy)


# Alias público de compatibilidad para llamadores que usaran el nombre viejo.
_bbox_recorte_centrado = _bbox_recorte_ideal


def _bbox_hoja_a_pixeles(hoja, bbox_hoja, img_w, img_h, margen_ratio=0.0):
    """
    Convierte un bbox en coordenadas de hoja (cm) a coordenadas de imagen
    (px), CLAMPEANDO a los bordes físicos del sheet. Se usa para saber qué
    parte del JPG exportado por Inventor debemos recortar. Si el bbox ideal
    se sale del sheet, ``_recortar_exportacion_jpg`` se encarga después de
    compensar con padding blanco.
    """
    sheet_w = float(hoja.Width)
    sheet_h = float(hoja.Height)

    minx, maxx, miny, maxy = bbox_hoja

    if margen_ratio > 0:
        bbox_w = max(1e-6, maxx - minx)
        bbox_h = max(1e-6, maxy - miny)
        margen_x = min(bbox_w * 0.40, max(0.4, bbox_w * margen_ratio))
        margen_y = min(bbox_h * 0.40, max(0.4, bbox_h * margen_ratio))
        minx -= margen_x
        maxx += margen_x
        miny -= margen_y
        maxy += margen_y

    minx = max(0.0, minx)
    maxx = min(sheet_w, maxx)
    miny = max(0.0, miny)
    maxy = min(sheet_h, maxy)

    left_px = int((minx / sheet_w) * img_w)
    right_px = int((maxx / sheet_w) * img_w)

    upper_px = int(((sheet_h - maxy) / sheet_h) * img_h)
    lower_px = int(((sheet_h - miny) / sheet_h) * img_h)

    left_px = max(0, min(left_px, img_w - 1))
    right_px = max(left_px + 1, min(right_px, img_w))
    upper_px = max(0, min(upper_px, img_h - 1))
    lower_px = max(upper_px + 1, min(lower_px, img_h))

    return left_px, upper_px, right_px, lower_px


def _blanquear_fondo(img, umbral=180):
    """
    Convierte a blanco puro todo píxel con valor de gris >= ``umbral``.
    Esto elimina el fondo beige del sheet de Inventor y las líneas grises
    tenues que Inventor pinta en el borde físico del sheet al exportar
    (esas líneas grises son las franjas que se veían en los JPGs y que
    hacían parecer que la pieza se salía del sheet).

    Las líneas del dibujo (negro puro o casi negro) se conservan.
    """
    rgb = img.convert("RGB")
    gris = rgb.convert("L")
    mascara = gris.point(lambda x: 255 if x >= umbral else 0, mode="L")
    blanco = Image.new("RGB", rgb.size, (255, 255, 255))
    return Image.composite(blanco, rgb, mascara)


def _recortar_exportacion_jpg(hoja, ruta_temporal, ruta_final):
    """
    Recorta el JPG exportado dejando la pieza CENTRADA en la imagen final,
    con padding blanco si el bbox ideal se salía del sheet físico. Además
    blanquea el fondo beige del sheet para que no se vean las franjas
    grises del contorno físico del sheet.

    Estrategia:
    1. ``_bbox_recorte_ideal(hoja)`` calcula el rectángulo ideal centrado en
       la pieza, expandido para cubrir cotas y con margen de respiro. Este
       bbox NO se clampea al sheet.
    2. Se blanquea el fondo del JPG exportado (beige + grises tenues).
    3. Se recorta el JPG a la intersección bbox_ideal ∩ sheet.
    4. Se crea un canvas blanco del tamaño proporcional al bbox ideal
       completo y se pega el recorte en la posición correcta dentro del
       canvas, de forma que la pieza quede CENTRADA en la imagen final.
    """
    bbox_ideal = _bbox_recorte_ideal(hoja)

    if not bbox_ideal:
        try:
            img_tmp = Image.open(ruta_temporal)
            _blanquear_fondo(img_tmp).save(
                ruta_final, quality=95, subsampling=0
            )
            img_tmp.close()
            try:
                os.remove(ruta_temporal)
            except OSError:
                pass
        except Exception:
            os.replace(ruta_temporal, ruta_final)
        return

    img_raw = Image.open(ruta_temporal)
    img = _blanquear_fondo(img_raw)
    img_raw.close()
    img_w, img_h = img.size

    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
    except Exception:
        os.replace(ruta_temporal, ruta_final)
        img.close()
        return

    if sheet_w <= 0 or sheet_h <= 0:
        os.replace(ruta_temporal, ruta_final)
        img.close()
        return

    minx_i, maxx_i, miny_i, maxy_i = bbox_ideal
    ancho_ideal_cm = max(1e-6, maxx_i - minx_i)
    alto_ideal_cm = max(1e-6, maxy_i - miny_i)

    # Salvaguarda anti-"recorte de más": si el bbox ideal cubre ya casi todo
    # el sheet (>=85 % en ambos ejes), es más seguro entregar el sheet
    # completo tal cual (ya blanqueado el fondo).
    cobertura_x = ancho_ideal_cm / sheet_w
    cobertura_y = alto_ideal_cm / sheet_h
    if cobertura_x >= 0.85 and cobertura_y >= 0.85:
        img.save(ruta_final, quality=95, subsampling=0)
        img.close()
        try:
            os.remove(ruta_temporal)
        except OSError:
            pass
        return

    # Intersección con el sheet físico (lo que realmente hay en el JPG).
    minx_c = max(0.0, minx_i)
    maxx_c = min(sheet_w, maxx_i)
    miny_c = max(0.0, miny_i)
    maxy_c = min(sheet_h, maxy_i)

    if maxx_c <= minx_c or maxy_c <= miny_c:
        os.replace(ruta_temporal, ruta_final)
        img.close()
        return

    # Bbox del recorte real dentro del JPG exportado.
    left_px = int((minx_c / sheet_w) * img_w)
    right_px = int((maxx_c / sheet_w) * img_w)
    upper_px = int(((sheet_h - maxy_c) / sheet_h) * img_h)
    lower_px = int(((sheet_h - miny_c) / sheet_h) * img_h)

    left_px = max(0, min(left_px, img_w - 1))
    right_px = max(left_px + 1, min(right_px, img_w))
    upper_px = max(0, min(upper_px, img_h - 1))
    lower_px = max(upper_px + 1, min(lower_px, img_h))

    recorte = img.crop((left_px, upper_px, right_px, lower_px)).convert("RGB")
    img.close()

    # Escala: cm -> px (misma escala del JPG exportado por Inventor).
    escala_x = img_w / sheet_w
    escala_y = img_h / sheet_h

    canvas_w = max(recorte.size[0], int(round(ancho_ideal_cm * escala_x)))
    canvas_h = max(recorte.size[1], int(round(alto_ideal_cm * escala_y)))

    # Padding requerido a cada lado (cm) para llegar del bbox real al ideal.
    pad_izq_cm = max(0.0, minx_c - minx_i)
    pad_sup_cm = max(0.0, maxy_i - maxy_c)

    offset_x = int(round(pad_izq_cm * escala_x))
    offset_y = int(round(pad_sup_cm * escala_y))

    # Ajuste por posibles redondeos: no dejar que el recorte se salga del
    # canvas.
    offset_x = max(0, min(offset_x, canvas_w - recorte.size[0]))
    offset_y = max(0, min(offset_y, canvas_h - recorte.size[1]))

    canvas = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    canvas.paste(recorte, (offset_x, offset_y))
    canvas.save(ruta_final, quality=95, subsampling=0)

    try:
        os.remove(ruta_temporal)
    except OSError:
        pass


# ============================================================
# EXPORTACIÓN A JPG
# ============================================================
def _cargar_piezas_solidas():
    """Lee el listado de piezas identificadas como cilindros sólidos."""
    piezas = set()
    try:
        if not os.path.exists(RUTA_PIEZAS_SOLIDAS):
            return piezas
        with open(RUTA_PIEZAS_SOLIDAS, "r", encoding="utf-8") as f:
            for linea in f:
                nombre = linea.strip()
                if nombre:
                    piezas.add(nombre.upper())
    except Exception:
        pass
    return piezas


def _limpiar_jpgs_diametro_interior_huerfanos(carpeta_salida, piezas_solidas):
    """
    Borra cualquier JPG ``*_DIAMETRO_INTERIOR_*.jpg`` que pertenezca a una
    pieza identificada como sólida.

    Diseñado para corridas incrementales, donde los JPGs viejos generados
    ANTES de que ``diametro.py`` supiera eliminar la hoja _FRENTE_2 pueden
    persistir en disco y contaminar el reporte final. Cubre la estructura
    plana y la reorganizada por cara (``LEFT/PIEZA/xxx.jpg``).
    """
    if not piezas_solidas or not carpeta_salida:
        return 0
    if not os.path.isdir(carpeta_salida):
        return 0

    borrados = 0

    def _es_diametro_interior(nombre_archivo, piezas_up):
        base = os.path.splitext(nombre_archivo)[0].upper()
        if "_DIAMETRO_INTERIOR" not in base:
            return False
        for p in piezas_up:
            if p and p in base:
                return True
        return False

    for raiz, _dirs, archivos in os.walk(carpeta_salida):
        for nombre in archivos:
            if not nombre.lower().endswith(".jpg"):
                continue
            if _es_diametro_interior(nombre, piezas_solidas):
                ruta = os.path.join(raiz, nombre)
                try:
                    os.remove(ruta)
                    borrados += 1
                except OSError:
                    continue

    if borrados:
        print(
            f"🧹 Limpieza: {borrados} JPG(s) _DIAMETRO_INTERIOR huérfanos "
            f"borrados (piezas cilíndricas sólidas)."
        )
    return borrados


def exportar_hojas_jpg(
    inv_app,
    doc,
    carpeta_salida=None,
    nombres_permitidos=None,
    hojas_lado_sin_thk=None,
):
    """
    Exporta cada hoja con vistas a JPG, excluyendo hojas vacías.
    Guarda en una carpeta junto al archivo machote o en la ruta explícita.
    Además recorta automáticamente la hoja para dejar solo el contenido útil.

    `nombres_permitidos` (opcional): set con nombres base (upper) de hojas a
    exportar. Si no se provee, se exportan todas.

    `hojas_lado_sin_thk` (opcional): lista de hojas ``_LADO`` cuyo THK no se
    pudo cotar. Se usa para NO renombrar el archivo a ``_THK``; conserva el
    sufijo ``_LADO`` para evitar dar la falsa impresión de que la cota existe.
    """
    print("⏳ Paso 4/4: Exportando hojas a JPG...")

    try:
        draw_doc = win32com.client.CastTo(doc, "DrawingDocument")
    except:
        try:
            draw_doc = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
        except Exception as e:
            print(f"❌ No se pudo convertir el documento activo a DrawingDocument: {e}")
            return

    try:
        ruta_dibujo = draw_doc.FullFileName
    except:
        ruta_dibujo = ""

    if not ruta_dibujo:
        print("❌ No se pudo obtener la ruta del archivo machote. Guarda el dibujo antes de exportar.")
        return

    if carpeta_salida is None:
        carpeta_base = os.path.dirname(ruta_dibujo)
        carpeta_salida = os.path.join(carpeta_base, CARPETA_EXPORTACION)
    else:
        carpeta_salida = os.path.abspath(carpeta_salida)
    os.makedirs(carpeta_salida, exist_ok=True)

    # Limpieza previa: borrar JPG _DIAMETRO_INTERIOR huérfanos de piezas
    # que ``diametro.py`` identificó como cilindros sólidos. Ejecutado ANTES
    # de exportar para que el reporte final no vuelva a duplicar el archivo.
    try:
        piezas_solidas = _cargar_piezas_solidas()
        if piezas_solidas:
            _limpiar_jpgs_diametro_interior_huerfanos(
                carpeta_salida, piezas_solidas
            )
    except Exception as exc_limp:
        print(f"AVISO: limpieza de _DIAMETRO_INTERIOR huérfanos falló: {exc_limp}")

    try:
        white = inv_app.TransientObjects.CreateColor(255, 255, 255)
    except Exception as e:
        print(f"❌ No se pudo crear el color de fondo blanco: {e}")
        return

    exportadas = 0
    omitidas = 0

    try:
        total_hojas = draw_doc.Sheets.Count
    except Exception as e:
        print(f"❌ No se pudieron leer las hojas del dibujo: {e}")
        return

    # Cachear catálogos de diámetro una sola vez (leer hojas del ensamble por
    # cada iteración desperdiciaba varios segundos por foto).
    try:
        hojas_diametro_visibles, hojas_diametro_bases = _cargar_hojas_para_diametro()
    except Exception:
        hojas_diametro_visibles, hojas_diametro_bases = [], []

    hojas_lado_sin_thk_bases = set()
    if hojas_lado_sin_thk:
        for nombre in hojas_lado_sin_thk:
            base, _ = _separar_nombre_hoja(str(nombre))
            hojas_lado_sin_thk_bases.add(base.upper())

    permitidos_up = None
    if nombres_permitidos is not None:
        permitidos_up = {str(x).upper() for x in nombres_permitidos}

    for i in range(1, total_hojas + 1):
        try:
            hoja = draw_doc.Sheets.Item(i)
        except Exception as e:
            print(f"⚠️ No se pudo acceder a la hoja #{i}: {e}")
            continue

        try:
            nombre_hoja_actual_visible = str(hoja.Name)
            nombre_actual_base, _ = _separar_nombre_hoja(nombre_hoja_actual_visible)
            if permitidos_up is not None and nombre_actual_base.upper() not in permitidos_up:
                continue
            nombre_hoja = _convertir_nombre_tecnico_hoja(
                nombre_hoja_actual_visible,
                hojas_diametro_visibles,
                hojas_diametro_bases,
                hojas_lado_sin_thk_bases,
            )
        except:
            nombre_hoja = f"Hoja_{i}"

        # Excluir hojas vacías / machote sin vistas
        try:
            if hoja.DrawingViews.Count == 0:
                print(f"⏭️ Omitiendo hoja vacía: {nombre_hoja}")
                omitidas += 1
                continue
        except:
            print(f"⏭️ Omitiendo hoja sin acceso a vistas: {nombre_hoja}")
            omitidas += 1
            continue

        try:
            # Un solo ciclo Activate + Update + sleep breve (antes había dos
            # rondas de 0.3s y un Fit() innecesario para vistas ya escaladas).
            hoja.Activate()
            try:
                inv_app.ActiveView.Update()
            except Exception:
                pass
            time.sleep(0.15)

            nombre_archivo = _limpiar_nombre_archivo(nombre_hoja) + ".jpg"
            ruta_jpg_final = os.path.join(carpeta_salida, nombre_archivo)

            nombre_temporal = "_tmp_" + _limpiar_nombre_archivo(nombre_hoja) + ".jpg"
            ruta_jpg_temporal = os.path.join(carpeta_salida, nombre_temporal)

            inv_app.ActiveView.Camera.SaveAsBitmap(
                ruta_jpg_temporal,
                ANCHO_EXPORTACION,
                ALTO_EXPORTACION,
                white
            )

            _recortar_exportacion_jpg(hoja, ruta_jpg_temporal, ruta_jpg_final)

            print(f"🖼️ Exportado: {ruta_jpg_final}")
            exportadas += 1

            # Respiro periódico para vaciar la cola COM (evita saturación tipo
            # RPC crash cuando se exportan cientos de hojas seguidas).
            if exportadas % 25 == 0:
                try:
                    _actualizar_inventor(inv_app)
                except Exception:
                    pass
                time.sleep(0.4)

        except Exception as e:
            print(f"⚠️ No se pudo exportar la hoja '{nombre_hoja}': {e}")

    print(f"✅ Exportación terminada. JPG creados: {exportadas} | Hojas omitidas: {omitidas}")
    print(f"📁 Carpeta de salida: {carpeta_salida}")


# ============================================================
# FLUJO COMPLETO (Desde la App con UI y API Pura)
# ============================================================
TAM_LOTE_PIEZAS = 20


def _chunks(secuencia, n):
    for i in range(0, len(secuencia), n):
        yield secuencia[i:i + n]


def _nombre_hoja_machote(doc):
    try:
        return str(doc.ActiveSheet.Name)
    except Exception:
        return None


def _forzar_compute_hojas(doc, inv_app, nombres_hojas, log_fn=None):
    """
    Fuerza que las vistas de las hojas del lote calculen su geometría 2D
    antes de que corran cotas.py / THK.py.

    Sin esta pasada, las funciones de cotado a veces ven ``DrawingCurves``
    vacío porque Inventor aún no dibujó la vista, y todas las hojas se
    marcan como "sin geometría lineal" y se envían al camino de diámetro,
    resultando en JPG sin cotas.

    OJO con el matching: Inventor agrega automáticamente ``:N`` al nombre
    cuando ya hay otra hoja con el mismo base ('X_FRENTE_1' -> 'X_FRENTE_1:54').
    Comparamos siempre contra la BASE del nombre para que los lotes sigan
    matcheando aunque Inventor haya numerado internamente.
    """
    if not nombres_hojas:
        return
    permitidos_up = {str(n).upper() for n in nombres_hojas}
    procesadas = 0
    sin_curvas = 0
    try:
        total = doc.Sheets.Count
    except Exception:
        return
    for i in range(1, total + 1):
        try:
            hoja = doc.Sheets.Item(i)
            nombre_visible = str(hoja.Name)
            base_visible, _ = _separar_nombre_hoja(nombre_visible)
        except Exception:
            continue
        if base_visible.upper() not in permitidos_up:
            continue
        try:
            hoja.Activate()
        except Exception:
            pass
        try:
            inv_app.ActiveView.Update()
        except Exception:
            pass
        try:
            for j in range(1, hoja.DrawingViews.Count + 1):
                v = hoja.DrawingViews.Item(j)
                try:
                    v.Update()
                except Exception:
                    pass
                try:
                    _ = v.DrawingCurves.Count
                except Exception:
                    sin_curvas += 1
        except Exception:
            pass
        procesadas += 1
    if log_fn:
        log_fn(
            f"  Compute forzado: {procesadas} hojas del lote listas "
            f"(sin_curvas={sin_curvas})"
        )


def ejecutar_flujo_desde_app(
    inv_app,
    ensamble_doc,
    doc,
    carpeta_salida=None,
    incremental=None,
    tam_lote=None,
):
    """
    Flujo por lotes (modo D + F):

    - Recolecta piezas del ensamble.
    - Si `incremental` está activo (o env ``PIEZAS_INCREMENTAL=1``), salta
      piezas cuyos JPG ya existan en ``carpeta_salida``.
    - Para cada lote de ``tam_lote`` piezas: crear vistas -> acotar -> renombrar
      -> exportar JPG -> borrar hojas del lote.

    Parámetros
    ----------
    incremental : bool | None
        Si None, se toma de env ``PIEZAS_INCREMENTAL`` (default False).
    tam_lote : int | None
        Si None, se toma de env ``PIEZAS_TAM_LOTE`` (default ``TAM_LOTE_PIEZAS``).
    """
    import traceback
    # IMPORTANTE: usar _importar_modulo para forzar reload desde disco.
    # ``import cotas`` / ``import THK`` cachean en sys.modules y perpetuan
    # versiones antiguas del código, lo que hace que los fixes NO se apliquen
    # hasta reiniciar Inventor por completo (efecto observado en corridas
    # 07:33 y 07:52 del 2026-08-13 donde THK.py tenía bugs corregidos en
    # disco pero seguían apareciendo en el log).
    _agregar_carpeta_actual_al_path()
    cotas = _importar_modulo("cotas")
    THK = _importar_modulo("THK")

    if incremental is None:
        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in ("1", "true", "TRUE", "yes")
    if tam_lote is None:
        try:
            tam_lote = int(os.environ.get("PIEZAS_TAM_LOTE", str(TAM_LOTE_PIEZAS)))
        except ValueError:
            tam_lote = TAM_LOTE_PIEZAS
    tam_lote = max(1, int(tam_lote))

    if getattr(sys, 'frozen', False):
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(log_dir, "error_log.txt")

    try:
        with open(log_path, "w") as f:
            f.write("Starting...\n")
    except:
        pass

    def log(msg):
        print(msg)
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(str(msg) + "\n")
                # Flush explícito: si Inventor crashea después de esta línea,
                # queremos que el log ya tenga escrita la última información
                # (permite saber dónde exactamente murió la corrida).
                try:
                    f.flush()
                    os.fsync(f.fileno())
                except Exception:
                    pass
        except:
            pass

    @contextlib.contextmanager
    def capturar_stdout_a_log(prefix=""):
        """
        Captura stdout de un bloque y lo vuelca al log del flujo con un
        prefijo opcional. Sirve para que los ``print`` de cotas.py / THK.py
        (que no usan `log()`) aparezcan también en ``error_log.txt`` y así
        se puedan diagnosticar corridas sin consola visible.
        """
        buf = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf
        try:
            yield buf
        finally:
            sys.stdout = old_stdout
            texto = buf.getvalue()
            if texto:
                for linea in texto.splitlines():
                    if prefix:
                        log(f"{prefix}{linea}")
                    else:
                        log(linea)

    log(f"Iniciando flujo para ensamble: {ensamble_doc.DisplayName}")
    log(f"Documento activo (Plano): {doc.DisplayName}")
    log(f"Modo por lotes: tam_lote={tam_lote} | incremental={incremental}")

    # Reinicia el registro global de pendientes THK para esta corrida.
    try:
        THK.reset_pendientes_thk()
    except Exception:
        pass

    # Nombre de la hoja machote (protegida contra borrado en lotes).
    nombre_machote = _nombre_hoja_machote(doc)
    if nombre_machote:
        log(f"Hoja machote protegida: {nombre_machote}")

    silent_prev = False
    screen_prev = True
    try:
        silent_prev = inv_app.SilentOperation
        screen_prev = inv_app.ScreenUpdating
        inv_app.SilentOperation = True
        inv_app.ScreenUpdating = True
    except Exception:
        pass

    _agregar_carpeta_actual_al_path()

    try:
        try:
            piezas = creador_vistas.recolectar_piezas_unicas(ensamble_doc)
        except Exception as e:
            log(f"❌ No se pudieron recolectar piezas: {e}")
            log(traceback.format_exc())
            return False

        if not piezas:
            log("AVISO: no hay piezas para procesar.")
            return True

        piezas_a_saltar = set()
        if incremental:
            piezas_a_saltar = _piezas_ya_exportadas(carpeta_salida)
            log(f"Modo incremental: {len(piezas_a_saltar)} piezas ya con JPG en destino.")

        # Filtro opcional para diagnóstico: PIEZAS_FILTRO="p1,p2" limita el
        # flujo a esas piezas (match case-insensitive contra el nombre corto
        # o el nombre completo del part). Útil para debug rápido sin re-correr
        # el tanque entero.
        filtro_raw = os.environ.get("PIEZAS_FILTRO", "").strip()
        filtro_up = set()
        if filtro_raw:
            filtro_up = {
                t.strip().upper()
                for t in filtro_raw.replace(";", ",").split(",")
                if t.strip()
            }
            log(f"Filtro de piezas activo (PIEZAS_FILTRO): {sorted(filtro_up)}")

        piezas_pendientes = []
        for part_doc, part_name in piezas:
            base = creador_vistas.obtener_nombre_base_corto(part_name)
            if filtro_up:
                pn_up = str(part_name or "").upper()
                base_up = str(base or "").upper()
                if not any(f in pn_up or f in base_up for f in filtro_up):
                    continue
            if base in piezas_a_saltar:
                log(f"  ⏭️ {part_name} (base={base}) — ya exportada, se salta.")
                continue
            piezas_pendientes.append((part_doc, part_name))

        total_pendientes = len(piezas_pendientes)
        log(f"Piezas pendientes: {total_pendientes} de {len(piezas)} totales.")
        if total_pendientes == 0:
            log("✅ Nada por hacer (todas las piezas ya tenían JPG). Fin.")
            return True

        contador_global = 0
        primer_lote = True
        for idx_lote, lote in enumerate(_chunks(piezas_pendientes, tam_lote), start=1):
            log(f"\n===== LOTE {idx_lote} — {len(lote)} piezas ({contador_global + 1}..{contador_global + len(lote)} de {total_pendientes}) =====")

            try:
                inv_app.ScreenUpdating = True
            except Exception:
                pass

            # 1) Crear vistas del lote
            log(f"  [chk] LOTE {idx_lote}: creando vistas...")
            try:
                ok_lote, nombres_hojas = creador_vistas.crear_vistas_lote(
                    inv_app,
                    ensamble_doc,
                    doc,
                    lote,
                    contador_inicio=contador_global,
                )
                contador_global += len(lote)
                if not ok_lote or not nombres_hojas:
                    log(f"  AVISO lote {idx_lote}: sin hojas creadas, se salta.")
                    continue
            except Exception as e:
                log(f"❌ Error creando vistas del lote {idx_lote}: {e}")
                log(traceback.format_exc())
                continue

            # Mantener ScreenUpdating=True mientras se aplican cotas / THK.
            # Con ScreenUpdating=False Inventor puede diferir el cálculo de
            # `DrawingCurves` y todas las hojas terminan sin geometría 2D
            # visible para las funciones de cotado.
            try:
                inv_app.ScreenUpdating = True
            except Exception:
                pass

            _actualizar_inventor(inv_app)
            time.sleep(0.3)

            # Forzar computación de todas las hojas del lote antes de cotarlas.
            _forzar_compute_hojas(doc, inv_app, nombres_hojas, log_fn=log)

            # 2) Cotas frentes
            log(f"  [chk] LOTE {idx_lote}: iniciando cotas linales...")
            try:
                with capturar_stdout_a_log(prefix="  [cotas] "):
                    cotas.acotar_planos(
                        nombres_permitidos=nombres_hojas,
                        reset_diametro=primer_lote,
                    )
                _actualizar_inventor(inv_app)
                time.sleep(0.3)
            except Exception as e:
                log(f"❌ Error en cotas.acotar_planos (lote {idx_lote}): {e}")
                log(traceback.format_exc())
            log(f"  [chk] LOTE {idx_lote}: cotas linales terminadas.")

            # 3) Cotas THK (LADO) + posible 4ª hoja ALTO (perfiles U/L)
            log(f"  [chk] LOTE {idx_lote}: iniciando THK...")
            hojas_extra = set()
            try:
                with capturar_stdout_a_log(prefix="  [THK] "):
                    resultado_thk = THK.acotar_thk(nombres_permitidos=nombres_hojas)
                if resultado_thk:
                    hojas_extra = set(resultado_thk)
                _actualizar_inventor(inv_app)
                time.sleep(0.3)
            except Exception as e:
                log(f"❌ Error en THK.acotar_thk (lote {idx_lote}): {e}")
                log(traceback.format_exc())
            log(f"  [chk] LOTE {idx_lote}: THK terminado.")

            nombres_para_rename = set(nombres_hojas) | hojas_extra

            # Snapshot de pendientes THK acumulados HASTA el lote actual, para
            # que el renombrado y la exportación conserven ``_LADO`` en las
            # hojas que quedaron sin cota real de THK.
            try:
                pendientes_snapshot = list(
                    getattr(THK, "LAST_PENDIENTES_THK", []) or []
                )
            except Exception:
                pendientes_snapshot = []

            # 4) Renombrar hojas del lote (obtenemos nombres finales para export/borrar).
            log(f"  [chk] LOTE {idx_lote}: renombrando hojas...")
            mapeo = {}
            try:
                mapeo = renombrar_hojas_finales(
                    doc,
                    nombres_permitidos=nombres_para_rename,
                    hojas_lado_sin_thk=pendientes_snapshot,
                ) or {}
            except Exception as e:
                log(f"AVISO renombrando hojas del lote {idx_lote}: {e}")

            nombres_finales = set()
            for original in nombres_para_rename:
                nombre_final = mapeo.get(original.upper())
                if nombre_final:
                    base_final, _ = _separar_nombre_hoja(nombre_final)
                    nombres_finales.add(base_final.upper())
                else:
                    nombres_finales.add(original.upper())

            _actualizar_inventor(inv_app)
            time.sleep(0.3)

            # 5) Exportar JPG del lote
            log(f"  [chk] LOTE {idx_lote}: exportando JPGs ({len(nombres_finales)} hojas)...")
            try:
                exportar_hojas_jpg(
                    inv_app,
                    doc,
                    carpeta_salida=carpeta_salida,
                    nombres_permitidos=nombres_finales,
                    hojas_lado_sin_thk=pendientes_snapshot,
                )
            except Exception as e:
                log(f"❌ Error exportando JPG del lote {idx_lote}: {e}")
                log(traceback.format_exc())
            log(f"  [chk] LOTE {idx_lote}: JPGs exportados.")

            # 6) Borrar hojas del lote para liberar memoria
            try:
                borrar_hojas_por_nombres(
                    doc,
                    nombres_finales,
                    nombre_machote_protegido=nombre_machote,
                )
            except Exception as e:
                log(f"AVISO borrando hojas del lote {idx_lote}: {e}")

            _actualizar_inventor(inv_app)
            time.sleep(0.4)

            # Checkpoint de reporte INCREMENTAL: si Inventor crashea a la
            # mitad de la corrida, el usuario ya tiene un reporte parcial
            # con las piezas sin THK acumuladas hasta este lote.
            try:
                _volcar_reporte_pendientes(THK, log)
            except Exception:
                pass

            primer_lote = False
            log(f"  [chk] LOTE {idx_lote}: terminado.")

        _volcar_reporte_pendientes(THK, log, cabecera_final=True)

        log("\n🎉 Flujo por lotes terminado.")
        return True
    finally:
        try:
            _volcar_reporte_pendientes(THK, log, cabecera_final=True)
        except Exception:
            pass
        try:
            inv_app.SilentOperation = silent_prev
            inv_app.ScreenUpdating = screen_prev
        except Exception:
            pass


def _volcar_reporte_pendientes(THK, log, cabecera_final=False):
    """Persiste ``piezas_sin_cotas.txt`` con las hojas ``_LADO`` cuyo THK no
    fue resuelto. Se llama incrementalmente al final de cada lote (para no
    perder el reporte si Inventor crashea a la mitad) y al final del flujo.
    """
    try:
        pendientes_thk = list(getattr(THK, "LAST_PENDIENTES_THK", []) or [])
    except Exception:
        pendientes_thk = []

    if cabecera_final and pendientes_thk:
        log("")
        log(f"⚠️ Piezas SIN cota THK ({len(pendientes_thk)}):")
        for h in pendientes_thk:
            log(f"   - {h}")

    try:
        ruta_reporte = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "piezas_sin_cotas.txt",
        )
        if pendientes_thk:
            with open(ruta_reporte, "w", encoding="utf-8") as f:
                f.write(
                    f"Piezas SIN cota THK - {len(pendientes_thk)} hoja(s)\n"
                )
                f.write("=" * 60 + "\n")
                for h in pendientes_thk:
                    f.write(h + "\n")
            if cabecera_final:
                log(f"  Reporte guardado en: {ruta_reporte}")
        else:
            # Si no hay pendientes, dejar archivo vacío o borrado.
            if os.path.exists(ruta_reporte):
                try:
                    os.remove(ruta_reporte)
                except Exception:
                    pass
    except Exception as err:
        if cabecera_final:
            log(f"  AVISO al escribir reporte: {err}")


# ============================================================
# FLUJO COMPLETO (Legacy con iLogic)
# ============================================================
def ejecutar_flujo_completo(carpeta_salida=None):
    print("Conectando con Inventor...")
    inv_app = conectar_inventor()

    doc = _obtener_o_activar_plano(inv_app)

    if doc is None:
        print("ERROR: No hay ningun plano (.idw / .dwg) abierto en Inventor.")
        print("   Abre el machote de planos e intenta de nuevo.")
        return False

    try:
        print(f"Documento activo: {doc.DisplayName}")
    except Exception:
        pass

    # 12292 = kDrawingDocumentObject
    if doc.DocumentType != 12292:
        print("ERROR: Debes tener un plano de Inventor (.idw / .dwg) abierto.")
        return False

    print("🚀 Buscando el motor de iLogic en tu sistema...")
    iLogicAutomation = _obtener_ilogic_automation(inv_app)

    if iLogicAutomation is None:
        print("❌ Error: No se encontró iLogic. ¿Está activado en tu Inventor?")
        return False

    print("✅ Motor iLogic encontrado.")

    print(f"⏳ Paso 1/4: Ejecutando regla '{NOMBRE_REGLA_ILOGIC}'...")
    try:
        iLogicAutomation.RunRule(doc, NOMBRE_REGLA_ILOGIC)
        _actualizar_inventor(inv_app)
        time.sleep(0.5)
        print("✅ Vistas generadas correctamente.")
    except Exception as e:
        print(f"❌ Hubo un problema al correr la regla '{NOMBRE_REGLA_ILOGIC}': {e}")
        return False

    carpeta_actual = _agregar_carpeta_actual_al_path()

    print("⏳ Paso 2/4: Ejecutando cotas de frentes con 'cotas.py'...")
    try:
        cotas = _importar_modulo("cotas")

        if not hasattr(cotas, "acotar_planos"):
            print("❌ Error: 'cotas.py' no tiene la función 'acotar_planos()'.")
            return False

        cotas.acotar_planos()
        _actualizar_inventor(inv_app)
        time.sleep(0.5)
        print("✅ Cotas de frentes terminadas.")
    except Exception as e:
        print(f"❌ Error al ejecutar 'cotas.py': {e}")
        print(f"📁 Carpeta usada para importar módulos: {carpeta_actual}")
        return False

    print("⏳ Paso 3/4: Ejecutando cotas de lado con 'THK.py'...")
    try:
        THK = _importar_modulo("THK")

        if not hasattr(THK, "acotar_thk"):
            print("❌ Error: 'THK.py' no tiene la función 'acotar_thk()'.")
            return False

        THK.acotar_thk()
        _actualizar_inventor(inv_app)
        time.sleep(0.5)
        print("✅ Cotas de lado terminadas.")
    except Exception as e:
        print(f"❌ Error al ejecutar 'THK.py': {e}")
        print(f"📁 Carpeta usada para importar módulos: {carpeta_actual}")
        return False

    print("⏳ Paso 3.5/4: Renombrando hojas finales...")
    renombrar_hojas_finales(doc)
    _actualizar_inventor(inv_app)
    time.sleep(0.5)

    exportar_hojas_jpg(inv_app, doc, carpeta_salida=carpeta_salida)

    print("\n🎉 Flujo completo terminado:")
    print("   1) Generar vistas")
    print("   2) Cotar frentes")
    print("   3) Cotar lados")
    print("   4) Exportar JPG")
    return True


if __name__ == "__main__":
    ok = ejecutar_flujo_completo()
    sys.exit(0 if ok else 1)