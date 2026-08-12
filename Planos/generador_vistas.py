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
    print(r'C:\Users\jose_rosales\AppData\Local\Programs\Python\Python314\python.exe -m pip install Pillow')
    raise

import creador_vistas


# ============================================================
# CONFIGURACIÓN
# ============================================================
NOMBRE_REGLA_ILOGIC = "GenerarVistas"
CARPETA_EXPORTACION = "JPG"
RUTA_HOJAS_DIAMETRO = r"C:\Temp\hojas_para_diametro.txt"

ANCHO_EXPORTACION = 3000
ALTO_EXPORTACION = 2200
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
        time.sleep(0.05)
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
    Lee C:\\Temp\\hojas_para_diametro.txt y devuelve:
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


def _convertir_nombre_tecnico_hoja(nombre: str, hojas_diametro_visibles=None, hojas_diametro_bases=None) -> str:
    """
    Convierte nombres finales de hoja.

    Regla:
    - Si la hoja fue mandada a diámetro:
        FRENTE_1 -> DIAMETRO_EXTERIOR
        FRENTE_2 -> DIAMETRO_INTERIOR
    - Si no fue mandada a diámetro:
        FRENTE_1 -> LARGO
        FRENTE_2 -> ANCHO
    - LADO siempre -> THK
    """
    if hojas_diametro_visibles is None:
        hojas_diametro_visibles = set()

    if hojas_diametro_bases is None:
        hojas_diametro_bases = set()

    base, sufijo = _separar_nombre_hoja(nombre)

    base_up = base.upper()
    visible_up = nombre.upper()

    es_diametro = (
        visible_up in hojas_diametro_visibles
        or base_up in hojas_diametro_bases
    )

    if "_LADO" in base:
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


def renombrar_hojas_finales(doc):
    """
    Renombra hojas al final del flujo:
    - Circulares:
        FRENTE_1 -> DIAMETRO_EXTERIOR
        FRENTE_2 -> DIAMETRO_INTERIOR
    - No circulares:
        FRENTE_1 -> LARGO
        FRENTE_2 -> ANCHO
    - LADO -> THK
    """
    try:
        draw_doc = win32com.client.CastTo(doc, "DrawingDocument")
    except:
        draw_doc = doc

    hojas_diametro_visibles, hojas_diametro_bases = _cargar_hojas_para_diametro()

    cambios = 0

    try:
        total_hojas = draw_doc.Sheets.Count
    except Exception as e:
        print(f"⚠️ No se pudieron leer las hojas para renombrar: {e}")
        return

    for i in range(1, total_hojas + 1):
        try:
            hoja = draw_doc.Sheets.Item(i)

            nombre_actual_visible = str(hoja.Name)
            nombre_actual_base, _ = _separar_nombre_hoja(nombre_actual_visible)

            nombre_nuevo_visible = _convertir_nombre_tecnico_hoja(
                nombre_actual_visible,
                hojas_diametro_visibles,
                hojas_diametro_bases
            )

            nombre_nuevo_base, _ = _separar_nombre_hoja(nombre_nuevo_visible)

            if nombre_nuevo_base != nombre_actual_base:
                hoja.Name = nombre_nuevo_base
                time.sleep(0.02)

                try:
                    nombre_resultado = str(hoja.Name)
                except:
                    nombre_resultado = nombre_nuevo_visible

                print(f"📝 Renombrada hoja: {nombre_actual_visible}  ->  {nombre_resultado}")
                cambios += 1

        except Exception as e:
            print(f"⚠️ No se pudo renombrar una hoja: {e}")

    print(f"✅ Renombrado final terminado. Hojas cambiadas: {cambios}")    
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


def _obtener_bbox_hoja(hoja):
    """
    Obtiene el área útil de la hoja con base en:
    - DrawingViews
    - GeneralDimensions
    Así ignoramos el marco del machote y el title block.
    """
    bbox = None

    try:
        total_views = hoja.DrawingViews.Count
    except:
        total_views = 0

    for i in range(1, total_views + 1):
        try:
            view = hoja.DrawingViews.Item(i)

            left = float(view.Left)
            top = float(view.Top)
            width = float(view.Width)
            height = float(view.Height)

            right = left + width
            bottom = top - height

            bbox = _expandir_bbox(bbox, left, right, bottom, top)
        except:
            pass

    try:
        total_dims = hoja.DrawingDimensions.GeneralDimensions.Count
    except:
        total_dims = 0

    for i in range(1, total_dims + 1):
        try:
            dim = hoja.DrawingDimensions.GeneralDimensions.Item(i)
            box = dim.RangeBox

            minx = float(box.MinPoint.X)
            maxx = float(box.MaxPoint.X)
            miny = float(box.MinPoint.Y)
            maxy = float(box.MaxPoint.Y)

            bbox = _expandir_bbox(bbox, minx, maxx, miny, maxy)
        except:
            pass

    return bbox


def _bbox_hoja_a_pixeles(hoja, bbox_hoja, img_w, img_h, margen_ratio=0.05):
    """
    Convierte el bbox de coordenadas de hoja a coordenadas de imagen.
    """
    sheet_w = float(hoja.Width)
    sheet_h = float(hoja.Height)

    minx, maxx, miny, maxy = bbox_hoja

    margen_x = sheet_w * margen_ratio
    margen_y = sheet_h * margen_ratio

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


def _recortar_exportacion_jpg(hoja, ruta_temporal, ruta_final):
    """
    Recorta el JPG exportado usando solo el contenido útil de la hoja.
    """
    bbox_hoja = _obtener_bbox_hoja(hoja)

    if not bbox_hoja:
        os.replace(ruta_temporal, ruta_final)
        return

    img = Image.open(ruta_temporal)
    img_w, img_h = img.size

    left_px, upper_px, right_px, lower_px = _bbox_hoja_a_pixeles(
        hoja, bbox_hoja, img_w, img_h, margen_ratio=MARGEN_RECORTE
    )

    recorte = img.crop((left_px, upper_px, right_px, lower_px))
    recorte = recorte.convert("RGB")
    recorte.save(ruta_final, quality=95, subsampling=0)

    img.close()
    try:
        os.remove(ruta_temporal)
    except:
        pass


# ============================================================
# EXPORTACIÓN A JPG
# ============================================================
def exportar_hojas_jpg(inv_app, doc, carpeta_salida=None):
    """
    Exporta cada hoja con vistas a JPG, excluyendo hojas vacías.
    Guarda en una carpeta junto al archivo machote o en la ruta explícita.
    Además recorta automáticamente la hoja para dejar solo el contenido útil.
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

    for i in range(1, total_hojas + 1):
        try:
            hoja = draw_doc.Sheets.Item(i)
        except Exception as e:
            print(f"⚠️ No se pudo acceder a la hoja #{i}: {e}")
            continue

        try:
            nombre_hoja = str(hoja.Name)

            hojas_diametro_visibles, hojas_diametro_bases = _cargar_hojas_para_diametro()
            nombre_hoja = _convertir_nombre_tecnico_hoja(
                nombre_hoja,
                hojas_diametro_visibles,
                hojas_diametro_bases
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
            hoja.Activate()
            _actualizar_inventor(inv_app)
            time.sleep(0.05)

            try:
                inv_app.ActiveView.Fit()
            except:
                pass

            _actualizar_inventor(inv_app)
            time.sleep(0.05)

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

        except Exception as e:
            print(f"⚠️ No se pudo exportar la hoja '{nombre_hoja}': {e}")

    print(f"✅ Exportación terminada. JPG creados: {exportadas} | Hojas omitidas: {omitidas}")
    print(f"📁 Carpeta de salida: {carpeta_salida}")


# ============================================================
# FLUJO COMPLETO (Desde la App con UI y API Pura)
# ============================================================
def ejecutar_flujo_desde_app(
    inv_app, ensamble_doc, doc, carpeta_salida=None
):
    import traceback
    import cotas
    import THK
    
    if getattr(sys, 'frozen', False):
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(log_dir, "error_log.txt")
    
    try:
        with open(log_path, "w") as f: f.write("Starting...\n")
    except:
        pass
        
    def log(msg):
        print(msg)
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(str(msg) + "\n")
        except:
            pass

    log(f"Iniciando flujo para ensamble: {ensamble_doc.DisplayName}")
    log(f"Documento activo (Plano): {doc.DisplayName}")

    # Modo silencioso: suprime redibujado de la UI mientras dura el flujo
    # completo de piezas. Reduce >20% del tiempo en tanques grandes.
    silent_prev = False
    screen_prev = True
    try:
        silent_prev = inv_app.SilentOperation
        screen_prev = inv_app.ScreenUpdating
        inv_app.SilentOperation = True
        inv_app.ScreenUpdating = False
    except Exception:
        pass

    try:
        try:
            ok_vistas = creador_vistas.crear_vistas(inv_app, ensamble_doc, doc)
            if not ok_vistas:
                log("❌ Hubo un problema al generar las vistas con Python.")
                return False
            _actualizar_inventor(inv_app)
            time.sleep(0.1)
        except Exception as e:
            log(f"❌ Error al ejecutar 'creador_vistas.py': {e}")
            log(traceback.format_exc())
            return False

        carpeta_actual = _agregar_carpeta_actual_al_path()

        log("⏳ Paso 2/4: Ejecutando cotas de frentes con 'cotas.py'...")
        try:
            cotas.acotar_planos()
            _actualizar_inventor(inv_app)
            time.sleep(0.1)
            log("✅ Cotas de frentes terminadas.")
        except Exception as e:
            log(f"❌ Error al ejecutar 'cotas.py': {e}")
            log(traceback.format_exc())
            return False

        log("⏳ Paso 3/4: Ejecutando cotas de lado con 'THK.py'...")
        try:
            THK.acotar_thk()
            _actualizar_inventor(inv_app)
            time.sleep(0.1)
            log("✅ Cotas de lado terminadas.")
        except Exception as e:
            log(f"❌ Error al ejecutar 'THK.py': {e}")
            log(traceback.format_exc())
            return False

        log("⏳ Paso 3.5/4: Renombrando hojas finales...")
        try:
            renombrar_hojas_finales(doc)
        except Exception as e:
            log(f"Error en renombrar_hojas_finales: {e}")
            log(traceback.format_exc())

        _actualizar_inventor(inv_app)
        time.sleep(0.1)

        try:
            exportar_hojas_jpg(
                inv_app, doc, carpeta_salida=carpeta_salida
            )
        except Exception as e:
            log(f"Error en exportar_hojas_jpg: {e}")
            log(traceback.format_exc())

        log("\n🎉 Flujo completo terminado (Vía App):")
        return True
    finally:
        try:
            inv_app.SilentOperation = silent_prev
            inv_app.ScreenUpdating = screen_prev
        except Exception:
            pass


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
        time.sleep(0.1)
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
        time.sleep(0.1)
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
        time.sleep(0.1)
        print("✅ Cotas de lado terminadas.")
    except Exception as e:
        print(f"❌ Error al ejecutar 'THK.py': {e}")
        print(f"📁 Carpeta usada para importar módulos: {carpeta_actual}")
        return False

    print("⏳ Paso 3.5/4: Renombrando hojas finales...")
    renombrar_hojas_finales(doc)
    _actualizar_inventor(inv_app)
    time.sleep(0.1)

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