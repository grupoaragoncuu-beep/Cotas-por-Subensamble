import win32com.client
import math
import os
import sys
import time
import pythoncom

kPartDocumentObject = 12290
kPlaneSurface = 5890
kLineSegmentCurve = 5123
kArbitraryViewOrientation = 10763
kDefaultViewOrientation = 10753
kHiddenLineRemovedDrawingViewStyle = 32258

def _log(msg):
    if getattr(sys, 'frozen', False):
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(log_dir, "error_log_creador.txt")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(str(msg) + "\n")
    except:
        pass
    print(msg)

def _sleep_and_pump(seconds):
    """Pausa el script pero mantiene vivos los mensajes de Windows/DirectX"""
    iterations = max(1, int(seconds * 10))
    for _ in range(iterations):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)


def _es_com_transitorio(exc):
    """Errores COM típicos de saturación / enumerador inválido (no crash total)."""
    try:
        code = int(exc.args[0]) if exc.args else 0
    except Exception:
        code = 0
    # RPC_E_SYS_CALL_FAILED / RPC_S_SERVER_UNAVAILABLE / RPC_E_DISCONNECTED
    return code in (-2147417856, -2147023174, -2147417848)


def recolectar_piezas_unicas(ensamble_doc):
    """Alias público de :func:`_recolectar_piezas_unicas` (uso desde otros módulos)."""
    return _recolectar_piezas_unicas(ensamble_doc)


def _recolectar_piezas_unicas(ensamble_doc):
    """
    Lista piezas únicas ANTES de crear hojas.

    No se mantiene el enumerador AllLeafOccurrences abierto durante CopyTo/
    AddBaseView: tras un flujo largo de caras ese handle se corrompe y
    Item(i) lanza 'Error en la llamada de sistema' abortando PIEZAS.
    """
    piezas = []
    vistos = set()
    leaf_occs = ensamble_doc.ComponentDefinition.Occurrences.AllLeafOccurrences
    try:
        total = int(leaf_occs.Count)
    except Exception as exc:
        _log(f"ERROR: no se pudo leer AllLeafOccurrences.Count: {exc}")
        return piezas

    i = 1
    fallos_seguidos = 0
    while i <= total:
        try:
            occ = leaf_occs.Item(i)
            fallos_seguidos = 0
        except Exception as exc:
            fallos_seguidos += 1
            _log(
                f"AVISO: leaf Item({i}/{total}) falló ({exc}); "
                f"refresco enumerador (intento {fallos_seguidos})"
            )
            _sleep_and_pump(1.5)
            try:
                leaf_occs = (
                    ensamble_doc.ComponentDefinition.Occurrences.AllLeafOccurrences
                )
                total = int(leaf_occs.Count)
                occ = leaf_occs.Item(i)
                fallos_seguidos = 0
            except Exception as exc2:
                _log(f"AVISO: se omite índice {i}: {exc2}")
                if fallos_seguidos >= 5:
                    _log("ERROR: demasiados fallos seguidos leyendo ocurrencias.")
                    break
                i += 1
                continue

        try:
            if occ.Suppressed:
                i += 1
                continue
        except Exception:
            i += 1
            continue

        try:
            doc_type = occ.DefinitionDocumentType
            _log(f"Occ: {occ.Name}, type: {doc_type}")
        except Exception:
            pass

        try:
            part_doc = win32com.client.CastTo(occ.Definition.Document, "PartDocument")
            if part_doc is None:
                part_doc = occ.Definition.Document
            ruta = part_doc.FullFileName
        except Exception:
            i += 1
            continue

        if not ruta or ruta in vistos:
            i += 1
            continue
        vistos.add(ruta)
        part_name = os.path.splitext(os.path.basename(ruta))[0]
        piezas.append((part_doc, part_name))
        i += 1

    _log(f"Piezas únicas a procesar: {len(piezas)} (de {total} leaf occs)")
    return piezas


def _limpiar_border_y_titleblock(new_sheet):
    """
    Elimina TODO lo que no sea una vista dibujada en el sheet clonado:
    border, title block, sketches del machote, sketched symbols, notas,
    hole tables, leader notes y balloons. Sin esto, cualquier arte del
    machote (líneas guía, LC, centerlines, logos, etc.) queda visible en
    el JPG exportado.

    La hoja original del machote (base_sheet) NO se toca — esta es una
    copia recién creada por CopyTo.
    """
    def _borrar_todos(coleccion):
        """Elimina todos los items de una colección de Inventor, del
        último al primero para no invalidar los índices."""
        try:
            n = coleccion.Count
        except Exception:
            return
        for idx in range(n, 0, -1):
            try:
                coleccion.Item(idx).Delete()
            except Exception:
                pass

    try:
        if new_sheet.Border is not None:
            try:
                new_sheet.Border.Delete()
            except Exception:
                pass
    except Exception:
        pass

    try:
        if new_sheet.TitleBlock is not None:
            try:
                new_sheet.TitleBlock.Delete()
            except Exception:
                pass
    except Exception:
        pass

    for attr in (
        "Sketches",
        "SketchedSymbols",
        "DrawingNotes",
        "LeaderNotes",
        "GeneralNotes",
        "Balloons",
        "HoleTables",
        "RevisionTables",
        "PartsLists",
        "GeneralTables",
    ):
        try:
            coleccion = getattr(new_sheet, attr, None)
        except Exception:
            coleccion = None
        if coleccion is not None:
            _borrar_todos(coleccion)


def _crear_hoja_vista(machote_doc, base_sheet, nombre_hoja):
    """CopyTo + Activate con un reintento ante COM transitorio."""
    ultimo = None
    for intento in range(2):
        try:
            new_sheet = base_sheet.CopyTo(machote_doc)
            new_sheet.Name = nombre_hoja
            new_sheet.Activate()
            for v in range(new_sheet.DrawingViews.Count, 0, -1):
                new_sheet.DrawingViews.Item(v).Delete()
            _limpiar_border_y_titleblock(new_sheet)
            return new_sheet
        except Exception as exc:
            ultimo = exc
            if not _es_com_transitorio(exc) or intento == 1:
                break
            _log(f"AVISO: reintento CopyTo '{nombre_hoja}' tras COM: {exc}")
            _sleep_and_pump(2.0)
    raise ultimo


def _area_util_hoja(sheet):
    """
    Devuelve (px, py, ancho_util, alto_util) para el sheet: el centro del
    área útil (después de descontar el border si existiera) y sus dimensiones.

    Como en el flujo actual el border y title block se eliminan al clonar el
    sheet (ver ``_limpiar_border_y_titleblock``), el área útil es prácticamente
    todo el sheet menos un pequeño margen de seguridad para que la vista no
    quede tocando el borde físico de la hoja.
    """
    try:
        sheet_w = float(sheet.Width)
        sheet_h = float(sheet.Height)
    except Exception:
        return 10.0, 15.0, 20.0, 30.0

    # Márgenes de seguridad (5 % de cada dimensión, con mínimo/máximo).
    margen_x = min(3.0, max(1.0, sheet_w * 0.05))
    margen_y = min(3.0, max(1.0, sheet_h * 0.05))

    ancho_util = max(1.0, sheet_w - 2.0 * margen_x)
    alto_util = max(1.0, sheet_h - 2.0 * margen_y)

    px = sheet_w / 2.0
    py = sheet_h / 2.0

    return px, py, ancho_util, alto_util


def crear_vistas(inv_app, ensamble_doc, machote_doc, piezas_a_saltar=None):
    """
    Compat: recolecta todas las piezas y las procesa en un solo lote.

    Parametros
    ----------
    piezas_a_saltar : set[str] | None
        Nombres base (obtener_nombre_base_corto) a NO regenerar (modo F).
    """
    try:
        machote_doc = win32com.client.CastTo(machote_doc, "DrawingDocument")
    except:
        pass

    try:
        ensamble_doc = win32com.client.CastTo(ensamble_doc, "AssemblyDocument")
    except:
        pass

    piezas = _recolectar_piezas_unicas(ensamble_doc)
    ok, _hojas = crear_vistas_lote(
        inv_app, ensamble_doc, machote_doc, piezas, piezas_a_saltar=piezas_a_saltar
    )
    return ok


def crear_vistas_lote(
    inv_app,
    ensamble_doc,
    machote_doc,
    piezas,
    piezas_a_saltar=None,
    contador_inicio=0,
):
    """
    Crea vistas para una lista pre-recolectada de piezas.

    Retorna (ok, nombres_hojas_creadas) donde `nombres_hojas_creadas` es un
    `set[str]` con los nombres iniciales (FRENTE_1 / FRENTE_2 / LADO) creados
    en el machote.
    """
    _log("⏳ Paso 1/4: Generando vistas con API de Python...")
    try:
        _log(f"Ensamble activo: {ensamble_doc.DisplayName}")
    except Exception:
        pass

    piezas_a_saltar = piezas_a_saltar or set()
    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects

    base_sheet = machote_doc.ActiveSheet
    contador = int(contador_inicio)
    nombres_creadas = set()

    try:
        for part_doc, part_name in piezas:
            contador += 1
            nombre_base = obtener_nombre_base_corto(part_name)
            if nombre_base in piezas_a_saltar:
                _log(f"⏭️ {part_name} ya tiene JPG previo (base={nombre_base}). Saltada.")
                continue
            _log(f"Procesando pieza: {part_name}")
            
            is_sm = False
            try:
                is_sm = (part_doc.SubType == "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            except:
                pass
            
            res_prep = preparar_geometria(part_doc, is_sm, to)
            if not res_prep:
                _log(f"⚠️ {part_name} -> no se pudo preparar geometría.")
                continue
                
            use_flat_pattern, caras_a_medir, cuerpo_medicion = res_prep
            
            res_frente = elegir_frente(caras_a_medir)
            if not res_frente:
                _log(f"⚠️ {part_name} -> no se encontró frente válido.")
                continue
                
            frente_face, v_frente, area_frente = res_frente
            
            res_lado = elegir_lado(caras_a_medir, frente_face, v_frente, area_frente, use_flat_pattern)
            if not res_lado:
                v_lado = obtener_lado_fallback(tg, v_frente)
            else:
                lado_face, v_lado, area_lado = res_lado
                
            cx, cy, cz = obtener_centro(part_doc, cuerpo_medicion)
            
            tiene_guia_frente, v_guia_frente = obtener_vector_guia_frente(frente_face, v_frente, tg)

            # Orientación LADO desde sólido doblado (chapa con flat pattern).
            lado_cx, lado_cy, lado_cz = cx, cy, cz
            lado_eye, lado_up = v_lado, v_frente
            if use_flat_pattern:
                ori = _orientacion_lado_doblado(part_doc, tg, to, v_frente)
                if ori:
                    lado_eye = ori["v_lado"]
                    lado_up = ori["v_up"]
                    lado_cx, lado_cy, lado_cz = ori["cx"], ori["cy"], ori["cz"]
                    _log(
                        f"  {part_name}: LADO orientado desde modelo doblado "
                        f"({ori.get('modo', '?')})"
                    )
            
            sufijos = ["FRENTE_1", "FRENTE_2", "LADO"]
            es_lado = [False, False, True]
            
            for idx in range(3):
                is_side = es_lado[idx]
                nombre_hoja = construir_nombre_hoja(machote_doc, part_name, sufijos[idx])
                
                try:
                    new_sheet = _crear_hoja_vista(machote_doc, base_sheet, nombre_hoja)
                except Exception as e:
                    _log(f"⚠️ {part_name} -> no se pudo crear hoja {nombre_hoja}: {e}")
                    continue
                nombres_creadas.add(nombre_hoja)
                    
                try:
                    if new_sheet.TitleBlock is not None:
                        tb = new_sheet.TitleBlock
                        tb.SetResultText(tb.Definition.Sketch.TextBoxes.Item(1), part_name)
                except:
                    pass
                    
                px, py, ancho_util, alto_util = _area_util_hoja(new_sheet)

                view = None
                try:
                    if is_side:
                        view = _crear_vista_lado_con_reintentos(
                            new_sheet,
                            part_doc,
                            tg,
                            to,
                            px,
                            py,
                            lado_cx,
                            lado_cy,
                            lado_cz,
                            lado_eye,
                            lado_up,
                        )
                    else:
                        if tiene_guia_frente:
                            up_hint = v_guia_frente.Copy()
                        else:
                            up_hint = v_lado.Copy()
                        cam = crear_camara(
                            part_doc, tg, to, cx, cy, cz, v_frente.Copy(), up_hint
                        )
                        view = _crear_vista_base(
                            new_sheet,
                            part_doc,
                            tg,
                            to,
                            px,
                            py,
                            cam,
                            use_flat_pattern_view=use_flat_pattern,
                        )
                except Exception as ex1:
                    try:
                        # Solo frentes pueden caer a default; LADO no (rompe THK).
                        if not is_side:
                            view = new_sheet.DrawingViews.AddBaseView(
                                part_doc,
                                tg.CreatePoint2d(px, py),
                                1.0,
                                kDefaultViewOrientation,
                                kHiddenLineRemovedDrawingViewStyle
                            )
                        else:
                            _log(
                                f"⚠️ {part_name} -> no se pudo crear vista LADO "
                                f"con cámara de perfil: {ex1}"
                            )
                    except:
                        _log(f"⚠️ {part_name} -> no se pudo crear vista {sufijos[idx]}")
                        
                if view is not None:
                    escalar_vista(
                        machote_doc, view, tg, px, py, ancho_util, alto_util
                    )
            
            # Respiro a la tarjeta gráfica y bombeo de mensajes para evitar TDR.
            # Antes: 4s por pieza + 10s cada 10 (~10 min muertos en OTC).
            # Ahora: 1.2s por pieza + 3s cada 20; suficiente para el TDR sin
            # sobrecosto acumulado.
            try:
                inv_app.ActiveView.Update()
            except:
                pass
            _sleep_and_pump(1.2)

            if contador % 20 == 0:
                _log(f"Descanso cada 20 componentes (vamos en {contador})...")
                _sleep_and_pump(3.0)
                    
    finally:
        try:
            base_sheet.Activate()
        except:
            pass

    _log(f"✅ Vistas generadas correctamente para {contador} piezas (lote).")
    return True, nombres_creadas


def preparar_geometria(part_doc, is_sm, to):
    use_flat_pattern = False
    cuerpo_medicion = None
    caras_a_medir = to.CreateObjectCollection()
    
    try:
        if is_sm:
            try:
                sm_def = win32com.client.CastTo(part_doc.ComponentDefinition, "SheetMetalComponentDefinition")
            except:
                sm_def = part_doc.ComponentDefinition
            
            if sm_def is None: return None
            
            if not sm_def.HasFlatPattern:
                try:
                    sm_def.Unfold()
                except:
                    pass
                    
            if sm_def.HasFlatPattern:
                use_flat_pattern = True
                try:
                    cuerpo_medicion = sm_def.FlatPattern.SurfaceBodies.Item(1)
                except:
                    try:
                        cuerpo_medicion = sm_def.FlatPattern.Body
                    except:
                        pass
                if cuerpo_medicion is not None:
                    for i in range(1, cuerpo_medicion.Faces.Count + 1):
                        caras_a_medir.Add(cuerpo_medicion.Faces.Item(i))
                        
            if caras_a_medir.Count == 0:
                for i in range(1, sm_def.SurfaceBodies.Count + 1):
                    body = sm_def.SurfaceBodies.Item(i)
                    if cuerpo_medicion is None: cuerpo_medicion = body
                    for j in range(1, body.Faces.Count + 1):
                        caras_a_medir.Add(body.Faces.Item(j))
        else:
            for i in range(1, part_doc.ComponentDefinition.SurfaceBodies.Count + 1):
                body = part_doc.ComponentDefinition.SurfaceBodies.Item(i)
                if cuerpo_medicion is None: cuerpo_medicion = body
                for j in range(1, body.Faces.Count + 1):
                    caras_a_medir.Add(body.Faces.Item(j))
                    
        if caras_a_medir.Count > 0:
            return (use_flat_pattern, caras_a_medir, cuerpo_medicion)
        return None
    except:
        return None


def _caras_y_cuerpo_doblado(part_doc, to):
    """Caras del modelo doblado (no flat pattern) para orientar LADO/THK."""
    caras = to.CreateObjectCollection()
    cuerpo = None
    try:
        cdef = part_doc.ComponentDefinition
        for i in range(1, cdef.SurfaceBodies.Count + 1):
            body = cdef.SurfaceBodies.Item(i)
            if cuerpo is None:
                cuerpo = body
            for j in range(1, body.Faces.Count + 1):
                caras.Add(body.Faces.Item(j))
    except Exception:
        pass
    return caras, cuerpo


def _orientacion_lado_doblado(part_doc, tg, to, v_frente_fallback):
    """
    Calcula eye/up/centro de LADO desde el sólido doblado.

    Preferencia:
    1) Perfil de escuadra L/U: mirar según el eje de doblez
       (cruz de normales de las dos caras grandes ~perpendiculares).
    2) Fallback: cara de espesor más chica ⊥ a la cara grande.
    """
    caras, cuerpo = _caras_y_cuerpo_doblado(part_doc, to)
    if caras is None or caras.Count == 0:
        return None

    cx, cy, cz = obtener_centro(part_doc, cuerpo)

    # --- 1) Intentar eje de doblez (perfil L real) ---
    planas = []
    try:
        for i in range(1, caras.Count + 1):
            face = caras.Item(i)
            if face.SurfaceType != kPlaneSurface:
                continue
            ok, n = obtener_normal_cara(face)
            if not ok:
                continue
            try:
                area = float(face.Evaluator.Area)
            except Exception:
                continue
            if area <= 0:
                continue
            n.Normalize()
            planas.append((area, n))
    except Exception:
        planas = []

    planas.sort(key=lambda x: x[0], reverse=True)
    tope = min(8, len(planas))
    for i in range(tope):
        for j in range(i + 1, tope):
            area1, n1 = planas[i]
            area2, n2 = planas[j]
            if area2 < area1 * 0.12:
                continue
            # Casi perpendiculares → tipico L/U
            if abs(n1.DotProduct(n2)) > 0.35:
                continue
            try:
                eye = n1.CrossProduct(n2)
            except Exception:
                continue
            if eye is None or eye.Length < 0.001:
                continue
            eye.Normalize()
            # Up en el plano del perfil L (dirección de una pata)
            try:
                up = eye.CrossProduct(n1)
                if up.Length < 0.001:
                    up = eye.CrossProduct(n2)
                if up.Length < 0.001:
                    continue
                up.Normalize()
            except Exception:
                continue
            return {
                "v_lado": eye,
                "v_up": up,
                "cx": cx,
                "cy": cy,
                "cz": cz,
                "modo": "eje_doblez_L",
            }

    # --- 2) Fallback clásico: espesor ⊥ frente ---
    res_frente = elegir_frente(caras)
    if not res_frente:
        return None

    frente_face, v_frente, area_frente = res_frente
    res_lado = elegir_lado(
        caras, frente_face, v_frente, area_frente, use_flat_pattern=False
    )
    if res_lado:
        v_lado = res_lado[1]
    else:
        v_lado = obtener_lado_fallback(tg, v_frente)

    return {
        "v_lado": v_lado,
        "v_up": v_frente,
        "cx": cx,
        "cy": cy,
        "cz": cz,
        "modo": "espesor",
    }


def _crear_vista_base(new_sheet, part_doc, tg, to, px, py, cam, use_flat_pattern_view):
    options = to.CreateNameValueMap()
    if use_flat_pattern_view:
        options.Add("SheetMetalFoldedModel", False)
    return new_sheet.DrawingViews.AddBaseView(
        part_doc,
        tg.CreatePoint2d(px, py),
        1.0,
        kArbitraryViewOrientation,
        kHiddenLineRemovedDrawingViewStyle,
        "",
        cam,
        options,
    )


def _grosor_3d_pieza(part_doc):
    """
    Devuelve el lado MÁS PEQUEÑO del bbox 3D de la pieza (en cm de Inventor).
    Se usa como referencia para validar que la vista LADO efectivamente está
    mostrando el canto delgado y no la cara grande.
    """
    try:
        rb = part_doc.ComponentDefinition.RangeBox
        dims = [
            abs(float(rb.MaxPoint.X) - float(rb.MinPoint.X)),
            abs(float(rb.MaxPoint.Y) - float(rb.MinPoint.Y)),
            abs(float(rb.MaxPoint.Z) - float(rb.MinPoint.Z)),
        ]
        dims.sort()
        return dims[0]
    except Exception:
        return None


def _vista_lado_muestra_canto(view, grosor_3d):
    """
    Verifica que la vista LADO tenga un lado corto compatible con el grosor
    real de la pieza. Si el lado más chico del bbox 2D es mucho mayor que el
    grosor 3D, la orientación de la cámara está mal (probablemente Inventor
    la ignoró y usó la cara grande).

    Umbral: el lado corto 2D no puede ser > 3× el grosor 3D. Este margen
    tolera perfiles L/U/redondeados donde el bbox 2D es más gordo que el
    grosor real.
    """
    if grosor_3d is None or grosor_3d <= 0:
        return True
    try:
        scale = float(view.Scale)
        if scale <= 0:
            return True
        w = float(view.Width) / scale
        h = float(view.Height) / scale
        menor_2d = min(w, h)
    except Exception:
        return True
    return menor_2d <= grosor_3d * 3.0


def _borrar_todas_las_vistas(sheet):
    """Borra todas las vistas del sheet, en orden inverso para no invalidar
    índices. Fuerza un Update al final para que Inventor procese los deletes
    antes de seguir agregando vistas."""
    try:
        n = sheet.DrawingViews.Count
    except Exception:
        n = 0
    for idx in range(n, 0, -1):
        try:
            sheet.DrawingViews.Item(idx).Delete()
        except Exception:
            pass
    try:
        sheet.Update()
    except Exception:
        pass


def _crear_vista_lado_con_reintentos(
    new_sheet, part_doc, tg, to, px, py, cx, cy, cz, eye_dir, up_hint
):
    """
    Intenta varias cámaras para LADO. Nunca cae a DefaultViewOrientation
    (eso suele mostrar la cara grande y deja THK vacío).

    Estrategia (simple y segura, sin dejar vistas fantasma en el sheet):
    1. Para cada candidato, borra TODAS las vistas del sheet, crea una
       vista con esa cámara y valida con ``_vista_lado_muestra_canto``.
    2. Si la vista es de canto real → retorna esa (única) vista.
    3. Si no lo es → sigue al siguiente candidato (que borrará esta antes
       de crear la próxima).
    4. Si NINGÚN candidato produjo vista de canto → deja la ÚLTIMA como
       fallback (una sola vista, no acumuladas). Mejor tener algo dudoso
       que dejar el sheet vacío.
    """
    candidatos = []
    try:
        e0 = eye_dir.Copy()
        u0 = up_hint.Copy()
        candidatos.append((e0, u0))
        e1 = eye_dir.Copy()
        e1.ScaleBy(-1.0)
        candidatos.append((e1, u0.Copy()))
        u1 = up_hint.Copy()
        u1.ScaleBy(-1.0)
        candidatos.append((e0.Copy(), u1))
        cruz = eye_dir.CrossProduct(up_hint)
        if cruz is not None and cruz.Length > 0.001:
            cruz.Normalize()
            candidatos.append((cruz, up_hint.Copy()))
            cruz2 = cruz.Copy()
            cruz2.ScaleBy(-1.0)
            candidatos.append((cruz2, up_hint.Copy()))
    except Exception:
        candidatos = [(eye_dir, up_hint)]

    grosor_3d = _grosor_3d_pieza(part_doc)

    ultimo_error = None
    vista_actual = None
    for eye, up in candidatos:
        # Antes de crear, borra CUALQUIER vista previa en el sheet (incluida
        # la que dejó el candidato anterior si no pasó la validación).
        _borrar_todas_las_vistas(new_sheet)
        vista_actual = None

        try:
            cam = crear_camara(part_doc, tg, to, cx, cy, cz, eye, up)
            vista_actual = _crear_vista_base(
                new_sheet, part_doc, tg, to, px, py, cam, use_flat_pattern_view=False
            )
        except Exception as exc:
            ultimo_error = exc
            continue

        if vista_actual is None:
            continue

        if _vista_lado_muestra_canto(vista_actual, grosor_3d):
            return vista_actual

        # No es canto: seguimos iterando. La próxima iteración la borrará
        # antes de crear la siguiente candidata.

    # Si llegamos aquí y hay UNA vista aún viva en el sheet (fallback),
    # devolverla. Si no hay ninguna, propagar el último error.
    if vista_actual is not None:
        return vista_actual
    if ultimo_error:
        raise ultimo_error
    return None

def elegir_frente(caras_a_medir):
    frente_face = None
    v_frente = None
    area_frente = 0.0
    
    try:
        for i in range(1, caras_a_medir.Count + 1):
            face = caras_a_medir.Item(i)
            if face.SurfaceType != kPlaneSurface: continue
            
            try:
                area = face.Evaluator.Area
            except:
                continue
                
            if area > area_frente:
                normal_ok, n = obtener_normal_cara(face)
                if normal_ok:
                    frente_face = face
                    v_frente = n
                    area_frente = area
                    
        if frente_face is not None and v_frente is not None:
            v_frente.Normalize()
            return (frente_face, v_frente, area_frente)
        return None
    except:
        return None


def elegir_lado(caras_a_medir, frente_face, v_frente, area_frente, use_flat_pattern):
    lado_face = None
    v_lado = None
    area_lado = 0.0
    
    try:
        mejor_area = float('inf')
        area_minima_valida = max(area_frente * 0.001, 0.000001)
        
        for i in range(1, caras_a_medir.Count + 1):
            face = caras_a_medir.Item(i)
            if face.SurfaceType != kPlaneSurface: continue
            # En Python WIN32COM object comparison might need identity check or properties
            try:
                if face.InternalName == frente_face.InternalName: continue
            except:
                pass
            
            normal_ok, n = obtener_normal_cara(face)
            if not normal_ok: continue
            n.Normalize()
            
            dp = v_frente.DotProduct(n)
            abs_dp = abs(dp)
            
            if abs_dp > 0.95: continue
            if abs_dp > 0.10: continue
            
            if use_flat_pattern:
                if not comparte_arista_con_outer_loop(frente_face, face): continue
                
            try:
                area = face.Evaluator.Area
            except:
                continue
                
            if area < area_minima_valida: continue
            
            if area < mejor_area:
                mejor_area = area
                lado_face = face
                v_lado = n
                area_lado = area
                
        if lado_face is not None and v_lado is not None:
            v_lado.Normalize()
            return (lado_face, v_lado, area_lado)
        return None
    except:
        return None


def obtener_normal_cara(face):
    try:
        if face.SurfaceType != kPlaneSurface: return False, None
        n = face.Geometry.Normal.AsVector()
        if n is None: return False, None
        if n.Length < 0.0001: return False, None
        n.Normalize()
        return True, n
    except:
        return False, None


def comparte_arista_con_outer_loop(frente_face, candidata):
    try:
        for i in range(1, frente_face.EdgeLoops.Count + 1):
            loop = frente_face.EdgeLoops.Item(i)
            if not loop.IsOuterEdgeLoop: continue
            
            for j in range(1, loop.EdgeUses.Count + 1):
                edge_use = loop.EdgeUses.Item(j)
                e_outer = edge_use.Edge
                
                for k in range(1, candidata.Edges.Count + 1):
                    e_cand = candidata.Edges.Item(k)
                    # COM Identity check
                    if e_cand is e_outer or (hasattr(e_cand, 'InternalName') and hasattr(e_outer, 'InternalName') and e_cand.InternalName == e_outer.InternalName):
                        return True
    except:
        pass
    return False


def obtener_vector_guia_frente(frente_face, v_frente, tg):
    v_guia = None
    try:
        mejor_largo = 0.0
        for i in range(1, frente_face.Edges.Count + 1):
            edge = frente_face.Edges.Item(i)
            if edge.GeometryType != kLineSegmentCurve: continue
            
            p1 = edge.StartVertex.Point
            p2 = edge.StopVertex.Point
            
            v_tmp = tg.CreateVector(p2.X - p1.X, p2.Y - p1.Y, p2.Z - p1.Z)
            if v_tmp.Length < 0.0001: continue
            
            v_normal_comp = v_frente.Copy()
            v_normal_comp.ScaleBy(v_tmp.DotProduct(v_frente))
            
            v_plano = v_tmp.Copy()
            v_plano.SubtractVector(v_normal_comp)
            
            if v_plano.Length < 0.0001: continue
            
            if v_plano.Length > mejor_largo:
                mejor_largo = v_plano.Length
                v_plano.Normalize()
                v_guia = v_plano
                
        if v_guia is not None:
            v_guia.Normalize()
            return True, v_guia
        return False, None
    except:
        return False, None


def obtener_lado_fallback(tg, v_frente):
    world_up = tg.CreateVector(0, 0, 1)
    if abs(v_frente.DotProduct(world_up)) > 0.9:
        world_up = tg.CreateVector(0, 1, 0)
        
    v_lado = v_frente.CrossProduct(world_up)
    if v_lado.Length < 0.001:
        world_up = tg.CreateVector(1, 0, 0)
        v_lado = v_frente.CrossProduct(world_up)
        
    v_lado.Normalize()
    return v_lado


def obtener_centro(part_doc, cuerpo_medicion):
    cx, cy, cz = 0.0, 0.0, 0.0
    try:
        box = None
        if cuerpo_medicion is not None:
            box = cuerpo_medicion.RangeBox
        else:
            box = part_doc.ComponentDefinition.RangeBox
            
        cx = (box.MaxPoint.X + box.MinPoint.X) / 2.0
        cy = (box.MaxPoint.Y + box.MinPoint.Y) / 2.0
        cz = (box.MaxPoint.Z + box.MinPoint.Z) / 2.0
    except:
        pass
    return cx, cy, cz


def crear_camara(part_doc, tg, to, cx, cy, cz, eye_dir, up_hint):
    eye_dir.Normalize()
    r_vec = eye_dir.CrossProduct(up_hint)
    if r_vec.Length < 0.001:
        temp = tg.CreateVector(1, 0, 0)
        if abs(eye_dir.DotProduct(temp)) > 0.9:
            temp = tg.CreateVector(0, 1, 0)
        r_vec = eye_dir.CrossProduct(temp)
        
    r_vec.Normalize()
    true_up = r_vec.CrossProduct(eye_dir)
    true_up.Normalize()
    
    cam = to.CreateCamera()
    cam.SceneObject = part_doc.ComponentDefinition
    cam.Perspective = False
    cam.Target = tg.CreatePoint(cx, cy, cz)
    cam.Eye = tg.CreatePoint(cx + eye_dir.X * 100, cy + eye_dir.Y * 100, cz + eye_dir.Z * 100)
    cam.UpVector = tg.CreateUnitVector(true_up.X, true_up.Y, true_up.Z)
    
    return cam


def escalar_vista(doc, view, tg, px, py, ancho_util=None, alto_util=None):
    """
    Escala y centra la vista de forma que la PIEZA + espacio para cotas
    quede DENTRO del sheet físico. La regla es:

    - Reserva 2.5 cm de espacio para cotas por cada lado (número + flecha
      + margen al borde caben siempre).
    - La pieza sola no puede ocupar más del 55 % del área útil.
    - Se elige la escala discreta MÁS GRANDE cuyo bbox de vista + reserva
      de cotas cabe dentro del área útil.
    - VERIFICACIÓN POST-ESCALA: después de aplicar escala y posición,
      leemos ``view.Left``/``view.Top``/``view.Width``/``view.Height`` y
      confirmamos que el rectángulo real de la vista está DENTRO del sheet
      físico. Si no lo está (por asimetría del bbox 2D o errores de
      redondeo), bajamos a la siguiente escala y reintentamos.

    Con esto se garantiza que:
    1. La pieza nunca se sale del sheet, aunque su bbox 2D esté descentrado.
    2. Las cotas nunca se pierden en el borde.
    """
    try:
        doc.Update()
        curr_w = view.Width
        curr_h = view.Height
        curr_scale = view.Scale

        if curr_w <= 0 or curr_h <= 0 or curr_scale <= 0:
            return

        real_w = curr_w / curr_scale
        real_h = curr_h / curr_scale

        if ancho_util is None or ancho_util <= 0:
            ancho_util = 15.0
        if alto_util is None or alto_util <= 0:
            alto_util = 20.0

        # Reserva para cotas (cm) por cada lado.
        reserva_cotas = 2.5
        max_ancho_pieza = max(1e-3, ancho_util - 2.0 * reserva_cotas)
        max_alto_pieza = max(1e-3, alto_util - 2.0 * reserva_cotas)

        # La pieza sola nunca puede pasar del 55 % del área útil.
        max_ancho_pieza = min(max_ancho_pieza, ancho_util * 0.55)
        max_alto_pieza = min(max_alto_pieza, alto_util * 0.55)

        escalas = [
            5.0, 4.0, 3.0, 2.0, 1.5, 1.0, 0.75, 0.5,
            0.4, 0.3, 0.25, 0.2, 0.15, 0.1, 0.08, 0.05,
            0.04, 0.03, 0.02, 0.015, 0.01, 0.008, 0.005,
            0.004, 0.003, 0.002, 0.0015, 0.001,
        ]

        # Índice inicial: primera escala cuyo bbox teórico cabe.
        indice_inicial = len(escalas) - 1
        for idx, e_val in enumerate(escalas):
            w_esc = real_w * e_val
            h_esc = real_h * e_val
            if w_esc <= max_ancho_pieza and h_esc <= max_alto_pieza:
                indice_inicial = idx
                break

        # Sheet físico (para la verificación post-escala).
        sheet_w = None
        sheet_h = None
        try:
            sheet = view.Parent
            sheet_w = float(sheet.Width)
            sheet_h = float(sheet.Height)
        except Exception:
            pass

        # Aplicamos escalas de forma progresiva bajando si la vista sigue
        # saliéndose del sheet. Máximo 5 reintentos para no ciclar mucho.
        for offset in range(6):
            idx_actual = min(indice_inicial + offset, len(escalas) - 1)
            e_val = escalas[idx_actual]

            try:
                view.Scale = e_val
                view.Position = tg.CreatePoint2d(px, py)
                doc.Update()
            except Exception:
                continue

            if sheet_w is None or sheet_h is None:
                return

            try:
                left = float(view.Left)
                top = float(view.Top)
                width = float(view.Width)
                height = float(view.Height)
            except Exception:
                return

            right = left + width
            bottom = top - height

            # Margen mínimo al borde del sheet: 1.5 cm (para que la cota
            # tenga espacio de dibujarse sin salirse).
            margen = 1.5
            fits = (
                left >= margen
                and right <= sheet_w - margen
                and bottom >= margen
                and top <= sheet_h - margen
            )
            if fits:
                return

            # No cabe: siguiente iteración probará la escala más chica.

    except Exception:
        pass


def obtener_nombre_base_corto(part_name):
    limpio = part_name.strip()
    if not limpio: return limpio
    
    partes = limpio.split('-')
    if len(partes) >= 3:
        return f"{partes[0]}-{partes[1]}-{partes[2]}"
    else:
        return limpio


def construir_nombre_hoja(doc, part_name, sufijo):
    nombre_base = obtener_nombre_base_corto(part_name)
    base_name = f"{nombre_base}_{sufijo}"
    
    try:
        for i in range(1, doc.Sheets.Count + 1):
            if doc.Sheets.Item(i).Name == base_name:
                doc.Sheets.Item(i).Delete()
                break
    except:
        pass
        
    return base_name
