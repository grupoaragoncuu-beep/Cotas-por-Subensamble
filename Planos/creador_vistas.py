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
    iterations = int(seconds * 10)
    for _ in range(iterations):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)

def crear_vistas(inv_app, ensamble_doc, machote_doc):
    _log("⏳ Paso 1/4: Generando vistas con API de Python...")
    _log(f"Ensamble activo: {ensamble_doc.DisplayName}")
    
    try:
        machote_doc = win32com.client.CastTo(machote_doc, "DrawingDocument")
    except:
        pass
        
    try:
        ensamble_doc = win32com.client.CastTo(ensamble_doc, "AssemblyDocument")
    except:
        pass
    
    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    
    base_sheet = machote_doc.ActiveSheet
    processed_docs = set()
    contador = 0
    
    try:
        leaf_occs = ensamble_doc.ComponentDefinition.Occurrences.AllLeafOccurrences
        for i in range(1, leaf_occs.Count + 1):
            occ = leaf_occs.Item(i)
            if occ.Suppressed:
                continue
                
            try:
                doc_type = occ.DefinitionDocumentType
                _log(f"Occ: {occ.Name}, type: {doc_type}")
            except:
                pass
                
            part_doc = win32com.client.CastTo(occ.Definition.Document, "PartDocument")
            if part_doc is None:
                try:
                    part_doc = occ.Definition.Document
                except:
                    continue
                
            if part_doc.FullFileName in processed_docs:
                continue
                
            processed_docs.add(part_doc.FullFileName)
            
            part_name = os.path.splitext(os.path.basename(part_doc.FullFileName))[0]
            contador += 1
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
            uso_fallback_lado = False
            if not res_lado:
                v_lado = obtener_lado_fallback(tg, v_frente)
                uso_fallback_lado = True
                area_lado = 0.0
                lado_face = None
            else:
                lado_face, v_lado, area_lado = res_lado
                
            cx, cy, cz = obtener_centro(part_doc, cuerpo_medicion)
            
            tiene_guia_frente, v_guia_frente = obtener_vector_guia_frente(frente_face, v_frente, tg)
            
            sufijos = ["FRENTE_1", "FRENTE_2", "LADO"]
            es_lado = [False, False, True]
            
            for idx in range(3):
                is_side = es_lado[idx]
                nombre_hoja = construir_nombre_hoja(machote_doc, part_name, sufijos[idx])
                
                try:
                    new_sheet = base_sheet.CopyTo(machote_doc)
                    new_sheet.Name = nombre_hoja
                    new_sheet.Activate()
                    
                    for v in range(new_sheet.DrawingViews.Count, 0, -1):
                        new_sheet.DrawingViews.Item(v).Delete()
                except Exception as e:
                    _log(f"⚠️ {part_name} -> no se pudo crear hoja {nombre_hoja}: {e}")
                    continue
                    
                try:
                    if new_sheet.TitleBlock is not None:
                        tb = new_sheet.TitleBlock
                        tb.SetResultText(tb.Definition.Sketch.TextBoxes.Item(1), part_name)
                except:
                    pass
                    
                px, py = 10.0, 15.0
                
                if is_side:
                    eye_dir = v_lado.Copy()
                    up_hint = v_frente.Copy()
                else:
                    eye_dir = v_frente.Copy()
                    if tiene_guia_frente:
                        up_hint = v_guia_frente.Copy()
                    else:
                        up_hint = v_lado.Copy()
                        
                cam = crear_camara(part_doc, tg, to, cx, cy, cz, eye_dir, up_hint)
                
                view = None
                try:
                    options = to.CreateNameValueMap()
                    if use_flat_pattern:
                        options.Add("SheetMetalFoldedModel", False)
                        
                    view = new_sheet.DrawingViews.AddBaseView(
                        part_doc,
                        tg.CreatePoint2d(px, py),
                        1.0,
                        kArbitraryViewOrientation,
                        kHiddenLineRemovedDrawingViewStyle,
                        "",
                        cam,
                        options
                    )
                except Exception as ex1:
                    try:
                        view = new_sheet.DrawingViews.AddBaseView(
                            part_doc,
                            tg.CreatePoint2d(px, py),
                            1.0,
                            kDefaultViewOrientation,
                            kHiddenLineRemovedDrawingViewStyle
                        )
                    except:
                        _log(f"⚠️ {part_name} -> no se pudo crear vista {sufijos[idx]}")
                        
                if view is not None:
                    escalar_vista(machote_doc, view, tg, px, py)
            
            # Dar respiro a la tarjeta gráfica y bombear mensajes para evitar TDR (cuelgue DX12)
            try:
                inv_app.ActiveView.Update()
            except:
                pass
            _sleep_and_pump(4)
            
            if contador % 10 == 0:
                _log(f"Descanso de 10 segundos por cada 10 componentes (vamos en {contador})...")
                _sleep_and_pump(10)
                    
    finally:
        try:
            base_sheet.Activate()
        except:
            pass

    _log(f"✅ Vistas generadas correctamente para {contador} piezas.")
    return True


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


def escalar_vista(doc, view, tg, px, py):
    try:
        doc.Update()
        curr_w = view.Width
        curr_h = view.Height
        curr_scale = view.Scale
        
        if curr_w <= 0 or curr_h <= 0 or curr_scale <= 0: return
        
        real_w = curr_w / curr_scale
        real_h = curr_h / curr_scale
        
        ratio = min(15.0 / real_w, 20.0 / real_h) * 0.8
        escalas = [5.0, 4.0, 3.0, 2.0, 1.5, 1.0, 0.75, 0.5, 0.25, 0.2, 0.15, 0.1, 0.05, 0.02, 0.01]
        final_scale = 0.01
        
        for e_val in escalas:
            if ratio >= e_val:
                final_scale = e_val
                break
                
        view.Scale = final_scale
        view.Position = tg.CreatePoint2d(px, py)
    except:
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
