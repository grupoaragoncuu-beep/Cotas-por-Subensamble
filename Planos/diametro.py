import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota


def acotar_diametros(hojas_pendientes=None):
    print("⭕ diametro.py: Iniciando escáner de límites (Ext/Int)...")
    inv_app = conectar_inventor()
    
    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, 'DrawingDocument')
    except:
        print("❌ Error: No hay un plano de Inventor abierto.")
        return list(hojas_pendientes) if hojas_pendientes is not None else []

    tg = inv_app.TransientGeometry
    procesadas = 0
    procesadas_nombres = set()

    hojas_objetivo = None
    objetivo_set = None

    if hojas_pendientes is not None:
        hojas_objetivo = [str(h).upper() for h in hojas_pendientes]
        objetivo_set = set(hojas_objetivo)

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja = str(hoja.Name).upper()
        
        if objetivo_set is not None and nombre_hoja not in objetivo_set:
            continue
            
        if "_FRENTE" not in nombre_hoja:
            continue

        if hoja.DrawingViews.Count == 0:
            continue
        
        vista = hoja.DrawingViews.Item(1)
        anillos = []

        for j in range(1, vista.DrawingCurves.Count + 1):
            curva = vista.DrawingCurves.Item(j)
            try:
                caja = curva.Evaluator2D.RangeBox
                ancho = abs(caja.MaxPoint.X - caja.MinPoint.X)
                alto = abs(caja.MaxPoint.Y - caja.MinPoint.Y)
                
                if ancho < 0.1 or alto < 0.1:
                    continue
                
                if abs(ancho - alto) < (ancho * 0.15):
                    cx = (caja.MaxPoint.X + caja.MinPoint.X) / 2.0
                    cy = (caja.MaxPoint.Y + caja.MinPoint.Y) / 2.0
                    anillos.append({
                        'curva': curva,
                        'tamaño': ancho,
                        'cx': cx,
                        'cy': cy
                    })
            except:
                pass

        if not anillos:
            continue

        try:
            anillo_objetivo = None
            etiqueta = ""

            anillos_validos = [a for a in anillos if a['tamaño'] > 0.3]

            if "_FRENTE_1" in nombre_hoja:
                if anillos_validos:
                    anillo_objetivo = max(anillos_validos, key=lambda x: x['tamaño'])
                    etiqueta = "EXTERIOR"
                
            elif "_FRENTE_2" in nombre_hoja:
                if anillos_validos:
                    anillo_exterior = max(anillos_validos, key=lambda x: x['tamaño'])

                    tolerancia_centro = max(0.15, anillo_exterior['tamaño'] * 0.05)

                    interiores_concentricos = []
                    for a in anillos_validos:
                        if a is anillo_exterior:
                            continue

                        dx_c = abs(a['cx'] - anillo_exterior['cx'])
                        dy_c = abs(a['cy'] - anillo_exterior['cy'])

                        if dx_c <= tolerancia_centro and dy_c <= tolerancia_centro and a['tamaño'] < anillo_exterior['tamaño']:
                            interiores_concentricos.append(a)

                    if interiores_concentricos:
                        # Para _FRENTE_2 queremos el círculo MÁS PEQUEÑO
                        # del mismo centro, o sea el límite interior real.
                        anillo_objetivo = min(interiores_concentricos, key=lambda x: x['tamaño'])
                        etiqueta = "INTERIOR"
                    else:
                        print(f"↪️ {nombre_hoja}: no tiene límite interior concéntrico real; se deja pendiente")

            if anillo_objetivo:
                intencion = hoja.CreateGeometryIntent(anillo_objetivo['curva'])
                offset = (anillo_objetivo['tamaño'] / 2.0) + 1.0
                punto_texto = tg.CreatePoint2d(
                    anillo_objetivo['cx'] + offset,
                    anillo_objetivo['cy'] + offset
                )

                dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
                aplicar_estilo_cota(dim, hoja=hoja)
                print(f"✅ {nombre_hoja}: Límite {etiqueta} (Ø {anillo_objetivo['tamaño']:.2f})")
                procesadas += 1
                procesadas_nombres.add(nombre_hoja)

        except Exception:
            print(f"⚠️ {nombre_hoja}: Inventor rechazó el diámetro.")

    print(f"✅ diametro.py finalizado. Piezas acotadas: {procesadas}")

    if hojas_objetivo is not None:
        no_resueltas = [h for h in hojas_objetivo if h not in procesadas_nombres]
        return no_resueltas

    return []


if __name__ == "__main__":
    acotar_diametros()