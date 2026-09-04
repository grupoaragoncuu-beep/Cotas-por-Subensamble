import os
import re
import win32com.client
from inventor_com import conectar_inventor
from cota_estilo import aplicar_estilo_cota
from rutas_runtime import ruta_piezas_solidas


# Archivo donde publicamos los nombres BASE de piezas identificadas como
# "cilindros sólidos sin interior concéntrico". El flujo de exportación
# (``exportar_hojas_jpg``) lee este archivo para borrar cualquier JPG
# ``_DIAMETRO_INTERIOR_*.jpg`` residual que haya quedado de corridas viejas
# antes de que se agregara la eliminación de hojas huérfanas.
# Portable: Planos/.runtime/ (antes C:\Temp\...)
RUTA_PIEZAS_SOLIDAS = ruta_piezas_solidas()


def _publicar_piezas_solidas(nombres_hojas_solidas):
    """Escribe los nombres BASE de piezas sólidas para consumo del exportador."""
    if not nombres_hojas_solidas:
        return
    try:
        os.makedirs(os.path.dirname(RUTA_PIEZAS_SOLIDAS), exist_ok=True)
    except Exception:
        pass
    piezas = set()
    for hoja_nombre in nombres_hojas_solidas:
        base = str(hoja_nombre).split(":", 1)[0]
        base = re.sub(r"_FRENTE_[12]$", "", base, flags=re.IGNORECASE)
        base = re.sub(r"_DIAMETRO_(INTERIOR|EXTERIOR)$", "", base, flags=re.IGNORECASE)
        piezas.add(base.strip().upper())

    if not piezas:
        return

    existentes = set()
    try:
        if os.path.exists(RUTA_PIEZAS_SOLIDAS):
            with open(RUTA_PIEZAS_SOLIDAS, "r", encoding="utf-8") as f:
                for linea in f:
                    linea = linea.strip()
                    if linea:
                        existentes.add(linea.upper())
    except Exception:
        existentes = set()

    piezas.update(existentes)
    try:
        with open(RUTA_PIEZAS_SOLIDAS, "w", encoding="utf-8") as f:
            for p in sorted(piezas):
                f.write(p + "\n")
    except Exception:
        pass


def _clampear_punto_hoja(hoja, tg, x, y, margen=1.2):
    """Fuerza el Point2d de texto a caer dentro del rectángulo físico de la
    hoja para que la cámara del JPG lo capture. Margen de 1.2 cm para dejar
    espacio al número y la flecha del diámetro."""
    try:
        sheet_w = float(hoja.Width)
        sheet_h = float(hoja.Height)
        x = max(margen, min(sheet_w - margen, x))
        y = max(margen, min(sheet_h - margen, y))
    except Exception:
        pass
    return tg.CreatePoint2d(x, y)


def _anillos_en_vista(vista, min_tam=0.25, max_frac=0.55):
    """
    Círculos/arcos casi circulares (barrenos por arista circular).

    Placas OTC (p.ej. P12_2) tienen aristas circulares sin HoleFeature ni
    cara Cylinder: hay que detectarlas en el HLR 2D.
    """
    anillos = []
    try:
        n = int(vista.DrawingCurves.Count)
    except Exception:
        return anillos

    minx = miny = None
    maxx = maxy = None
    for j in range(1, n + 1):
        try:
            caja = vista.DrawingCurves.Item(j).Evaluator2D.RangeBox
            x0, x1 = float(caja.MinPoint.X), float(caja.MaxPoint.X)
            y0, y1 = float(caja.MinPoint.Y), float(caja.MaxPoint.Y)
            minx = x0 if minx is None else min(minx, x0)
            maxx = x1 if maxx is None else max(maxx, x1)
            miny = y0 if miny is None else min(miny, y0)
            maxy = y1 if maxy is None else max(maxy, y1)
        except Exception:
            continue
    if minx is None:
        return anillos
    span = max(maxx - minx, maxy - miny)
    tope = span * max_frac if span > 0 else 1e9

    for j in range(1, n + 1):
        try:
            curva = vista.DrawingCurves.Item(j)
            caja = curva.Evaluator2D.RangeBox
            ancho = abs(float(caja.MaxPoint.X) - float(caja.MinPoint.X))
            alto = abs(float(caja.MaxPoint.Y) - float(caja.MinPoint.Y))
            if ancho < min_tam or alto < min_tam:
                continue
            if abs(ancho - alto) > max(ancho, alto) * 0.18:
                continue
            tam = (ancho + alto) * 0.5
            if tam > tope:
                continue
            anillos.append(
                {
                    "curva": curva,
                    "tamaño": tam,
                    "cx": (float(caja.MaxPoint.X) + float(caja.MinPoint.X)) / 2.0,
                    "cy": (float(caja.MaxPoint.Y) + float(caja.MinPoint.Y)) / 2.0,
                }
            )
        except Exception:
            continue
    return anillos


def _agrupar_diametros_unicos(anillos, tol_frac=0.06, max_grupos=8):
    if not anillos:
        return []
    ordenados = sorted(anillos, key=lambda a: a["tamaño"], reverse=True)
    grupos = []
    for a in ordenados:
        colocado = False
        for g in grupos:
            ref = g[0]["tamaño"]
            if abs(a["tamaño"] - ref) <= max(0.05, ref * tol_frac):
                g.append(a)
                colocado = True
                break
        if not colocado:
            grupos.append([a])
        if len(grupos) >= max_grupos:
            break
    return [max(g, key=lambda x: x["tamaño"]) for g in grupos]


def acotar_barrenos_placas(nombres_frente_ok=None):
    """
    Tras LARGO/ANCHO, crea hojas ``*_DIAMETRO_Hnn`` por tamaño de barreno.
    """
    print("⭕ diametro.py: barrenos en placas (aristas circulares)...")
    inv_app = conectar_inventor()
    try:
        plano = win32com.client.CastTo(inv_app.ActiveDocument, "DrawingDocument")
    except Exception:
        print("❌ Error: No hay un plano de Inventor abierto.")
        return []

    tg = inv_app.TransientGeometry
    objetivo = None
    if nombres_frente_ok is not None:
        objetivo = {str(h).upper().rsplit(":", 1)[0] for h in nombres_frente_ok}

    creadas = 0
    creadas_nombres = []
    for i in range(1, plano.Sheets.Count + 1):
        try:
            hoja = plano.Sheets.Item(i)
        except Exception:
            continue
        nombre = str(hoja.Name)
        nombre_up = nombre.upper()
        if "_FRENTE_1" not in nombre_up:
            continue
        base_cmp = nombre_up.rsplit(":", 1)[0]
        if objetivo is not None and base_cmp not in objetivo:
            continue
        if hoja.DrawingViews.Count < 1:
            continue
        vista = hoja.DrawingViews.Item(1)
        unicos = _agrupar_diametros_unicos(_anillos_en_vista(vista))
        if not unicos:
            continue

        base_nombre = nombre.rsplit(":", 1)[0]
        base_nombre = re.sub(r"_FRENTE_1$", "", base_nombre, flags=re.IGNORECASE)

        for idx, anillo in enumerate(unicos, start=1):
            nombre_nueva = f"{base_nombre}_DIAMETRO_H{idx:02d}"
            try:
                for j in range(plano.Sheets.Count, 0, -1):
                    try:
                        h = plano.Sheets.Item(j)
                        if str(h.Name).upper().startswith(nombre_nueva.upper()):
                            h.Delete()
                    except Exception:
                        continue
            except Exception:
                pass

            try:
                nueva = hoja.CopyTo(plano)
            except Exception as e:
                print(f"⚠️ {nombre_nueva}: CopyTo falló ({e})")
                continue
            try:
                nueva.Name = nombre_nueva
            except Exception:
                pass
            try:
                nueva.Activate()
            except Exception:
                pass
            try:
                dims = nueva.DrawingDimensions.GeneralDimensions
                for k in range(dims.Count, 0, -1):
                    try:
                        dims.Item(k).Delete()
                    except Exception:
                        pass
            except Exception:
                pass

            if nueva.DrawingViews.Count < 1:
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            vista_n = nueva.DrawingViews.Item(1)
            anillos_n = _anillos_en_vista(vista_n)
            if not anillos_n:
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            objetivo_a = min(
                anillos_n, key=lambda a: abs(a["tamaño"] - anillo["tamaño"])
            )
            if abs(objetivo_a["tamaño"] - anillo["tamaño"]) > max(
                0.08, anillo["tamaño"] * 0.12
            ):
                try:
                    nueva.Delete()
                except Exception:
                    pass
                continue
            try:
                intent = nueva.CreateGeometryIntent(objetivo_a["curva"])
                offset = (objetivo_a["tamaño"] / 2.0) + 1.0
                pt = _clampear_punto_hoja(
                    nueva, tg,
                    objetivo_a["cx"] + offset,
                    objetivo_a["cy"] + offset,
                )
                dim = nueva.DrawingDimensions.GeneralDimensions.AddDiameter(
                    pt, intent
                )
                aplicar_estilo_cota(dim, hoja=nueva)
                print(
                    f"✅ {nombre_nueva}: barreno Ø "
                    f"{objetivo_a['tamaño'] / 2.54:.3f} in"
                )
                creadas += 1
                creadas_nombres.append(str(nueva.Name).rsplit(":", 1)[0])
            except Exception as e:
                print(f"⚠️ {nombre_nueva}: AddDiameter falló ({e})")
                try:
                    nueva.Delete()
                except Exception:
                    pass

    print(f"✅ diametro.py barrenos: {creadas} hojas DIAMETRO_H* creadas")
    return creadas_nombres


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

    # Nombres de hojas _FRENTE_2 a ELIMINAR al final del loop porque la pieza
    # cilíndrica no tiene borde interior concéntrico (es una barra/pin/stud
    # sólido) y no tiene sentido exportar un JPG con solo el contorno
    # exterior duplicado.
    hojas_a_eliminar = []

    for i in range(1, plano.Sheets.Count + 1):
        hoja = plano.Sheets.Item(i)
        nombre_hoja_original = str(hoja.Name)
        nombre_hoja = nombre_hoja_original.upper()

        if objetivo_set is not None and nombre_hoja not in objetivo_set:
            continue

        if "_FRENTE" not in nombre_hoja:
            continue

        if hoja.DrawingViews.Count == 0:
            continue

        vista = hoja.DrawingViews.Item(1)
        anillos = _anillos_en_vista(vista, min_tam=0.1, max_frac=0.98)

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
                        # Sin interior concéntrico → pieza cilíndrica SÓLIDA
                        # (barra, pin, stud). No tiene diámetro interior,
                        # así que eliminamos la hoja para que no aparezca un
                        # JPG mudo con solo el contorno exterior.
                        print(
                            f"🗑️ {nombre_hoja}: pieza cilíndrica sólida sin "
                            f"interior concéntrico; se elimina la hoja "
                            f"_FRENTE_2 (no aplica DIAMETRO_INTERIOR)."
                        )
                        hojas_a_eliminar.append(nombre_hoja_original)
                        procesadas_nombres.add(nombre_hoja)
                        continue

            if anillo_objetivo:
                intencion = hoja.CreateGeometryIntent(anillo_objetivo['curva'])
                offset = (anillo_objetivo['tamaño'] / 2.0) + 1.0
                punto_texto = _clampear_punto_hoja(
                    hoja, tg,
                    anillo_objetivo['cx'] + offset,
                    anillo_objetivo['cy'] + offset,
                )

                dim = hoja.DrawingDimensions.GeneralDimensions.AddDiameter(punto_texto, intencion)
                aplicar_estilo_cota(dim, hoja=hoja)
                print(f"✅ {nombre_hoja}: Límite {etiqueta} (Ø {anillo_objetivo['tamaño']:.2f})")
                procesadas += 1
                procesadas_nombres.add(nombre_hoja)

        except Exception:
            print(f"⚠️ {nombre_hoja}: Inventor rechazó el diámetro.")

    # Eliminar hojas _FRENTE_2 de piezas cilíndricas sólidas. Iteramos por
    # nombre y de atrás hacia adelante en el arreglo de sheets para no
    # invalidar índices.
    for nombre_borrar in hojas_a_eliminar:
        try:
            for j in range(plano.Sheets.Count, 0, -1):
                try:
                    if str(plano.Sheets.Item(j).Name) == nombre_borrar:
                        plano.Sheets.Item(j).Delete()
                        break
                except Exception:
                    continue
        except Exception as e:
            print(f"⚠️ No se pudo eliminar la hoja '{nombre_borrar}': {e}")

    print(f"✅ diametro.py finalizado. Piezas acotadas: {procesadas}")
    if hojas_a_eliminar:
        print(f"🗑️ Hojas _FRENTE_2 eliminadas por ser sólidas: {len(hojas_a_eliminar)}")
        # Publicar la lista de piezas sólidas para que el flujo de exportación
        # borre cualquier JPG _DIAMETRO_INTERIOR residual de corridas anteriores.
        try:
            _publicar_piezas_solidas(hojas_a_eliminar)
        except Exception as pub_err:
            print(f"AVISO: no se pudo publicar piezas sólidas: {pub_err}")

    if hojas_objetivo is not None:
        no_resueltas = [h for h in hojas_objetivo if h not in procesadas_nombres]
        return no_resueltas

    return []


if __name__ == "__main__":
    acotar_diametros()