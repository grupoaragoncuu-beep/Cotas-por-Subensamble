"""
Flujo integrado de COTAS ABIGAIL.

1. Cotas por referencia de las 5 caras del tanque (FRONT, BACK, LEFT, RIGHT, TOP).
2. Flujo original COTAS_ILOGIC_ABIGAIL por pieza, reorganizado en subcarpetas
   por cara.

Ambas salidas se organizan bajo Planos/JPG/<tanque>/ en carpetas separadas.
"""

import os
import re
import shutil
import sys
import time

import pythoncom

import generador_caras_tanque
from generador_caras_tanque import (
    _carpeta_salida_tanque,
    _eliminar_hojas_piezas_residuales,
    _encontrar_hoja_machote,
    _limpiar_machote,
    _obtener_ensamble_principal,
    _obtener_plano_activo,
    _resolver_contenedor_paredes,
    ejecutar as ejecutar_caras,
)
from generador_vistas import ejecutar_flujo_desde_app
from inventor_com import conectar_inventor


CARPETA_PIEZAS_ACOTADAS = "PIEZAS_ACOTADAS"
# Carpetas activas con selección manual (COTAS_ILOGIC_ABIGAIL).
SUBCARPETAS_CARA_SELECCION = (
    "SEGM1",
    "SEGM2",
    "SEGM3",
    "SEGM4",
    "TOP",
    "BASE",
)
# Solo limpieza/compat de corridas viejas PQart; NO se crean vacías.
SUBCARPETAS_CARA_LEGACY = (
    "FRONT",
    "BACK",
    "LEFT",
    "RIGHT",
)
SUBCARPETAS_CARA_PIEZAS = SUBCARPETAS_CARA_SELECCION + SUBCARPETAS_CARA_LEGACY
SUBCARPETA_OTROS_PIEZAS = "OTROS"

# --- Clasificación por proceso (iProperty escrita por el iLogic Colorimetria).
# Se usa como subcarpetas del flujo PIEZAS_ACOTADAS en lugar de las caras
# geométricas del tanque. El orden fija la salida de logs.
SUBCARPETAS_CLASIFICACION_PIEZAS = (
    "Almacén",
    "Corte",
    "Maquinado",
    "Doblado",
    "Plasma",
    "Plasma Doblado",
)
SUBCARPETA_SIN_CLASIFICAR = "SIN CLASIFICACION"

# Árbol acordado (GIGA BOARD / cobre):
#   Corte/Plasma y Laser/Corte metal/<PIEZA>/   ← flat metal
#   Corte/Plasma y Laser/Corte Busbar/<PIEZA>/  ← (secundario)
#   Corte/Maquinado/Maquinados metal/<PIEZA>/
#   Corte/Maquinado/Corte Busbar/<PIEZA>/      ← flat cobre
#   Doblado/Metal/<PIEZA>/  |  Doblado/Busbar/<PIEZA>/
#   Estañado Busbar/*.jpg                      ← cobre isométrico SIN_COTA
# Legacy (migración): Corte/Corte|Doblado|Estañado
SUBCARPETAS_CORTE_PLASMA = ("Corte metal", "Corte Busbar")
SUBCARPETAS_CORTE_MAQUINADO = ("Maquinados metal", "Corte Busbar")
SUBCARPETAS_DOBLADO_ANIDADAS = ("Metal", "Busbar")
SUBCARPETA_ESTANIADO_BUSBAR = "Estañado Busbar"
# Legacy names kept for cleanup/migración de carpetas viejas.
SUBCARPETAS_CORTE_ANIDADAS = ("Corte", "Doblado", "Estañado")
STAGING_DESPLIEGUE = "_STAGING_DESPLIEGUE"
STAGING_ESTANIADO = "_STAGING_ESTANIADO"

# Claves de conteo / clasificacion (ruta lógica).
CLASIFICACIONES_LOG = (
    "Almacén",
    "Corte/Plasma y Laser/Corte metal",
    "Corte/Plasma y Laser/Corte Busbar",
    "Corte/Maquinado/Maquinados metal",
    "Corte/Maquinado/Corte Busbar",
    "Doblado/Metal",
    "Doblado/Busbar",
    SUBCARPETA_ESTANIADO_BUSBAR,
    "Maquinado",
    "Doblado",
    "Plasma",
    "Plasma Doblado",
    SUBCARPETA_SIN_CLASIFICAR,
)

# Sufijos que agrega el flujo de piezas al final del nombre del JPG. Se usan
# para separar "pieza" de "tipo de cota" y así crear una carpeta por pieza.
# ORDEN IMPORTA: los sufijos compuestos (LARGO_PATA, DIAMETRO_*) deben ir
# ANTES de sus prefijos simples (LARGO, DIAMETRO) para que la alternancia
# del regex los reconozca primero y no truncue el nombre.
_SUFIJOS_JPG_PIEZA = (
    "LENGTH",
    "WIDTH",
    "BROAD",  # legacy
    "THK",
    "HEIGHT",
    "LEG",
    "OD",
    "ID",
    "XCENTRO_TYP",
    "YCENTRO_TYP",
    "XCENTRO",
    "YCENTRO",
    "XMIN_TYP",
    "YMIN_TYP",
    "XMIN",
    "YMIN",
    "CUT_LENGTH",
    "CUT_WIDTH",
    "HOLE\\d{2}",
    "DIAMETRO_EXTERIOR",
    "DIAMETRO_INTERIOR",
    "DIAMETRO_H\\d{2}",
    "LARGO_PATA",
    "ANCHO",
    "LARGO",
    "ALTO",
)
_RE_SUFIJO_PIEZA = re.compile(
    r"^(?P<pieza>.+?)_(?P<tipo>"
    + "|".join(_SUFIJOS_JPG_PIEZA)
    + r")_\d+(?:\.\d+)?$",
    re.IGNORECASE,
)
_RE_SUFIJO_PIEZA_JOB = re.compile(
    r"^.+?__(?P<pieza>.+?)__(?P<tipo>"
    + "|".join(_SUFIJOS_JPG_PIEZA)
    + r")_\d+(?:\.\d+)?$",
    re.IGNORECASE,
)


def _limpiar_exportacion_piezas(carpeta, incremental=False):
    """
    Vacía la carpeta de piezas y sus subcarpetas antes de exportar.

    Si `incremental` es True se respeta lo que ya haya en la carpeta (modo F):
    solo se crea la estructura si no existe.

    Limpia AMBOS esquemas de subcarpetas:
    - Cara geométrica (FRONT/BACK/LEFT/RIGHT/TOP/OTROS) — flujo legacy.
    - Clasificación por proceso (Almacén/Corte/.../SIN CLASIFICACION) — flujo
      nuevo basado en iProperty.

    Así una migración de esquema no deja carpetas viejas mezcladas con
    las nuevas.
    """
    os.makedirs(carpeta, exist_ok=True)
    if incremental:
        return

    # Nombres válidos a limpiar: mantenemos también comparación con acentos
    # para las nuevas categorías, así "Almacén" hace match aunque el sistema
    # de archivos lo devuelva con la misma codificación.
    caras_upper = {c.upper() for c in SUBCARPETAS_CARA_PIEZAS}
    caras_upper.add(SUBCARPETA_OTROS_PIEZAS.upper())
    clasificaciones_ci = {
        c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS
    }
    clasificaciones_ci.add(SUBCARPETA_SIN_CLASIFICAR.casefold())
    clasificaciones_ci.add(SUBCARPETA_ESTANIADO_BUSBAR.casefold())
    staging_ci = {
        STAGING_DESPLIEGUE.casefold(),
        STAGING_ESTANIADO.casefold(),
    }

    for nombre in os.listdir(carpeta):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isdir(ruta):
            if (
                nombre.upper() in caras_upper
                or nombre.casefold() in clasificaciones_ci
                or nombre.casefold() in staging_ci
            ):
                try:
                    shutil.rmtree(ruta)
                except OSError as err:
                    print(f"AVISO: no se pudo vaciar {nombre}/: {err}")
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        try:
            os.remove(ruta)
        except OSError:
            pass


def _clave_pieza(texto):
    """Normaliza un nombre para comparar por token/subcadena robustamente."""
    if not texto:
        return ""
    limpio = str(texto).upper()
    # Quitar extensión y sufijos comunes de nombre de hoja.
    limpio = re.sub(r"\.JPG$", "", limpio)
    limpio = re.sub(r"[\s\-_:()\[\]{},.]+", "", limpio)
    return limpio


def _clave_familia_pieza(texto):
    """
    Familia de pieza (SP-852_1 y SP-852_2 → misma clave).

    Misma idea que ``_clave_tipo_pieza`` del flujo de caras, local aquí
    para no importar generador_caras (ciclo).
    """
    base = str(texto or "").upper().strip()
    if "|" in base:
        base = base.rsplit("|", 1)[-1]
    base = base.split(":")[0].strip()
    # Quitar sufijos de cota del JPG de piezas.
    base = re.sub(
        r"^(?:.+?__)(.+?)__"
        r"(?:LENGTH|WIDTH|BROAD|THK|HEIGHT|LEG|OD|ID|HOLE\d{2}|"
        r"XCENTRO_TYP|YCENTRO_TYP|XCENTRO|YCENTRO|"
        r"DIAMETRO_EXTERIOR|DIAMETRO_INTERIOR|DIAMETRO_H\d{2}|"
        r"LARGO_PATA|ANCHO|LARGO|ALTO)_\d+$",
        r"\1",
        base,
        flags=re.IGNORECASE,
    )
    base = re.sub(
        r"_(DIAMETRO_EXTERIOR|DIAMETRO_INTERIOR|DIAMETRO_H\d{2}|LARGO_PATA|"
        r"ANCHO|LARGO|THK|ALTO|LENGTH|WIDTH|BROAD|HEIGHT|LEG|OD|ID|HOLE\d{2}|"
        r"XCENTRO_TYP|YCENTRO_TYP|XCENTRO|YCENTRO)_\d+$",
        "",
        base,
        flags=re.IGNORECASE,
    )
    m = re.match(r"^(\d+-\d+-[AP]\d+)", base)
    if m:
        return _clave_pieza(m.group(1))
    base = re.sub(r"_\d+$", "", base)
    return _clave_pieza(base)


def _claves_match_pieza(texto):
    """Claves full + familia para matching multi-cara."""
    claves = set()
    c = _clave_pieza(texto)
    if c:
        claves.add(c)
    f = _clave_familia_pieza(texto)
    if f:
        claves.add(f)
    return claves


def _match_claves_pieza(ca, cp):
    """
    Empareja claves normalizadas de JPG vs mapa.

    Cubrir:
    - nombre completo vs completo
    - JPG corto ``VT-5657-R1`` vs mapa ``VT-5657-R1-D3000 LIFTING LUG…``
    - JPG con prefijo JOB ``MODELO…__ITEM__LENGTH_…`` (usar ITEM extraído)
    """
    if not ca or not cp:
        return False
    if ca == cp:
        return True
    # Prefijo: el token corto del JPG es el inicio del nombre en el mapa.
    if len(ca) >= 6 and cp.startswith(ca):
        return True
    if len(cp) >= 6 and ca.startswith(cp):
        return True
    # Contención solo con claves largas (evita falsos positivos tipo "R1").
    if len(ca) >= 10 and ca in cp:
        return True
    if len(cp) >= 10 and cp in ca:
        return True
    return False


def _candidatos_nombre_pieza_jpg(nombre_archivo):
    """
    Nombres a probar al matchear un JPG contra mapas de cara/clasificación.

    Prioriza el ITEM de la nomenclatura ``JOB__ITEM__MEDIDA_VALOR``.
    """
    base = os.path.splitext(os.path.basename(str(nombre_archivo)))[0]
    out = []
    try:
        from nomenclatura_capturas import extraer_item_de_captura_pieza

        item = extraer_item_de_captura_pieza(nombre_archivo)
        if item and item not in out:
            out.append(item)
    except Exception:
        pass
    if base and base not in out:
        out.append(base)
    return out


def _cara_para_pieza(nombre_archivo, mapa_por_cara):
    """Devuelve UNA cara (match más específico). Compat con callers viejos."""
    caras = _caras_para_pieza(nombre_archivo, mapa_por_cara)
    return caras[0] if caras else None


def _caras_para_pieza(nombre_archivo, mapa_por_cara):
    """
    Cara(s) destino para el JPG de pieza.

    Política (feedback OTC 62201):
    - Como máximo **una** carpeta ``SEGM*`` (dueño primario = mejor match).
      Evita que la base de SEGM1 (P14_1 / P14_2) se copie también a SEGM2.
    - ``TOP`` / ``BASE`` pueden coexistir con un SEGM si el catálogo lo dice
      (p. ej. accesorio en tapa y pared); no se duplican SEGM entre sí.
    """
    if not mapa_por_cara:
        return []
    claves_archivo = set()
    for cand in _candidatos_nombre_pieza_jpg(nombre_archivo):
        claves_archivo |= _claves_match_pieza(cand)
    if not claves_archivo:
        return []
    halladas = []
    for cara, piezas in mapa_por_cara.items():
        mejor_len = 0
        for pieza in piezas:
            claves_pieza = _claves_match_pieza(pieza)
            if not claves_pieza:
                continue
            for ca in claves_archivo:
                for cp in claves_pieza:
                    if _match_claves_pieza(ca, cp):
                        mejor_len = max(mejor_len, len(cp), len(ca))
        if mejor_len > 0:
            halladas.append((mejor_len, str(cara).upper()))
    if not halladas:
        return []
    orden = {c: i for i, c in enumerate(SUBCARPETAS_CARA_SELECCION)}
    segms = [(n, c) for n, c in halladas if str(c).startswith("SEGM")]
    otras = [(n, c) for n, c in halladas if not str(c).startswith("SEGM")]
    vistas = []
    if segms:
        # Mejor match primero; empate → orden SEGM1..4.
        segms.sort(key=lambda item: (-item[0], orden.get(item[1], 99), item[1]))
        vistas.append(segms[0][1])
    otras.sort(key=lambda item: (orden.get(item[1], 99), -item[0], item[1]))
    for _n, cara in otras:
        if cara not in vistas:
            vistas.append(cara)
    return vistas


def deduplicar_mapa_piezas_por_cara(mapa_por_cara):
    """
    Deja cada nombre de pieza en un solo SEGM* (primero en SEGM1..4).

    Además quita de SEGM* todo lo que también esté en el catálogo BASE
    (placa base del tanque no es “pieza de pared”).
    """
    if not mapa_por_cara:
        return mapa_por_cara
    mapa = {str(k).upper(): set(v or set()) for k, v in mapa_por_cara.items()}
    base = set(mapa.get("BASE") or set())
    if base:
        claves_base = set()
        for p in base:
            claves_base |= _claves_match_pieza(p)
        for cara in list(mapa):
            if not cara.startswith("SEGM"):
                continue
            quitar = set()
            for p in mapa[cara]:
                if any(
                    _match_claves_pieza(ca, cp)
                    for ca in _claves_match_pieza(p)
                    for cp in claves_base
                ):
                    quitar.add(p)
            if quitar:
                mapa[cara] -= quitar
                print(
                    f"  Dedupe SEGM: {cara} pierde {len(quitar)} "
                    f"pieza(s) que son de BASE"
                )

    vistos = {}  # clave_familia -> SEGM dueño
    for cara in ("SEGM1", "SEGM2", "SEGM3", "SEGM4"):
        if cara not in mapa:
            continue
        quitar = set()
        for p in list(mapa[cara]):
            fam = _clave_familia_pieza(p) or _clave_pieza(p)
            if not fam:
                continue
            if fam in vistos and vistos[fam] != cara:
                quitar.add(p)
            else:
                vistos[fam] = cara
        if quitar:
            mapa[cara] -= quitar
            print(
                f"  Dedupe SEGM: {cara} pierde {len(quitar)} "
                f"pieza(s) ya asignadas a otro SEGM"
            )
    return mapa


def _extraer_pieza_de_jpg(nombre_archivo):
    """
    Devuelve el nombre de pieza contenido en un JPG del flujo por pieza.

    Formato nuevo: `{JOB}__{ITEM}__{LENGTH|WIDTH|THK|…}_{n}.jpg`
    Legacy: `{ITEM}_{LARGO|ANCHO|…}_{n}.jpg`
    """
    try:
        from nomenclatura_capturas import extraer_item_de_captura_pieza

        return extraer_item_de_captura_pieza(nombre_archivo)
    except Exception:
        pass
    base = os.path.splitext(os.path.basename(nombre_archivo))[0]
    # Colisiones de aplanar: ``...__dup1`` no debe crear carpeta basura.
    base = re.sub(r"__(?:dup|v)\d+$", "", base, flags=re.IGNORECASE)
    match = _RE_SUFIJO_PIEZA_JOB.match(base)
    if match:
        return match.group("pieza")
    match = _RE_SUFIJO_PIEZA.match(base)
    if not match:
        return base
    return match.group("pieza")


def _nombre_carpeta_pieza(nombre_pieza):
    """Sanea el nombre de pieza para usarlo como carpeta en Windows."""
    limpio = re.sub(r'[<>:"/\\|?*]+', "_", str(nombre_pieza).strip())
    limpio = limpio.rstrip(". ")
    return limpio or "PIEZA"


def _nombre_jpg_destino(nombre_archivo):
    """
    Nombre de archivo limpio para el árbol final: quita ``__dupN`` / ``__vN``
    de colisiones para no crear carpetas/archivos basura.
    """
    base, ext = os.path.splitext(os.path.basename(nombre_archivo))
    base = re.sub(r"__(?:dup|v)\d+$", "", base, flags=re.IGNORECASE)
    return f"{base}{ext or '.jpg'}"


def _mover_jpg_a_destino(ruta, destino_dir, nombre_archivo):
    """
    Mueve JPG a destino_dir con nombre sin sufijos de colisión.
    Si ya existe el destino: mismo tamaño → borra origen; distinto → reemplaza.
    """
    os.makedirs(destino_dir, exist_ok=True)
    nombre_limpio = _nombre_jpg_destino(nombre_archivo)
    destino = os.path.join(destino_dir, nombre_limpio)
    if os.path.abspath(ruta) == os.path.abspath(destino):
        return True
    if os.path.exists(destino):
        try:
            if os.path.getsize(ruta) == os.path.getsize(destino):
                os.remove(ruta)
                return True
        except OSError:
            pass
        try:
            os.remove(destino)
        except OSError:
            pass
    shutil.move(ruta, destino)
    return True


def _carpetas_cara_activas(mapa_por_cara=None):
    """
    Carpetas a crear/usar: SEGM*/TOP/BASE (+ OTROS), nunca FRONT/BACK/LEFT/RIGHT.
    """
    activas = list(SUBCARPETAS_CARA_SELECCION)
    if mapa_por_cara:
        extras = [
            str(k).upper()
            for k in mapa_por_cara.keys()
            if str(k).upper() in SUBCARPETAS_CARA_SELECCION
        ]
        if extras:
            activas = [c for c in SUBCARPETAS_CARA_SELECCION if c in extras]
            for c in extras:
                if c not in activas:
                    activas.append(c)
    return tuple(activas) + (SUBCARPETA_OTROS_PIEZAS,)


def _limpiar_carpetas_cara_vacias_y_legacy(carpeta_piezas):
    """Borra FRONT/BACK/LEFT/RIGHT y cualquier cara activa que haya quedado vacía."""
    if not os.path.isdir(carpeta_piezas):
        return
    candidatas = set(SUBCARPETAS_CARA_LEGACY) | set(SUBCARPETAS_CARA_SELECCION)
    candidatas.add(SUBCARPETA_OTROS_PIEZAS)
    for nombre in list(candidatas):
        ruta = os.path.join(carpeta_piezas, nombre)
        if not os.path.isdir(ruta):
            continue
        # Legacy siempre se elimina si está vacío; si tiene JPG (corrida vieja)
        # se deja, pero FRONT/BACK/LEFT/RIGHT vacíos no deben residualizar.
        try:
            tiene_jpg = False
            for _root, _dirs, files in os.walk(ruta):
                if any(f.lower().endswith(".jpg") for f in files):
                    tiene_jpg = True
                    break
            if not tiene_jpg:
                shutil.rmtree(ruta, ignore_errors=True)
                print(f"  Carpeta residual vacía eliminada: {nombre}/")
            elif nombre in SUBCARPETAS_CARA_LEGACY:
                print(
                    f"  AVISO: {nombre}/ aún tiene JPG de esquema PQart viejo; "
                    "no se borró automáticamente."
                )
        except OSError as err:
            print(f"  AVISO: no se pudo limpiar {nombre}/: {err}")
    _podar_directorios_sin_jpg(carpeta_piezas)


def _podar_directorios_sin_jpg(carpeta_piezas):
    """
    Elimina subcarpetas sin ningún JPG (p. ej. ``SEGM1/P71`` vacío tras
    mover a ``SEGM1/Doblado/Metal/P71``, o residuos ``__dup`` / ``__v2``).
    """
    if not os.path.isdir(carpeta_piezas):
        return
    borradas = 0
    # Bottom-up: deepest paths first.
    for dirpath, dirnames, _files in os.walk(carpeta_piezas, topdown=False):
        if os.path.abspath(dirpath) == os.path.abspath(carpeta_piezas):
            continue
        try:
            tiene_jpg = False
            for _r, _d, fs in os.walk(dirpath):
                if any(f.lower().endswith(".jpg") for f in fs):
                    tiene_jpg = True
                    break
            if not tiene_jpg:
                shutil.rmtree(dirpath, ignore_errors=True)
                borradas += 1
        except OSError:
            pass
    if borradas:
        print(f"  Carpetas vacías podadas: {borradas}")


def _reorganizar_piezas_por_cara(carpeta_piezas, mapa_por_cara):
    """
    Coloca cada JPG en `<CARA>/<PIEZA>/<archivo>.jpg`.

    Destino por ``_caras_para_pieza`` (como máximo un SEGM*). Si aún
    hubiera más de una cara no-SEGM, copia. Sin match → ``OTROS/<PIEZA>/``.
    NO crea FRONT/BACK/LEFT/RIGHT (esquema PQart residual).
    """
    mapa_por_cara = deduplicar_mapa_piezas_por_cara(mapa_por_cara)
    activas = _carpetas_cara_activas(mapa_por_cara)
    validas_destino = set(SUBCARPETAS_CARA_SELECCION) | {SUBCARPETA_OTROS_PIEZAS}
    try:
        os.makedirs(carpeta_piezas, exist_ok=True)
        for sub in activas:
            os.makedirs(os.path.join(carpeta_piezas, sub), exist_ok=True)
    except OSError as err:
        print(f"AVISO: no se pudieron preparar subcarpetas de PIEZAS_ACOTADAS: {err}")
        return {}

    conteo_cara = {sub: 0 for sub in activas}
    piezas_por_cara = {sub: set() for sub in activas}
    try:
        entradas = list(os.listdir(carpeta_piezas))
    except OSError as err:
        print(f"AVISO: no se pudo listar {carpeta_piezas}: {err}")
        return conteo_cara

    for nombre in entradas:
        ruta = os.path.join(carpeta_piezas, nombre)
        if os.path.isdir(ruta):
            continue
        if not nombre.lower().endswith(".jpg"):
            continue
        caras = [
            c for c in _caras_para_pieza(nombre, mapa_por_cara)
            if c in validas_destino
        ]
        if not caras:
            caras = [SUBCARPETA_OTROS_PIEZAS]
        pieza = _extraer_pieza_de_jpg(nombre)
        pieza_folder = _nombre_carpeta_pieza(pieza)
        ruta_fuente = None
        for idx_cara, destino_sub in enumerate(caras):
            if destino_sub not in conteo_cara:
                conteo_cara[destino_sub] = 0
                piezas_por_cara[destino_sub] = set()
            destino_dir = os.path.join(carpeta_piezas, destino_sub, pieza_folder)
            try:
                os.makedirs(destino_dir, exist_ok=True)
                destino = os.path.join(destino_dir, nombre)
                if os.path.exists(destino):
                    os.remove(destino)
                if idx_cara == 0:
                    shutil.move(ruta, destino)
                    ruta_fuente = destino
                else:
                    shutil.copy2(ruta_fuente, destino)
                conteo_cara[destino_sub] += 1
                piezas_por_cara[destino_sub].add(pieza_folder)
            except OSError as err:
                print(
                    f"AVISO: no se pudo colocar '{nombre}' en "
                    f"{destino_sub}/{pieza_folder}/: {err}"
                )
        if len(caras) > 1:
            print(
                f"  Pieza multi-cara '{pieza_folder}': "
                f"copiada a {', '.join(caras)}"
            )

    _limpiar_carpetas_cara_vacias_y_legacy(carpeta_piezas)

    print("  PIEZAS_ACOTADAS por cara:")
    for sub in list(activas):
        if sub not in conteo_cara:
            continue
        print(
            f"    {sub}: {conteo_cara[sub]} JPG en "
            f"{len(piezas_por_cara.get(sub, ()))} piezas"
        )
    return conteo_cara


def _clasificacion_para_pieza(nombre_archivo, mapa_por_clasificacion):
    """
    Devuelve la clasificación asignada al JPG según el mapa
    ``{clasificación: set(nombres_pieza)}`` construido leyendo el iProperty.

    Usa el ITEM de ``JOB__ITEM__MEDIDA_VALOR`` (no el nombre completo del
    archivo) y admite JPG cortos (``VT-5657-R1``) contra el nombre completo
    del mapa (``VT-5657-R1-D3000 LIFTING LUG…``).
    """
    if not mapa_por_clasificacion:
        return None

    claves_archivo = set()
    for cand in _candidatos_nombre_pieza_jpg(nombre_archivo):
        c = _clave_pieza(cand)
        if c:
            claves_archivo.add(c)
    if not claves_archivo:
        return None

    mejor_clase = None
    mejor_len = 0
    sin_cf = SUBCARPETA_SIN_CLASIFICAR.casefold()
    for clase, piezas in mapa_por_clasificacion.items():
        # No preferir "SIN CLASIFICACION" si hay otro match válido.
        if str(clase).casefold() == sin_cf:
            continue
        for pieza in piezas:
            clave_pieza = _clave_pieza(pieza)
            if not clave_pieza:
                continue
            for ca in claves_archivo:
                if _match_claves_pieza(ca, clave_pieza):
                    if len(clave_pieza) > mejor_len:
                        mejor_clase = clase
                        mejor_len = len(clave_pieza)
    return mejor_clase


def _es_jpg_despliegue(nombre_archivo, staging_marker=None):
    """True si la captura viene del flat pattern (corte plano)."""
    marker = str(staging_marker or "").upper()
    nombre_u = str(nombre_archivo or "").upper()
    if STAGING_DESPLIEGUE.upper() in marker or "DESPLIEGUE" in marker:
        return True
    return "DESPLIEGUE" in nombre_u


def _es_jpg_estanado(nombre_archivo, staging_marker=None, nombre_pieza=None):
    """True si debe ir a Estañado Busbar (cobre SIN_COTA / hoja ESTANIADO)."""
    marker = str(staging_marker or "").upper()
    nombre_u = str(nombre_archivo or "").upper()
    if STAGING_ESTANIADO.upper() in marker or "ESTANIADO" in marker:
        return True
    if "SIN_COTA" not in nombre_u:
        return False
    try:
        from piezas_cobre import es_pieza_cobre

        pieza = nombre_pieza or _extraer_pieza_de_jpg(nombre_archivo)
        return bool(es_pieza_cobre(pieza))
    except Exception:
        return "SIN_COTA" in nombre_u


def _es_cobre_nombre(nombre_archivo, nombre_pieza=None):
    """True si el JPG pertenece a pieza cobre (ABB/GENE/RLG)."""
    try:
        from piezas_cobre import es_pieza_cobre

        pieza = nombre_pieza or _extraer_pieza_de_jpg(nombre_archivo)
        return bool(es_pieza_cobre(pieza))
    except Exception:
        return False


def _iter_jpgs_para_reorg(carpeta_piezas):
    """
    JPG en la raíz de PIEZAS_ACOTADAS + staging ``_STAGING_*``.

    Yields ``(ruta_abs, nombre_archivo, staging_marker|None)``.
    """
    try:
        entradas = list(os.listdir(carpeta_piezas))
    except OSError:
        return
    for nombre in entradas:
        ruta = os.path.join(carpeta_piezas, nombre)
        if os.path.isfile(ruta) and nombre.lower().endswith(".jpg"):
            yield ruta, nombre, None
            continue
        if not os.path.isdir(ruta):
            continue
        if nombre not in (STAGING_DESPLIEGUE, STAGING_ESTANIADO):
            continue
        try:
            hijos = list(os.listdir(ruta))
        except OSError:
            continue
        for hijo in hijos:
            if not hijo.lower().endswith(".jpg"):
                continue
            yield os.path.join(ruta, hijo), hijo, nombre


def _destino_dirs_clasificacion(
    carpeta_piezas, destino_sub, nombre_archivo, staging_marker=None
):
    """
    Devuelve ``(destino_dir, clave_conteo, con_carpeta_pieza)``.

    Árbol acordado (ver ``CLASIFICACIONES_LOG``). Estañado Busbar: JPG sueltos.
    """
    pieza = _extraer_pieza_de_jpg(nombre_archivo)
    pieza_folder = _nombre_carpeta_pieza(pieza)
    cobre = _es_cobre_nombre(nombre_archivo, pieza)
    dest = str(destino_sub or "").casefold()

    if _es_jpg_estanado(nombre_archivo, staging_marker, pieza):
        destino_dir = os.path.join(carpeta_piezas, SUBCARPETA_ESTANIADO_BUSBAR)
        return destino_dir, SUBCARPETA_ESTANIADO_BUSBAR, False

    # Flat (despliegue): metal → Plasma y Laser; cobre → Maquinado/Corte Busbar.
    if _es_jpg_despliegue(nombre_archivo, staging_marker) or dest in (
        "plasma",
        "plasma doblado",
    ):
        if cobre:
            destino_dir = os.path.join(
                carpeta_piezas,
                "Corte",
                "Maquinado",
                "Corte Busbar",
                pieza_folder,
            )
            return destino_dir, "Corte/Maquinado/Corte Busbar", True
        destino_dir = os.path.join(
            carpeta_piezas,
            "Corte",
            "Plasma y Laser",
            "Corte metal",
            pieza_folder,
        )
        return destino_dir, "Corte/Plasma y Laser/Corte metal", True

    if dest == "maquinado":
        if cobre:
            destino_dir = os.path.join(
                carpeta_piezas,
                "Corte",
                "Maquinado",
                "Corte Busbar",
                pieza_folder,
            )
            return destino_dir, "Corte/Maquinado/Corte Busbar", True
        destino_dir = os.path.join(
            carpeta_piezas,
            "Corte",
            "Maquinado",
            "Maquinados metal",
            pieza_folder,
        )
        return destino_dir, "Corte/Maquinado/Maquinados metal", True

    # Doblado (iProp Doblado) o Corte doblado histórico → Doblado/Metal|Busbar.
    if dest in ("doblado", "corte"):
        if cobre:
            destino_dir = os.path.join(
                carpeta_piezas, "Doblado", "Busbar", pieza_folder
            )
            return destino_dir, "Doblado/Busbar", True
        destino_dir = os.path.join(
            carpeta_piezas, "Doblado", "Metal", pieza_folder
        )
        return destino_dir, "Doblado/Metal", True

    destino_dir = os.path.join(carpeta_piezas, destino_sub, pieza_folder)
    return destino_dir, destino_sub, True


def _preparar_arbol_clasificacion(carpeta_piezas):
    """Crea carpetas raíz + anidadas del árbol acordado."""
    os.makedirs(carpeta_piezas, exist_ok=True)
    for sub in SUBCARPETAS_CLASIFICACION_PIEZAS + (SUBCARPETA_SIN_CLASIFICAR,):
        os.makedirs(os.path.join(carpeta_piezas, sub), exist_ok=True)
    os.makedirs(
        os.path.join(carpeta_piezas, SUBCARPETA_ESTANIADO_BUSBAR), exist_ok=True
    )
    for nest in SUBCARPETAS_CORTE_PLASMA:
        os.makedirs(
            os.path.join(carpeta_piezas, "Corte", "Plasma y Laser", nest),
            exist_ok=True,
        )
    for nest in SUBCARPETAS_CORTE_MAQUINADO:
        os.makedirs(
            os.path.join(carpeta_piezas, "Corte", "Maquinado", nest),
            exist_ok=True,
        )
    for nest in SUBCARPETAS_DOBLADO_ANIDADAS:
        os.makedirs(
            os.path.join(carpeta_piezas, "Doblado", nest), exist_ok=True
        )


def _reorganizar_piezas_por_clasificacion(carpeta_piezas, mapa_por_clasificacion):
    """
    Mueve cada JPG a su carpeta del árbol acordado.

    Estructura resultante en PIEZAS_ACOTADAS:
        Almacén/<PIEZA>/*.jpg
        Corte/Plasma y Laser/Corte metal/<PIEZA>/*.jpg
        Corte/Maquinado/Corte Busbar/<PIEZA>/*.jpg
        Corte/Maquinado/Maquinados metal/<PIEZA>/*.jpg
        Doblado/Metal/<PIEZA>/*.jpg | Doblado/Busbar/<PIEZA>/*.jpg
        Estañado Busbar/*.jpg
        SIN CLASIFICACION/<PIEZA>/*.jpg

    Piezas sin match en el mapa caen en ``SIN CLASIFICACION/``.
    """
    try:
        _preparar_arbol_clasificacion(carpeta_piezas)
    except OSError as err:
        print(
            f"AVISO: no se pudieron preparar subcarpetas de PIEZAS_ACOTADAS "
            f"por clasificación: {err}"
        )
        return {}

    conteo = {sub: 0 for sub in CLASIFICACIONES_LOG}
    piezas_por_clase = {sub: set() for sub in CLASIFICACIONES_LOG}

    clases_validas_cf = {c.casefold() for c in SUBCARPETAS_CLASIFICACION_PIEZAS}

    for ruta, nombre, staging in _iter_jpgs_para_reorg(carpeta_piezas):
        clase = _clasificacion_para_pieza(nombre, mapa_por_clasificacion)
        if clase and clase.casefold() in clases_validas_cf:
            destino_sub = next(
                s for s in SUBCARPETAS_CLASIFICACION_PIEZAS
                if s.casefold() == clase.casefold()
            )
        else:
            destino_sub = SUBCARPETA_SIN_CLASIFICAR
        destino_dir, clave_conteo, _con_pieza = _destino_dirs_clasificacion(
            carpeta_piezas, destino_sub, nombre, staging
        )
        pieza_folder = _nombre_carpeta_pieza(_extraer_pieza_de_jpg(nombre))
        try:
            if _mover_jpg_a_destino(ruta, destino_dir, nombre):
                conteo[clave_conteo] = conteo.get(clave_conteo, 0) + 1
                if clave_conteo.startswith("Corte/"):
                    conteo["Corte"] = conteo.get("Corte", 0) + 1
                piezas_por_clase.setdefault(clave_conteo, set()).add(pieza_folder)
                if clave_conteo.startswith("Corte/"):
                    piezas_por_clase.setdefault("Corte", set()).add(pieza_folder)
        except OSError as err:
            print(
                f"AVISO: no se pudo mover '{nombre}' a "
                f"{clave_conteo}/: {err}"
            )

    # Vaciar staging vacío.
    for staging in (STAGING_DESPLIEGUE, STAGING_ESTANIADO):
        staging_dir = os.path.join(carpeta_piezas, staging)
        if os.path.isdir(staging_dir):
            try:
                if not os.listdir(staging_dir):
                    os.rmdir(staging_dir)
                else:
                    shutil.rmtree(staging_dir, ignore_errors=True)
            except OSError:
                pass

    print("  PIEZAS_ACOTADAS por clasificación:")
    for sub in CLASIFICACIONES_LOG:
        n = conteo.get(sub, 0)
        n_piezas = len(piezas_por_clase.get(sub, ()))
        print(f"    {sub}: {n} JPG en {n_piezas} piezas")
    return conteo


def _recuperar_antes_de_piezas(inv_app, plano):
    """
    Tras COTAS_POR_REFERENCIA Inventor queda saturado; sin este respiro
    creador_vistas falla al leer AllLeafOccurrences / CopyTo.
    """
    print("  Recuperando Inventor antes de PIEZAS_ACOTADAS...")
    try:
        _eliminar_hojas_piezas_residuales(plano)
    except Exception as err:
        print(f"  AVISO limpiando hojas residuales de piezas: {err}")
    try:
        inv_app.SilentOperation = False
        inv_app.ScreenUpdating = True
    except Exception:
        pass
    try:
        inv_app.ActiveDocument.Update()
    except Exception:
        pass
    try:
        inv_app.ActiveView.Update()
    except Exception:
        pass
    try:
        inv_app.UserInterfaceManager.DoEvents()
    except Exception:
        pass
    for _ in range(20):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.1)
    print("  Recuperación lista.")


def _reactivar_machote(inv_app):
    """El flujo legado termina en la última pieza; dejar visible el machote."""
    try:
        plano = _obtener_plano_activo(inv_app)
        hoja = _encontrar_hoja_machote(plano)
        if hoja is not None:
            hoja.Activate()
    except Exception:
        pass


def _prevalidar_cuatro_caras(ensamble):
    """
    Comprueba que existan cuatro paredes físicas verificables.

    Acepta segmentos nombrados en la raíz (Vantran) o un subensamble
    estructural que demuestre geométricamente sus cuatro laterales (OTC).
    Nunca habilita el flujo solo porque haya cuatro IAM arbitrarios.
    """
    try:
        return _resolver_contenedor_paredes(ensamble, registrar=False)
    except Exception:
        return {"valido": False, "motivo": "error durante prevalidación"}


def _limpiar_salida_parcial_caras(carpeta_tanque):
    """Borra únicamente la salida de caras que no es válida para este tanque."""
    carpeta = os.path.join(carpeta_tanque, "COTAS_POR_REFERENCIA")
    try:
        shutil.rmtree(carpeta)
    except FileNotFoundError:
        pass
    except OSError as error:
        print(f"AVISO: no se pudo limpiar salida parcial de caras: {error}")


def ejecutar():
    print("=" * 62)
    print(" COTAS ABIGAIL - TANQUE COMPLETO")
    print("=" * 62)

    pythoncom.CoInitialize()
    inv_app = None
    ok = False
    try:
        inv_app = conectar_inventor()
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            print(
                "ERROR: no se pudo recuperar el plano o ensamble "
                "para cotas por pieza."
            )
            return False

        from producto_tipo import (
            aplicar_unidad_producto,
            clasificar_producto,
            redirigir_si_board,
        )

        try:
            info = clasificar_producto(ensamble)
            aplicar_unidad_producto(ensamble=ensamble, info=info)
        except Exception as exc_u:
            print(f"AVISO unidades: {exc_u}")

        desvio = redirigir_si_board(
            inv_app,
            plano,
            ensamble,
            origen_flujo="TANQUE_COMPLETO",
            gestionar_com_board=False,
            limpiar=True,
        )
        if desvio is not None:
            return bool(desvio)

        carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        alcance = _prevalidar_cuatro_caras(ensamble)
        if alcance.get("valido"):
            contenedor = alcance.get("contenedor")
            fuente = (
                str(contenedor.Name)
                if contenedor is not None
                else "segmentos nombrados en la raíz"
            )
            print(
                "Paso 1/2: Cotando accesorios por caras del tanque "
                f"(alcance: {fuente})..."
            )
            # El orquestador ya posee el contexto COM: no hacer CoUninitialize
            # intermedio (eso dejó Inventor inestable al pasar a piezas).
            if not ejecutar_caras(gestionar_com=False):
                print("ERROR: no se completaron las cotas por cara.")
                return False
            # La rutina de caras reactiva el machote; refrescar handles.
            plano = _obtener_plano_activo(inv_app)
            ensamble = _obtener_ensamble_principal(inv_app)
            carpeta_tanque = _carpeta_salida_tanque(plano, ensamble)
        else:
            print(
                "Paso 1/2: Caras omitidas; no se identificaron cuatro "
                "paredes físicas verificables "
                f"({alcance.get('motivo', 'sin detalle')})."
            )
            _limpiar_machote(plano, inv_app)
            _limpiar_salida_parcial_caras(carpeta_tanque)

        carpeta_piezas = os.path.join(
            carpeta_tanque, CARPETA_PIEZAS_ACOTADAS
        )
        incremental = os.environ.get("PIEZAS_INCREMENTAL", "").strip() in ("1", "true", "TRUE", "yes")
        _limpiar_exportacion_piezas(carpeta_piezas, incremental=incremental)
        # Refrescar handles tras caras + vaciar hojas residuales + respiro COM.
        plano = _obtener_plano_activo(inv_app)
        ensamble = _obtener_ensamble_principal(inv_app)
        if plano is None or ensamble is None:
            print(
                "ERROR: se perdió el plano o ensamble al pasar a "
                "cotas por pieza."
            )
            return False
        _recuperar_antes_de_piezas(inv_app, plano)
        print("Paso 2/2: Ejecutando cotas originales por pieza...")
        print(f"  Carpeta de piezas acotadas: {carpeta_piezas}")
        if incremental:
            print("  Modo incremental ACTIVO (PIEZAS_INCREMENTAL=1).")

        # Detectar clasificación por proceso (iProperty) ANTES de exportar
        # para que quede persistida y disponible al reorganizar.
        print(
            "  Leyendo clasificación (iProperty) de cada pieza del ensamble..."
        )
        try:
            mapa_clasificacion = (
                generador_caras_tanque.detectar_mapa_piezas_por_clasificacion(
                    inv_app, ensamble
                )
            )
            if mapa_clasificacion:
                ruta_mapa = (
                    generador_caras_tanque
                    .guardar_mapa_piezas_por_clasificacion(
                        carpeta_tanque, mapa_clasificacion
                    )
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
                            f"  Corte anidado: {len(nombres_corte)} piezas → "
                            "flat + Doblado (+ Estañado cobre)."
                        )
                except Exception as exc_corte:
                    print(f"  AVISO configurar piezas Corte: {exc_corte}")
        except Exception as err:
            print(f"AVISO: fallo detectando clasificación por iProperty: {err}")

        # Equivalente funcional de COTAS_ILOGIC_ABIGAIL, pero sin llamar a
        # iLogic GenerarVistas (esa llamada queda bloqueada en este tanque).
        # Conserva creador_vistas + cotas.py + THK.py + exportación JPG.
        ok = bool(
            ejecutar_flujo_desde_app(
                inv_app,
                ensamble,
                plano,
                carpeta_salida=carpeta_piezas,
                incremental=incremental,
            )
        )
        if ok:
            try:
                _reorganizar_piezas_por_clasificacion(
                    carpeta_piezas,
                    dict(
                        generador_caras_tanque.LAST_PIEZAS_POR_CLASIFICACION
                    ),
                )
            except Exception as err:
                print(
                    f"AVISO: fallo en reorganización por clasificación de "
                    f"PIEZAS_ACOTADAS: {err}"
                )
        return ok
    finally:
        if inv_app is not None:
            _reactivar_machote(inv_app)
        pythoncom.CoUninitialize()
        if ok:
            print("PROCESO COMPLETO: caras y piezas acotadas exportadas.")


if __name__ == "__main__":
    sys.exit(0 if ejecutar() else 1)
