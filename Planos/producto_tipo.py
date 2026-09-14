# -*- coding: utf-8 -*-
"""
Clasificación de producto para COTAS ABIGAIL.

Separa TANQUE (Vantran/OTC/SWE/…) de BOARD/GIGA tablero eléctrico.
Las reglas iLogic y los generadores llaman ``clasificar_producto`` /
``redirigir_si_board`` antes del flujo de caras de tanque.
"""

from __future__ import annotations

import os
import re
from typing import Any

TIPO_TANQUE = "TANQUE"
TIPO_BOARD = "BOARD"
TIPO_DESCONOCIDO = "DESCONOCIDO"

# Nombre / ruta del IAM (tablero).
_RE_BOARD_NOMBRE = re.compile(
    r"(BOARD)|(\d{3,5}\s*[-_]?\s*BOARD)|(GIGA\s+BOARD)",
    re.IGNORECASE,
)
# Señales de tanque tradicional (ganan sobre GIGA suelto).
_RE_TANQUE_NOMBRE = re.compile(
    r"(VANTRAN)|(SUNBELT)|(\bOTC\b)|(\bSWE\b)|"
    r"(SEGMENTO)|(TOP[\s_-]?COVER)|(1246|1247|1248)|"
    r"(\bTANK\b)|(CASCO)|(SOLERA)",
    re.IGNORECASE,
)
# Piezas típicas de board eléctrico en 1er nivel.
_RE_BOARD_PIEZA = re.compile(
    r"(^GENE[-_])|(^GEN1[-_])|(^ABB-42-BCK)|(^9919-[FMP]-)|"
    r"(AcuCT)|(CITEL)|(PM8000)|(carriage\s+bolt\s+stack)|"
    r"(\bstack\.iam\b)|(\bisolator\s+stack)",
    re.IGNORECASE,
)
_RE_HW_NOMBRE = re.compile(
    r"(WASHER|NUT|BOLT|SCREW|RIVET|(^HW[-_])|STACK|ISOLATOR)",
    re.IGNORECASE,
)


def _texto_ensamble(ensamble) -> str:
    partes = []
    for attr in ("DisplayName", "FullFileName", "FullDocumentName"):
        try:
            v = getattr(ensamble, attr, None)
            if v:
                partes.append(str(v))
        except Exception:
            pass
    try:
        ff = str(getattr(ensamble, "FullFileName", "") or "")
        if ff:
            partes.append(os.path.basename(ff))
            partes.append(os.path.dirname(ff))
    except Exception:
        pass
    return " | ".join(partes)


def _muestra_nombres_raiz(ensamble, limite: int = 80) -> list[str]:
    """Nombres de ocurrencias de 1er nivel (barato, sin recursión)."""
    nombres: list[str] = []
    try:
        occs = ensamble.ComponentDefinition.Occurrences
        n = int(occs.Count)
    except Exception:
        return nombres
    for i in range(1, min(n, limite) + 1):
        try:
            occ = occs.Item(i)
            if getattr(occ, "Suppressed", False):
                continue
            nom = str(occ.Name or "")
            try:
                doc = occ.Definition.Document
                base = os.path.splitext(
                    os.path.basename(str(doc.FullFileName or ""))
                )[0]
                if base:
                    nom = base
            except Exception:
                pass
            if nom:
                nombres.append(nom)
        except Exception:
            continue
    return nombres


def _puntaje_board_desde_hijos(nombres: list[str]) -> tuple[int, int, list[str]]:
    """(hits_board, hits_hw, señales)."""
    hits_board = 0
    hits_hw = 0
    senales: list[str] = []
    for nom in nombres:
        if _RE_BOARD_PIEZA.search(nom):
            hits_board += 1
            if len(senales) < 8:
                senales.append(nom)
        if _RE_HW_NOMBRE.search(nom):
            hits_hw += 1
    return hits_board, hits_hw, senales


def clasificar_producto(ensamble, escanear_hijos: bool = True) -> dict[str, Any]:
    """
    Devuelve ``{tipo, familia, motivo, senales}``.

    Prioridad:
      1. Nombre BOARD / GIGA BOARD → BOARD
      2. Señales claras de tanque en nombre → TANQUE
      3. Árbol 1er nivel muy “eléctrico” (GENE/ABB/9919-F…) → BOARD
      4. GIGA sin TANK/BOARD → DESCONOCIDO (no forzar; puede ser tanque GIGA)
      5. Default → TANQUE (flujos actuales)
    """
    texto = _texto_ensamble(ensamble)
    senales: list[str] = []

    if _RE_BOARD_NOMBRE.search(texto):
        return {
            "tipo": TIPO_BOARD,
            "familia": "BOARD",
            "motivo": "nombre/ruta contiene BOARD",
            "senales": senales,
        }

    if _RE_TANQUE_NOMBRE.search(texto):
        fam = "TANQUE"
        tu = texto.upper()
        for marca in (
            "VANTRAN",
            "SUNBELT",
            "OTC",
            "SWE",
            "PTT",
            "GIGA",
        ):
            if marca in tu:
                fam = marca
                break
        return {
            "tipo": TIPO_TANQUE,
            "familia": fam,
            "motivo": "nombre/ruta con señales de tanque",
            "senales": senales,
        }

    hits_board = 0
    hits_hw = 0
    if escanear_hijos:
        nombres = _muestra_nombres_raiz(ensamble)
        hits_board, hits_hw, senales = _puntaje_board_desde_hijos(nombres)
        # Umbral: varios componentes eléctricos típicos de board.
        if hits_board >= 4 and hits_board >= max(3, hits_hw // 8):
            return {
                "tipo": TIPO_BOARD,
                "familia": "BOARD",
                "motivo": (
                    f"arbol 1er nivel electrico "
                    f"(board={hits_board}, hw~{hits_hw})"
                ),
                "senales": senales,
            }

    if re.search(r"\bGIGA\b", texto, re.IGNORECASE):
        return {
            "tipo": TIPO_DESCONOCIDO,
            "familia": "GIGA",
            "motivo": "GIGA sin BOARD/TANK claro; no se fuerza desvío",
            "senales": senales,
        }

    return {
        "tipo": TIPO_TANQUE,
        "familia": "TANQUE",
        "motivo": "default tanque (sin señales de board)",
        "senales": senales,
    }


def es_board(ensamble, escanear_hijos: bool = True) -> bool:
    return clasificar_producto(ensamble, escanear_hijos=escanear_hijos)[
        "tipo"
    ] == TIPO_BOARD


def unidad_para_producto(ensamble=None, info=None) -> str:
    """
    ``mm`` para GIGA / BOARD; ``in`` para tanques Vantran/OTC/….
    """
    if info is None and ensamble is not None:
        info = clasificar_producto(ensamble)
    info = info or {}
    tipo = str(info.get("tipo") or "")
    familia = str(info.get("familia") or "").upper()
    if tipo == TIPO_BOARD or familia in ("BOARD", "GIGA"):
        return "mm"
    return "in"


def aplicar_unidad_producto(ensamble=None, info=None) -> str:
    """Setea ``cota_estilo`` y modo de nombre de pieza según producto."""
    from cota_estilo import set_unidad_cota

    u = unidad_para_producto(ensamble=ensamble, info=info)
    set_unidad_cota(u)
    try:
        from creador_vistas import set_nombre_pieza_completo

        # GIGA/BOARD: cada IPT distinto = carpeta/pieza propia (sin truncar).
        set_nombre_pieza_completo(u == "mm")
    except Exception:
        pass
    print(f"  Unidades de cota: {u}", flush=True)
    return u


def redirigir_si_board(
    inv_app,
    plano,
    ensamble,
    origen_flujo: str = "",
    gestionar_com_board: bool = False,
    limpiar: bool = True,
) -> bool | None:
    """
    Si el ensamble es BOARD, ejecuta ``generador_board`` y devuelve bool.

    Si no es BOARD, devuelve ``None`` para que el caller siga su flujo normal.
    """
    info = clasificar_producto(ensamble)
    tipo = info.get("tipo")
    tag = f"[{origen_flujo}] " if origen_flujo else ""
    print(
        f"{tag}Producto: {tipo} / {info.get('familia')} "
        f"— {info.get('motivo')}",
        flush=True,
    )
    try:
        aplicar_unidad_producto(ensamble=ensamble, info=info)
    except Exception:
        pass
    if tipo != TIPO_BOARD:
        return None

    print(
        f"{tag}Desvío → flujo BOARD (no se usan caras TOP/SEGM/BASE de tanque).",
        flush=True,
    )
    if info.get("senales"):
        print(
            f"{tag}Señales: {', '.join(str(s) for s in info['senales'][:6])}",
            flush=True,
        )

    from generador_board import ejecutar as ejecutar_board

    return bool(
        ejecutar_board(
            gestionar_com=gestionar_com_board,
            inv_app=inv_app,
            plano=plano,
            ensamble=ensamble,
            limpiar=limpiar,
        )
    )
