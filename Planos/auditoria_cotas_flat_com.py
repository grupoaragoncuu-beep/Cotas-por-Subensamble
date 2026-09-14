# -*- coding: utf-8 -*-
"""
Auditoría COM post-flujo (flat Corte/Corte):

  - Relee cada hoja XCENTRO / YCENTRO / DIAMETRO_H*
  - Compara el texto dibujado vs geometría exacta (cm→mm sin redondeo corto)
  - Detecta cotas "paso entre barrenos" (origen caído en un agujero)
  - Re-acota fallos y deja listos para re-export JPG

Uso:
  from auditoria_cotas_flat_com import auditar_y_reparar_plano
  ok, reparadas = auditar_y_reparar_plano(inv_app, plano)
"""
from __future__ import annotations

import re
from typing import Iterable

from cota_estilo import texto_cota_limpio, asegurar_unidad_cota, get_unidad_cota
from diametro import _silueta_placa_vista, _agrupar_diametros_grupos, _barrenos_en_vista

_RE_NUM = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
# Tol. mm: no aceptar ni el ultimo digito "a ojo"
_TOL_MM = 0.0005
_MAX_PASADAS = 2


def _escala(vista) -> float:
    try:
        esc = abs(float(vista.Scale))
        if esc > 1e-12:
            return esc
    except Exception:
        pass
    return 1.0


def _cm_hoja_a_mm(cm_hoja: float, vista) -> float:
    """Distancia en hoja (cm) → mm de modelo, exacto (sin round intermedio)."""
    cm_modelo = abs(float(cm_hoja)) / _escala(vista)
    if get_unidad_cota() == "mm":
        return cm_modelo * 10.0
    return cm_modelo / 2.54 * 25.4


def _parse_mm(texto) -> float | None:
    if texto is None:
        return None
    m = _RE_NUM.search(str(texto).replace(",", "."))
    if not m:
        return None
    try:
        return abs(float(m.group(0)))
    except ValueError:
        return None


def _leer_texto_cota_hoja(hoja) -> str:
    """Prioridad: nota XY= → TextBoxes de sketch → GeneralDimensions."""
    try:
        notes = hoja.DrawingNotes.GeneralNotes
        for i in range(1, int(notes.Count) + 1):
            try:
                txt = str(notes.Item(i).Text or "").strip()
            except Exception:
                continue
            up = txt.upper()
            if up.startswith("XY=") or "TYP" in up:
                if _parse_mm(txt) is not None:
                    return txt
    except Exception:
        pass

    sketches = None
    for attr in ("Sketches", "DrawingSketches"):
        try:
            sketches = getattr(hoja, attr)
            break
        except Exception:
            continue
    if sketches is not None:
        mejor = ""
        try:
            for i in range(1, int(sketches.Count) + 1):
                sk = sketches.Item(i)
                try:
                    boxes = sk.TextBoxes
                except Exception:
                    continue
                for j in range(1, int(boxes.Count) + 1):
                    try:
                        t = str(boxes.Item(j).Text or "").strip()
                    except Exception:
                        continue
                    if _parse_mm(t) is None:
                        continue
                    if len(t) >= len(mejor):
                        mejor = t
        except Exception:
            pass
        if mejor:
            return mejor

    try:
        dims = hoja.DrawingDimensions.GeneralDimensions
        for i in range(1, int(dims.Count) + 1):
            dim = dims.Item(i)
            try:
                return texto_cota_limpio(dim.ModelValue, hoja)
            except Exception:
                try:
                    return str(dim.Text.Text or "")
                except Exception:
                    continue
    except Exception:
        pass
    return ""


def _niveles_esperados_mm(vista, hoja, eje: str) -> list[float]:
    """Distancias distintas origen-placa → centro, en mm exactos."""
    from barrenos_xy_despliegue import _centros_barrenos

    sil = _silueta_placa_vista(vista)
    if not sil:
        return []
    ox, _mx, oy, _my, _ = sil
    try:
        tg = hoja.Parent.Application.TransientGeometry
    except Exception:
        return []
    barrenos = _centros_barrenos(vista, tg)
    vals = []
    for b in barrenos:
        try:
            cx, cy = float(b["cx"]), float(b["cy"])
        except Exception:
            continue
        d_cm = (cx - ox) if eje == "X" else (cy - oy)
        if d_cm < 0.05:
            continue
        vals.append(_cm_hoja_a_mm(d_cm, vista))
    vals.sort()
    unicos = []
    for v in vals:
        if not unicos or abs(v - unicos[-1]) > _TOL_MM:
            unicos.append(v)
    return unicos


def _pasos_entre_barrenos_mm(vista, hoja, eje: str) -> list[float]:
    """Pasos centro-centro en el eje (para detectar cotas malas)."""
    from barrenos_xy_despliegue import _centros_barrenos

    try:
        tg = hoja.Parent.Application.TransientGeometry
    except Exception:
        return []
    barrenos = _centros_barrenos(vista, tg)
    coords = []
    for b in barrenos:
        try:
            coords.append(float(b["cx"] if eje == "X" else b["cy"]))
        except Exception:
            continue
    coords = sorted(set(round(c, 6) for c in coords))
    pasos = []
    for i in range(len(coords) - 1):
        d_cm = abs(coords[i + 1] - coords[i])
        if d_cm < 0.05:
            continue
        pasos.append(_cm_hoja_a_mm(d_cm, vista))
    return pasos


def _coincide(valor: float, candidatos: Iterable[float], tol: float = _TOL_MM) -> bool:
    return any(abs(valor - c) <= tol for c in candidatos)


def _hoja_xy_falla(hoja, vista, eje: str) -> tuple[bool, str]:
    texto = _leer_texto_cota_hoja(hoja)
    medido = _parse_mm(texto)
    if medido is None:
        return True, "sin texto numerico"

    esperados = _niveles_esperados_mm(vista, hoja, eje)
    if not esperados:
        return True, "sin barrenos/silueta placa"

    if _coincide(medido, esperados):
        pasos = _pasos_entre_barrenos_mm(vista, hoja, eje)
        err_esp = min(abs(medido - e) for e in esperados)
        if pasos:
            err_paso = min(abs(medido - p) for p in pasos)
            if err_paso + 1e-9 < err_esp and err_paso <= _TOL_MM:
                return True, f"parece PASO entre barrenos ({medido})"
        if "." in str(texto):
            dec = str(texto).split(".", 1)[1]
            dec = re.split(r"\D", dec)[0]
            if len(dec) < 3:
                return True, f"pocos decimales ({texto})"
        return False, "ok"

    pasos = _pasos_entre_barrenos_mm(vista, hoja, eje)
    if _coincide(medido, pasos, tol=max(_TOL_MM, 0.05)):
        return True, f"PASO entre barrenos {medido} (esperado origen→centro)"
    preview = ", ".join(f"{e:.3f}" for e in esperados[:5])
    return True, f"valor {medido} no cuadra con niveles [{preview}]"


def _reparar_hoja_xy(hoja, vista, inv_app, tg, eje: str) -> bool:
    """Re-dibuja la cota origen-placa → centro del nivel."""
    from barrenos_xy_despliegue import (
        _centros_barrenos,
        _dibujar_cota_centro_sketch,
        _limpiar_dims_y_sketches,
        _valor_desde_hoja,
        _tol_coincidencia_hoja,
    )
    from diametro import _marcar_barrenos_azules

    sil = _silueta_placa_vista(vista)
    if not sil:
        return False
    ox, _mx, oy, _my, _ = sil
    barrenos = _centros_barrenos(vista, tg)
    if not barrenos:
        return False

    texto = _leer_texto_cota_hoja(hoja)
    medido = _parse_mm(texto)
    mejor = None
    mejor_score = 1e18
    nivel_cx = nivel_cy = None
    for b in barrenos:
        try:
            cx, cy = float(b["cx"]), float(b["cy"])
        except Exception:
            continue
        d_cm = (cx - ox) if eje == "X" else (cy - oy)
        if d_cm < 0.05:
            continue
        mm = _cm_hoja_a_mm(d_cm, vista)
        score = abs(mm - medido) if medido is not None else mm
        if score < mejor_score:
            mejor_score = score
            mejor = b
            nivel_cx, nivel_cy = cx, cy

    if mejor is None:
        return False

    esperados = _niveles_esperados_mm(vista, hoja, eje)
    pasos = _pasos_entre_barrenos_mm(vista, hoja, eje)
    if (
        medido is not None
        and _coincide(medido, pasos, tol=0.05)
        and not _coincide(medido, esperados, tol=0.05)
    ):
        target = min(esperados) if esperados else None
        if target is not None:
            for b in barrenos:
                try:
                    cx, cy = float(b["cx"]), float(b["cy"])
                except Exception:
                    continue
                d_cm = (cx - ox) if eje == "X" else (cy - oy)
                if abs(_cm_hoja_a_mm(d_cm, vista) - target) <= max(_TOL_MM, 0.05):
                    mejor = b
                    nivel_cx, nivel_cy = cx, cy
                    break

    cx, cy = float(nivel_cx), float(nivel_cy)
    tol = _tol_coincidencia_hoja(vista, barrenos)
    dist_ref = (cx - ox) if eje == "X" else (cy - oy)
    miembros = []
    for b in barrenos:
        dist = (float(b["cx"]) - ox) if eje == "X" else (float(b["cy"]) - oy)
        if abs(dist - dist_ref) <= tol:
            miembros.append(b)
    if not miembros:
        miembros = [mejor]
    es_typ = len(miembros) >= 2

    if eje == "X":
        val_txt = _valor_desde_hoja(vista, hoja, ox, cx)
    else:
        val_txt = _valor_desde_hoja(vista, hoja, oy, cy)
    texto_vista = asegurar_unidad_cota(
        f"{val_txt} TYP" if es_typ else str(val_txt)
    )

    _limpiar_dims_y_sketches(hoja)
    ok = _dibujar_cota_centro_sketch(
        hoja,
        vista,
        tg,
        inv_app,
        eje,
        ox,
        oy,
        cx,
        cy,
        texto_vista,
        miembros_typ=miembros if es_typ else None,
    )
    if ok and es_typ:
        try:
            _marcar_barrenos_azules(hoja, vista, miembros, tg, inv_app)
        except Exception:
            pass
    try:
        for i in range(int(hoja.DrawingNotes.GeneralNotes.Count), 0, -1):
            try:
                n = hoja.DrawingNotes.GeneralNotes.Item(i)
                if str(n.Text or "").upper().startswith("XY="):
                    n.Delete()
            except Exception:
                pass
        nota = hoja.DrawingNotes.GeneralNotes.AddFitted(
            tg.CreatePoint2d(0.4, 0.4), f"XY={val_txt}"
        )
        try:
            nota.Visible = False
        except Exception:
            pass
    except Exception:
        pass
    return bool(ok)


def _hoja_hole_falla(hoja, vista) -> tuple[bool, str]:
    texto = _leer_texto_cota_hoja(hoja)
    medido = _parse_mm(texto)
    if medido is None:
        return True, "sin texto Ø"
    grupos = _agrupar_diametros_grupos(_barrenos_en_vista(vista))
    if not grupos:
        return True, "sin barrenos Ø"
    nombre = str(hoja.Name).upper()
    m = re.search(r"_H(\d+)", nombre)
    idx = int(m.group(1)) - 1 if m else 0
    if idx < 0 or idx >= len(grupos):
        idx = 0
    try:
        ref = max(grupos[idx], key=lambda x: float(x.get("tamaño") or 0))
        tam_hoja = float(ref.get("tamaño") or 0)
        if tam_hoja <= 0:
            return False, "sin ref Ø (skip)"
        # tamaño en cm de hoja → cm modelo → texto mm exacto
        tam_cm = tam_hoja / _escala(vista)
        esperado = _parse_mm(texto_cota_limpio(tam_cm, hoja))
        if esperado is None:
            return False, "sin ref Ø (skip)"
        if abs(medido - esperado) <= max(_TOL_MM, 0.01):
            return False, "ok"
        return True, f"Ø {medido} != {esperado:.6f}"
    except Exception as exc:
        return True, f"Ø check error: {exc}"


def auditar_y_reparar_plano(inv_app, plano, max_pasadas: int = _MAX_PASADAS):
    """
    Audita hojas XCENTRO/YCENTRO (y revisa HOLE). Repara XY fallidas.

    Returns
    -------
    (ok_global, nombres_reparados)
    """
    tg = inv_app.TransientGeometry
    reparadas: list[str] = []

    for pasada in range(1, max_pasadas + 1):
        print(f"\n[AUDIT COM] pasada {pasada}/{max_pasadas}")
        pendientes = []
        try:
            n_hojas = int(plano.Sheets.Count)
        except Exception:
            n_hojas = 0

        for i in range(1, n_hojas + 1):
            try:
                hoja = plano.Sheets.Item(i)
            except Exception:
                continue
            nombre = str(hoja.Name)
            up = nombre.upper()
            if int(getattr(hoja.DrawingViews, "Count", 0) or 0) < 1:
                continue
            vista = hoja.DrawingViews.Item(1)

            if "XCENTRO" in up or "YCENTRO" in up:
                eje = "X" if "XCENTRO" in up else "Y"
                falla, motivo = _hoja_xy_falla(hoja, vista, eje)
                if falla:
                    print(f"  FAIL {nombre.rsplit(':', 1)[0]}: {motivo}")
                    pendientes.append((hoja, vista, eje, nombre))
                else:
                    print(f"  OK   {nombre.rsplit(':', 1)[0]}")
            elif "DIAMETRO_H" in up:
                falla, motivo = _hoja_hole_falla(hoja, vista)
                if falla:
                    print(f"  WARN Ø {nombre.rsplit(':', 1)[0]}: {motivo}")

        if not pendientes:
            print("[AUDIT COM] todas las XY OK")
            break

        for hoja, vista, eje, nombre in pendientes:
            try:
                hoja.Activate()
            except Exception:
                pass
            ok = _reparar_hoja_xy(hoja, vista, inv_app, tg, eje)
            base = nombre.rsplit(":", 1)[0]
            if ok:
                print(f"  REPARADA {base}")
                if base not in reparadas:
                    reparadas.append(base)
            else:
                print(f"  NO SE PUDO REPARAR {base}")

    fallos_finales = 0
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper()
            if "XCENTRO" not in up and "YCENTRO" not in up:
                continue
            if int(hoja.DrawingViews.Count) < 1:
                continue
            eje = "X" if "XCENTRO" in up else "Y"
            falla, motivo = _hoja_xy_falla(hoja, hoja.DrawingViews.Item(1), eje)
            if falla:
                fallos_finales += 1
                print(f"  FINAL FAIL {hoja.Name}: {motivo}")
        except Exception:
            continue

    ok_global = fallos_finales == 0
    print(
        f"[AUDIT COM] reparadas={len(reparadas)} "
        f"fallos_finales={fallos_finales} ok={ok_global}"
    )
    return ok_global, reparadas


if __name__ == "__main__":
    import pythoncom
    import win32com.client
    from inventor_com import conectar_inventor

    pythoncom.CoInitialize()
    try:
        inv = conectar_inventor()
        plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
        ok, reps = auditar_y_reparar_plano(inv, plano)
        print("RESULT", ok, reps)
    finally:
        pythoncom.CoUninitialize()
