# -*- coding: utf-8 -*-
"""
Auditoría COM post-flujo (flat Corte/Corte) — TODAS las piezas:

  1) XCENTRO / YCENTRO  → texto vs geometría exacta; detecta "paso" entre barrenos
  2) DIAMETRO_H / HOLE   → Ø vs barrenos en vista
  3) DESPLIEGUE_LADO/THK → debe ser CANTO FLAT desdoblado (no perfil L/U);
                           valor vs Sheet Metal Thickness

Si falla: repara por COM (re-dibuja XY / recrea DESPLIEGUE_LADO flat + reacota THK)
y deja las hojas listas para re-export JPG.

Uso:
  from auditoria_cotas_flat_com import auditar_y_reparar_plano
  ok, reparadas = auditar_y_reparar_plano(inv_app, plano)
"""
from __future__ import annotations

import os
import re
from typing import Iterable

from cota_estilo import (
    texto_cota_limpio,
    asegurar_unidad_cota,
    get_unidad_cota,
    set_unidad_cota,
)
from diametro import _silueta_placa_vista, _agrupar_diametros_grupos, _barrenos_en_vista

_RE_NUM = re.compile(r"[-+]?\d+(?:[.,]\d+)?")
_RE_PULGADAS = re.compile(r"(?:^|\s)(in|pulg\.?)(?:\s|$)", re.IGNORECASE)
_TOL_MM = 0.0005
_TOL_THK_MM = 0.05  # Thickness vs cota: margen de lectura
_MAX_PASADAS = 2


def _forzar_mm_flat() -> None:
    """GIGA Board flat: cotas siempre en mm (nunca in)."""
    os.environ["SOLO_FLAT_CORTE"] = "1"
    set_unidad_cota("mm")


def _texto_en_pulgadas(texto) -> bool:
    t = str(texto or "")
    if _RE_PULGADAS.search(t):
        return True
    # Número chico típico de pulgada cuando la unidad activa es mm
    # (p.ej. 3.5 in mostrado sin sufijo) — se valida aparte vs esperados.
    return False


def _escala(vista) -> float:
    try:
        esc = abs(float(vista.Scale))
        if esc > 1e-12:
            return esc
    except Exception:
        pass
    return 1.0


def _cm_hoja_a_mm(cm_hoja: float, vista) -> float:
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
    """Nota XY=/THK= → TextBoxes sketch → GeneralDimensions (ModelValue)."""
    try:
        notes = hoja.DrawingNotes.GeneralNotes
        for i in range(1, int(notes.Count) + 1):
            try:
                txt = str(notes.Item(i).Text or "").strip()
            except Exception:
                continue
            up = txt.upper()
            if (
                up.startswith("XY=")
                or up.startswith("THK")
                or "TYP" in up
                or "Ø" in txt
                or "MM" in up
                or " IN" in up
            ):
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


def _coincide(valor: float, candidatos: Iterable[float], tol: float = _TOL_MM) -> bool:
    return any(abs(valor - c) <= tol for c in candidatos)


def _tg_desde_hoja_vista(hoja, vista):
    """TransientGeometry robusto; puede ser None (HLR no lo necesita)."""
    for getter in (
        lambda: hoja.Parent.Application.TransientGeometry,
        lambda: vista.Parent.Parent.Application.TransientGeometry,
        lambda: getattr(vista, "Application", None)
        and vista.Application.TransientGeometry,
    ):
        try:
            tg = getter()
            if tg is not None:
                return tg
        except Exception:
            continue
    try:
        from inventor_com import conectar_inventor

        return conectar_inventor().TransientGeometry
    except Exception:
        return None


def _niveles_esperados_mm(vista, hoja, eje: str) -> list[float]:
    """
    Niveles X/Y esperados en mm desde origen IL.
    NUNCA retornar [] solo porque falle TransientGeometry: HLR basta.
    """
    from barrenos_xy_despliegue import (
        _centros_barrenos,
        _centros_barrenos_hlr,
        _referencias_oval_xy,
    )
    from diametro import _origen_il_pieza, _barrenos_en_vista

    sil = _silueta_placa_vista(vista)
    origen = _origen_il_pieza(vista)
    if origen is not None:
        ox, oy = origen
    elif sil:
        ox, _mx, oy, _my, _ = sil
    else:
        return []

    # HLR primero (no depende de tg) — bug previo: return [] si fallaba tg
    barrenos = []
    try:
        barrenos = list(_centros_barrenos_hlr(vista)) + list(
            _referencias_oval_xy(vista)
        )
    except Exception:
        barrenos = []

    tg = _tg_desde_hoja_vista(hoja, vista)
    if tg is not None:
        try:
            fused = list(_centros_barrenos(vista, tg) or [])
            if fused:
                barrenos = fused
        except Exception:
            pass

    if not barrenos:
        try:
            for a in _barrenos_en_vista(vista) or []:
                barrenos.append(
                    {
                        "cx": float(a["cx"]),
                        "cy": float(a["cy"]),
                        "tamaño": float(a.get("tamaño") or 0.2),
                        "tipo": str(a.get("tipo") or "circulo"),
                    }
                )
        except Exception:
            pass

    vals = []
    for b in barrenos:
        try:
            cx, cy = float(b["cx"]), float(b["cy"])
        except Exception:
            continue
        d_cm = (cx - ox) if eje == "X" else (cy - oy)
        if abs(d_cm) < 0.02:
            continue
        try:
            vals.append(_cm_hoja_a_mm(d_cm, vista))
        except Exception:
            continue
    vals.sort()
    unicos = []
    for val in vals:
        if not unicos or abs(val - unicos[-1]) > _TOL_MM:
            unicos.append(val)
    return unicos


def _pasos_entre_barrenos_mm(vista, hoja, eje: str) -> list[float]:
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


def _activar_hoja(hoja) -> None:
    """Activa la hoja y espera a que la vista tenga curvas (HLR listo)."""
    import time

    try:
        hoja.Activate()
    except Exception:
        pass
    try:
        vista = hoja.DrawingViews.Item(1)
    except Exception:
        return
    # Update() no existe en todos los wrappers; no tumbar la ref
    for meth in ("_Update", "Update"):
        try:
            getattr(vista, meth)()
            break
        except Exception:
            pass
    for _ in range(8):
        try:
            if int(vista.DrawingCurves.Count) > 0:
                return
        except Exception:
            pass
        time.sleep(0.05)


def _hoja_xy_falla(hoja, vista, eje: str) -> tuple[bool, str]:
    from diametro import (
        _origen_il_pieza,
        _punto_cerca_contorno,
        _origen_flota_en_vacio,
    )

    _activar_hoja(hoja)
    try:
        vista = hoja.DrawingViews.Item(1)
    except Exception:
        pass

    if str(hoja.Name).upper().startswith("COPIA DE"):
        return True, "hoja Copia de (basura)"

    # 1) Origen DEBE existir y tocar geometría — nunca AABB flotante
    origen = _origen_il_pieza(vista)
    sil = _silueta_placa_vista(vista)
    if origen is None:
        return True, "sin origen IL sobre geometria"
    ox, oy = origen
    if _origen_flota_en_vacio(vista, ox, oy) or not _punto_cerca_contorno(
        vista, ox, oy
    ):
        return True, "origen IL flota en vacio (no toca pieza)"

    # Si AABB ≠ IL: el valor NO puede ser el del AABB (cota vieja mala)
    if sil is not None:
        aabb_ox, aabb_oy = float(sil[0]), float(sil[2])
        aabb_flota = _origen_flota_en_vacio(vista, aabb_ox, aabb_oy)
        if aabb_flota and (
            abs(aabb_ox - ox) > 0.05 or abs(aabb_oy - oy) > 0.05
        ):
            # forzar que el texto coincida con IL, no con AABB
            pass  # se valida abajo con esperados desde IL

    texto = _leer_texto_cota_hoja(hoja)
    if _texto_en_pulgadas(texto):
        return True, f"unidad IN (debe ser mm): {texto!r}"
    if get_unidad_cota() != "mm":
        _forzar_mm_flat()
    medido = _parse_mm(texto)
    if medido is None:
        return True, "sin texto numerico"
    # Valor en pulgadas colado con sufijo mm / sin unidad (p.ej. 3.5 en vez de 88.9)
    if medido < 12.0 and "MM" not in str(texto).upper():
        # sospechoso: se confirma abajo si cuadra con esperados/25.4
        pass

    # 2) Fantasmas
    try:
        from barrenos_xy_despliegue import (
            _centros_barrenos_modelo,
            _centros_barrenos,
        )

        tg = hoja.Parent.Application.TransientGeometry
        modelo = _centros_barrenos_modelo(vista, tg)
        actual = _centros_barrenos(vista, tg)
        if modelo and len(actual) > len(modelo) + 1:
            return True, (
                f"barrenos fantasma sospechosos "
                f"(vista={len(actual)} modelo={len(modelo)})"
            )
    except Exception:
        pass

    esperados = _niveles_esperados_mm(vista, hoja, eje)
    if not esperados:
        nb = -1
        try:
            from barrenos_xy_despliegue import _centros_barrenos

            tg = hoja.Parent.Application.TransientGeometry
            nb = len(_centros_barrenos(vista, tg) or [])
        except Exception:
            try:
                from diametro import _barrenos_en_vista

                nb = len(_barrenos_en_vista(vista) or [])
            except Exception:
                nb = -1
        if nb <= 0:
            return True, "sin barrenos detectados en vista (HLR no listo)"
        return True, (
            f"origen mal ubicado (barrenos={nb} pero ninguna cota >0 "
            f"desde IL)"
        )
    # Si el número cuadra con esperados/25.4 → se dibujó en pulgadas
    if medido < 50.0:
        en_in = [e / 25.4 for e in esperados if e >= 0.05]
        if en_in and _coincide(medido, en_in, tol=0.02):
            return True, f"valor en pulgadas ({medido}) — debe ser mm"

    # 3) Valor debe ser desde IL real
    if not _coincide(medido, esperados, tol=max(_TOL_MM, 0.05)):
        # ¿Cuadra con AABB flotante? → origen malo explícito
        if sil is not None:
            aabb_ox, aabb_oy = float(sil[0]), float(sil[2])
            try:
                tg = hoja.Parent.Application.TransientGeometry
                from barrenos_xy_despliegue import _centros_barrenos

                bars = _centros_barrenos(vista, tg)
                aabb_vals = []
                for b in bars:
                    d = (
                        float(b["cx"]) - aabb_ox
                        if eje == "X"
                        else float(b["cy"]) - aabb_oy
                    )
                    if d >= 0.05:
                        aabb_vals.append(_cm_hoja_a_mm(d, vista))
                if aabb_vals and _coincide(medido, aabb_vals, tol=0.05):
                    return True, (
                        "origen AABB flotante "
                        "(no toca esquina inferior-izquierda)"
                    )
            except Exception:
                pass
        pasos = _pasos_entre_barrenos_mm(vista, hoja, eje)
        if _coincide(medido, pasos, tol=max(_TOL_MM, 0.05)):
            return True, f"PASO entre barrenos {medido}"
        preview = ", ".join(f"{e:.3f}" for e in esperados[:5])
        return True, f"valor {medido} desfazado vs IL [{preview}]"

    pasos = _pasos_entre_barrenos_mm(vista, hoja, eje)
    err_esp = min(abs(medido - e) for e in esperados)
    if pasos:
        err_paso = min(abs(medido - p) for p in pasos)
        if err_paso + 1e-9 < err_esp and err_paso <= _TOL_MM:
            return True, f"parece PASO entre barrenos ({medido})"
    # Decimales: exigir al menos 3 si hay punto (no fallar por tener 6)
    if "." in str(texto):
        dec = str(texto).split(".", 1)[1]
        dec = re.split(r"\D", dec)[0]
        if len(dec) < 6:
            return True, f"pocos decimales (se exigen 6): {texto}"
    return False, "ok"


def _reparar_hoja_xy(hoja, vista, inv_app, tg, eje: str) -> bool:
    _forzar_mm_flat()
    from barrenos_xy_despliegue import (
        _centros_barrenos,
        _dibujar_cota_centro_sketch,
        _limpiar_dims_y_sketches,
        _valor_desde_hoja,
        _tol_coincidencia_hoja,
    )
    from diametro import _marcar_barrenos_azules, _origen_il_pieza

    sil = _silueta_placa_vista(vista)
    origen = _origen_il_pieza(vista)
    from diametro import _origen_flota_en_vacio

    if origen is None or _origen_flota_en_vacio(vista, origen[0], origen[1]):
        return False
    ox, oy = origen
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
        if abs(d_cm) < 0.05:
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
    if _texto_en_pulgadas(texto):
        return True, f"Ø en IN (debe mm): {texto!r}"
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
        tam_cm = tam_hoja / _escala(vista)
        esperado = _parse_mm(texto_cota_limpio(tam_cm, hoja))
        if esperado is None:
            return False, "sin ref Ø (skip)"
        if abs(medido - esperado) <= max(_TOL_MM, 0.01):
            return False, "ok"
        return True, f"Ø {medido} desfazado != {esperado:.6f}"
    except Exception as exc:
        return True, f"Ø check error: {exc}"


def _es_hoja_thk(up: str) -> bool:
    if "_DESPLIEGUE_" not in up:
        return False
    return (
        "_LADO" in up
        or "_THK" in up
        or "DESPLIEGUE_THK" in up
        or "DESPLIEGUE_LADO" in up
    )


def _hoja_thk_falla(hoja, vista) -> tuple[bool, str]:
    """
    THK flat: silueta = franja de canto desdoblada; valor ≈ Thickness.
    Falla si perfil L/U (doblado), cara plana (desfazada) o cota ausente/rara.
    """
    from THK import (
        _obtener_curvas_validas,
        _parece_silueta_franja_canto,
        _es_perfil_u_o_l,
        _espesor_chapa_desde_vista,
    )

    datos = _obtener_curvas_validas(vista)
    if not datos:
        return True, "THK sin curvas (vista vacia)"

    if _es_perfil_u_o_l(datos, None):
        return True, "THK NO DESDOBLADO: perfil L/U doblado"

    if not _parece_silueta_franja_canto(datos):
        return True, "THK DESFAZADO: no es canto flat (parece cara/escuadra)"

    texto = _leer_texto_cota_hoja(hoja)
    if _texto_en_pulgadas(texto):
        return True, f"THK en IN (debe mm): {texto!r}"
    medido = _parse_mm(texto)
    thk_cm = _espesor_chapa_desde_vista(vista)
    if thk_cm and thk_cm > 1e-9:
        esperado = _parse_mm(texto_cota_limpio(thk_cm, hoja))
        if medido is None:
            return True, f"THK sin texto (esperado Thickness {esperado})"
        if esperado is not None and abs(medido - esperado) > _TOL_THK_MM:
            return True, f"THK desfazado {medido} != Thickness {esperado}"
        return False, "ok"

    if medido is None:
        return True, "THK sin texto ni Thickness de chapa"
    return False, "ok"


def _part_doc_de_vista(vista):
    for getter in (
        lambda: vista.ReferencedDocumentDescriptor.ReferencedDocument,
        lambda: vista.ReferencedFile.DocumentDescriptor.ReferencedDocument,
    ):
        try:
            doc = getter()
            if doc is not None:
                return doc
        except Exception:
            continue
    return None


def _pieza_base_desde_hoja(nombre: str) -> str:
    """PART_DESPLIEGUE_LADO / PART_DESPLIEGUE_THK → PART."""
    base = str(nombre).rsplit(":", 1)[0]
    base = re.sub(
        r"_DESPLIEGUE_(?:THK|LADO).*$",
        "",
        base,
        flags=re.IGNORECASE,
    )
    return base


def _borrar_hojas_thk_pieza(plano, pieza_base: str) -> list[str]:
    """Borra DESPLIEGUE_LADO / THK de una pieza. Devuelve nombres borrados."""
    borradas = []
    up_pieza = pieza_base.upper()
    try:
        n = int(plano.Sheets.Count)
    except Exception:
        return borradas
    for i in range(n, 0, -1):
        try:
            hoja = plano.Sheets.Item(i)
            nom = str(hoja.Name)
            up = nom.upper()
        except Exception:
            continue
        if up_pieza not in up:
            continue
        if "_DESPLIEGUE_" not in up:
            continue
        if "_LADO" not in up and "_THK" not in up:
            continue
        # No tocar FRENTE / XCENTRO / YCENTRO / HOLE
        if any(t in up for t in ("XCENTRO", "YCENTRO", "DIAMETRO", "FRENTE")):
            continue
        try:
            hoja.Delete()
            borradas.append(nom.rsplit(":", 1)[0])
        except Exception as exc:
            print(f"  AVISO borrar THK {nom}: {exc}")
    return borradas


def _recrear_despliegue_lado_flat(inv_app, plano, part_doc, part_name: str) -> str | None:
    """
    Recrea SOLO DESPLIEGUE_LADO con Flat Pattern (SheetMetalFoldedModel=False).
    Devuelve nombre de hoja creada o None.
    """
    import creador_vistas as cv
    from generador_caras_tanque import _encontrar_hoja_machote

    tg = inv_app.TransientGeometry
    to = inv_app.TransientObjects
    base_sheet = _encontrar_hoja_machote(plano)
    if base_sheet is None:
        try:
            base_sheet = plano.Sheets.Item(1)
        except Exception:
            return None

    res_flat = cv.preparar_geometria_flat(part_doc, True, to)
    if not res_flat:
        print(f"  REPAIR THK {part_name}: sin Flat Pattern")
        return None
    _use_fp, caras_fp, cuerpo_fp = res_flat
    res_frente = cv.elegir_frente(caras_fp)
    if not res_frente:
        print(f"  REPAIR THK {part_name}: flat sin cara frente")
        cv._asegurar_modelo_doblado(part_doc, True)
        return None
    frente_face, v_frente, area_frente = res_frente
    res_lado = cv.elegir_lado(
        caras_fp, frente_face, v_frente, area_frente, use_flat_pattern=True
    )
    if not res_lado:
        v_lado = cv.obtener_lado_fallback(tg, v_frente)
    else:
        _lf, v_lado, _al = res_lado
    cx, cy, cz = cv.obtener_centro(part_doc, cuerpo_fp)

    nombre_hoja = cv.construir_nombre_hoja(plano, part_name, "DESPLIEGUE_LADO")
    try:
        new_sheet = cv._crear_hoja_vista(plano, base_sheet, nombre_hoja)
    except Exception as exc:
        print(f"  REPAIR THK {part_name}: no se pudo crear hoja ({exc})")
        cv._asegurar_modelo_doblado(part_doc, True)
        return None

    try:
        px, py, ancho_util, alto_util = cv._area_util_hoja(new_sheet)
        cam = cv.crear_camara(part_doc, tg, to, cx, cy, cz, v_lado, v_frente)
        view = cv._crear_vista_base(
            new_sheet,
            part_doc,
            tg,
            to,
            px,
            py,
            cam,
            use_flat_pattern_view=True,
        )
        if view is not None:
            cv.escalar_vista(plano, view, tg, px, py, ancho_util, alto_util)
    except Exception as exc:
        print(f"  REPAIR THK {part_name}: vista flat fallo ({exc})")
        try:
            new_sheet.Delete()
        except Exception:
            pass
        cv._asegurar_modelo_doblado(part_doc, True)
        return None
    finally:
        try:
            cv._asegurar_modelo_doblado(part_doc, True)
        except Exception:
            pass

    return nombre_hoja


def _reparar_hoja_thk(inv_app, plano, hoja, nombre: str) -> list[str]:
    """
    Borra LADO/THK malos, recrea DESPLIEGUE_LADO flat y reacota THK.
    Devuelve nombres de hojas resultantes a re-exportar.
    """
    os.environ["SOLO_FLAT_CORTE"] = "1"
    vista = None
    try:
        if int(hoja.DrawingViews.Count) >= 1:
            vista = hoja.DrawingViews.Item(1)
    except Exception:
        vista = None
    part_doc = _part_doc_de_vista(vista) if vista is not None else None
    pieza = _pieza_base_desde_hoja(nombre)
    if part_doc is None:
        print(f"  REPAIR THK {pieza}: sin part_doc")
        return []

    # Nombre corto tipo Inventor (sin path)
    try:
        part_name = str(
            getattr(part_doc, "DisplayName", None)
            or getattr(part_doc, "FullFileName", pieza)
        )
        part_name = os.path.splitext(os.path.basename(part_name))[0]
        if ":" in part_name:
            part_name = part_name.split(":", 1)[0]
    except Exception:
        part_name = pieza

    _borrar_hojas_thk_pieza(plano, pieza)
    # También por DisplayName por si el base del sheet difiere
    if part_name.upper() != pieza.upper():
        _borrar_hojas_thk_pieza(plano, part_name)

    nueva = _recrear_despliegue_lado_flat(inv_app, plano, part_doc, part_name)
    if not nueva:
        return []

    from THK import acotar_thk

    try:
        acotar_thk(nombres_permitidos={nueva.upper(), f"{nueva.upper()}"})
    except Exception as exc:
        print(f"  REPAIR THK {part_name}: acotar_thk fallo ({exc})")
        return [nueva]

    # Recoger hojas LADO/THK actuales de la pieza
    out = []
    up_p = part_name.upper()
    up_pieza = pieza.upper()
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            h = plano.Sheets.Item(i)
            up = str(h.Name).upper()
        except Exception:
            continue
        if up_p not in up and up_pieza not in up:
            continue
        if _es_hoja_thk(up):
            out.append(str(h.Name).rsplit(":", 1)[0])
    return out or [nueva]


def auditar_y_reparar_plano(inv_app, plano, max_pasadas: int = _MAX_PASADAS):
    """
    Audita TODAS las cotas flat (XY + HOLE + THK) por COM y repara fallos.

    Returns
    -------
    (ok_global, nombres_reparados)
    """
    _forzar_mm_flat()
    print(f"  Unidades de cota: {get_unidad_cota()} (flat GIGA = mm)")
    tg = inv_app.TransientGeometry
    reparadas: list[str] = []

    for pasada in range(1, max_pasadas + 1):
        _forzar_mm_flat()
        print(f"\n[AUDIT COM] pasada {pasada}/{max_pasadas} — XY + HOLE + THK")
        pend_xy = []
        pend_thk = []
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
            if up.startswith("COPIA DE"):
                try:
                    hoja.Delete()
                    print(f"  DEL  basura {nombre.rsplit(':', 1)[0]}")
                except Exception:
                    pass
                continue
            if int(getattr(hoja.DrawingViews, "Count", 0) or 0) < 1:
                continue
            # Solo flat DESPLIEGUE (nunca hojas dobladas residuales)
            if "_DESPLIEGUE_" not in up and "XCENTRO" not in up and "YCENTRO" not in up:
                # XCENTRO/YCENTRO siempre llevan DESPLIEGUE en nombre de prod;
                # si no, igual auditar si el token está.
                if "DIAMETRO_H" not in up:
                    continue
            _activar_hoja(hoja)
            vista = hoja.DrawingViews.Item(1)

            if "XCENTRO" in up or "YCENTRO" in up or "XMIN" in up or "YMIN" in up:
                eje = "X" if ("XCENTRO" in up or "XMIN" in up) else "Y"
                falla, motivo = _hoja_xy_falla(hoja, vista, eje)
                if falla:
                    print(f"  FAIL XY  {nombre.rsplit(':', 1)[0]}: {motivo}")
                    pend_xy.append((hoja, vista, eje, nombre))
                else:
                    print(f"  OK   XY  {nombre.rsplit(':', 1)[0]}")
            elif "DIAMETRO_H" in up:
                falla, motivo = _hoja_hole_falla(hoja, vista)
                if falla:
                    print(f"  WARN Ø   {nombre.rsplit(':', 1)[0]}: {motivo}")
                else:
                    print(f"  OK   Ø   {nombre.rsplit(':', 1)[0]}")
            elif _es_hoja_thk(up):
                falla, motivo = _hoja_thk_falla(hoja, vista)
                if falla:
                    print(f"  FAIL THK {nombre.rsplit(':', 1)[0]}: {motivo}")
                    pend_thk.append((hoja, nombre))
                else:
                    print(f"  OK   THK {nombre.rsplit(':', 1)[0]}")

        if not pend_xy and not pend_thk:
            print("[AUDIT COM] todas las cotas OK")
            break

        for hoja, vista, eje, nombre in pend_xy:
            try:
                hoja.Activate()
            except Exception:
                pass
            ok = _reparar_hoja_xy(hoja, vista, inv_app, tg, eje)
            base = nombre.rsplit(":", 1)[0]
            if ok:
                print(f"  REPARADA XY  {base}")
                if base not in reparadas:
                    reparadas.append(base)
            else:
                print(f"  NO SE PUDO REPARAR XY  {base}")

        # Agrupar THK por pieza (una sola recreacion por pieza)
        piezas_thk_hechas = set()
        for hoja, nombre in pend_thk:
            pieza = _pieza_base_desde_hoja(nombre).upper()
            if pieza in piezas_thk_hechas:
                continue
            piezas_thk_hechas.add(pieza)
            try:
                hoja.Activate()
            except Exception:
                pass
            nuevas = _reparar_hoja_thk(inv_app, plano, hoja, nombre)
            if nuevas:
                print(f"  REPARADA THK {pieza} → {nuevas}")
                for n in nuevas:
                    if n not in reparadas:
                        reparadas.append(n)
            else:
                print(f"  NO SE PUDO REPARAR THK {pieza}")

    # Verificacion final completa
    fallos_finales = 0
    try:
        n_hojas = int(plano.Sheets.Count)
    except Exception:
        n_hojas = 0
    for i in range(1, n_hojas + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper()
            if int(hoja.DrawingViews.Count) < 1:
                continue
            vista = hoja.DrawingViews.Item(1)
            falla = False
            motivo = ""
            if "XCENTRO" in up or "YCENTRO" in up or "XMIN" in up or "YMIN" in up:
                eje = "X" if ("XCENTRO" in up or "XMIN" in up) else "Y"
                falla, motivo = _hoja_xy_falla(hoja, vista, eje)
            elif _es_hoja_thk(up):
                falla, motivo = _hoja_thk_falla(hoja, vista)
            elif "DIAMETRO_H" in up:
                falla, motivo = _hoja_hole_falla(hoja, vista)
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
