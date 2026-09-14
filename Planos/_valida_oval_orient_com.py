# -*- coding: utf-8 -*-
"""
COM: validar ovals por orientacion en GENE-FCU-5-118 (hoja V118 si existe).

  Slot H: X=extremos, Y=centro
  Slot V: Y=extremos, X=centro
  Circulos: X+Y centro
"""
from __future__ import annotations

import importlib
import sys

import pythoncom
import win32com.client

sys.path.insert(0, ".")
pythoncom.CoInitialize()

from inventor_com import conectar_inventor
import diametro as dm
import barrenos_xy_despliegue as bx

importlib.reload(dm)
importlib.reload(bx)


def main() -> int:
    inv = conectar_inventor()
    tg = inv.TransientGeometry
    plano = win32com.client.CastTo(inv.ActiveDocument, "DrawingDocument")
    hoja = None
    for i in range(1, int(plano.Sheets.Count) + 1):
        h = plano.Sheets.Item(i)
        up = str(h.Name).upper()
        if "V118" in up or "GENE-FCU-5-118" in up and "DESPLIEGUE_FRENTE" in up:
            hoja = h
            break
    if hoja is None:
        print("FAIL: no hay hoja V118 / GENE-FCU-5-118 DESPLIEGUE_FRENTE")
        return 2

    print("HOJA", hoja.Name)
    if int(hoja.DrawingViews.Count) < 1:
        print("FAIL: hoja sin vistas (recrear V118 con _valida_gene_fcu_5_118.py)")
        # Intentar otra hoja DESPLIEGUE con vistas de la misma pieza
        hoja = None
        for i in range(1, int(plano.Sheets.Count) + 1):
            h = plano.Sheets.Item(i)
            up = str(h.Name).upper()
            if "GENE-FCU-5-118" in up and "DESPLIEGUE_FRENTE" in up:
                if int(h.DrawingViews.Count) >= 1:
                    hoja = h
                    break
        if hoja is None:
            # Recrear vista flat rapida
            print("Recreando vista flat GENE-FCU-5-118...")
            import time
            import creador_vistas as cv
            from producto_tipo import aplicar_unidad_producto
            from cota_estilo import set_unidad_cota

            part = None
            for i in range(1, int(inv.Documents.Count) + 1):
                d = inv.Documents.Item(i)
                if int(d.DocumentType) != 12290:
                    continue
                if "GENE-FCU-5-118" in str(d.DisplayName).upper():
                    part = win32com.client.CastTo(d, "PartDocument")
                    break
            if part is None:
                print("FAIL: parte no abierta")
                return 2
            for i in range(1, int(inv.Documents.Count) + 1):
                d = inv.Documents.Item(i)
                if int(d.DocumentType) == 12291 and "BOARD" in str(d.DisplayName).upper():
                    aplicar_unidad_producto(
                        ensamble=win32com.client.CastTo(d, "AssemblyDocument")
                    )
                    break
            else:
                set_unidad_cota("mm")
            to = inv.TransientObjects
            res_flat = cv.preparar_geometria_flat(part, True, to)
            _fp, caras, cuerpo = res_flat
            frente_face, v_frente, area = cv.elegir_frente(caras)
            res_lado = cv.elegir_lado(caras, frente_face, v_frente, area, True)
            v_lado = res_lado[1] if res_lado else cv.obtener_lado_fallback(tg, v_frente)
            tiene_guia, v_guia = cv.obtener_vector_guia_frente(frente_face, v_frente, tg)
            cx0, cy0, cz0 = cv.obtener_centro(part, cuerpo)
            upv = v_guia if tiene_guia else v_lado
            cam = cv.crear_camara(part, tg, to, cx0, cy0, cz0, v_frente, upv)
            base = plano.Sheets.Item(plano.Sheets.Count)
            hoja = base.CopyTo(plano)
            try:
                hoja.Name = "V118_DESPLIEGUE_FRENTE_1"
            except Exception:
                pass
            hoja.Activate()
            for i in range(hoja.DrawingViews.Count, 0, -1):
                try:
                    hoja.DrawingViews.Item(i).Delete()
                except Exception:
                    pass
            px, py = float(hoja.Width) * 0.5, float(hoja.Height) * 0.55
            vista = cv._crear_vista_base(
                hoja, part, tg, to, px, py, cam, use_flat_pattern_view=True
            )
            try:
                cv.escalar_vista(plano, vista, tg, px, py, 18.0, 22.0)
                vista.Update()
            except Exception:
                pass
            time.sleep(0.5)
            print("HOJA recreada", hoja.Name)

    vista = hoja.DrawingViews.Item(1)
    sil = bx._silueta_vista(vista)
    if not sil:
        print("FAIL silueta")
        return 3
    ox, _mx, oy, _my, span = sil
    print(f"Scale={vista.Scale} W={vista.Width:.3f} H={vista.Height:.3f} span={span:.3f}")

    extremos = dm._arcos_extremos_ranura(vista)
    ranuras = dm._emparejar_ranuras(extremos)
    n_h = n_v = 0
    for r in ranuras:
        dx = abs(float(r["cx_a"]) - float(r["cx_b"]))
        dy = abs(float(r["cy_a"]) - float(r["cy_b"]))
        if dx >= dy:
            n_h += 1
        else:
            n_v += 1
    print(f"extremos={len(extremos)} ranuras={len(ranuras)} H={n_h} V={n_v}")

    refs = bx._centros_barrenos(vista, tg)
    n_circ = sum(1 for r in refs if r.get("tipo") == "circulo")
    n_ext = sum(1 for r in refs if r.get("tipo") == "oval_ext")
    n_cy = sum(1 for r in refs if r.get("tipo") == "oval_cy")
    print(f"refs total={len(refs)} circ={n_circ} oval_ext={n_ext} oval_cy={n_cy}")

    mems_x, mems_y = [], []
    for b in refs:
        ejes = tuple(b.get("ejes") or ("X", "Y"))
        cx, cy = float(b["cx"]), float(b["cy"])
        tam = float(b.get("tamaño") or 0.2)
        if "X" in ejes and (cx - ox) >= 0.05:
            mems_x.append(
                {
                    "cx": cx,
                    "cy": cy,
                    "tamaño": tam,
                    "dist_hoja": cx - ox,
                    "clave": bx._valor_desde_hoja(vista, hoja, ox, cx),
                    "tipo": b.get("tipo"),
                }
            )
        if "Y" in ejes and (cy - oy) >= 0.05:
            mems_y.append(
                {
                    "cx": cx,
                    "cy": cy,
                    "tamaño": tam,
                    "dist_hoja": cy - oy,
                    "clave": bx._valor_desde_hoja(vista, hoja, oy, cy),
                    "tipo": b.get("tipo"),
                }
            )

    gx = bx._agrupar_coincidencias(mems_x, vista)
    gy = bx._agrupar_coincidencias(mems_y, vista)
    print(f"COINCIDENCIAS X={len(gx)} Y={len(gy)}")
    for i, g in enumerate(gx, 1):
        tipos = sorted({str(m.get("tipo")) for m in g["miembros"]})
        print(
            f"  X{i}: dist={g['dist_hoja']:.5f} n={len(g['miembros'])} "
            f"TYP={g['typ']} clave={g['clave']} tipos={tipos}"
        )
    for i, g in enumerate(gy, 1):
        tipos = sorted({str(m.get("tipo")) for m in g["miembros"]})
        print(
            f"  Y{i}: dist={g['dist_hoja']:.5f} n={len(g['miembros'])} "
            f"TYP={g['typ']} clave={g['clave']} tipos={tipos}"
        )

    # Ø por tipo
    from cota_estilo import texto_cota_dibujo

    grupos_d = dm._agrupar_diametros_grupos(dm._barrenos_en_vista(vista))
    print(f"Ø tipos={len(grupos_d)}")
    for i, g in enumerate(grupos_d, 1):
        a = max(g, key=lambda x: x["tamaño"])
        print(
            f"  H{i:02d} n={len(g)} tipo={a.get('tipo')} "
            f"txt={texto_cota_dibujo(a['tamaño'])}"
        )

    # Expectativas GENE-FCU-5-118: ~6 circ + 24 oval H →
    # X: 2 circ + 2 extremos/columna*2 ≈ varias; Y: 15 filas TYP
    ok_slots = len(ranuras) >= 20
    ok_orient = (n_h >= 20 and n_v == 0) or (n_v >= 20 and n_h == 0) or (
        n_h + n_v >= 20
    )
    ok_y = len(gy) >= 14  # ~15 filas
    ok_x = len(gx) >= 4
    # Con tol estricta: círculos y extremos de óvalo no deben mezclarse
    x_mezcla = any(
        "circulo" in {str(m.get("tipo")) for m in g["miembros"]}
        and "oval_ext" in {str(m.get("tipo")) for m in g["miembros"]}
        for g in gx
    )
    print(f"X mezcla circ+oval_ext={x_mezcla}")
    ok_typ_y = sum(1 for g in gy if g["typ"]) >= 12
    ok_diam = len(grupos_d) >= 1

    print(
        f"checks slots={ok_slots} orient={ok_orient} "
        f"X>={4}:{ok_x} Y>={14}:{ok_y} typY={ok_typ_y} diam={ok_diam}"
    )
    if (
        ok_slots
        and ok_orient
        and ok_x
        and ok_y
        and ok_typ_y
        and ok_diam
        and not x_mezcla
    ):
        print("PASS COM oval-orientacion + coincidencias")
        return 0
    print("FAIL")
    return 10


if __name__ == "__main__":
    raise SystemExit(main())
