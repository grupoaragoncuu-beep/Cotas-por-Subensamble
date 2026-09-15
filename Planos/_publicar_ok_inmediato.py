# -*- coding: utf-8 -*-
"""
Publica YA al dossier Corte\\Corte las piezas con staging completo
que pasen auditoría COM (XY + THK). Una pieza a la vez → sube al momento.
"""
from __future__ import annotations

import os
import re
import shutil
import sys

import pythoncom

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from rutas_arbol_giga import (  # noqa: E402
    ROOT_JPGS_BOARD2,
    carpeta_pieza_flat,
    catalogo_piezas_flat,
    dest_flat,
    roots_flat,
)

# Ruta canónica: DOSSIER FILES\JPGS
SHARE_JPGS_ROOT = ROOT_JPGS_BOARD2
SHARE_DOSSIER = os.path.join(
    SHARE_JPGS_ROOT, "Corte", "Plasma y Laser", "Corte metal"
)
SHARE_JPGS = SHARE_DOSSIER
JOB = "9919-Board 2"
_TOKENS = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")
# Solo archivos con valor a 6 decimales (25.400000); descartar 3 (25.400)
_RE_VAL_6DEC = re.compile(r"\.\d{6}(?:_|\.|$)")
_RE_VAL_3DEC_ONLY = re.compile(r"\.\d{3}(?:_|\.|$)")


def _es_jpg_6dec(fn: str) -> bool:
    """True si el nombre lleva valor con exactamente 6 decimales."""
    base = os.path.basename(fn)
    up = base.upper()
    if not any(t in up for t in _TOKENS):
        return False
    return bool(_RE_VAL_6DEC.search(base))


def _es_jpg_3dec_corto(fn: str) -> bool:
    """True si parece valor a 3 decimales (sin llegar a 6)."""
    base = os.path.basename(fn)
    if _RE_VAL_6DEC.search(base):
        return False
    return bool(_RE_VAL_3DEC_ONLY.search(base))


def _filtrar_solo_6dec(archivos: list[str]) -> list[str]:
    return [a for a in archivos if _es_jpg_6dec(a)]


def _staging() -> str:
    jpg = os.path.join(ROOT, "JPG")
    cands = []
    if os.path.isdir(jpg):
        for d in os.listdir(jpg):
            st = os.path.join(jpg, d, "PIEZAS_ACOTADAS", "_STAGING_DESPLIEGUE")
            if os.path.isdir(st):
                cands.append(st)
    if not cands:
        return ""

    def mt(p):
        try:
            fs = os.listdir(p)
            if not fs:
                return os.path.getmtime(p)
            return max(os.path.getmtime(os.path.join(p, f)) for f in fs)
        except Exception:
            return 0

    return sorted(cands, key=mt, reverse=True)[0]


def _map_staging(st: str) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    if not st or not os.path.isdir(st):
        return out
    for fn in os.listdir(st):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        up = fn.upper()
        if not any(t in up for t in _TOKENS):
            continue
        parts = fn.split("__")
        if len(parts) < 2:
            continue
        out.setdefault(parts[1], []).append(os.path.join(st, fn))
    return out


def _completo(archivos: list[str]) -> bool:
    ups = [os.path.basename(a).upper() for a in archivos]
    return (
        any("XCENTRO" in u or "YCENTRO" in u for u in ups)
        and any("__THK_" in u for u in ups)
    )


def _catalogo(raiz: str) -> list[str]:
    if not os.path.isdir(raiz):
        return []
    return sorted(
        n
        for n in os.listdir(raiz)
        if os.path.isdir(os.path.join(raiz, n))
        and not n.upper().startswith("COPIA DE")
        and not n.startswith("_")
    )


def _publicar(pieza: str, archivos: list[str], roots: list[str]) -> int:
    # Solo 6 decimales; las de 3 se descartan
    archivos = _filtrar_solo_6dec(archivos)
    if not archivos:
        return 0
    n = 0
    for root in roots:
        dst = dest_flat(root, pieza)
        os.makedirs(dst, exist_ok=True)
        for fn in list(os.listdir(dst)):
            if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            if not any(t in fn.upper() for t in _TOKENS):
                continue
            # Limpiar destino (incluye borrar las de 3 decimales)
            try:
                os.remove(os.path.join(dst, fn))
            except OSError:
                pass
        for src in archivos:
            fn = os.path.basename(src)
            try:
                shutil.copy2(src, os.path.join(dst, fn))
                n += 1
            except OSError as e:
                print(f"  AVISO copy {fn}: {e}")
    return n


def _auditar_pieza(plano, pieza: str) -> tuple[bool, str]:
    from auditoria_cotas_flat_com import (
        _activar_hoja,
        _hoja_xy_falla,
        _hoja_thk_falla,
        _es_hoja_thk,
    )

    up_p = pieza.upper()
    vistos = 0
    for i in range(1, int(plano.Sheets.Count) + 1):
        try:
            hoja = plano.Sheets.Item(i)
            up = str(hoja.Name).upper().rsplit(":", 1)[0]
        except Exception:
            continue
        if up_p not in up or "_DESPLIEGUE_" not in up:
            continue
        if up.startswith("COPIA DE"):
            continue
        if int(hoja.DrawingViews.Count) < 1:
            continue
        _activar_hoja(hoja)
        vista = hoja.DrawingViews.Item(1)
        if "XCENTRO" in up or "YCENTRO" in up:
            eje = "X" if "XCENTRO" in up else "Y"
            falla, motivo = _hoja_xy_falla(hoja, vista, eje)
            vistos += 1
            if falla:
                return False, f"{up}: {motivo}"
        elif _es_hoja_thk(up):
            falla, motivo = _hoja_thk_falla(hoja, vista)
            vistos += 1
            if falla:
                return False, f"{up}: {motivo}"
    if vistos == 0:
        return False, "sin hojas XY/THK en plano"
    return True, f"ok ({vistos} hojas)"


def main() -> int:
    print("=" * 60)
    print(" PUBLICAR OK INMEDIATO → flat (Plasma metal / Maquinado Busbar)")
    print("=" * 60)
    os.environ["SOLO_FLAT_CORTE"] = "1"
    os.environ["COTAS_JOB_OVERRIDE"] = JOB
    from cota_estilo import set_unidad_cota

    set_unidad_cota("mm")

    cat = catalogo_piezas_flat(SHARE_JPGS_ROOT)
    st = _staging()
    por = _map_staging(st)
    roots = [SHARE_JPGS_ROOT]
    print(f"catalogo={len(cat)} staging={st}")
    print(f"piezas_staging={len(por)} destino=JPGS flat")

    pythoncom.CoInitialize()
    try:
        from inventor_com import conectar_inventor
        from generador_caras_tanque import (
            _encontrar_hoja_machote,
            _obtener_ensamble_principal,
            _obtener_plano_activo,
        )

        inv = conectar_inventor()
        plano = _obtener_plano_activo(inv)
        ensamble = _obtener_ensamble_principal(inv)
        if plano is None:
            print("ERROR: sin plano")
            return 2
        try:
            h = _encontrar_hoja_machote(plano)
            if h:
                h.Activate()
        except Exception:
            pass
        try:
            from producto_tipo import aplicar_unidad_producto

            if ensamble is not None:
                aplicar_unidad_producto(ensamble=ensamble)
        except Exception:
            pass
        set_unidad_cota("mm")

        # Barrer restos de 3 decimales en shares (dejar solo 6)
        for root in roots:
            for flat in roots_flat(root):
                if not os.path.isdir(flat):
                    continue
                for dirpath, _dirs, files in os.walk(flat):
                    for fn in files:
                        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
                            continue
                        if not any(t in fn.upper() for t in _TOKENS):
                            continue
                        if _es_jpg_3dec_corto(fn):
                            try:
                                os.remove(os.path.join(dirpath, fn))
                                print(f"  DEL 3dec {fn}")
                            except OSError:
                                pass

        ok_n = fail_n = skip_n = pubs = 0
        # Orden: catálogo primero, luego extras de staging
        items = []
        for c in cat:
            items.append(c)
        for k in sorted(por):
            if not any(k.upper() == c.upper() for c in items):
                items.append(k)

        for pieza in items:
            archivos = por.get(pieza) or next(
                (v for k, v in por.items() if k.upper() == pieza.upper()),
                [],
            )
            if not archivos:
                print(f"  SKIP (sin staging): {pieza}")
                skip_n += 1
                continue
            n3 = sum(1 for a in archivos if _es_jpg_3dec_corto(a))
            archivos = _filtrar_solo_6dec(archivos)
            if n3:
                print(f"  (descarta {n3} jpg de 3 decimales)")
            if not archivos:
                print(f"  SKIP (sin jpg a 6 dec): {pieza}")
                skip_n += 1
                continue
            if not _completo(archivos):
                print(f"  SKIP (incompleto 6dec): {pieza} n={len(archivos)}")
                skip_n += 1
                continue
            print(f"\n[{pieza}] audit COM…")
            bien, motivo = _auditar_pieza(plano, pieza)
            if not bien:
                print(f"  FAIL → no sube: {motivo}")
                fail_n += 1
                continue
            n = _publicar(pieza, archivos, roots)
            if n <= 0:
                print(f"  SKIP (nada a 6 dec tras filtro): {pieza}")
                skip_n += 1
                continue
            print(f"  SUBE OK → {n} archivos 6dec ({motivo})")
            ok_n += 1
            pubs += n

        print("\n" + "=" * 60)
        print(f"RESULTADO: OK_subidas={ok_n} FAIL={fail_n} SKIP={skip_n} files={pubs}")
        print(f"destino=flat árbol GIGA")
        return 0 if ok_n else 3
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback

        traceback.print_exc()
        return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
