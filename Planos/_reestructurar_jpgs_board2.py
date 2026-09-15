# -*- coding: utf-8 -*-
"""
Migra JPGS (+ gemelo dossier) al árbol acordado GIGA BOARD.

  Corte/Plasma y Laser/Corte metal/   ← flat metal (ex Corte/Corte no-cobre)
  Corte/Maquinado/Corte Busbar/       ← flat cobre (ex Corte/Corte cobre)
  Corte/Maquinado/Maquinados metal/   ← ex Maquinado no-cobre
  Doblado/Metal/ | Doblado/Busbar/    ← ex Doblado + Corte/Doblado
  Estañado Busbar/                    ← ex Corte/Estañado
  Almacén/ | SIN CLASIFICACION/       ← sin cambio

No toca: Nueva carpeta, _STAGING_*.
"""
from __future__ import annotations

import os
import shutil
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

from piezas_cobre import es_pieza_cobre  # noqa: E402

DOSSIER_FILES = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES"
)
ROOT_JPGS = os.path.join(DOSSIER_FILES, "JPGS")
ROOT_TWIN = os.path.join(DOSSIER_FILES, "9919-BOARD2_2")

DRY_RUN = "--dry-run" in sys.argv


def _mkdir(p: str) -> None:
    if not DRY_RUN:
        os.makedirs(p, exist_ok=True)


def _move(src: str, dst: str) -> None:
    if os.path.abspath(src) == os.path.abspath(dst):
        return
    _mkdir(os.path.dirname(dst))
    if DRY_RUN:
        print(f"  DRY  {src}\n    -> {dst}")
        return
    if os.path.exists(dst):
        if os.path.isdir(dst) and os.path.isdir(src):
            # Fusionar carpeta pieza: mover archivos hijos.
            for fn in os.listdir(src):
                s2 = os.path.join(src, fn)
                d2 = os.path.join(dst, fn)
                if os.path.exists(d2):
                    if os.path.isfile(d2):
                        os.remove(d2)
                    else:
                        shutil.rmtree(d2, ignore_errors=True)
                shutil.move(s2, d2)
            try:
                os.rmdir(src)
            except OSError:
                shutil.rmtree(src, ignore_errors=True)
            return
        if os.path.isfile(dst):
            os.remove(dst)
        else:
            shutil.rmtree(dst, ignore_errors=True)
    shutil.move(src, dst)


def _rm_empty(path: str) -> None:
    if DRY_RUN or not os.path.isdir(path):
        return
    try:
        for dirpath, dirnames, filenames in os.walk(path, topdown=False):
            if not dirnames and not filenames:
                try:
                    os.rmdir(dirpath)
                except OSError:
                    pass
        if os.path.isdir(path) and not os.listdir(path):
            os.rmdir(path)
    except OSError:
        pass


def _preparar_arbol(raiz: str) -> None:
    for rel in (
        r"Corte\Plasma y Laser\Corte metal",
        r"Corte\Plasma y Laser\Corte Busbar",
        r"Corte\Maquinado\Maquinados metal",
        r"Corte\Maquinado\Corte Busbar",
        r"Doblado\Metal",
        r"Doblado\Busbar",
        "Estañado Busbar",
        "Almacén",
        "SIN CLASIFICACION",
    ):
        _mkdir(os.path.join(raiz, rel))


def _dest_flat(raiz: str, pieza: str) -> str:
    if es_pieza_cobre(pieza):
        return os.path.join(raiz, "Corte", "Maquinado", "Corte Busbar", pieza)
    return os.path.join(raiz, "Corte", "Plasma y Laser", "Corte metal", pieza)


def _dest_doblado(raiz: str, pieza: str) -> str:
    if es_pieza_cobre(pieza):
        return os.path.join(raiz, "Doblado", "Busbar", pieza)
    return os.path.join(raiz, "Doblado", "Metal", pieza)


def _dest_maquinado(raiz: str, pieza: str) -> str:
    if es_pieza_cobre(pieza):
        return os.path.join(raiz, "Corte", "Maquinado", "Corte Busbar", pieza)
    return os.path.join(
        raiz, "Corte", "Maquinado", "Maquinados metal", pieza
    )


def _estanado_dir(raiz: str) -> str:
    buen = os.path.join(raiz, "Estañado Busbar")
    mal = os.path.join(raiz, "Estanado Busbar")
    if os.path.isdir(buen):
        return buen
    if os.path.isdir(mal):
        return mal
    return buen


def _mover_piezas_en(src_dir: str, dest_fn) -> tuple[int, int]:
    """Mueve subcarpetas-pieza. dest_fn(pieza)->ruta destino."""
    if not os.path.isdir(src_dir):
        return 0, 0
    n_dir, n_jpg = 0, 0
    for nombre in sorted(os.listdir(src_dir)):
        src = os.path.join(src_dir, nombre)
        if not os.path.isdir(src):
            continue
        if nombre.startswith("_") or nombre.upper().startswith("COPIA DE"):
            continue
        # Evitar mover ya-anidadas (Metal/Busbar bajo Doblado, etc.)
        if nombre.casefold() in (
            "metal",
            "busbar",
            "corte",
            "doblado",
            "estañado",
            "estanado",
            "plasma y laser",
            "maquinado",
            "corte metal",
            "corte busbar",
            "maquinados metal",
        ):
            continue
        dst = dest_fn(nombre)
        jpgs = [
            f
            for f in os.listdir(src)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
        n_dir += 1
        n_jpg += len(jpgs)
        _move(src, dst)
    return n_dir, n_jpg


def _mover_estanado(src_dir: str, raiz: str) -> int:
    if not os.path.isdir(src_dir):
        return 0
    dst_dir = _estanado_dir(raiz)
    _mkdir(dst_dir)
    n = 0
    for fn in sorted(os.listdir(src_dir)):
        src = os.path.join(src_dir, fn)
        if not os.path.isfile(src):
            continue
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        dst = os.path.join(dst_dir, fn)
        _move(src, dst)
        n += 1
    return n


def migrar_raiz(raiz: str, etiqueta: str) -> None:
    print(f"\n=== {etiqueta} ===")
    print(raiz)
    if not os.path.isdir(raiz):
        print("  (no existe, skip)")
        return
    _preparar_arbol(raiz)

    # 1) Flat legacy Corte/Corte
    d, j = _mover_piezas_en(
        os.path.join(raiz, "Corte", "Corte"),
        lambda p: _dest_flat(raiz, p),
    )
    print(f"  Corte/Corte → flat: {d} piezas / {j} jpg")

    # 2) Corte/Doblado → Doblado/Metal|Busbar
    d, j = _mover_piezas_en(
        os.path.join(raiz, "Corte", "Doblado"),
        lambda p: _dest_doblado(raiz, p),
    )
    print(f"  Corte/Doblado → Doblado: {d} piezas / {j} jpg")

    # 3) Estañado
    n = _mover_estanado(os.path.join(raiz, "Corte", "Estañado"), raiz)
    n2 = _mover_estanado(os.path.join(raiz, "Corte", "Estanado"), raiz)
    print(f"  Corte/Estañado → Estañado Busbar: {n + n2} jpg")

    # 4) Doblado plano (raíz) → Doblado/Metal|Busbar
    d, j = _mover_piezas_en(
        os.path.join(raiz, "Doblado"),
        lambda p: _dest_doblado(raiz, p),
    )
    print(f"  Doblado/ → Doblado/Metal|Busbar: {d} piezas / {j} jpg")

    # 5) Maquinado raíz → Corte/Maquinado/...
    d, j = _mover_piezas_en(
        os.path.join(raiz, "Maquinado"),
        lambda p: _dest_maquinado(raiz, p),
    )
    print(f"  Maquinado/ → Corte/Maquinado: {d} piezas / {j} jpg")

    # 6) Plasma / Plasma Doblado → Corte metal
    for sub in ("Plasma", "Plasma Doblado"):
        d, j = _mover_piezas_en(
            os.path.join(raiz, sub),
            lambda p: os.path.join(
                raiz, "Corte", "Plasma y Laser", "Corte metal", p
            ),
        )
        if d:
            print(f"  {sub}/ → Corte metal: {d} piezas / {j} jpg")

    # Limpieza vacíos legacy
    for rel in (
        r"Corte\Corte",
        r"Corte\Doblado",
        r"Corte\Estañado",
        r"Corte\Estanado",
        "Maquinado",
        "Plasma",
        "Plasma Doblado",
    ):
        _rm_empty(os.path.join(raiz, rel))


def _conteo(raiz: str) -> None:
    if not os.path.isdir(raiz):
        return
    print(f"\n--- conteo {os.path.basename(raiz)} ---")
    targets = [
        r"Corte\Plasma y Laser\Corte metal",
        r"Corte\Plasma y Laser\Corte Busbar",
        r"Corte\Maquinado\Corte Busbar",
        r"Corte\Maquinado\Maquinados metal",
        r"Doblado\Metal",
        r"Doblado\Busbar",
        "Estañado Busbar",
        "Estanado Busbar",
        "Almacén",
        "Almacen",
        "SIN CLASIFICACION",
        r"Corte\Corte",
        r"Corte\Doblado",
        r"Corte\Estañado",
        "Doblado",
    ]
    for rel in targets:
        p = os.path.join(raiz, rel)
        if not os.path.isdir(p):
            continue
        n = 0
        for _dp, _dns, fs in os.walk(p):
            n += sum(1 for f in fs if f.lower().endswith((".jpg", ".jpeg")))
        # Solo reportar Doblado raíz si no es Metal/Busbar count doble
        if rel == "Doblado":
            # contar solo JPG directamente bajo piezas sueltas (no Metal/Busbar)
            n = 0
            for nombre in os.listdir(p):
                sub = os.path.join(p, nombre)
                if not os.path.isdir(sub):
                    continue
                if nombre.casefold() in ("metal", "busbar"):
                    continue
                for _dp, _dns, fs in os.walk(sub):
                    n += sum(
                        1 for f in fs if f.lower().endswith((".jpg", ".jpeg"))
                    )
        print(f"  {n:5d}  {rel}")


def main() -> int:
    print("Reestructuración JPGS Board2", "DRY_RUN" if DRY_RUN else "APPLY")
    # Solo JPGS es canónico.
    migrar_raiz(ROOT_JPGS, "JPGS")
    _conteo(ROOT_JPGS)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
