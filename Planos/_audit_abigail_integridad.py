# -*- coding: utf-8 -*-
"""
Auditoría de integridad Abigail (COM ↔ PIEZAS_ACOTADAS).

Compara el ensamble abierto en Inventor contra lo exportado en disco:
  - mapa cara (seleccion_caras.json)
  - mapa clasificación (iProperty)
  - árbol PIEZAS_ACOTADAS/<cara>/<clase…>/<pieza>/*.jpg
  - reglas TANQUE: Corte sin HOLE/XCENTRO (sí XMIN/YMIN/CUT_* si hay hueco);
    Doblado con dims generales

Uso (Inventor con 62223-1246-A01.iam abierto):
  python _audit_abigail_integridad.py
  python _audit_abigail_integridad.py --job 62223-1246-A01
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

JOB_DEFAULT = "62223-1246-A01"
SEL_PATH = os.path.join(ROOT, "seleccion_caras.json")

# Medidas ilegales en Corte TANQUE (barrenos Ø / centro).
# XMIN/YMIN/CUT_* sí se permiten si hay hueco rectangular.
_RE_FLAT_ILEGAL_CORTE = re.compile(
    r"__(?:XCENTRO(?:_TYP)?|YCENTRO(?:_TYP)?|HOLE\d{2})_",
    re.I,
)
_RE_FLAT = _RE_FLAT_ILEGAL_CORTE  # compat nombre histórico
_RE_DIM_GEN = re.compile(
    r"__(?:LENGTH|WIDTH|BROAD|THK|HEIGHT|LEG|OD|ID)_",
    re.I,
)


def _log(msg: str = "") -> None:
    print(msg, flush=True)


def _cargar_sel(path: str) -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def _walk_jpgs(carpeta: str) -> list[tuple[str, str]]:
    """[(ruta_rel_desde_carpeta, nombre_archivo), ...]"""
    out = []
    if not os.path.isdir(carpeta):
        return out
    for dp, _dn, fs in os.walk(carpeta):
        for fn in fs:
            if not fn.lower().endswith(".jpg"):
                continue
            rel = os.path.relpath(os.path.join(dp, fn), carpeta)
            out.append((rel.replace("/", "\\"), fn))
    return out


def _clase_desde_ruta(rel: str) -> str:
    """Infiera clasificación lógica desde ruta relativa bajo una cara."""
    parts = rel.split("\\")
    if not parts:
        return "?"
    p0 = parts[0]
    if p0.casefold() == "almacén" or p0.casefold() == "almacen":
        return "Almacén"
    if p0.casefold() == "sin clasificacion":
        return "SIN CLASIFICACION"
    if p0.casefold() == "estañado busbar" or p0.casefold() == "estanado busbar":
        return "Estañado Busbar"
    if p0.casefold() == "doblado":
        return "Doblado"
    if p0.casefold() == "corte":
        # Corte/Maquinado/... o Corte/Plasma...
        joined = "\\".join(parts[:3]).lower()
        if "maquinado" in joined:
            return "Maquinado"
        if "plasma" in joined:
            return "Plasma"
        return "Corte"
    if p0.casefold() in ("maquinado", "plasma", "plasma doblado"):
        return p0
    return p0


def _pieza_folder_desde_ruta(rel: str) -> str | None:
    parts = rel.split("\\")
    if len(parts) < 2:
        return None
    # .../<pieza>/<archivo.jpg>
    return parts[-2]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--job", default=JOB_DEFAULT)
    ap.add_argument(
        "--seleccion",
        default=SEL_PATH,
        help="JSON de picks Top+SEGM+BASE",
    )
    args = ap.parse_args()

    import pythoncom

    pythoncom.CoInitialize()

    from inventor_com import conectar_inventor
    from generador_caras_tanque import (
        construir_mapa_piezas_desde_seleccion,
        detectar_mapa_piezas_por_clasificacion,
        _obtener_ensamble_principal,
    )
    from generador_tanque_completo import (
        SUBCARPETAS_CARA_SELECCION,
        _caras_para_pieza,
        _clave_pieza,
        _extraer_pieza_de_jpg,
        _match_claves_pieza,
    )
    from nomenclatura_capturas import extraer_item_de_captura_pieza

    job_dir = os.path.join(ROOT, "JPG", args.job)
    piezas_dir = os.path.join(job_dir, "PIEZAS_ACOTADAS")
    ref_dir = os.path.join(job_dir, "COTAS_POR_REFERENCIA")
    ens_dir = os.path.join(job_dir, "ENSAMBLES_INDEPENDIENTES")

    _log("=" * 70)
    _log(f" AUDITORÍA ABIGAIL: {args.job}")
    _log("=" * 70)

    inv = conectar_inventor()
    ensamble = _obtener_ensamble_principal(inv)
    _log(f"Ensamble COM: {ensamble.DisplayName}")

    seleccion = _cargar_sel(args.seleccion)
    _log(f"Selección: {args.seleccion}")
    if str(seleccion.get("ensamble") or "").upper() not in str(
        ensamble.DisplayName
    ).upper():
        _log(
            f"  AVISO: seleccion.ensamble={seleccion.get('ensamble')!r} "
            f"≠ activo={ensamble.DisplayName!r}"
        )

    mapa_cara = construir_mapa_piezas_desde_seleccion(ensamble, seleccion)
    mapa_cls = detectar_mapa_piezas_por_clasificacion(inv, ensamble)

    # --- Persistidos vs vivos ---
    def _load_json(name: str) -> dict:
        path = os.path.join(job_dir, name)
        if not os.path.isfile(path):
            return {}
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)

    disk_cara = _load_json(".piezas_por_cara.json")
    disk_cls = _load_json(".piezas_por_clasificacion.json")

    _log("\n--- Mapas COM (vivos) ---")
    for cara in list(SUBCARPETAS_CARA_SELECCION):
        n = len(mapa_cara.get(cara) or [])
        _log(f"  {cara}: {n} piezas en catálogo")
    for cls, s in sorted(mapa_cls.items()):
        if s:
            _log(f"  class {cls}: {len(s)}")

    hallazgos: list[str] = []

    # Diff maps disk vs live
    for cara in SUBCARPETAS_CARA_SELECCION:
        live = {_clave_pieza(p) for p in (mapa_cara.get(cara) or [])}
        disk = {_clave_pieza(p) for p in (disk_cara.get(cara) or [])}
        if live != disk:
            hallazgos.append(
                f"MAPA_CARA drift {cara}: live={len(live)} disk={len(disk)} "
                f"solo_live={len(live - disk)} solo_disk={len(disk - live)}"
            )

    # --- Index JPG on disk ---
    jpgs = _walk_jpgs(piezas_dir)
    _log(f"\n--- Disco PIEZAS_ACOTADAS: {len(jpgs)} JPG ---")

    # structure: cara -> pieza_key -> {rels, classes, measures}
    por_cara_pieza: dict[str, dict[str, dict]] = defaultdict(
        lambda: defaultdict(lambda: {"rels": [], "clases": set(), "files": []})
    )
    sueltos = []
    estructura_rara = []
    for rel, fn in jpgs:
        parts = rel.split("\\")
        if len(parts) < 2:
            sueltos.append(rel)
            continue
        cara = parts[0].upper()
        if cara not in set(SUBCARPETAS_CARA_SELECCION) | {"OTROS"}:
            estructura_rara.append(rel)
            continue
        try:
            item = extraer_item_de_captura_pieza(fn)
        except Exception:
            item = _extraer_pieza_de_jpg(fn)
        clave = _clave_pieza(item)
        clase = _clase_desde_ruta("\\".join(parts[1:]))
        folder = _pieza_folder_desde_ruta(rel)
        por_cara_pieza[cara][clave]["rels"].append(rel)
        por_cara_pieza[cara][clave]["clases"].add(clase)
        por_cara_pieza[cara][clave]["files"].append(fn)
        por_cara_pieza[cara][clave]["folder"] = folder
        # folder name should match item (allowing family)
        if folder and not _match_claves_pieza(
            _clave_pieza(folder), clave
        ) and "__" not in folder:
            # folder might be exact piece name
            if _clave_pieza(folder) != clave:
                estructura_rara.append(f"folder≠item: {rel} item={item}")

    if sueltos:
        hallazgos.append(f"JPG sueltos en raíz/rara: {len(sueltos)}")
        for s in sueltos[:8]:
            _log(f"  SUELTO: {s}")

    # Conteos por cara
    _log("\n--- JPG por cara (disco) ---")
    for cara in list(SUBCARPETAS_CARA_SELECCION) + ["OTROS"]:
        n_jpg = sum(
            len(v["files"]) for v in por_cara_pieza.get(cara, {}).values()
        )
        n_pz = len(por_cara_pieza.get(cara, {}))
        _log(f"  {cara}: {n_jpg} JPG / {n_pz} piezas")

    # --- 1) Catálogo cara → ¿hay JPG? ---
    _log("\n--- Cobertura catálogo cara → JPG ---")
    for cara in SUBCARPETAS_CARA_SELECCION:
        catalogo = mapa_cara.get(cara) or set()
        presentes = set(por_cara_pieza.get(cara, {}).keys())
        # También aceptar si está en otra cara por multi-copy TOP/BASE
        faltan = []
        for p in sorted(catalogo, key=str):
            ck = _clave_pieza(p)
            if ck in presentes:
                continue
            # ¿aparece en alguna cara?
            hallado = any(
                ck in por_cara_pieza.get(c, {})
                for c in list(SUBCARPETAS_CARA_SELECCION) + ["OTROS"]
            )
            if not hallado:
                faltan.append(p)
        _log(f"  {cara}: catálogo={len(catalogo)} con_JPG_en_cara={len(presentes)} sin_JPG_global={len(faltan)}")
        if faltan:
            # hardware/nuts often skip — still report
            hallazgos.append(
                f"FALTAN JPG ({cara}): {len(faltan)} piezas del catálogo "
                f"sin ninguna captura (ej: {', '.join(faltan[:6])})"
            )
            for f in faltan[:12]:
                _log(f"    - {f}")

    # --- 2) JPG en OTROS: ¿deberían tener cara? ---
    _log("\n--- OTROS: ¿huérfanos o mal ubicados? ---")
    otros = por_cara_pieza.get("OTROS", {})
    mal_otros = []
    for ck, info in sorted(otros.items()):
        # reconstruct a fake filename for _caras_para_pieza
        sample = info["files"][0] if info["files"] else ck
        caras = _caras_para_pieza(sample, mapa_cara)
        caras_ok = [c for c in caras if c in SUBCARPETAS_CARA_SELECCION]
        if caras_ok:
            mal_otros.append((ck, caras_ok, len(info["files"])))
    _log(f"  OTROS piezas: {len(otros)}; con cara esperada: {len(mal_otros)}")
    if mal_otros:
        hallazgos.append(
            f"OTROS mal ubicados: {len(mal_otros)} piezas deberían estar en "
            f"SEGM/TOP/BASE (ej: {mal_otros[0][0]}→{mal_otros[0][1]})"
        )
        for ck, caras_ok, n in mal_otros[:15]:
            _log(f"    {ck}: {n} JPG → esperado {caras_ok}")

    # --- 3) Clasificación iProp vs ruta ---
    _log("\n--- Clasificación iProp vs carpeta ---")
    # Invert mapa: pieza_key -> clase esperada
    clase_esperada: dict[str, str] = {}
    for cls, piezas in mapa_cls.items():
        for p in piezas:
            clase_esperada[_clave_pieza(p)] = cls

    # Mapa ruta→clase esperada de destino Abigail:
    # Almacén→Almacén; Doblado→Doblado; Corte→Doblado (dims plegadas, sin flat);
    # Maquinado→Maquinado path; Plasma→Plasma path; SIN→SIN
    def _clase_ruta_esperada(iprop: str) -> set[str]:
        c = (iprop or "").strip()
        if not c or c.upper() == "SIN CLASIFICACION":
            return {"SIN CLASIFICACION"}
        if c.casefold() == "almacén" or c.casefold() == "almacen":
            return {"Almacén"}
        if c.casefold() == "doblado":
            return {"Doblado"}
        if c.casefold() == "corte":
            # TANQUE: Corte se omite flat; dims generales viven bajo Doblado/
            return {"Doblado", "Corte"}
        if c.casefold() == "maquinado":
            return {"Maquinado", "Corte"}
        if c.casefold() in ("plasma", "plasma doblado"):
            return {"Plasma", "Corte", "Doblado"}
        return {c}

    mal_clase = []
    for cara, piezas in por_cara_pieza.items():
        for ck, info in piezas.items():
            esp = clase_esperada.get(ck)
            if not esp:
                # no en mapa (kits/hardware) — ok si SIN o OTROS
                continue
            ok_set = _clase_ruta_esperada(esp)
            actual = set(info["clases"])
            if not (actual & ok_set):
                mal_clase.append((cara, ck, esp, sorted(actual), len(info["files"])))

    _log(f"  mismatches clase: {len(mal_clase)}")
    if mal_clase:
        hallazgos.append(
            f"Clasificación incorrecta en disco: {len(mal_clase)} "
            f"(ej: {mal_clase[0][1]} iProp={mal_clase[0][2]} ruta={mal_clase[0][3]})"
        )
        for row in mal_clase[:20]:
            _log(f"    {row[0]}/{row[1]}: iProp={row[2]} disco={row[3]} ({row[4]} JPG)")

    # --- 4) Regla TANQUE Corte: sin HOLE/XCENTRO (CUT/XMIN OK) ---
    _log("\n--- Regla TANQUE: Corte iProp sin HOLE/XCENTRO ---")
    corte_keys = {_clave_pieza(p) for p in (mapa_cls.get("Corte") or [])}
    viol_corte = []
    for cara, piezas in por_cara_pieza.items():
        for ck, info in piezas.items():
            if ck not in corte_keys:
                continue
            flats = [f for f in info["files"] if _RE_FLAT_ILEGAL_CORTE.search(f)]
            if flats:
                viol_corte.append((cara, ck, flats[:5], len(flats)))
    _log(f"  piezas Corte con HOLE/XCENTRO ilegal: {len(viol_corte)}")
    if viol_corte:
        hallazgos.append(
            f"TANQUE/Corte con HOLE/XCENTRO: {len(viol_corte)} piezas "
            f"(XMIN/YMIN/CUT_* sí permitidos)"
        )
        for cara, ck, samples, n in viol_corte[:15]:
            _log(f"    {cara}/{ck}: {n} ilegal → {samples[0]}")

    # --- 5) Doblado: al menos dims generales ---
    _log("\n--- Doblado: ¿tiene LENGTH/WIDTH/THK? ---")
    doblado_keys = {_clave_pieza(p) for p in (mapa_cls.get("Doblado") or [])}
    sin_dim = []
    for cara, piezas in por_cara_pieza.items():
        for ck, info in piezas.items():
            if ck not in doblado_keys:
                continue
            if not any(_RE_DIM_GEN.search(f) for f in info["files"]):
                sin_dim.append((cara, ck, info["files"][:3]))
    _log(f"  Doblado sin dims generales: {len(sin_dim)}")
    if sin_dim:
        hallazgos.append(
            f"Doblado sin LENGTH/WIDTH/THK: {len(sin_dim)} "
            f"(ej: {sin_dim[0][1]})"
        )
        for row in sin_dim[:12]:
            _log(f"    {row[0]}/{row[1]}: {row[2]}")

    # --- 6) Piezas en mapa clase sin ningún JPG ---
    _log("\n--- Piezas clasificadas sin JPG en absoluto ---")
    all_jpg_keys = set()
    for piezas in por_cara_pieza.values():
        all_jpg_keys |= set(piezas.keys())
    for cls in ("Almacén", "Corte", "Doblado", "Maquinado"):
        falt = []
        for p in sorted(mapa_cls.get(cls) or [], key=str):
            if _clave_pieza(p) not in all_jpg_keys:
                falt.append(p)
        _log(f"  {cls}: {len(falt)} sin JPG")
        if falt and cls in ("Doblado", "Corte", "Almacén"):
            hallazgos.append(
                f"Clase {cls}: {len(falt)} piezas iProp sin JPG "
                f"(ej: {', '.join(falt[:5])})"
            )
            for f in falt[:8]:
                _log(f"    - {f}")

    # --- 7) Multi-SEGM (misma pieza en 2 SEGM) ---
    _log("\n--- Misma pieza en más de un SEGM (disco) ---")
    segm_owners: dict[str, list[str]] = defaultdict(list)
    for cara in ("SEGM1", "SEGM2", "SEGM3", "SEGM4"):
        for ck in por_cara_pieza.get(cara, {}):
            segm_owners[ck].append(cara)
    multi = {k: v for k, v in segm_owners.items() if len(v) > 1}
    _log(f"  multi-SEGM: {len(multi)}")
    if multi:
        hallazgos.append(f"Piezas en múltiples SEGM: {len(multi)}")
        for ck, caras in list(multi.items())[:12]:
            _log(f"    {ck}: {caras}")

    # --- 8) Refs / Ensambles smoke ---
    n_ref = len(_walk_jpgs(ref_dir))
    n_ens = len(_walk_jpgs(ens_dir))
    _log(f"\n--- Otros flujos ---")
    _log(f"  COTAS_POR_REFERENCIA: {n_ref} JPG")
    _log(f"  ENSAMBLES_INDEPENDIENTES: {n_ens} JPG")
    base_ref = os.path.join(ref_dir, "BASE")
    n_base_ref = len(_walk_jpgs(base_ref)) if os.path.isdir(base_ref) else 0
    n_base_map = len(mapa_cara.get("BASE") or [])
    _log(f"  BASE ref JPG={n_base_ref} (mapa cara BASE={n_base_map}; 0 en ref es OK)")

    # --- 9) piezas_sin_cotas ---
    sin_cotas = os.path.join(job_dir, "piezas_sin_cotas.txt")
    if os.path.isfile(sin_cotas):
        lines = [
            ln.strip()
            for ln in open(sin_cotas, encoding="utf-8", errors="ignore")
            if ln.strip()
        ]
        _log(f"  piezas_sin_cotas.txt: {len(lines)} líneas")
        if lines:
            hallazgos.append(f"piezas_sin_cotas.txt: {len(lines)} entradas")
            for ln in lines[:10]:
                _log(f"    {ln}")

    # --- Resumen ---
    _log("\n" + "=" * 70)
    _log(f" RESUMEN: {len(hallazgos)} hallazgo(s)")
    _log("=" * 70)
    if not hallazgos:
        _log("OK — sin inconsistencias detectadas por las reglas auditadas.")
        return 0
    for i, h in enumerate(hallazgos, 1):
        _log(f"  {i}. {h}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
