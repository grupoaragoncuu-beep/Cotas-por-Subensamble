import os, re, shutil
ST = r"JPG\9919-Board 1\PIEZAS_ACOTADAS\_STAGING_DESPLIEGUE"
SHARE = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\9919-BOARD2_2\Corte\Corte"
)
SHARE2 = (
    r"\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals"
    r"\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA"
    r"\9919-BOARD2_2\DOSSIER FILES\JPGS\Corte\Corte"
)
PIEZA = "GEN1-OP-20-112"
# Solo las cotas de la corrida limpia (14 barrenos interiores)
KEEP = {
    "THK_1.750000",
    "XCENTRO_TYP_35.130000",
    "XCENTRO_TYP_36.000000",
    "XCENTRO_TYP_96.000000",
    "XCENTRO_TYP_96.870000",
    "YCENTRO_TYP_79.203999",
    "YCENTRO_TYP_303.003999",
    "YCENTRO_TYP_603.003999",
    "YCENTRO_TYP_903.003999",
    "YCENTRO_TYP_1203.003999",
    "YCENTRO_TYP_1503.003999",
    "YCENTRO_TYP_1726.803999",
}
# limpiar staging basura de esta pieza
for fn in list(os.listdir(ST)):
    if PIEZA not in fn:
        continue
    if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
        continue
    keep = False
    for k in KEEP:
        if f"__{k}." in fn.replace(".jpg", ".") or fn.endswith(f"__{k}.jpg"):
            keep = True
            break
        # match token_value
        if f"__{k}.jpg" in fn or fn.endswith(f"{k}.jpg"):
            keep = True
            break
    # simpler: basename contains exact medida_valor
    base = fn
    ok = any(k in base for k in KEEP)
    if not ok:
        path = os.path.join(ST, fn)
        print("DEL staging", fn)
        try:
            os.remove(path)
        except OSError as e:
            print(" ", e)

archs = []
for fn in os.listdir(ST):
    if PIEZA not in fn or not fn.lower().endswith((".jpg", ".jpeg", ".png")):
        continue
    if not any(k in fn for k in KEEP):
        continue
    if fn.split("__")[1] != PIEZA:
        continue
    archs.append(os.path.join(ST, fn))
print("keep", len(archs))
for a in sorted(archs):
    print(" ", os.path.basename(a))
for share in (SHARE, SHARE2):
    dst = os.path.join(share, PIEZA)
    os.makedirs(dst, exist_ok=True)
    for fn in list(os.listdir(dst)):
        if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        if any(t in fn.upper() for t in ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")):
            try:
                os.remove(os.path.join(dst, fn))
                print("DEL share", fn)
            except OSError:
                pass
    for src in archs:
        shutil.copy2(src, os.path.join(dst, os.path.basename(src)))
print("SUBE limpio", len(archs))
