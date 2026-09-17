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
RE6 = re.compile(r"\.\d{6}(?:_|\.|$)")
TOK = ("XCENTRO", "YCENTRO", "XMIN", "YMIN", "__THK_", "__HOLE")
archs = []
for fn in os.listdir(ST):
    if PIEZA not in fn:
        continue
    if not fn.lower().endswith((".jpg", ".jpeg", ".png")):
        continue
    if not any(t in fn.upper() for t in TOK):
        continue
    if not RE6.search(fn):
        continue
    if fn.split("__")[1] != PIEZA:
        continue
    archs.append(os.path.join(ST, fn))
print("n", len(archs))
for a in sorted(archs):
    print(" ", os.path.basename(a))
ups = [os.path.basename(a).upper() for a in archs]
ok = any("XCENTRO" in u or "YCENTRO" in u for u in ups) and any("__THK_" in u for u in ups)
print("ok", ok)
if not ok:
    raise SystemExit(3)
for share in (SHARE, SHARE2):
    dst = os.path.join(share, PIEZA)
    os.makedirs(dst, exist_ok=True)
    for fn in list(os.listdir(dst)):
        if fn.lower().endswith((".jpg", ".jpeg", ".png")) and any(
            t in fn.upper() for t in TOK
        ):
            try:
                os.remove(os.path.join(dst, fn))
            except OSError:
                pass
    for src in archs:
        shutil.copy2(src, os.path.join(dst, os.path.basename(src)))
print("SUBE", PIEZA, len(archs))
