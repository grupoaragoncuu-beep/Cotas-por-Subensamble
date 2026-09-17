# configurable.md — Cursor ↔ Autodesk Inventor (COM)

**Propósito:** que **otra conversación de Cursor, en otro PC**, pueda conectar Cursor con Inventor **igual que aquí**: crear/editar `.ipt`, `.iam` y drawings, y ejecutar operaciones por API en segundo plano.

**Cómo usarlo en el otro equipo:**
1. Copia este archivo a la raíz del workspace (o adjúntalo con `@configurable.md`).
2. En el primer mensaje: *«Lee @configurable.md. Conecta Cursor con Inventor por COM y verifica con el smoke test.»*
3. No operes Inventor a mano. El agente escribe un `.py` y lo ejecuta; Inventor responde por API.

**Referencia de implementación:** proyecto Enclosure (`Automation/scripts/*.py`, `Drawings/fase 2/inventor_bridge.py`).

---

## 0. Qué es este stack (no negociable)

```
Cursor (agente)
    → escribe / ejecuta Python en la misma máquina
        → pywin32 COM  (Inventor.Application)
            → Autodesk Inventor (sesión Windows)
                → .ipt  .iam  .dwg/.idw
```

- Cursor **no** tiene un plugin nativo de Inventor. El puente es **Python + COM**.
- Inventor y Cursor deben correr **en el mismo Windows**. No hay puente Linux/Mac.
- El agente **no hace clic** en la UI. Escribe un script, lo lanza, lee `stdout`/`stderr`, itera.

---

## 1. Requisitos en el PC destino (una vez)

| Requisito | Detalle |
|----------|---------|
| OS | Windows 10/11 |
| Inventor | Instalado y **licenciado**. Versión 2021+ (validado 2024/2025). |
| Python | **64-bit**, 3.11+ (misma bitness que Inventor; Inventor es 64-bit) |
| pywin32 | `py -m pip install pywin32` |
| Cursor | Workspace abierto en la carpeta del proyecto CAD |

Verificar bitness (debe decir 64 bit):

```bat
py -c "import struct; print(struct.calcsize('P')*8, 'bit')"
```

Si sale `32 bit`, COM fallará o lanzará otra instancia. Instala Python 64-bit y usa `py -3.11` (o el launcher 64-bit).

---

## 2. Modo de conexión (elige uno)

| Modo | Cuándo | Código | Cuidado |
|------|--------|--------|---------|
| **A — Attach** (preferido) | Inventor **ya está abierto** con el documento de trabajo | `GetActiveObject` | No lanza Inventor. Falla si no hay sesión. |
| **B — Dispatch** | Inventor cerrado, batch, o smoke | `Dispatch("Inventor.Application")` | Puede **abrir una segunda instancia** si ya hay una. Evitar si el usuario está modelando. |

**Regla de este equipo:** Inventor abierto → modo A. `Dispatch` solo si A falla **y** el usuario acepta una instancia nueva. Nunca `Quit()` en modo interactivo.

Patrón canónico:

```python
import pythoncom
import win32com.client
import win32com.client.dynamic as dyn

def connect_inventor(*, attach_only: bool = True):
    try:
        raw = win32com.client.GetActiveObject("Inventor.Application")
    except Exception:
        if attach_only:
            raise RuntimeError(
                "Inventor no está abierto. Ábrelo, carga el .iam/.ipt/.dwg y reintenta."
            )
        raw = win32com.client.Dispatch("Inventor.Application")
        raw.Visible = True
    inv = dyn.Dispatch(raw._oleobj_)
    inv.Visible = True
    inv.SilentOperation = True
    return inv
```

Después de conectar, **siempre** deja Inventor visible. Scripts batch ocultos (`Visible = False` + `Quit`) son otra familia; no mezclarlos con una sesión de diseño.

---

## 3. El truco COM que no se puede omitir

Los objetos que devuelve Inventor hay que **re-envolver**. Sin esto, métodos fallan con `getattr` / `TypeError` / COM opacos.

```python
def com(obj):
    try:
        return win32com.client.Dispatch(obj._oleobj_)
    except AttributeError:
        return obj
```

Usar `com(...)` en **todo** lo que sale de colecciones:

```python
doc  = com(inv.ActiveDocument)
cdef = doc.ComponentDefinition
occ  = com(cdef.Occurrences.Item(1))
feat = com(cdef.Features.Item(1))
```

Colecciones Inventor son **1-based** (`Item(1)` … `Count`). Nunca `Item(0)`.

---

## 4. Unidades (el bug #1)

| Superficie | Unidad |
|-----------|--------|
| `Parameter.Expression` | Texto de usuario: `"4 in"`, `"10 mm"`, `"3 ul"` |
| `Parameter.Value` (API) | **cm** (unidad de base Inventor), siempre |
| `TransientGeometry` / matrices / puntos 3D | **cm** |
| Bocetos 2D (`Point2d`) | **cm** |

Constante de este repo:

```python
CM = 2.54  # 1 in = 2.54 cm

def inch_of(param) -> float:
    return float(param.Value) / CM

def cm_from_in(inches: float) -> float:
    return float(inches) * CM
```

Nunca pases `4` (pulgadas) a `CreatePoint` / `SetTranslation` / `SetDistanceExtent` sin convertir. `SetDistanceExtent` también espera **cm** si le das un `double`; si le das un Parameter object, Inventor usa su Value (cm).

Parámetros de usuario:

```python
def ensure_user(user_params, name: str, expr: str, unit: str = "in"):
    try:
        p = user_params.Item(name)
        p.Expression = expr
    except Exception:
        seed = "1 ul" if unit == "ul" else "1 in"
        p = user_params.AddByExpression(name, seed, unit)
        p.Expression = expr
    return p
```

Unidades típicas: `"in"`, `"mm"`, `"ul"` (unitless), `"Boolean"`.

---

## 5. Tipos de documento

```python
K_PART     = 12290  # .ipt
K_ASSEMBLY = 12291  # .iam
K_DRAWING  = 12292  # .dwg / .idw
```

Abrir o reutilizar (nunca abrir duplicado):

```python
def find_or_open(inv, path: str, visible: bool = True):
    pl = path.lower()
    for i in range(1, inv.Documents.Count + 1):
        d = com(inv.Documents.Item(i))
        fn = str(getattr(d, "FullFileName", "") or "")
        if fn.lower() == pl:
            return d
    return com(inv.Documents.Open(path, visible))
```

Crear IPT / IAM nuevos:

```python
def create_part(inv, path: str):
    doc = com(inv.Documents.Add(K_PART, "", True))
    doc.SaveAs(path, False)
    return doc

def create_assembly(inv, path: str):
    doc = com(inv.Documents.Add(K_ASSEMBLY, "", True))
    doc.SaveAs(path, False)
    return doc
```

Drawing con plantilla del usuario (no documento vacío ciego):

```python
tmpl = inv.FileManager.GetTemplateFile(K_DRAWING)  # Standard.dwg / .idw
draw = com(inv.Documents.Add(K_DRAWING, tmpl, True))
```

Planos: este equipo usa sobre todo **`.dwg` Inventor** (no AutoCAD puro). El documento activo debe ser `DocumentType == 12292`.

Enums de features que se reusan:

```python
K_NEW_BODY = 20485
K_JOIN     = 20481
K_CUT      = 20482
K_POS      = 20993  # PositiveExtentDirection
K_NEG      = 20994
K_SYM      = 20995
```

WorkPlanes de pieza (origen): `Item(1)=YZ`, `Item(2)=XZ`, `Item(3)=XY`.

---

## 6. Operaciones típicas (lo que Cursor hace aquí)

### 6.1 Inspeccionar sesión

```python
inv = connect_inventor(attach_only=True)
print("Inventor", inv.SoftwareVersion.DisplayVersion)
print("Docs", inv.Documents.Count)
ad = com(inv.ActiveDocument)
print("Active", ad.DisplayName, ad.FullFileName, "type", int(ad.DocumentType))
```

### 6.2 Crear geometría en un IPT

1. `Documents.Add` o abrir existente.
2. User parameters (`ensure_user`).
3. Sketch en un WorkPlane → `Profiles.AddForSolid()` → `ExtrudeFeatures`.
4. Bind model params (`d*`) a nombres (`p.Expression = "Longitud"`).
5. `doc.Update2(True)` → `doc.Save2(True)`.

### 6.3 Ensamblar (IAM)

```python
occ = cdef.Occurrences.Add(path_ipt, inv.TransientGeometry.CreateMatrix())
occ.Name = "Bracket-Top:1"   # nombres estables para iLogic
occ.Grounded = True
```

Posar en pulgadas:

```python
def set_pose_in(inv, occ, x_in, y_in, z_in):
    tg = inv.TransientGeometry
    m = tg.CreateMatrix()
    m.SetTranslation(tg.CreateVector(x_in * CM, y_in * CM, z_in * CM), False)
    occ.Transformation = m
```

### 6.4 Drawings (segundo plano)

- Adjuntar al `.dwg` **abierto** (`--use-active-drawing` en este repo).
- Vistas: `sheet.DrawingViews.AddBaseView(asm, point2d, scale, orient, style)`.
- Cotas, sheets, notes: ver `Drawings/` (fases 1–6). No reinventar el motor de cotas si el workspace ya lo tiene.

iLogic (`.iLogicVb`) vive **dentro** de Inventor. Python lo complementa: Python crea/audita/rebuild; iLogic propaga parámetros en el raíz. No mezclar “enlace nativo `:1`” entre raíz y sub-IAM (dependencias circulares).

---

## 7. Cómo debe trabajar el agente Cursor

1. **Inventor abierto** con el documento correcto **antes** de scripts de diseño.
2. Escribir un `.py` en el workspace (no un one-liner opaco de 200 líneas en la terminal).
3. Ejecutarlo con el Shell:

   ```bat
   py -3 "ruta\al_script.py"
   ```

4. Leer el output. Inventor es la fuente de verdad; si el script dice PASS y el modelo no cambió, el wrap COM o el documento activo están mal.
5. **No** `app.Quit()`. **No** `Close` del documento que el usuario está usando, salvo que se pida.
6. `SilentOperation = True` durante writes; no desactivar dialogs de licencia.
7. Paths **absolutos** Windows (`R:\Proyecto\Pieza.ipt`). Nada de OneDrive hardcodeado de otro usuario.
8. Tras crear archivos CAD, `SaveAs` / `Save2`. El disco es el entregable.
9. Loguear nombres de occs, features y parámetros. No asumir que `Item(1)` es “el de siempre”.
10. Cambios mínimos al modelo. Rebuild destructivo (`Delete` de todos los features) solo con contrato explícito.

Salida UTF-8 (Windows):

```python
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
```

---

## 8. Smoke test (obligatorio en el otro PC)

Guardar como `inventor_smoke.py` en el workspace. Inventor **debe estar abierto**.

```python
# inventor_smoke.py
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import win32com.client
import win32com.client.dynamic as dyn

def com(obj):
    try:
        return win32com.client.Dispatch(obj._oleobj_)
    except AttributeError:
        return obj

raw = win32com.client.GetActiveObject("Inventor.Application")
inv = dyn.Dispatch(raw._oleobj_)
inv.Visible = True
print("OK connect", inv.SoftwareVersion.DisplayVersion)
print("docs", inv.Documents.Count)
ad = com(inv.ActiveDocument)
print("active", ad.DisplayName)
print("file", ad.FullFileName)
print("type", int(ad.DocumentType), "(12290 ipt / 12291 iam / 12292 dwg)")
print("SMOKE PASS")
```

```bat
py -3 inventor_smoke.py
```

**PASS** = imprime versión + documento activo + `SMOKE PASS`.  
**FAIL típico:** `Invalid class string` / `GetActiveObject` → Inventor no está abierto, o Python es 32-bit.

---

## 9. Fallos conocidos

| Síntoma | Causa | Qué hacer |
|---------|--------|-----------|
| `GetActiveObject` falla | Inventor cerrado, o Python 32-bit | Abrir Inventor; verificar 64-bit |
| Se abre **otro** Inventor | `Dispatch` con sesión ya viva | Usar `GetActiveObject`; no mezclar |
| Método COM “no existe” | Falta `com()` / `_oleobj_` | Re-envolver el objeto |
| Geometría 2.54× más grande/chica | Pulgadas vs cm | `* 2.54` en API Value/puntos |
| Script cuelga / UI congelada | Dialog modal, o `Pick` sin foco | `SilentOperation`; para Pick, traer el HWND al frente |
| Documento vacío al Add | Template `""` a veces ok en ipt/iam; drawings no | `GetTemplateFile` para 12292 |
| iLogic no ve la pieza | Nombre de ocurrencia distinto | Renombrar `occ.Name` estable (`Pieza:1`) |
| Parámetro no actualiza hijos | Ecuación duplicada en IPT | Empujar desde el IAM raíz (iLogic) |
| `Update2` sucio / transacción | Feature a medias | `try/except` + no dejar sketches en edit |

Threading: cada hilo extra necesita `pythoncom.CoInitialize()` y su propio `GetActiveObject`. El hilo principal del script no.

---

## 10. Checklist del otro equipo (antes de modelar)

- [ ] Inventor abierto, un documento activo
- [ ] `py -c "import struct; print(struct.calcsize('P')*8)"` → 64
- [ ] `py -c "import win32com.client; print('pywin32 ok')"`
- [ ] `py -3 inventor_smoke.py` → `SMOKE PASS`
- [ ] Workspace Cursor = carpeta del CAD (o carpeta padre)
- [ ] Este archivo `@configurable.md` en el primer mensaje

---

## 11. Prompt para la conversación nueva

```
Lee @configurable.md.

Este PC tiene Autodesk Inventor. Quiero que Cursor se conecte por COM
(Python + pywin32), igual que en el proyecto Enclosure.

1. Verifica Python 64-bit y pywin32.
2. Corre inventor_smoke.py (Inventor ya está abierto).
3. Si SMOKE PASS: [describe la tarea: crear IPT, editar IAM, plano, audit…]
4. Trabaja solo por scripts .py. No pidas que yo haga clic en Inventor
   salvo Pick interactivo explícito.
5. Unidades API = cm. Expresiones de parámetro = "N in" / "N ul".
6. Envuelve cada objeto COM con com() / Dispatch(_oleobj_).
7. No hagas app.Quit().
```

---

## 12. Dónde está el código vivo en Enclosure (si copias este repo)

| Qué | Dónde |
|-----|--------|
| Attach + `com()` drawings | `Drawings/fase 2/inventor_bridge.py` |
| Crear IPT/IAM + params + extrude | `Automation/scripts/create_bot_bracket_l.py` |
| Rebuild piezas en ensamble | `Automation/scripts/rebuild_coil_assembly_windings.py` |
| Keys / params en raíz | `Automation/scripts/ensure_root_sheet101_note_keys.py` |
| Motor de planos | `Drawings/` (ver `Drawings/AGENTS.md`) |
| Metodología paramétrica (no el puente) | `CURSOR_CONTEXTO_METODOLOGIA.md` |

Este archivo cubre **el puente Cursor–Inventor**. La metodología de parámetros (una sola fuente, sin mágicos, iLogic por zona) está en `CURSOR_CONTEXTO_METODOLOGIA.md`.

---

*Versión 1.0 — Enclosure — 2026-09-09*
