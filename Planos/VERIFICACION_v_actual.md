# Verificación de la versión actual — Cotas por Subensamble

Documento operativo para corroborar que el flujo funciona correctamente **en Windows local con Inventor** (piloto NvidiaSpark cerrado por no viabilidad, ver `NvidiaSpark_Win11_Inventor/04_INFORME_JEFATURA.md`).

Aplica al alcance ampliado 2026-08-12:
- Cotas por **subensamble** (caras).
- Cotas de **piezas individuales**.
- **División por segmento** (FRONT/BACK/LEFT/RIGHT/TOP + OTROS).
- Nueva cara **TOP** (tapa) con las mismas reglas que las 4 paredes.

Fuente de verdad de reglas: [`../DEBER_SER_COTAS_CARAS.md`](../DEBER_SER_COTAS_CARAS.md).

---

## 0. Alcance de la verificación

| Frente | Qué se prueba |
|--------|----------------|
| A. Caras (subensamble) | 5 carpetas en `COTAS_POR_REFERENCIA/`: FRONT, BACK, LEFT, RIGHT, TOP |
| B. Piezas individuales | `PIEZAS_ACOTADAS/` dividido por FRONT/BACK/LEFT/RIGHT/TOP (+ OTROS si aplica) |
| C. Universalidad | Correr sobre OTC (`62176-...`) y (idealmente) un Vantran (Segmentos claros) |
| D. Higiene | Machote limpio, sin quedarse en Model, sin hojas `TANQUE_DATUM_*` visibles |
| E. Log | Mapeo `cara <- segmento` para 5 caras + conteo por subcarpeta de piezas |

---

## 1. Pre-requisitos en Windows

- [ ] Autodesk Inventor abierto.
- [ ] `MACHOTE PLANOS.dwg` **activo** como documento actual.
- [ ] Ensamble del tanque completo **abierto** (ej. `62176-1246-A01 LIMPIO Y MARCADO.iam`).
- [ ] Repo sincronizado (`git pull`) para tener los cambios de código más recientes.
- [ ] Sin corridas previas contaminando `Planos/JPG/<ensamble>/`. Si dudas, borra la carpeta del tanque a probar.

---

## 2. Corrida guiada

1. Desde `Planos/`, ejecutar el flujo completo:
   ```bat
   python generador_tanque_completo.py
   ```
   o desde `interfaz_app.py` (botón **Generar Planos**).
2. Esperar a que termine ambos pasos (caras + piezas).
3. Verificar que la consola termina con:
   ```
   PROCESO COMPLETO: caras y piezas acotadas exportadas.
   ```

---

## 3. Checklist post-corrida (marcar todo)

### A. Cotas por subensamble (`COTAS_POR_REFERENCIA/`)
- [ ] Existen las 5 subcarpetas: `FRONT`, `BACK`, `LEFT`, `RIGHT`, `TOP`.
- [ ] Cada subcarpeta contiene JPG numerados y con eje/extremo en el nombre.
- [ ] Al menos 1 imagen inspeccionada por cara: origen (0,0) coherente, cotas H/V con líneas y flechas fuera del tanque, sin abanico.
- [ ] Piezas circulares aparecen solo con `Xcentro` / `Ycentro`.
- [ ] Piezas rectangulares tienen sus 4 referencias.
- [ ] No hay JPG con marco del machote / título del plano recortado.
- [ ] **No** existe carpeta de la base.

### B. Piezas individuales (`PIEZAS_ACOTADAS/`)
- [ ] Existen subcarpetas `FRONT/ BACK/ LEFT/ RIGHT/ TOP/`.
- [ ] Si hay piezas no mapeables, existe `OTROS/` con esas piezas y quedan reportadas en el log.
- [ ] Ninguna pieza está simultáneamente en dos carpetas.
- [ ] La suma de JPG por subcarpeta coincide con el conteo total de piezas del tanque (menos placa madre y elementos excluidos por reglas).
- [ ] Muestreo: 2 piezas al azar por carpeta se ven correctamente acotadas (según reglas del DEBER_SER).

### C. Universalidad (battery de tanques locales en `Tanques/`)

Pre-diagnóstico automático disponible en `Planos/_analisis_tanques_locales_resultado.txt` (se regenera con `python _analisis_tanques_locales.py`).

Comportamiento **esperado** por tanque (base contra la que comparar el log real):

| Tanque | Familia | Ruta esperada de detección | ¿Debe generar TOP? | Notas |
|--------|---------|----------------------------|--------------------|-------|
| `MODELO VANTRAN 251007.stp` | VANTRAN | Raíz nombrada (13 `Segmento*`) + tapa por nombre (`Top_Cover_Assembly1` u otro) | **SÍ** | Caso de oro. Es el que debe validar el flujo entero de forma más limpia. |
| `SUNBELT TANK 3,750KVA.stp` | SUNBELT | Contenedor geométrico de 4 paredes + tapa por nombre; hay `Tapa trasera ATC` que NO es superior — la selección por altura (+cover) debe elegir la correcta | **SÍ** | Universalidad no-Vantran, español. Si TOP sale con "Tapa trasera", es bug: revisar `_detectar_tapa_como_segmento`. |
| `62154-1246-A01.step` | OTC 62154 | Contenedor geométrico + tapa por nombre débil (`AC-CG-02_COVER ON`) o fallback +Y | **SÍ** | OTC con match tenue por nombre. |
| `62176-1246-A01 LIMPIO Y MARCADO.iam` | OTC 62176 | Contenedor geométrico + **fallback geométrico +Y** (la tapa `62176-1247-A01.iam` no tiene keyword) | **SÍ** (crítico) | Caso crítico del fallback geométrico. Si TOP se omite aquí, el fallback no funciona y hay que ajustarlo. |
| `9919-Board 1.STEP` | BOARD (tablero) | Debe fallar en "no se detectaron cuatro paredes reales" | **N/A** | Edge case. Confirmar que el flujo rechaza correctamente y no exporta basura. |

Checklist:
- [ ] Vantran corrido y aprobado como caso de oro.
- [ ] Sunbelt: TOP corresponde a la tapa superior real (no a `Tapa trasera`).
- [ ] OTC 62154: TOP generado.
- [ ] **OTC 62176: TOP generado por fallback geométrico** (verificar log: debe decir `TOP <- ... (por geometría (+cover), ...)`).
- [ ] Board 9919: rechazado con mensaje claro, sin JPG parciales.

### D. Higiene del machote
- [ ] Al terminar, la hoja plantilla del machote es la activa.
- [ ] No hay hojas `TANQUE_DATUM_*` visibles.
- [ ] La vista no quedó en `Model (AutoCAD)` (fondo negro).

### E. Log (`Planos/error_log_caras.txt`)
- [ ] Aparecen 5 líneas `FRONT/BACK/LEFT/RIGHT/TOP <- <segmento/subensamble>`.
- [ ] Aparece conteo por subcarpeta de `PIEZAS_ACOTADAS`.
- [ ] Sin `Traceback` ni `ERROR` no controlado en la última corrida.

---

## 4. Criterios go/no-go

| Nivel | Requisito para pasar |
|-------|----------------------|
| **Mínimo** | A, D, E completos en OTC 62176 |
| **Objetivo** | Todo A, B, D, E en OTC 62176 |
| **Blindado** | Objetivo + C (Vantran + tanque sin tapa) |

Si un requisito **mínimo** falla → registrar en el log/incidencia y abrir corrección antes de otra prueba.

---

## 5. Registro de resultados

| Fecha | Tanque probado | Resultado (min/obj/blin) | Incidencias | JPG revisados |
|-------|-----------------|--------------------------|-------------|----------------|
|       |                  |                          |             |                |

---

## 6. Incidencias conocidas (llenar durante la corrida)

### 2026-08-12 — Crash de Inventor durante `_exportar_caras_jpg` (LEFT, OTC 62176)

- **Síntoma:** después de exportar FRONT/BACK/RIGHT correctamente, a mitad de LEFT
  (foto #93 en adelante) todas las llamadas devuelven
  `pywintypes.com_error (-2147023174, 'El servidor RPC no está disponible.', ...)`.
  Cuando el flujo pasa a TOP, `vista.Scale` lanza el mismo error y aparece el
  diálogo _Autodesk Inventor Error Report — a software problem has caused
  Autodesk Inventor to close unexpectedly_.
- **Causa raíz:** los tiempos de espera (`time.sleep`) tras `Update()`/`Activate()`
  se habían reducido de 0.25–0.5 s a 0.05–0.1 s como parte de la Fase 2 de
  optimización. Con 5 caras + cientos de exportaciones `SaveAsBitmap` seguidas,
  Inventor no alcanzaba a procesar su cola COM y el proceso colapsaba.
- **Corrección aplicada:**
  1. Se revirtieron todos los `time.sleep` a sus valores originales
     (0.25–0.5 s) en `generador_caras_tanque.py` y `generador_vistas.py`.
  2. Dentro de `_exportar_caras_jpg` se añadió un _respiro_ periódico:
     cada 40 fotos exportadas se llama a `_actualizar_inventor(inv_app)`
     y se espera 0.4 s para que Inventor procese eventos pendientes.
  3. Se mantienen las Fases 1 (SilentOperation en piezas) y 3
     (cache `_bbox_occurrence` / `_centroide_occurrence`) que **no** son
     causa del crash y sí aportan ganancia estable.
- **Cómo re-probar:** reabrir Inventor, cargar el ensamble + el machote de
  planos, y volver a ejecutar el flujo. Debe completar las 5 caras sin
  errores `RPC no disponible` y sin diálogo de crash.

### 2026-08-12 — Fallo en PIEZAS_ACOTADAS tras COTAS_POR_REFERENCIA OK (OTC 62176)

- **Síntoma:** `COTAS_POR_REFERENCIA` terminó 508/508 (incluye TOP). Al pasar a
  piezas: `creador_vistas` aborta con
  `(-2147417856, 'Error en la llamada de sistema.')` en
  `AllLeafOccurrences.Item(i)` / `CopyTo`. `PIEZAS_ACOTADAS` quedó en 0 JPG.
- **Causa:** tras 508 `SaveAsBitmap`, el enumerador leaf se corrompe si se
  mantiene abierto durante la creación de hojas; además hojas residuales
  `_FRENTE_*`/`_LADO` no se limpiaban (solo se buscaban `_ANCHO`/`_LARGO`/`_THK`).
- **Corrección:**
  1. Recolectar piezas únicas primero; luego crear hojas (sin enumerator vivo).
  2. Reintento en `CopyTo` ante COM transitorio.
  3. Limpiar residuales `_FRENTE_1/_FRENTE_2/_LADO` + respiro COM antes de piezas.
  4. `ScreenUpdating=True` durante `crear_vistas`; se apaga después.
