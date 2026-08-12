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

### C. Universalidad
- [ ] OTC probado (`62176-1246-A01 LIMPIO Y MARCADO.iam`): OK / detalle: _____
- [ ] Vantran probado (`MODELO VANTRAN 251007` o similar): OK / detalle: _____
- [ ] Al menos un tanque **sin tapa detectable**: log dice “TOP omitido, motivo=...” y el resto se genera bien.

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

_(dejar vacío hasta que aparezcan)_
