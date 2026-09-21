# DEBER SER — Cotas por caras del tanque (COTAS ABIGAIL)

Documento de verdad del flujo **de caras** (ubicación vs 0,0). Si el código o una conversación contradicen esto, gana este archivo en UI/origen/agrupación de caras.

Contrato de **árbol JPGS / dossier / Abigail piezas / BOARD**: [`DEBER_SER_COTAS_FLUJOS.md`](DEBER_SER_COTAS_FLUJOS.md) (manda sobre publish y `cotas_dossier`).

Última actualización: 2026-09-21

---

## 1. Propósito

Generar fotografías de cotas desde `(0,0)` sobre las paredes del tanque, **agrupando piezas iguales** para reducir capturas: un JPG por tipo de pieza con todas sus cotas X e Y a la vez.

Reglas iLogic disponibles:

| Regla | Alcance | Script Python |
|-------|---------|---------------|
| `COTAS_POR_SUBENSAMBLE` | Cotas por caras + **Top Cover** (selección manual) | `generador_caras_tanque.py --seleccion …` |
| `COTAS_ILOGIC_ABIGAIL` | Piezas acotadas + **Top Cover** en mapa de caras | `generador_piezas.py --seleccion …` |
| `COTAS_CARAS_TANQUE` | Flujo **completo** (caras auto + piezas) | `generador_tanque_completo.py` |

---

## 2. Entrada en Inventor

Siempre:

1. Plano **machote** activo (`MACHOTE PLANOS.dwg` o equivalente).
2. **Tanque completo** (ensamble principal) abierto junto al machote.

### Selección manual (obligatoria en ambas reglas de producción)

Tras lanzar `COTAS_POR_SUBENSAMBLE` **o** `COTAS_ILOGIC_ABIGAIL`, el operador selecciona en orden:

1. **TOP COVER** (cara plana de la tapa) — se integra en ambos flujos.
2. Cara expuesta del **SEGMENTO 1**.
3. Cara expuesta del **SEGMENTO 2**.
4. Cara expuesta del **SEGMENTO 3**.
5. Cara expuesta del **SEGMENTO 4**.
6. **BASE** (piso / fondo) — obligatoria en tanque.

iLogic escribe `Planos/seleccion_caras.json` y lanza Python con `--seleccion`.

| Regla | Uso del Top Cover |
|-------|-------------------|
| `COTAS_POR_SUBENSAMBLE` | Vista `TOP/` con cotas H/V de accesorios sobre la tapa |
| `COTAS_ILOGIC_ABIGAIL` | Piezas (L/A/THK/Ø); local puede agrupar por cara; **dossier final** solo por proceso — ver FLUJOS |

### Compat automático (`COTAS_CARAS_TANQUE` / import sin JSON)

Si Python se llama **sin** `--seleccion`, caras usa mapeo geométrico FRONT/BACK/…; piezas solo clasificación.

---

## 3. Salida

- Carpeta raíz local: `Planos/JPG/<nombre_ensamble>/`.
- `COTAS_POR_REFERENCIA/` (regla caras — **sí** por cara):
  - Con selección: `SEGM1/`…`SEGM4/`, `TOP/`, `BASE/`.
  - Automático: `FRONT/`, `BACK/`, `LEFT/`, `RIGHT/`, `TOP/`.
  - Dentro de cada cara: **un JPG por tipo de pieza**, nombre `NNN_QTYK_<tipo>.jpg`.
- `PIEZAS_ACOTADAS/` (regla Abigail):
  - **Local (durante/tras corrida):** puede llevar `<SEGM*|TOP|BASE|OTROS>/<proceso>/<PIEZA>/` o solo `<proceso>/`.
  - **Dossier `DOSSIER FILES\JPGS\`:** solo `<proceso>/<…>/<PIEZA>/` (Corte / Doblado / …). El publicador **quita** `SEGM*`/`TOP`/`BASE`/`OTROS` bajo `PIEZAS_ACOTADAS`, **no copia** `_STAGING_*`, y **diferisce** JPG sueltos sin carpeta de proceso.
- Log: `Planos/error_log_caras.txt` / `error_log.txt` según flujo.
- Al terminar: borrar hojas temporales `TANQUE_DATUM_*` y dejar visible la hoja plantilla del machote (**nunca** quedarse en `Model (AutoCAD)`).

---

## 4. Origen (0,0)

En cada vista:

- Con selección manual: **(0,0)** = esquina **inferior-izquierda** (min X / min Y en hoja) de la **cara seleccionada**.
- Sin selección: **(0,0)** = esquina inferior-izquierda de la placa madre del segmento.
- El origen es un datum interno de cálculo: **no** se dibuja etiqueta, cruz ni indicador visual `(0,0)` en el JPG.

---

## 5. Cotas (cómo deben verse)

- Cotas **horizontales y verticales con líneas, extensiones y flechas**.
- Desde el origen se señalan inicio/fin (o centro si es circular) de cada instancia: `Xmin`/`Xmax`/`Ymin`/`Ymax` o `Xcentro`/`Ycentro`.
- **Agrupación:** piezas con el mismo código OTC `PROYECTO-FAMILIA-Pxx` (p. ej. `P17_403` y `P17_HOLE_404` → `P17`; `P35_Bottom Flange_372` → `P35`) o, si no es OTC, la misma base tras quitar `_\d+$` (`SP-852_1`/`_2`). Un JPG por tipo con QTY.
- **Inicio/fin obligatorio:** toda pieza lleva Xmin/Xmax e Ymin/Ymax desde (0,0). Nunca reemplazar por solo centro (salvo bore adicional en L845/SP-*).
- **Barrenos:** no sustituyen la envolvente (P35 bottom flange = solo inicio/fin). Centro de bore solo en L845/SP-*/*_HOLE como complemento.
- Valores idénticos dentro del grupo se fusionan en una sola línea con texto `TYP`. Las donas azules van en el **punto de extensión** de cada miembro tipico (no en el centro de la pieza).
- **Encuadre por grupo:** reserva dinámica según # cotas; stacks Y a izquierda y derecha si >8; texto vertical en stacks largos.
- Texto y líneas **fuera del tanque**, legibles, en stacks H/V. El JPG se recorta a las curvas del dibujo y a sus cotas.

### Prohibido

- Cotas **alineadas / inclinadas** (abanico).
- **Números flotantes** sin líneas.
- Acotar el **tamaño** de la pieza como dato principal; el dato principal es **dónde está** respecto a (0,0).
- Mezclar en una cara accesorios de **otro** segmento/pared.
- Depender de vistas Front/Right nativas de Inventor para nombrar segmentos cuando hay selección manual.

---

## 6. Qué se acota en cada cara

Se acotan **las piezas del IAM del segmento/cara seleccionado** (todas las capas colgadas de esa rama). Con selección, el catálogo sale del contenedor raíz de la rama; no se mezclan extras de BASE u otras caras.

### No se acota

- La **placa madre** usada como referencia de pared.
- Barrenos/agujeros que no sean ocurrencias independientes.
- Elementos de otra pared física.
- Piezas de otra cara ya mapeada (p. ej. BASE no entra en SEGM).

---

## 7. Segmentos y mapeo

### Con selección manual

| Carpeta | Origen |
|---------|--------|
| `TOP` | Pick 1 — Top Cover |
| `SEGM1`…`SEGM4` | Picks 2–5 — caras expuestas (rama IAM completa del segmento, no solo la placa partida) |
| `BASE` | Pick 6 — base / fondo del tanque |

Cámara: mira la cara (`eye` = normal saliente); `up` = normal del Top Cover (en TOP/BASE, `up` ≈ normal de SEGM1).

### Familia OTC (estructura real — global en piso)

Validado en **11 OPs OTC** bajo `ORDENES DE PRODUCCION` (planos `*-46/47/48.xx` en todas). Detalle: [`PATRONES_OTC_GLOBAL.md`](PATRONES_OTC_GLOBAL.md).

El rol lo dan los **últimos 2 dígitos** de la familia (`1246`→46, `1247`→47, `1248`→48), no el número de proyecto:

```
*-*46-A01.iam      ← TANQUE          (plano *-*46.00)
├── *-*47-A0N      ← TOP COVER       (plano *-*47.00 / 47.01; a veces A01..A07)
└── *-*48-A01      ← casco / shell   (plano *-*48.00; NO acotar como una cara)
    ├── *-*48-A02  ← BASE
    ├── *-*48-A03  ← pared / SEGM
    ├── *-*48-A04  ← pared / SEGM
    ├── *-*48-A05  ← pared / SEGM
    └── *-*48-A06  ← pared / SEGM
         (+ A07+ = subensambles de pared, no caras)
```

El pick de una cara debe resolver al IAM `48-A02`…`A06` (o cualquier `47-A*`), nunca a `48-A01` ni `46-A01`.

El contenedor de un SEGM es el **subensamble de esa cara** (p. ej. `A03`/`A0x` con sus capas), **no** el casco multi-cara `A01` que agrupa las 4 paredes. Subir al shell mezclaba BASE y piezas de laterales (SP-852).

Accesorios estándar recurrentes entre OTC: `SP-852`, `SP-855`, `SP-800`, `SP-752`, `L845.RADVLV`, `SF-NP-*`, etc. (planos SP sueltos en `4. Planos`).

### Automático (compat)

Mapear por centroide contra el marco PQart (`FRONT`/`BACK`/`LEFT`/`RIGHT`/`TOP`). No confiar en el número del nombre (`Segmento 1` ≠ FRONT siempre).

Log obligatorio, por ejemplo:  
`SEGM2 <- 62201-1248-A05` o `FRONT <- Assembly Segmento 2`.

---

## 8. Cámara y orientación

- Selección: normal de la cara elegida + up del Top Cover.
- Automático: criterio PQart (tapa ≈ +Y, cara ≈ +Z, lateral ≈ +X), cámaras a paredes reales **sin** snap a ejes mundo.
- Enderezar la vista en hoja si los bordes salen rotados unos grados.
- **Base del tanque:** carpeta `BASE` con pick manual (no mezclar en SEGM).

---

## 9. Familias reales (contexto de piso)

- Familias: **VANTRAN, OTC, SWE, GIGA, SUNBELT, PTT**, …
- Solo Vantran usa de forma estable `Assembly Segmento N`.
- OTC suele distinguir copias con sufijo `_NNN` → se agrupan por base.
- Otros tanques con el mismo nombre exacto se agrupan por igualdad.

---

## 10. Anti-regresiones (no volver a romper)

1. No dejar la hoja activa en **Model (AutoCAD)**.
2. No **snap** de cámara a ±X/±Z mundo si el ensamble viene inclinado.
3. No mezclar el catálogo de los **4** segmentos en **una** sola cara.
4. El JPG debe **encuadrar vista + cotas**.
5. No mutar `dim.Style` compartido del machote.
6. No reintroducir 1 JPG por referencia individual; las cotas coincidentes sí llevan `TYP`.
7. Con selección, no renombrar carpetas a FRONT/BACK “por costumbre” de Inventor.
8. En el **dossier** de piezas: no publicar `SEGM*`/`BASE`/`TOP` bajo `PIEZAS_ACOTADAS`, ni `_STAGING_*`, ni JPG sueltos en la raíz de PIEZAS (detalle en FLUJOS).
9. Nombres de pieza con decimal (`PIPE FLANGE 0.250`): no usar `splitext` a ciegas (no crear `PIPE FLANGE 0`).

---

## 11. Criterio de “listo” por corrida

| Check | OK |
|-------|----|
| Preparación | Machote + tanque + picks Top+SEGM1..4+**BASE** en subensamble **y** en piezas |
| Caras | 4 segmentos + TOP + BASE |
| Piezas | Local puede llevar cara; dossier JPGS solo por proceso (ver FLUJOS) |
| Cotas caras | Líneas H/V desde (0,0), 1 JPG por tipo |
| Agrupación | Piezas iguales en un solo JPG; nombre `QTY*` |
| TYP | Texto `TYP` + dona pequeña en los otros extremos del mismo valor (no en el centro de la pieza) |
| Machote | Limpio, plantilla visible |

Validar al menos: un **OTC** con sufijos `_NNN` y un tanque con nombres exactos repetidos.

---

## 12. Archivos del sistema

| Archivo | Rol |
|---------|-----|
| `Planos/generador_caras_tanque.py` | Vistas, selección, grupos, JPG |
| `Planos/seleccion_caras.json` | Salida temporal de los 5 picks iLogic |
| `Planos/orientacion_pqart.py` | Marco tapa/cara/lateral (modo auto) |
| `Planos/cota_estilo.py` | Color/negrita texto de cota |
| `Planos/error_log_caras.txt` | Log de la última corrida |
| Regla iLogic `COTAS_POR_SUBENSAMBLE` | Picks + lanza generador |

---

## 13. Flujo lógico (resumen)

```
Machote + tanque completo
  → iLogic: pick Top + SEGM1..4 → seleccion_caras.json
  → por cada SEGM/TOP: cámara por normal + vista HLR
  → origen (0,0) = esquina IL de la cara seleccionada
  → agrupar piezas por clave de tipo
  → 1 JPG por tipo con todas las cotas X+Y
  → limpiar hojas temporales del machote
```
