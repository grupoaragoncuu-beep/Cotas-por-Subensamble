# DEBER SER — Flujos principales COTAS ABIGAIL (Caras + Piezas) y dossier DB

Documento de verdad del **cómo debe comportarse** el sistema de cotas.

Si el código, un chat o un documento viejo contradicen esto, **gana este archivo**
(salvo que se actualice aquí primero).

| | |
|--|--|
| Actualizado | 2026-09-21 |
| Reglas iLogic principales | `COTAS_POR_SUBENSAMBLE` · `COTAS_ILOGIC_ABIGAIL` |
| Complemento caras (detalle UI) | [`DEBER_SER_COTAS_CARAS.md`](DEBER_SER_COTAS_CARAS.md) |
| Contrato DB legacy ANS | [`ALINEACION_COTAS_DOSSIER_ANS.md`](ALINEACION_COTAS_DOSSIER_ANS.md) — **este archivo manda** en árbol JPGS y `seleccionadas` |

---

## 0. Vista mental (dos rieles)

```text
Machote (.dwg/.idw) + ensamble principal (.iam) abiertos
        │
        ▼
  producto_tipo → TANQUE | BOARD | DESCONOCIDO
        │
        ├─────────────── TANQUE (OTC / Vantran / SWE / PTT / …) ───────────────┐
        │                                                                      │
        │   Picks obligatorios: TOP → SEGM1..4 → BASE                          │
        │                                                                      │
        │   ┌─ COTAS_POR_SUBENSAMBLE ──► COTAS_POR_REFERENCIA/                 │
        │   │     (ubicación vs 0,0; 1 JPG por tipo de pieza)                  │
        │   │                                                                  │
        │   └─ COTAS_ILOGIC_ABIGAIL ──► PIEZAS_ACOTADAS/                       │
        │         (dims de pieza L/A/THK/Ø; por cara + clasificación)          │
        │                                                                      │
        └─────────────── BOARD / GIGA tablero (9919-Board…) ───────────────────┐
                                                                               │
            Sin picks de tanque                                                │
                                                                               │
            ┌─ Reglas de CARAS / BOARD ──► generador_board.py                   │
            │     JPG/<job>/BOARD/<kit>/<FRONT|…>/                             │
            │                                                                  │
            └─ COTAS_ILOGIC_ABIGAIL ──► generador_piezas.py (sin picks)         │
                  PIEZAS_ACOTADAS/ por clasificación                            │
                  • Omite Almacén                                              │
                  • DESPLIEGUE si Corte o chapa con huecos                     │
                  • Cobre ABB/GENE/RLG → SIN_COTA + Estañado                    │
                  • Unidades mm                                                │
                                                                               │
        En ambos rieles (si dossier activo):                                   │
            publicar → …\<PRODUCTO>\<CLIENTE>\<JOB>\DOSSIER FILES\JPGS\        │
            INSERT  → public.cotas_dossier (NestingPro :5433)                   │
```

---

## 1. Las dos reglas principales

### 1.1 `COTAS_POR_SUBENSAMBLE` — cotas por caras (ubicación)

| | |
|--|--|
| Archivo | `Planos/ilogic/COTAS_POR_SUBENSAMBLE.iLogicVb` |
| Motor | `Planos/generador_caras_tanque.py --seleccion …` |
| Salida local | `Planos/JPG/<ensamble>/COTAS_POR_REFERENCIA/` |
| Pregunta que responde | *¿Dónde está esta pieza respecto al (0,0) de la cara?* |

**Deber ser:**

1. Machote activo + ensamble principal abierto.
2. Si el producto es **BOARD** → **no** pedir caras de tanque; desvío a `generador_board.py` (kits ViewCube).
3. Si es **TANQUE** → picks en orden:
   1. **TOP COVER**
   2. **SEGM1** … **SEGM4** (cara expuesta de cada segmento)
   3. **BASE** (obligatoria; OTC tiene tejado inclinado — la cámara se cuadra desde el piso)
4. iLogic escribe `Planos/seleccion_caras.json` y lanza Python.
5. Por cada cara: **(0,0)** = esquina inferior-izquierda de la silueta frontal de la cara seleccionada (datum interno; **no** dibujar cruz ni texto `(0,0)`).
6. **Un JPG por tipo de pieza** en esa cara, con **todas** las cotas X e Y desde (0,0) en la misma foto.
7. Cotas H/V con líneas, extensiones y flechas; TYP agrupa valores idénticos.
8. Al terminar: limpiar hojas temporales y dejar visible la plantilla del machote.

**Prueba rápida de una cara:** regla hermana `COTAS_POR_SEG` (`--solo SEGM2`, etc.).

---

### 1.2 `COTAS_ILOGIC_ABIGAIL` — cotas por piezas (fabricación)

| | |
|--|--|
| Archivo | `Planos/ilogic/COTAS_ILOGIC_ABIGAIL.iLogicVb` |
| Motor | `Planos/generador_piezas.py` (+ `--seleccion` solo en tanque) |
| Salida local | `Planos/JPG/<ensamble>/PIEZAS_ACOTADAS/` |
| Pregunta que responde | *¿Cuáles son las dims de fabricación de cada pieza?* |

**Deber ser — TANQUE:**

1. Mismos picks TOP + SEGM1–4 + BASE que caras.
2. Organiza salida **local** en `PIEZAS_ACOTADAS/<SEGM*|TOP|BASE|OTROS>/<clasificación>/<PIEZA>/` (o solo por clasificación tras reorg).
3. En el **dossier** `JPGS\PIEZAS_ACOTADAS\` queda solo `<clasificación>/<…>/<PIEZA>/` (sin carpetas de cara).
4. Dims tipicas: LENGTH / WIDTH / THK / HEIGHT / LEG / OD / HOLE…
5. Regla **Corte vs Doblado** (iProperty `Clasificación`) — ver §3.
6. Tras reorganizar: sync dossier (fail-soft; gates §4.2).

**Deber ser — BOARD / GIGA tablero:**

1. **Sin picks.** Mensaje iLogic: se acotan piezas sin selección de caras.
2. Alcance: **todas las piezas únicas excepto Almacén** (tornillería / comprados / McMaster con iProp Almacén).
3. Unidades: **mm**. Nombres de pieza **completos** (sin truncar a 3 segmentos).
4. Orientación: FRENTE = cara de mayor área; perfil L/canto aparte.
5. Barrenos flat (**DESPLIEGUE**): si iProp Corte **o** chapa con huecos → XCENTRO / YCENTRO / HOLE / THK.
6. Cobre (**ABB / GENE / RLG** únicamente):
   - Por pieza: `…__LENGTH_<v>.jpg` **+** `…__LENGTH_SIN_COTA_<v>.jpg` (solo LENGTH / FRENTE_1).
   - Vista isométrica → carpeta **Estañado Busbar**.
   - En DB: `seleccionadas = si` solo en las **2 primeras XCENTRO** y **2 primeras YCENTRO** distintas desde (0,0); resto `no`.

**Prueba rápida de piezas de una cara:** `COTAS_POR_SEG_PIEZAS`.

---

## 2. Clasificación de producto (`producto_tipo.py`)

Antes de acotar, el sistema decide **TANQUE** vs **BOARD**.

| Prioridad | Señal | Resultado | Familia típica |
|-----------|-------|-----------|----------------|
| 1 | Nombre/ruta con `BOARD`, `####-BOARD`, `GIGA BOARD` | **BOARD** | BOARD |
| 2 | VANTRAN, SUNBELT, OTC, SWE, SEGMENTO, TOP COVER, 1246/47/48, TANK, CASCO, SOLERA | **TANQUE** | OTC / VANTRAN / SWE / … |
| 3 | Árbol 1er nivel muy eléctrico (GENE-/ABB-42-BCK/9919-F|M|P/…) | **BOARD** | BOARD |
| 4 | Solo `\bGIGA\b` sin BOARD/TANK claro | **DESCONOCIDO** | GIGA (no forzar desvío) |
| 5 | Default | **TANQUE** | TANQUE |

### Unidades

| Tipo / familia | Unidad de cota | Nombre de pieza en capturas |
|----------------|----------------|-----------------------------|
| BOARD o familia BOARD/GIGA | **mm** | Completo |
| TANQUE (OTC, Vantran, …) | **in** | Convencional (puede truncar) |

### Desvío BOARD en reglas de *caras*

| Regla | Si es BOARD |
|-------|-------------|
| `COTAS_POR_SUBENSAMBLE` / `COTAS_POR_SEG` / `COTAS_CARAS_TANQUE` / `COTAS_BOARD` | → `generador_board.py` (kits 6 vistas) |
| `COTAS_ILOGIC_ABIGAIL` / `COTAS_POR_SEG_PIEZAS` | → `generador_piezas.py` **sin picks** (no usa `generador_board`) |

---

## 3. Qué se acota y qué se omite

### 3.1 Matriz DESPLIEGUE (flat + barrenos)

Implementación: `creador_vistas._debe_crear_despliegue`.

| Producto | iProp **Corte** | iProp **Doblado** | Otra chapa con huecos |
|----------|-----------------|-------------------|------------------------|
| **TANQUE** | **Nunca** flat / barrenos / cortes internos | Flat **solo si** hay barrenos o cortes pasantes | No (salvo Doblado) |
| **BOARD** | **Sí** DESPLIEGUE | Sí si aplica | **Sí** DESPLIEGUE |

Staging temporal: `_STAGING_DESPLIEGUE/` → luego árbol Corte/…  
Capturas flat: XCENTRO, YCENTRO, HOLE##, THK.

### 3.2 Almacén — no acotar (desde 2026-09-18)

| | |
|--|--|
| Detección | iProperty `Clasificación` = **Almacén** (o ruta bajo `Almacén/`) |
| Abigail | Se **excluye del catálogo** antes de `ejecutar_flujo_desde_app` |
| Dossier | **No se publica** ni se inserta en `cotas_dossier` |
| Motivo | Tornillos, arandelas, McMaster, comprados: ocupan espacio/tiempo y no requieren cota de fab |

### 3.3 Cobre GIGA (solo ABB / GENE / RLG)

| | |
|--|--|
| Detección | `piezas_cobre.es_pieza_cobre` — prefijo al inicio del nombre |
| Dual JPG | Solo en **LENGTH** (cara mayor): con cota + `SIN_COTA` |
| Estañado | Iso sin dims de fab → `Estañado Busbar/` (JPG sueltos, sin subcarpeta por pieza) |
| TYP letras A/B/C | Apagadas en cobre; resto de piezas BOARD pueden llevarlas |
| `seleccionadas` | Ver §5.3 |

### 3.4 Otras omisiones (ambos flujos)

- Hojas vacías / cotas no asociativas → no exportar JPG.
- Piezas madre de placa / paredes ajenas en caras → no acotar como accesorio (detalle en `DEBER_SER_COTAS_CARAS.md`).
- En kits BOARD: stacks HW típicos se omiten del kit ViewCube.

---

## 4. Árbol de carpetas (deber ser)

### 4.1 Local (siempre bajo el repo)

```text
Planos/JPG/<nombre_ensamble>/
├── COTAS_POR_REFERENCIA/          ← regla CARAS
│   ├── SEGM1/ … SEGM4/
│   ├── TOP/
│   └── BASE/
│
├── PIEZAS_ACOTADAS/               ← regla ABIGAIL
│   ├── [SEGM*|TOP|BASE|OTROS/]    ← solo tanque con --seleccion
│   │     └── <árbol clasificación>/…
│   ├── Corte/
│   │   ├── Plasma y Laser/
│   │   │   ├── Corte metal/<PIEZA>/*.jpg
│   │   │   └── Corte Busbar/<PIEZA>/*.jpg
│   │   └── Maquinado/
│   │       ├── Maquinados metal/<PIEZA>/*.jpg
│   │       └── Corte Busbar/<PIEZA>/*.jpg
│   ├── Doblado/
│   │   ├── Metal/<PIEZA>/*.jpg
│   │   └── Busbar/<PIEZA>/*.jpg
│   ├── Estañado Busbar/*.jpg      ← cobre SIN_COTA iso
│   ├── SIN CLASIFICACION/<PIEZA>/
│   ├── _STAGING_DESPLIEGUE/       ← temporal
│   └── _STAGING_ESTANIADO/        ← temporal
│
└── BOARD/<kit>/<FRONT|BACK|TOP|BOTTOM|RIGHT|LEFT>/   ← generador_board
```

**Nota:** `Almacén/` puede existir en corridas antiguas; **corridas nuevas no deben generar ni publicar** Almacén.

### 4.2 Corporativo (dossier)

Raíz canónica:

```text
\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals\
  ARGA METALS CORPORATE SYSTEM\
    <PRODUCTO>\
      <CLIENTE>\
        <JOB>\
          DOSSIER FILES\
            JPGS\          ← evidencias Cotas (única raíz de dossier)
              ├── COTAS_POR_REFERENCIA\SEGM*|TOP|BASE\…   ← caras (sí por cara)
              └── PIEZAS_ACOTADAS\<proceso>\…\            ← solo proceso
```

Ejemplos vivos:

| Job | Cliente | Producto | Ruta tipica JPGS |
|-----|---------|----------|------------------|
| `9919-Board 5` | GIGA | ENCLOSURES NEMA 1 | `\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA\9919-Board5\DOSSIER FILES\JPGS` |
| `9919-Board 11` | GIGA | ENCLOSURES NEMA 1 | `…\GIGA\9919-BOARD-11_1\DOSSIER FILES\JPG\` (o `JPGS\`) |
| `9919-BOARD2_2` | GIGA | ENCLOSURES NEMA 1 | `…\GIGA\9919-BOARD2_2\DOSSIER FILES\JPGS\` |
| Tanque OTC `62223-…` | OTC (u otro) | TANKS | `…\TANKS\<CLIENTE>\<JOB>\DOSSIER FILES\JPGS\` |
| Vantran `261093` | VANTRAN | TANKS | `…\TANKS\VANTRAN\261093\DOSSIER FILES\JPGS\` |

> **Board 5 (post-corrida):** la export overnight pudo haber publicado mal hacia OTC/62223. La ruta **canónica** a corregir es la de `9919-Board5\DOSSIER FILES\JPGS` arriba (cliente GIGA, producto ENCLOSURES NEMA 1). Luego quitar Almacén de disco + DB.

**Publicación relativa (`cotas_dossier_registro.publicar_*`):**

| Flujo | Qué se copia bajo `JPGS\` |
|-------|---------------------------|
| TANQUE caras | `COTAS_POR_REFERENCIA\SEGM*|TOP|BASE\…` (caras **sí** se conservan) |
| TANQUE Abigail | `PIEZAS_ACOTADAS\<proceso>\…` — **sin** `SEGM*`/`TOP`/`BASE`/`OTROS` en el path publicado |
| BOARD Abigail | Árbol de proceso (`Corte\…`, `Doblado\…`, `Estañado Busbar\…`) |

**Gates anti-basura (obligatorios):**

| Condición | Acción |
|-----------|--------|
| Ruta contiene `_STAGING_*` | **No publicar** ni registrar en DB |
| `Clasificación` / path Almacén | **No publicar** ni registrar (corridas nuevas) |
| JPG suelto `PIEZAS_ACOTADAS\file.jpg` o bajo SEGM sin proceso | **Diferir** hasta reorg por proceso |
| `PIEZAS_ACOTADAS\SEGM2\Corte\…` | Publicar como `PIEZAS_ACOTADAS\Corte\…` (sanear cara) |

Helpers de ruta GIGA: `Planos/rutas_arbol_giga.py` (`dest_flat`, `dest_doblado`, `dest_estanado`).

---

## 5. Base de datos `cotas_dossier` (deber ser)

### 5.1 Conexión

| Rol | Host | Puerto | DB | Tabla |
|-----|------|--------|-----|-------|
| Escritura evidencias | `192.168.2.80` | **5433** | `nestingpro_db` | `public.cotas_dossier` |
| Lectura cliente/producto/path | `192.168.2.80` | **5437** | `foldertree` | `public.jobs` |

Credenciales vía env `NESTING_DB_*` / `VSM_DB_*` (defaults en `cotas_dossier_registro.py`).

**Orden de negocio:** Cotas **antes** del nest. Fallo de red/DB **no** aborta el acotado (fail-soft).

### 5.2 Columnas

| Campo | Deber ser | Ejemplo |
|-------|-----------|---------|
| `cliente` | = VSM `jobs.client` | `GIGA`, `OTC` |
| `producto` | = VSM `jobs.product` | `ENCLOSURES NEMA 1`, `TANKS` |
| `job` | Token ensamble ≈ `jobs.job_number` | `9919-Board 5`, `62223-1246-A01` |
| `type` | Siempre **`TYP`** | `TYP` |
| `cantidad_spoteos` | `1` normal; `N` = repetir la misma foto N veces en dossier | `1` |
| `nombre_archivo` | `{JOB}__{ITEM}__{MEDIDA}_{valor:.6f}.jpg` | `9919-Board 5__GENE-…__LENGTH_12.500000.jpg` |
| `ruta` | UNC corporativa preferida; si no hay publish, local | `\\192.168.2.80\…\JPGS\…` |
| `clasificacion` | Ruta de proceso bajo JPGS (árbol §4) | `Corte/Plasma y Laser/Corte metal` |
| `seleccionadas` | `si` / `no` (cobre XY; ver §5.3) | `si` |
| `created_at` | Timestamp alta | |

Upsert lógico: `(job, nombre_archivo)`.

### 5.3 `seleccionadas` (cobre)

Archivo: `Planos/cotas_seleccionadas_cobre.py`.

| Condición | Valor |
|-----------|-------|
| Pieza **no** cobre | siempre `no` |
| Medida ≠ XCENTRO / YCENTRO | `no` |
| Cobre + XCENTRO/YCENTRO | `si` solo en las **2 primeras** posiciones distintas de X y las **2 primeras** de Y (valores numéricos del nombre, orden ascendente desde 0) |
| Resto de centros de la misma pieza | `no` |

Uso ANS/Nesting: filtrar “cotas de calidad” cobre sin perder el resto de evidencias.

### 5.4 Acomodo por familia de producto

| Familia | `cliente` típico | `producto` típico | Unidad | Regla caras | Regla Abigail | Notas DB |
|---------|------------------|-------------------|--------|-------------|---------------|----------|
| **GIGA Board** | `GIGA` | `ENCLOSURES NEMA 1` | mm | → `generador_board` | Piezas sin picks; sin Almacén; DESPLIEGUE amplio; cobre SIN_COTA | `seleccionadas` aplica a cobre |
| **OTC / Vantran / SWE tanque** | según VSM | `TANKS` (u otro) | in | Picks TOP+SEGM+BASE → `COTAS_POR_REFERENCIA` | Picks + `PIEZAS_ACOTADAS` por cara; Corte **sin** flat | Publicar con ancla `PIEZAS_ACOTADAS` / `COTAS_POR_REFERENCIA` |
| **GIGA suelto** (nombre GIGA sin BOARD/TANK) | — | — | — | No forzar BOARD | Tratar con cuidado (DESCONOCIDO) | Resolver job en VSM antes de publish |
| **Cabinet 9919-N** | `GIGA` | `ENCLOSURES NEMA 1` | — | No es el foco Abigail Board | — | Jobs VSM tipo `9919-11CABINET` |

Prioridad al resolver cliente/producto:

1. Env `COTAS_DOSSIER_CLIENTE` / `COTAS_DOSSIER_PRODUCTO`
2. Contexto sesión (`Planos/.runtime/dossier_contexto.json`)
3. VSM `foldertree.jobs`
4. Parse de ruta `…/PRODUCTO/CLIENTE/JOB`
5. Respaldo ANS `erp_jobs` / `diccionario_swo`

### 5.5 Flags de entorno

| Variable | Efecto |
|----------|--------|
| `COTAS_DOSSIER=0` | No registrar en DB |
| `COTAS_DOSSIER_PUBLISH=0` | No copiar al share (sí puede registrar local si aplica) |
| `COTAS_DOSSIER_PICKER=0` | No abrir explorador si falta JOB |
| `COTAS_VSM_JOB_ROOT` | Forzar carpeta JOB |
| `PIEZAS_INCREMENTAL=1` | No re-acotar piezas que ya tienen JPG |
| `PIEZAS_FILTRO=p1,p2` | Limitar a esas piezas |
| `SKIP_ENSAMBLES_IND=1` | No correr kits independientes tras Abigail tanque |

---

## 6. Nomenclatura de capturas

```text
{JOB}__{ITEM}__{MEDIDA}_{valor:.6f}.jpg
```

Separador: `__` (`nomenclatura_capturas.SEP`).

**Limpieza de tokens (`limpiar_token_archivo`):**

- **No** usar `os.path.splitext` ni `rstrip('.')` a ciegas sobre el nombre de pieza: `PIPE FLANGE 0.250` debe conservar el decimal (nunca `PIPE FLANGE 0`).
- Sí quitar `.iam` / `.ipt` embebidos en DisplayName (`MODELO VANTRAN.iam (Estado…)` → token de job limpio).
- Inventor ES: hojas `Copia de …` / `Copy of …` se normalizan al nombre real de la pieza antes de mapear THK/LADO.

Medidas frecuentes:

| Token | Uso |
|-------|-----|
| `LENGTH` / `WIDTH` / `THK` / `HEIGHT` / `LEG` / `OD` / `ID` | Dims generales |
| `LENGTH_SIN_COTA` | Cobre: misma vista LENGTH sin dimensión |
| `XCENTRO` / `YCENTRO` / `XCENTRO_TYP` / `YCENTRO_TYP` | Barrenos en DESPLIEGUE |
| `HOLE01`… | Diámetros |
| `TYP` / `TYP+` | Coincidencias tipadas (spoteos) |

**Cotas redondas / THK (anti-huecos):**

| Caso | Deber ser |
|------|-----------|
| Nipple / boss / tierra / brida (Ø exterior) | `OD` desde anillo de silueta (`solo_interiores=False` en export OD) |
| Solera / pata | `LEG` cuenta como ancho de pata |
| THK canto fino (bbox 2D pequeño) | Exportar si hay cota asociativa o nota THK (no tumbar por bbox ≥ 0.05) |

---

## 7. Checklist operador (deber ser en planta)

### Tanque (OTC / Vantran / …)

1. Abrir machote + IAM completo.
2. Colorimetría / iProperty `Clasificación` al día (Corte / Doblado / …).
3. Correr **`COTAS_POR_SUBENSAMBLE`** → picks TOP→SEGM→BASE → revisar `COTAS_POR_REFERENCIA`.
4. Correr **`COTAS_ILOGIC_ABIGAIL`** → mismos picks → revisar `PIEZAS_ACOTADAS`.
5. Confirmar filas en `cotas_dossier` para el `job` (cliente/producto VSM correctos).

### GIGA Board

1. Preferir IAM en Escritorio / Pack&Go (evitar UNC lento o job equivocado).
2. Confirmar que Inventor tiene el **Board correcto** (el de mayor leaf count si hay varios).
3. Sheet metal real (bends > 0) donde aplique.
4. Colorimetría; **Almacén** se ignorará al acotar.
5. Correr **`COTAS_ILOGIC_ABIGAIL`** (sin picks) → `PIEZAS_ACOTADAS` local → publish JPGS.
6. Verificar cobre: dual LENGTH + Estañado; `seleccionadas` en XY centros.
7. **No** esperar carpetas Almacén nuevas en dossier.

---

## 8. Reglas hermanas (no principales, pero relacionadas)

| Regla | Rol |
|-------|-----|
| `COTAS_POR_SEG` | Una cara (ubicación) |
| `COTAS_POR_SEG_PIEZAS` | Piezas de una cara |
| `COTAS_CARAS_TANQUE` | Caras auto + piezas en una corrida |
| `COTAS_BOARD` | Solo kits ViewCube BOARD |
| `COTAS_ENSAMBLES_INDEPENDIENTES` | Kits independientes (off por defecto en Abigail) |
| `COTAS_BARRENOS_FLAT_CORTE` | Re-acotar solo flat barrenos |

---

## 9. Archivos clave

| Archivo | Responsabilidad |
|---------|-----------------|
| `Planos/ilogic/COTAS_POR_SUBENSAMBLE.iLogicVb` | Entrada caras |
| `Planos/ilogic/COTAS_ILOGIC_ABIGAIL.iLogicVb` | Entrada piezas |
| `Planos/generador_caras_tanque.py` | Motor caras |
| `Planos/generador_piezas.py` | Motor piezas (+ omitir Almacén) |
| `Planos/generador_vistas.py` | Loop de export JPG |
| `Planos/creador_vistas.py` | Vistas, DESPLIEGUE, orientación BOARD |
| `Planos/producto_tipo.py` | TANQUE vs BOARD + unidades |
| `Planos/piezas_cobre.py` | Prefijos cobre + SIN_COTA |
| `Planos/cotas_seleccionadas_cobre.py` | Flag `seleccionadas` |
| `Planos/cotas_dossier_registro.py` | VSM + publish + INSERT |
| `Planos/rutas_arbol_giga.py` | Destinos flat/doblado/estañado GIGA |
| `Planos/generador_tanque_completo.py` | Reorg por clasificación |

---

## 10. Anti-regresiones (no romper)

1. **No** volver a acotar Almacén en Abigail ni a insertarlo en DB.
2. **No** hacer DESPLIEGUE de Corte en TANQUE.
3. **No** tratar todo GENE/ABB como cobre sin prefijo estricto ABB/GENE/RLG.
4. **No** usar árbol legacy `Corte/Corte` como destino nuevo (solo lectura/migración).
5. **No** inventar `cliente`/`producto`: salen de VSM (o override explícito).
6. **No** dejar el machote en hoja `Model (AutoCAD)` al terminar.
7. BASE es pick **obligatorio** en tanque (docs viejos que digan “5 picks” están obsoletos).
8. **No** publicar `_STAGING_*` al share ni en `cotas_dossier`.
9. **No** publicar `SEGM*`/`TOP`/`BASE`/`OTROS` bajo `JPGS\PIEZAS_ACOTADAS` (sí bajo `COTAS_POR_REFERENCIA`).
10. **No** publicar JPG sueltos en la raíz de `PIEZAS_ACOTADAS` (diferir hasta proceso).
11. **No** truncar nombres con decimal vía `splitext` (`PIPE FLANGE 0.250` ≠ `PIPE FLANGE 0`).
12. **No** filtrar OD solo a barrenos interiores cuando la pieza es silueta redonda (boss/nipple/tierra).

---

*Fin del deber ser. Actualizar la fecha del encabezado en cada cambio de contrato.*
