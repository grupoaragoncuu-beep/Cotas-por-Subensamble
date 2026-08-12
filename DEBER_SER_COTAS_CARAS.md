# DEBER SER — Cotas por caras del tanque (COTAS ABIGAIL)

Documento de verdad del flujo. Si el código o una conversación contradicen esto, gana este archivo.

Última actualización: 2026-08-12

---

## 1. Propósito

Generar una **fotografía por cada referencia dimensional** de las piezas en las cuatro paredes del tanque. Cada foto contiene una sola cota real desde `(0,0)` para que en piso se pueda identificar y medir sin mezclar dimensiones.

Regla iLogic: `COTAS_CARAS_TANQUE`  
Código principal: `Planos/generador_caras_tanque.py`

---

## 2. Entrada en Inventor

Siempre:

1. Plano **machote** activo (`MACHOTE PLANOS.dwg` o equivalente).
2. **Tanque completo** (ensamble principal) abierto junto al machote.

No es requisito abrir por separado `Assembly Segmento 1..4` u otros IAM de cara.  
El flujo obtiene los segmentos del **árbol del tanque completo**. Cada vista se compone desde el tanque principal, aislando el IAM del segmento de esa cara y cualquier accesorio raíz físicamente ubicado en la misma pared (por ejemplo un bracket agregado fuera del IAM por ingeniería). La placa madre y el `(0,0)` siguen saliendo exclusivamente del segmento.

---

## 3. Salida

- Carpeta raíz: `Planos/JPG/<nombre_ensamble>/`.
- `COTAS_POR_REFERENCIA/`: subcarpetas `FRONT/`, `BACK/`, `LEFT/`, `RIGHT/`, `TOP/`; cada una contiene una foto JPG por referencia de esa cara, con nombre secuencial, eje y extremo.
- `PIEZAS_ACOTADAS/`: salida independiente del flujo original `COTAS_ILOGIC_ABIGAIL`. Se **divide por cara** con las mismas subcarpetas `FRONT/ BACK/ LEFT/ RIGHT/ TOP/` y una carpeta `OTROS/` para piezas que no puedan mapearse a una cara. Cada JPG por pieza cae en la subcarpeta del segmento al que pertenece por geometría.
- Log: `Planos/error_log_caras.txt` (debe registrar `cara <- segmento/subensamble` para las 5 caras y el conteo por subcarpeta de `PIEZAS_ACOTADAS`).
- Al terminar: borrar hojas temporales `TANQUE_DATUM_*` y dejar visible la hoja plantilla del machote (**nunca** quedarse en `Model (AutoCAD)`).

---

## 4. Origen (0,0)

En cada vista de pared:

- **(0,0)** = esquina **inferior-izquierda de la placa madre del segmento** en esa cara.
- La placa madre es la superficie rectangular del segmento que el personal usa como referencia, no la base exterior, soleras de base ni otros elementos estructurales externos.
- **No** usar picos de lifting lug, bridas sueltas, la base del tanque ni ruido de HLR como datum.

El origen es un datum interno de cálculo: **no** se dibuja la etiqueta, cruz ni indicador visual `(0,0)` en el JPG.

---

## 5. Cotas (cómo deben verse)

- Cotas **horizontales y verticales con líneas, extensiones y flechas**.
- Desde el **origen (0,0)** se deben señalar el **inicio y el fin** de cada componente en ambos ejes: `Xmin`, `Xmax`, `Ymin` y `Ymax` de su proyección. No basta una sola esquina de referencia.
- Una pieza rectangular produce **cuatro referencias**: `Xmin`, `Xmax`, `Ymin`, `Ymax`. Cada JPG muestra una única referencia desde el datum oculto.
- Una pieza circular, como flange, boss o puerto circular visto de frente, produce solo **dos fotografías**: `Xcentro` y `Ycentro`. No se acotan sus extremos exterior izquierdo/derecho/superior/inferior.
- En piezas rectangulares, incluso soleras delgadas, deben existir las **cuatro referencias**: dos X y dos Y. No se pierde un extremo por una tolerancia aplicada en coordenadas de hoja.
- Cuando dos o más piezas distintas comparten exactamente una referencia, se exporta una sola cota con el sufijo **`TYP`** y puntos azules sobre todos los accesorios a los que aplica. Extremos distintos de una misma pieza nunca se fusionan.
- Texto y líneas **fuera del tanque**, legibles, en stacks H/V. La vista se escala y desplaza después de contar las cotas para reservarles el mayor espacio útil de la hoja. El JPG se recorta a las curvas del dibujo y a su cota; no muestra marco, título ni esquinas del plano.
- La envolvente 3D de la ocurrencia conserva las cuatro referencias de una solera aunque el HLR muestre pocas aristas. Al reencuadrar cada JPG, el origen y el extremo seleccionado se **releen de las curvas HLR vigentes**: la línea de extensión debe terminar exactamente sobre la arista visible, nunca en una coordenada transformada de forma aproximada.
- No existe un tope que descarte cotas por cantidad de accesorios: se conserva cada extremo y el encuadre adapta la escala/posición de la vista.
- Estilo: números azules, sin símbolos de diámetros/unidades que confundan.

La API nativa de dimensiones de Inventor ha sido inestable para estas vistas
arbitrarias (`AddLinear` generó abanicos; `OrdinateDimensions` dejó texto
suelto). La salida de producción usa un `DrawingSketch` que dibuja las líneas,
flechas y textos H/V; el valor real se calcula como distancia en hoja dividida
entre la escala ortográfica de la vista y se formatea con las unidades del
machote.

### Prohibido

- Cotas **alineadas / inclinadas** (abanico).
- **Números flotantes** sin líneas (Ordinate + `HideValue` o estilos que apaguen la cota).
- Acotar el **tamaño** de la pieza como si fuera el dato principal; el dato principal es **dónde está** respecto a (0,0).
- Mezclar en una cara accesorios de **otro** segmento/pared.

---

## 6. Qué se acota en cada cara

Se acotan **todas las piezas independientes que están sobre la superficie física de la cara**, con sus extremos de inicio/fin desde el datum. El nombre del `.iam` o de la pieza no decide si se acota.

Incluye (ejemplos transversales): flanges, nipples, ground pads, tierras, lugs, jacking pads, parking stands, patches, bosses, gauges, manways, bushing patches, pipes, valves, etc.

También piezas con **código opaco** (OTC `62201-…`, `SP-###`, PTT `AS000…`) si pertenecen al subensamble/segmento de esa pared, y accesorios independientes colgados directamente del tanque si su ubicación física corresponde a esa cara.

Los **barrenos, agujeros y curvas de una placa no son piezas** y no se acotan por sí mismos. Se excluye únicamente la placa madre de la pared (la ocurrencia de mayor área proyectada de la cara); se incluyen las soleras, bases, refuerzos y demás ocurrencias independientes aunque formen parte de la estructura.

### No se acota

- La **placa madre** usada para fijar el `(0,0)`.
- Barrenos, agujeros o aristas que pertenezcan a una pieza, pero que no sean una ocurrencia independiente.
- Elementos que estén en otra pared física.

Las soleras, bases, refuerzos y demás ocurrencias independientes **sí** se acotan si están en la superficie física de la cara.

---

## 7. Segmentos y mapeo a caras

### Vantran (y similares)

Suelen existir `Assembly Segmento 1..4` (o `Placa Segmento N`) **dentro** del IAM principal.

### Otras familias (OTC, SWE, GIGA, SUNBELT, PTT, …)

Muchas **no** usan la palabra “Segmento”. Hay que:

1. Detectar subensambles (o clusters) asociados a cada pared exterior, y/o  
2. Asignar por **geometría** (centroide vs normales de las 4 paredes).

### Mapeo FRONT / BACK / LEFT / RIGHT / TOP

No confiar en el número del nombre (`Segmento 1` ≠ FRONT siempre).

Mapear por centroide 3D del contenedor contra las normales de cámara estilo PQart:

- FRONT ↔ `+face`
- BACK ↔ `-face`
- RIGHT ↔ `+right`
- LEFT ↔ `-right`
- **TOP ↔ `+cover` (tapa)**

Log obligatorio por cara, por ejemplo:  
`FRONT <- Assembly Segmento 2` o `FRONT <- 62201-1248-A05 (auto)` o `TOP <- TAPA_A01 (por nombre)`.

---

## 8. Cámara y orientación

- Criterio PQart de postura: tapa/arriba ≈ **+Y**; cara principal ≈ **+Z** (`face`); lateral ≈ **+X** (`right`).
- Las cámaras de foto miran las **paredes reales** (normales de caras planas), **sin** forzar ejes mundo puros (eso deja el tanque “chueco” en el JPG).
- Enderezar la vista en hoja si los bordes largos salen rotados unos grados.
- Para la cara **TOP** la cámara mira desde `+Y` hacia abajo, usando la misma placa madre lógica: la superficie superior del ensamble de tapa (mayor área proyectada horizontal).

**Base (fondo) del tanque:** **NO** se acota. Solo se cubren las 4 paredes laterales + TOP.

---

## 9. Familias reales (contexto de piso)

Análisis de Órdenes de Producción (`Z:\…\ORDENES DE PRODUCCION`, carpeta `10. Solidos`):

- Familias vistas: **VANTRAN, OTC, SWE, GIGA, SUNBELT, PTT**, …
- ~**84%** de nombres `PRODUCT` en STEP son códigos sin keyword inglesa.
- Solo Vantran usa de forma estable `Assembly Segmento N`.
- Por eso: **geometría + árbol del ensamble**, no hardcode de un cliente.

Detalle del barrido: `Planos/_analisis_steps_op_resultado.txt` (auxiliar).

---

## 10. Anti-regresiones (no volver a romper)

1. No dejar la hoja activa en **Model (AutoCAD)** (fondo negro / machote sucio).
2. No **snap** de cámara a ±X/±Z mundo si el ensamble viene inclinado.
3. No borrar cotas H/V válidas con heurísticas de RangeBox “demasiado grandes”.
4. No mezclar el catálogo de los **4** segmentos en **una** sola cara.
5. El JPG debe **encuadrar vista + cotas** (si `cotas_detectadas=0` en el log, el recorte está mal).
6. No mutar `dim.Style` compartido del machote de forma que apague cotas.
7. OriginIndicator: no eliminarlo a lo bruto si las cotas dependen de él; si no se usa ordinate, igual puede marcar (0,0).

---

## 11. Criterio de “listo” por corrida

| Check | OK |
|-------|----|
| Preparación | Machote + tanque completo |
| Caras generadas | 5: FRONT, BACK, LEFT, RIGHT, TOP (una por carpeta con sus JPG) |
| Cotas | Líneas H/V desde (0,0) de cada cara, legibles |
| Contenido de cada cara | Accesorios del segmento + accesorios raíz físicamente asignados a esa cara |
| PIEZAS_ACOTADAS | Subcarpetas FRONT/ BACK/ LEFT/ RIGHT/ TOP/ (+ OTROS si aplica) con conteo coherente |
| Base | No hay carpeta ni JPG de la base |
| Log | Mapeo `cara <- segmento` para 5 caras y conteos de PIEZAS_ACOTADAS por subcarpeta |
| Machote | Limpio, plantilla visible |

Validar al menos: **Vantran** (segmentos claros) y un **OTC** (códigos, sin “Segmento”). Si un tanque no tiene tapa detectable, el flujo debe reportarlo en el log y **omitir TOP** sin abortar el resto.

---

## 12. Archivos del sistema

| Archivo | Rol |
|---------|-----|
| `Planos/generador_caras_tanque.py` | Vistas, segmentos, cotas, JPG |
| `Planos/orientacion_pqart.py` | Marco tapa/cara/lateral |
| `Planos/cota_estilo.py` | Color/negrita texto de cota |
| `Planos/error_log_caras.txt` | Log de la última corrida |
| Regla iLogic `COTAS_CARAS_TANQUE` | Lanza el generador |

---

## 13. Flujo lógico (resumen)

```
Machote + tanque completo
  → detectar 4 segmentos/contenedores en el IAM
  → mapear cada uno a FRONT/BACK/LEFT/RIGHT (centroide)
  → por cada cara: cámara PQart + vista HLR
  → acotar solo accesorios de ese segmento desde (0,0) H/V
  → exportar JPG encuadrado
  → limpiar hojas temporales del machote
```
