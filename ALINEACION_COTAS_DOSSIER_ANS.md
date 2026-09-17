# Alineación ANS ↔ COTAS ABIGAIL — tabla `cotas_dossier`

Documento para la pestaña de Cursor del proyecto
`C:\Proyectos\COTAS ABIGAIL\COTAS ABIGAIL` **y** la del ANS
(`C:\Proyectos\New Arga Nesting Suite`).

Describe la tabla en NestingPro y cómo Cotas **ya escribe** evidencias
orientadas al flujo VSM (Cotas → Incoming → Nest).

Actualizado: 2026-09-11 (columna `clasificacion` = carpeta JPGS)

---

## 0. Mensaje para la ventana ANS (checklist de empalme)

Cotas Abigail **ya inserta** en `public.cotas_dossier` al exportar JPG.

| Contrato | Cotas (hecho) | ANS debe |
|----------|---------------|----------|
| BD | `192.168.2.80:5433` / `nestingpro_db` | Misma tabla (DDL ya en `cotas_dossier_service.py`) |
| `cliente` / `producto` | Desde **VSM** `foldertree.jobs` (`client` / `product`), no inventados | Al armar dossier / nest, filtrar por esos mismos campos |
| `job` | Token del ensamble (= `jobs.job_number` VSM) | `WHERE job = …` |
| `type` | `TYP` (cotas tipadas) | Repetir imagen `cantidad_spoteos` veces en el dossier |
| `cantidad_spoteos` | `1` = cota normal; `N` = N coincidencias TYP | Armar N copias de la misma foto |
| `clasificacion` | Carpeta de proceso bajo `JPGS\` (`Almacén`, `Corte/Corte`, `Corte/Doblado`, `Corte/Estañado`, `Doblado`, `Maquinado`, `Plasma`, `Plasma Doblado`, `SIN CLASIFICACION`) | Filtrar / ordenar dossier por proceso |
| `ruta` | Ruta absoluta del JPG exportado | Leer archivo; ideal bajo `…\DOSSIER FILES\JPGS\<proceso>\` |
| Orden proceso | Cotas **antes** del nest | Nest no debe ser prerequisito del INSERT |

**Fuente de verdad cliente/producto:** VSM (`jobs.client`, `jobs.product`).
ANS `erp_jobs` / `diccionario_swo` son **respaldo** (mismos valores cuando
el JOB ya pasó por VSM). Cotas **no** casa el nesting como origen primario.

Ejemplo VSM vivo (GIGA):

| job_number | product | client | local_path (mount) |
|------------|---------|--------|--------------------|
| `9919-BOARD8` | `ENCLOSURES NEMA 1` | `GIGA` | `…/ENCLOSURES NEMA 1/GIGA/9919-BOARD8` |

Ruta de evidencias esperada en red ( Incoming / dossier ):

`X:\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA\<JOB>\DOSSIER FILES\JPGS\`  
subcarpetas: `Almacén`, `Corte/Corte`, `Corte/Doblado`, `Corte/Estañado` (cobre SIN_COTA iso), `Doblado`, `Maquinado`, `Plasma`, …

Parseo `…/PRODUCTO/CLIENTE/JOB` = misma regla que ANS
`_producto_cliente_desde_job_root`.

---

## 1. Dónde vive la tabla

| Dato | Valor |
|------|--------|
| Servidor | `192.168.2.80` |
| Puerto | `5433` |
| Base | `nestingpro_db` |
| Usuario | `postgres` (mismo `NESTING_DB_*` del ANS) |
| Schema | `public` |
| Tabla | **`cotas_dossier`** |

VSM (solo lectura de cliente/producto/ruta JOB):

| Dato | Valor |
|------|--------|
| Puerto | `5437` |
| Base | `foldertree` |
| Usuario | `user` |
| Tabla | `jobs` (`job_number`, `client`, `product`, `local_path`) |

---

## 2. Columnas (contrato)

| Campo | Tipo | Ejemplo | Uso |
|-------|------|---------|-----|
| `id` | SERIAL PK | `1` | ID único |
| `cliente` | TEXT NOT NULL | `GIGA` | = VSM `jobs.client` |
| `producto` | TEXT | `ENCLOSURES NEMA 1` | = VSM `jobs.product` |
| `job` | TEXT NOT NULL | `9919-BOARD2_2` | = VSM `jobs.job_number` |
| `type` | TEXT | `TYP` | Cota tipada (todas las evidencias Cotas) |
| `cantidad_spoteos` | INTEGER ≥ 0 | `1` / `3` | `1` = normal; `N` = N coincidencias TYP |
| `nombre_archivo` | TEXT | `…__LENGTH_12.5.jpg` | Nombre JPG |
| `ruta` | TEXT | UNC / `X:\…` | Ruta real |
| `clasificacion` | TEXT | `Almacén` / `Corte/Corte` / `Corte/Doblado` / `Corte/Estañado` / `Doblado` / … | Carpeta de proceso bajo `JPGS\` |
| `created_at` | TIMESTAMP | NOW() | Alta |

---

## 3. Lógica type + spoteos

1. `type = TYP` → cota tipada (evidencia Cotas).
2. `cantidad_spoteos = 1` → cota normal (una aparición).
3. `cantidad_spoteos = N` → la misma foto se coloca N veces (N piezas/coincidencias TYP).
3. Armado dossier:

```text
for cada fila del job:
    repetir cantidad_spoteos veces:
        insertar imagen(ruta)
```

---

## 4. Implementación Cotas (archivos)

| Archivo | Rol |
|---------|-----|
| [`Planos/cotas_dossier_registro.py`](Planos/cotas_dossier_registro.py) | Resolver VSM + INSERT + sync carpeta |
| [`Planos/generador_vistas.py`](Planos/generador_vistas.py) | `registrar_jpg` tras cada export (+ SIN_COTA) |
| [`Planos/generador_piezas.py`](Planos/generador_piezas.py) | `iniciar_sesion_dossier` + sync post-reorg |
| [`Planos/generador_board.py`](Planos/generador_board.py) | Sesión + sync BOARD |

Fail-soft: si no hay red/DB, el acotado **sigue**; solo log `AVISO dossier`.

Flags:

- `COTAS_DOSSIER=0` → desactiva registro.
- `COTAS_DOSSIER_CLIENTE` / `COTAS_DOSSIER_PRODUCTO` → override manual.
- `COTAS_VSM_JOB_ROOT` → carpeta JOB si no se resuelve por BD.
- `COTAS_VSM_DRIVE=X:` → traduce `/mnt/server_data/…` → `X:\…`.

---

## 5. Flujo objetivo (VSM)

```text
VSM crea JOB (client, product, local_path)
        ↓
Cotas acota → JPG (+ clasificación Almacén/Corte/…)
        ↓
INSERT cotas_dossier (cliente/producto desde VSM)
        ↓
Incoming / dossier lee tabla y arma evidencias (TYP × cantidad_spoteos)
        ↓
Nest / ANS (misma cliente/producto/job)
```

---

## 6. Verificación cruzada ANS ↔ Cotas

Desde ANS:

```powershell
cd "C:\Proyectos\New Arga Nesting Suite"
py -3.14 archive\tools\_migrate_cotas_dossier.py
```

Consulta tras una corrida Cotas:

```sql
SELECT id, cliente, producto, job, type, clasificacion, cantidad_spoteos, nombre_archivo, ruta
FROM public.cotas_dossier
WHERE job ILIKE '%9919-BOARD%'
ORDER BY id DESC
LIMIT 20;
```

Debe coincidir `cliente`/`producto` con:

```sql
-- VSM :5437 foldertree
SELECT job_number, client, product, local_path
FROM public.jobs
WHERE job_number ILIKE '%9919-BOARD%';
```

Smoke local Cotas (sin Inventor):

```powershell
cd "C:\Proyectos\COTAS ABIGAIL\COTAS ABIGAIL\Planos"
python cotas_dossier_registro.py
```

---

## 7. Publicación a DOSSIER FILES (anti-crash)

Al iniciar sesión (`iniciar_sesion_dossier`):

1. Busca la carpeta del JOB (VSM `local_path` → `X:\…`, env, contexto previo).
2. Si **no** existe / no es escribible → abre el explorador de Windows para
   elegir `DOSSIER FILES` (o el JOB), p. ej.  
   `\\192.168.2.80\Users\Administrator\Desktop\Grupo Arga Metals\ARGA METALS CORPORATE SYSTEM\ENCLOSURES NEMA 1\GIGA\9919-BOARD2_2\DOSSIER FILES`
3. Crea `DOSSIER FILES\JPGS` si falta y guarda `job_root` / `dossier_jpgs` en
   `Planos/.runtime/dossier_contexto.json`.

Tras reorganizar cotas:

1. Copia el árbol a `…\DOSSIER FILES\JPGS\<clasificacion>\…` (reintentos de red).
2. INSERT/UPDATE en `cotas_dossier.ruta` = ruta **corporativa** exacta.

Env: `COTAS_DOSSIER_PUBLISH=0` (solo DB local), `COTAS_DOSSIER_PICKER=0` (sin UI).

Pendiente opcional:

- Catálogo CHECK de `type` en ANS.
- Generador PDF/DOCX del dossier (ANS o Cotas).
