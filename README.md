# Cotas por Subensamble (COTAS ABIGAIL)

Automatización de planos y cotas para tanques en **Autodesk Inventor**: genera fotografías JPG por referencia dimensional (una cota por imagen) a partir del ensamble del tanque y un machote de planos.

Repositorio de respaldo y versionado para trabajo en PC Windows y clonado en **NvidiaSpark / DGX Spark**.

---

## Qué hace

1. Toma el **tanque completo** (ensamble) abierto en Inventor junto con el plano machote.
2. Identifica caras de trabajo (`SEGM1`–`SEGM4`, `TOP`, `BASE`) mediante selección guiada.
3. **Subensamble / caras:** cotas desde origen **(0,0)** en la silueta frontal del segmento; JPG por referencia.
4. **Piezas (Abigail):** LARGO / ANCHO / THK / diámetros (ext-int y barrenos) / ALTO-pata en perfiles.
5. Exporta a `Planos/JPG/<ensamble>/COTAS_POR_REFERENCIA/` y `.../PIEZAS_ACOTADAS/`.

Estado consolidado del corte actual: [`ESTADO_ACTUAL.md`](ESTADO_ACTUAL.md).

Documento de verdad del flujo de caras (reglas de cotas, origen, salida):

- [`DEBER_SER_COTAS_CARAS.md`](DEBER_SER_COTAS_CARAS.md)

Si el código y una conversación contradicen ese archivo, **gana el DEBER_SER**.
---

## Requisitos

| Requisito | Notas |
|-----------|--------|
| Windows 10/11 **x64** | Inventor corre en Windows (no nativo ARM) |
| Autodesk Inventor | Con licencia corporativa / válida |
| Python 3.x + `pywin32` | Solo si ejecutas scripts `.py` fuera del `.exe` |
| Plano machote | `Planos/MACHOTE PLANOS.dwg` (o equivalente) |
| Modelo 3D | Archivos STEP/STP en `Tanques/` o ensamble IAM de trabajo |

> **NvidiaSpark:** el host es Linux ARM. Inventor requiere una VM Windows 11 ARM (piloto documentado en `Planos/NvidiaSpark_Win11_Inventor/`). Eso es experimental y sin GPU passthrough.

---

## Estructura del repositorio

```text
.
├── DEBER_SER_COTAS_CARAS.md     # Reglas del flujo (fuente de verdad)
├── README.md                    # Este archivo
├── Planos/
│   ├── generador_caras_tanque.py
│   ├── generador_tanque_completo.py
│   ├── generador_vistas.py
│   ├── interfaz_app.py          # UI Tkinter "COTAS ABIGAIL"
│   ├── inventor_com.py
│   ├── instalar_boton_inventor.bat / .py
│   ├── MACHOTE PLANOS.dwg
│   ├── ilogic/                  # Reglas iLogic
│   ├── JPG/                     # Salidas / referencias por ensamble
│   └── NvidiaSpark_Win11_Inventor/  # Piloto VM Win11 + Inventor
├── Tanques/                     # Modelos STEP/STP de referencia
└── prueba 1/                    # Pruebas auxiliares
```

Artefactos **no** versionados (regenerables / temporales): `Planos/dist/`, `Planos/build/`, `Planos.zip`, `__pycache__`, locks AutoCAD `~*.tmp`. Ver [`.gitignore`](.gitignore).

---

## Cómo usarlo (PC Windows + Inventor)

> **Otra PC / clon nuevo:** sigue [`DEPLOY.md`](DEPLOY.md) (rutas portables, instalador, checklist).

### Opción A — Botón / flujo iLogic

1. Abre Inventor.
2. Abre el **tanque completo** y el **machote** de planos.
3. Si aún no está instalado el botón:
   ```bat
   python -m pip install -r requirements.txt
   Planos\instalar_boton_inventor.bat
   ```
   (cierra y reabre Inventor después)
4. Ejecuta la regla / botón:
   - `COTAS_POR_SUBENSAMBLE` — mapa completo de caras
   - `COTAS_POR_SEG` — una sola cara (prueba rápida)
   - `COTAS_ILOGIC_ABIGAIL` — piezas acotadas
   - `COTAS_POR_SEG_PIEZAS` — piezas de una sola cara (prueba rápida)
5. Revisa salidas en `Planos/JPG/<ensamble>/` y el log `Planos/error_log_caras.txt`.

Ver también [`ESTADO_ACTUAL.md`](ESTADO_ACTUAL.md) y [`DEPLOY.md`](DEPLOY.md).
### Opción B — Interfaz Python

```bat
cd Planos
python interfaz_app.py
```

1. Conecta a Inventor (debe estar abierto).
2. Elige el ensamble en la lista.
3. Pulsa **Generar Planos**.

### Opción C — Scripts directos

Desde `Planos/`, con Inventor abierto y documentos cargados:

```bat
python generador_caras_tanque.py
python generador_tanque_completo.py
```

Consulta siempre [`DEBER_SER_COTAS_CARAS.md`](DEBER_SER_COTAS_CARAS.md) para reglas de (0,0), `TYP`, circulares vs rectangulares y carpetas de salida.

---

## Clonar en NvidiaSpark (o cualquier máquina)

```bash
git clone https://github.com/grupoaragoncuu-beep/Cotas-por-Subensamble.git
cd Cotas-por-Subensamble
```

Piloto Windows 11 + Inventor en Spark:

```bash
cd Planos/NvidiaSpark_Win11_Inventor
# Ver README.md de esa carpeta (diagnóstico, dockur/windows-arm, checklist Inventor)
```

---

## Actualizar el respaldo

En la PC de desarrollo (con acceso de escritura al repo):

```bash
git add -A
git status
git commit -m "Describe el cambio"
git push origin main
```

En Spark / otra máquina:

```bash
git pull origin main
```

---

## Estado del piloto NvidiaSpark (resumen)

| Fase | Estado |
|------|--------|
| Diagnóstico host Spark | GO (KVM + Docker + espacio/RAM) |
| Contenedor Win11 ARM (`dockur/windows-arm`) | En curso / instalación Windows |
| Inventor dentro de la VM | Pendiente |
| Uso productivo CAD | No garantizado (ARM + sin GPU en VM) |

Detalle y bitácora: [`Planos/NvidiaSpark_Win11_Inventor/`](Planos/NvidiaSpark_Win11_Inventor/).

---

## Contacto / organización

Repositorio: [grupoaragoncuu-beep/Cotas-por-Subensamble](https://github.com/grupoaragoncuu-beep/Cotas-por-Subensamble)
