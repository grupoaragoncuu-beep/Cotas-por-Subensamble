# Despliegue en otra PC (Windows + Inventor)

Guía corta para clonar y usar COTAS ABIGAIL **sin rutas hardcodeadas** de la PC de desarrollo.

## Requisitos

| Requisito | Notas |
|-----------|--------|
| Windows 10/11 x64 | |
| Autodesk Inventor (licencia válida) | Misma familia de versión recomendada |
| Python 3.10+ en el PATH | Con `pip` |
| Repo clonado en cualquier carpeta | Ej. `D:\Apps\Cotas-por-Subensamble` |

## Instalación (una vez por PC)

```bat
cd <carpeta-del-repo>
python -m pip install -r requirements.txt

cd Planos
instalar_boton_inventor.bat
```

El instalador:

1. Escribe `Planos\ilogic\config_planos.txt` con **rutas de esta PC** (`PLANOS_DIR`, `PYTHON_EXE`).
2. Define la variable de usuario `COTAS_ABIGAIL_DIR` = carpeta `Planos`.
3. Registra las reglas iLogic externas en Inventor.

Luego **cierra y vuelve a abrir Inventor** y agrega los botones a la cinta (instrucciones en consola del instalador).

## Cómo resuelve rutas el flujo

Orden en las reglas iLogic:

1. Variable de entorno `COTAS_ABIGAIL_DIR`
2. `PLANOS_DIR` en `ilogic\config_planos.txt`
3. Carpetas de reglas externas de iLogic (`ExternalRuleDirectories`)

Archivos temporales entre módulos (diámetro / piezas sólidas):

- Antes: `C:\Temp\...`
- Ahora: `Planos\.runtime\` (local al repo, no se versiona)

Salida JPG: siempre relativa → `Planos\JPG\<tanque>\...`

## Uso diario

1. Abrir ensamble **nativo** (`.iam` / `.ipt`) con Colorimetria aplicada.
2. Abrir / activar el machote de planos.
3. Ejecutar la regla:
   - `COTAS_CARAS_TANQUE` — caras + piezas
   - `COTAS_POR_SUBENSAMBLE` — solo caras
   - `COTAS_ILOGIC_ABIGAIL` — solo piezas

## Qué NO copiar entre PCs

- `Planos\ilogic\config_planos.txt` (local; regenerarlo con el instalador)
- `Planos\.runtime\`
- `Planos\error_log*.txt`
- JPGs / logs de corridas locales (opcionales)

Plantilla: `Planos\ilogic\config_planos.example.txt`

## Scripts auxiliares (no son el flujo de producción)

- `_analisis_steps_op.py` escanea `Z:\` (servidor corporativo). En PCs sin `Z:` mapeada no aplica; no afecta iLogic ni cotas.
- `auditoria_tanques.py` y `_analisis_tanques_locales.py` usan rutas relativas al repo (`Tanques/`).

## Checklist rápido si falla en otra PC

1. ¿Corriste `instalar_boton_inventor.bat` en **esa** PC?
2. ¿Reiniciaste Inventor después?
3. ¿`echo %COTAS_ABIGAIL_DIR%` apunta a `...\Planos`?
4. ¿`python -c "import win32com; import PIL"` funciona?
5. ¿El ensamble tiene iProperty `Clasificación`? (si no, piezas → `SIN CLASIFICACION`)
6. ¿Existe `Planos\MACHOTE PLANOS.dwg` (o lo tienes abierto en Inventor)?
