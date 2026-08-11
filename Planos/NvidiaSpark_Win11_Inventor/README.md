# Cómo ejecutar el piloto CON Cursor dentro de NvidiaSpark

Este Cursor que trabaja el repo en `C:\Proyectos\...` es tu **PC Windows local**.  
Para que el agente ejecute comandos **en el Spark**, debes abrir Cursor **desde NvidiaSpark**.

## Opción recomendada (Cursor nativo en Spark)

1. En NVIDIA Sync / NvidiaSpark, lanza la app **Cursor**.
2. En ese Cursor remoto, abre o clona/copia este proyecto (o al menos la carpeta `Planos/NvidiaSpark_Win11_Inventor`).
3. Abre el chat del agente y escribe:
   > Continúa piloto Win11+Inventor. Ejecuta Fase 1 (`01_fase1_diagnostico.sh`) y pega/actualiza la bitácora.
4. El agente podrá correr `bash`/`docker` en el host Linux ARM.

## Opción alternativa (Terminal Sync)

1. Abre **Terminal** en NvidiaSpark.
2. Copia la carpeta `Planos/NvidiaSpark_Win11_Inventor` al Spark (scp, git, shared path, etc.).
3. Ejecuta:

```bash
cd Planos/NvidiaSpark_Win11_Inventor
chmod +x 01_fase1_diagnostico.sh 02_fase2_levantar_win11_arm.sh
bash 01_fase1_diagnostico.sh
```

4. Pega la salida aquí / guárdala en `evidencias/` y avisa para Fase 2.

## Comandos de control útiles (Fase 2+)

```bash
# Estado
docker ps -a --filter name=nvidiaSpark_win11_arm_inventor
docker logs -f nvidiaSpark_win11_arm_inventor

# Detener (sin borrar disco de la VM)
docker compose -f docker-compose.yml down

# Reiniciar
docker compose -f docker-compose.yml up -d
```

## Carpeta del expediente
`Planos/NvidiaSpark_Win11_Inventor/`

| Archivo | Uso |
|---------|-----|
| `00_ORDEN_Y_ALCANCE.md` | Marco de la orden y deslinde |
| `BITACORA.md` | Evidencia cronológica |
| `01_fase1_diagnostico.sh` | Diagnóstico Go/No-Go |
| `02_fase2_levantar_win11_arm.sh` | Levanta Win11 ARM |
| `03_fase3_checklist_inventor.md` | Checklist Inventor |
| `evidencias/` | Logs y capturas |
