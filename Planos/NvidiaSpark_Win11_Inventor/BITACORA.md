# Bitácora de evidencias — Win11 + Inventor en NvidiaSpark

Regla: cada fase se marca con fecha/hora, comando o acción, resultado (OK/FAIL/PARCIAL), y captura o salida adjunta.

---

## Fase 0 — Apertura del expediente
| Campo | Valor |
|-------|--------|
| Fecha/hora | 2026-08-11 |
| Acción | Se acepta orden de jefatura y se crea paquete operativo |
| Resultado | OK |
| Nota | Se deja explícito que se ejecutará el método ordenado para deslinde si no es viable |

---

## Fase 1 — Diagnóstico del host NvidiaSpark
| Campo | Valor |
|-------|--------|
| Fecha/hora | 2026-08-11 (~13:30–13:35 local) |
| Ejecutor | Cursor en NvidiaSpark (`spark-7e79`, proyecto `arga_dev`) |
| Host | `spark-7e79` — aarch64, `/dev/kvm` RW, Docker 29.2.1, ~3.2 TB libres, ~113 GiB RAM, nvidia-smi OK (GB10) |
| Resultado | OK |
| Go/No-Go | **GO** |
| Evidencia | En Spark: `~/NvidiaSpark_Win11_Inventor/evidencias/fase1_diagnostico.txt` (+ captura chat remoto) |

Checklist Go (todos deben cumplirse o justificarse):
- [x] Arquitectura `aarch64`
- [x] Existe `/dev/kvm` usable
- [x] Docker disponible
- [x] Espacio libre >= 100 GB (reportado ~3.2 TB)
- [x] RAM host suficiente (reportado ~113 GiB)
- [x] Riesgo aceptado documentado: sin GPU passthrough

---

## Fase 2 — Levantamiento Windows 11 ARM
| Campo | Valor |
|-------|--------|
| Fecha/hora | 2026-08-11 (~13:35) |
| Método | dockur/windows-arm (`ghcr.io/dockur/windows-arm`) |
| Contenedor | `nvidia_spark_win11_arm` — 16G RAM / 6 cores / 120G disco |
| Acceso | `http://localhost:8006` → HTTP 200 |
| Resultado | **OK** — escritorio Windows 11 ARM visible en `localhost:8006` (2026-08-11) |
| Evidencia | Compose + logs en Spark `~/NvidiaSpark_Win11_Inventor/` + captura desktop Win11 |

---

## Fase 3 — Instalación Inventor
| Campo | Valor |
|-------|--------|
| Fecha/hora | 2026-08-11 → 2026-08-12 |
| Versión Inventor | Autodesk Inventor Professional (instalador oficial copiado desde PC del usuario vía carpeta Shared del contenedor) |
| Resultado | **INSTALACIÓN OK / EJECUCIÓN FAIL** |
| Detalle | Inventor instaló correctamente dentro de Win11 ARM. Al abrirlo, la aplicación se cierra inesperadamente y muestra "Autodesk Inventor Error Report — A software problem has caused Autodesk Inventor to close unexpectedly." |
| Evidencia | Captura del diálogo de crash `Autodesk Inventor Error Report` (2026-08-12) |

---

## Fase 4 — Pruebas de usabilidad
| Prueba | Resultado | Observación |
|--------|-----------|-------------|
| Abrir Inventor | FAIL | Crash al inicio de la aplicación |
| Pieza simple | NO EJECUTABLE | No se llegó a la UI de modelado |
| Ensamble | NO EJECUTABLE | Idem |
| Orbit/pan | NO EJECUTABLE | Idem |
| Estabilidad 30 min | NO APLICA | La aplicación no se mantiene abierta |

Causa técnica más probable (coincide con lo pronosticado en la Fase 0):
- Windows 11 ARM + Inventor x64 vía emulación **Prism**.
- Sin GPU passthrough / vGPU en la VM → gráficos por software.
- Autodesk **no soporta oficialmente** Inventor sobre Windows on ARM ni sobre VM sin GPU certificada.
- Combinación no soportada → crash del proceso al iniciar.

---

## Fase 5 — Conclusión para jefatura
| Campo | Valor |
|-------|--------|
| Fecha/hora | 2026-08-12 |
| ¿Se siguió el método ordenado? | **SÍ** — VM Win11 en NvidiaSpark levantada, Inventor instalado |
| ¿Es viable para producción CAD? | **NO** — Inventor no arranca establemente en este entorno |
| Causa raíz | Combinación ARM64 (host) + Win11 ARM (VM) + Inventor x64 emulado + sin GPU dedicada = no soportado por Autodesk → crash |
| Recomendación | Continuar trabajo productivo con Inventor en **PC Windows x64 con GPU certificada** (equipo actual del usuario). Mantener NvidiaSpark para automatización, análisis de datos y scripts Python del proyecto. Piloto cerrado con evidencia; se ejecutó el método ordenado. |
| Deslinde | Se agotó la ruta ordenada por jefatura hasta el punto donde el fabricante (Autodesk) no soporta el entorno. Fallo por limitaciones de plataforma, no por omisión del procedimiento. |
