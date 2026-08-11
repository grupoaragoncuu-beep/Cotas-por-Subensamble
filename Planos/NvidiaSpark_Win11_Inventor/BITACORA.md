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
| Resultado | **OK parcial** — contenedor arriba; Windows aún en descarga/instalación automática |
| Evidencia | Compose + logs en Spark `~/NvidiaSpark_Win11_Inventor/` + captura chat remoto |

---

## Fase 3 — Instalación Inventor
| Campo | Valor |
|-------|--------|
| Fecha/hora | _(pendiente)_ |
| Versión Inventor | _(pendiente)_ |
| Resultado | PENDIENTE |
| Evidencia | capturas + log instalador |

---

## Fase 4 — Pruebas de usabilidad
| Prueba | Resultado | Tiempo / observación | Evidencia |
|--------|-----------|----------------------|-----------|
| Abrir Inventor | PENDIENTE | | |
| Pieza simple | PENDIENTE | | |
| Ensamble pequeño | PENDIENTE | | |
| Orbit/pan | PENDIENTE | | |
| Estabilidad 30 min | PENDIENTE | | |

---

## Fase 5 — Conclusión para jefatura
| Campo | Valor |
|-------|--------|
| Fecha/hora | _(pendiente)_ |
| ¿Se siguió el método ordenado? | SÍ (a completar) |
| ¿Es viable para producción CAD? | _(pendiente)_ |
| Causa raíz si falla | _(pendiente: ARM / sin GPU VM / emulación / permisos / etc.)_ |
| Recomendación | _(pendiente)_ |
