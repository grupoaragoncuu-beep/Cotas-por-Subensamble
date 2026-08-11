# Piloto ordenado: Windows 11 + Inventor en NvidiaSpark (DGX Spark)

## Orden
Instrucción directa de jefatura: investigar e implementar una máquina virtual que simule Windows 11 e instalar Autodesk Inventor **dentro de NvidiaSpark**, aunque el método no sea el más viable.

## Objetivo del piloto
1. Seguir el método ordenado (VM Windows en NvidiaSpark + intento de instalar Inventor).
2. Dejar bitácora y evidencias técnicas de cada fase.
3. Si el método falla o resulta inutilizable, documentar causa raíz (arquitectura/limitaciones de plataforma), no omisión del procedimiento ordenado.

## Alcance explícito
- **Dentro del alcance:** diagnóstico del host Spark, creación/arranque de Win11 ARM (KVM/Docker), intento de instalación y pruebas mínimas de Inventor, reporte con evidencias.
- **Fuera del alcance de “éxito productivo”:** garantizar rendimiento CAD igual a una PC Windows x64 con GPU dedicada. Eso depende de limitaciones conocidas de DGX Spark (ARM64, sin GPU passthrough a la VM, Inventor no nativo ARM).

## Criterios
| Tipo | Criterio |
|------|----------|
| Cumplimiento de orden | Se ejecutaron las fases del método ordenado y quedó evidencia |
| Éxito técnico mínimo | Win11 ARM arranca y se inicia el instalador de Inventor |
| Éxito técnico usable | Inventor abre pieza/ensamble con latencia de trabajo aceptable |
| Cierre por no viabilidad | Fallo documentado por limitaciones de plataforma pese a ejecutar el plan |

## Responsables de ejecución
- Ejecución técnica: equipo / agente Cursor siguiendo este paquete.
- Decisión final de continuidad: jefatura, con base en bitácora y resultados.

## Fecha de inicio del expediente
2026-08-11
