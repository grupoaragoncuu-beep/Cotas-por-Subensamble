# Informe a jefatura — Piloto Windows 11 + Inventor en NvidiaSpark

**Asunto:** Resultado del piloto ordenado — Windows 11 + Autodesk Inventor en NvidiaSpark (DGX Spark)
**Fecha de cierre:** 2026-08-12

---

## 1. Cumplimiento de la instrucción
Se ejecutó el método ordenado: máquina virtual Windows 11 dentro de NvidiaSpark, instalación de Autodesk Inventor Professional y prueba de ejecución. Evidencias completas en:

- `Planos/NvidiaSpark_Win11_Inventor/BITACORA.md`
- `Planos/NvidiaSpark_Win11_Inventor/evidencias/`
- Capturas: escritorio Win11 ARM en `localhost:8006` y crash de Inventor.

## 2. Resultado técnico

| Fase | Resultado |
|------|-----------|
| Diagnóstico host Spark (Fase 1) | **GO** — KVM + Docker + espacio + RAM |
| VM Windows 11 ARM levantada (Fase 2) | **OK** — `dockur/windows-arm`, escritorio accesible |
| Instalación Autodesk Inventor (Fase 3) | **OK** — instalador oficial ejecutado dentro de la VM |
| Ejecución de Inventor (Fase 4) | **FAIL** — la aplicación se cierra inesperadamente al iniciar (`Autodesk Inventor Error Report`) |
| Usabilidad CAD | **No aceptable** — Inventor no permanece abierto |

## 3. Causa raíz
Combinación no soportada por el fabricante:

- **Host:** NvidiaSpark = DGX Spark ARM64 (GB10). Virtualización no soportada oficialmente por NVIDIA en este equipo.
- **VM:** solo puede correr Windows 11 **ARM64** (no Windows 11 x64). Autodesk no soporta oficialmente productos de escritorio sobre Windows on ARM.
- **Inventor:** binario x64 → corre bajo la capa de emulación **Prism** de Windows on ARM.
- **GPU:** no hay passthrough / vGPU utilizable para CAD en la VM. Inventor cae en modo gráficos por software.
- **Resultado esperable:** crash al inicio, consistente con el diálogo de error observado.

El fallo se produjo **después** de haber ejecutado correctamente todas las fases previas del método ordenado. No es una omisión del procedimiento; es una limitación técnica de la plataforma.

## 4. Recomendación

1. **Cerrar** el piloto de Inventor sobre NvidiaSpark. Documentado y con evidencia.
2. **Continuar** el trabajo productivo de Inventor (proyecto **Cotas por Subensamble**) en la PC Windows x64 del usuario, que es el único entorno soportado por Autodesk para este flujo.
3. **Reutilizar** NvidiaSpark para lo que sí aprovecha su hardware: automatización, análisis de datos, procesado por scripts Python y respaldos del proyecto (repo GitHub ya versionado y clonable en Spark).

## 5. Anexos
- `Planos/NvidiaSpark_Win11_Inventor/BITACORA.md`
- `Planos/NvidiaSpark_Win11_Inventor/evidencias/`
- Repositorio de respaldo: https://github.com/grupoaragoncuu-beep/Cotas-por-Subensamble

## 6. Nota de deslinde
Se ejecutó el método ordenado por jefatura en todas sus fases hasta el punto donde el propio fabricante (Autodesk) no soporta el entorno. El resultado negativo obedece a limitaciones de plataforma, no a omisión del procedimiento.
