# Fase 3 — Checklist instalación Inventor (dentro de la VM Win11 ARM)

Marca cada paso. Adjunta capturas en `evidencias/`.

## Pre-requisitos en la VM
- [ ] Escritorio Windows 11 ARM accesible (noVNC :8006 o RDP)
- [ ] Red funciona (abrir un navegador / ping)
- [ ] Hora/fecha correctas
- [ ] Al menos ~40 GB libres dentro de Windows
- [ ] Cuenta Autodesk / instalador disponible (licencia corporativa o trial)

## Instalación
- [ ] Descargar instalador **x64** de Inventor (versión corporativa ordenada)
- [ ] Ejecutar instalador
- [ ] Resultado: OK / FAIL / PARCIAL
- [ ] Si FAIL: pegar mensaje de error exacto aquí:

```
(error)
```

## Post-instalación
- [ ] Inventor abre
- [ ] Mensaje de graphics mode (Hardware / Software): ________
- [ ] Abrir pieza simple: tiempo / comportamiento: ________
- [ ] Abrir ensamble pequeño: ________
- [ ] Orbit/pan fluido? Sí / No / Apenas
- [ ] Export/PDF o guardar: ________

## Datos para bitácora
| Campo | Valor |
|-------|--------|
| Fecha/hora | |
| Versión Inventor | |
| Modo gráficos reportado | |
| ¿Cumple uso productivo? | Sí / No |
| Comentario técnico | |

## Nota de deslinde (no borrar)
Se instaló / se intentó instalar Inventor **dentro de Windows 11 virtualizado en NvidiaSpark**, conforme a la instrucción de jefatura. Limitaciones conocidas de plataforma (ARM64 host, sin GPU en la VM, Inventor no nativo ARM) pueden hacer el entorno lento o no usable aunque la instalación tecnicamente ocurra.
