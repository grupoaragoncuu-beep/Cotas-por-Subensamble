# Estado actual — COTAS ABIGAIL

Actualizado: 2026-09-04

## Flujos listos para otras PCs

| Botón iLogic | Motor | Qué entrega |
|---|---|---|
| `COTAS_POR_SUBENSAMBLE` | `generador_caras_tanque.py` | Mapa de caras TOP/SEGM1–4/BASE, origen (0,0), JPG por referencia |
| `COTAS_POR_SEG` | mismo | Prueba rápida de **una** cara |
| `COTAS_ILOGIC_ABIGAIL` | `generador_piezas.py` | Piezas LARGO/ANCHO/THK/Ø/ALTO → `PIEZAS_ACOTADAS` |
| `COTAS_POR_SEG_PIEZAS` | mismo | Prueba rápida de piezas de **una** cara |
| `COTAS_ENSAMBLES_INDEPENDIENTES` | `generador_ensambles_instructivo.py` | Kits independientes, 6 vistas |
| `COTAS_BOARD` | `generador_board.py` | Tablero GIGA/Board → `JPG/<job>/BOARD/` |

Tras `git pull`: ejecutar `Planos\instalar_boton_inventor.bat` y reiniciar Inventor.

## Mejoras incluidas en este corte

- Origen (0,0) por silueta frontal del segmento (no tangente de doblez / HLR engañoso).
- Organización por cara SEGM/TOP/BASE (caras y piezas); copia de pieza multi-cara.
- Diámetros de barrenos en placas (`*_DIAMETRO_Hnn`) desde aristas circulares.
- Gate anti-JPG vacío (exige cota asociativa; ALTO/PATA sin cota se borran).
- Perfiles U/L: detector más estricto + reintentos de cámara transversal.
- Inventario tipológico de `4. Planos` en OPs reales (OTC/SWE/Vantran/GIGA/PTT).

## Cobertura por producto

- **Tanques** OTC / Vantran / SWE / PTT (casco): mismo motor; validar picks TOP/SEGM/BASE.
- **Board / GIGA tablero** (`9919-Board`, etc.):
  - Caras / ensambles → `generador_board.py` (`JPG/<job>/BOARD/…`).
  - **Abigail / piezas** → `generador_piezas.py` (sin picks). Piezas **cobre**
    (prefijos ABB, GENE, GEN1, GE3R, AcuCT, CITEL, 9919-F/M/P, … mapeados desde
    nesteos de cobre): **2 JPG por cota** (con dimensión + `*_SIN_COTA_*`).
- **No es objetivo** (aún): gabinetes ATC sueltos, STD/weld/BOM de PDF.

## Artefactos de análisis (referencia)

- `Planos/_analisis_planos_global_resultado.txt`
- `Planos/AUDITORIA_TANQUES.md`
- `Planos/_validar_cotas_silueta.py`
