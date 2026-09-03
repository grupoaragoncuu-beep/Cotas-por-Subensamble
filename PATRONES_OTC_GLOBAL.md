# Patrones globales OTC (Planos + ensamble)

Fuente: `Z:\...\ORDENES DE PRODUCCION\*OTC*\4. Planos` + STEPs en `10. Solidos` cuando existen.

**OPs OTC analizadas:** 11

## 1. Arbol de ensamble tipico (estable entre tanques)

Confirmado en los 11 OTC (62140, 62147, 62154, 62155, 62175, 62176, 62177, 62178, 62200, 62201, 62248). Todos tienen `4. Planos` con `*-46.00`, `*-47.00` y `*-48.00`.

```
*-*46-A01.iam          ← TANQUE (plano *-*46.00)
├── *-*47-A0N          ← TOP COVER (plano *-*47.00 / 47.01; a veces varios A01..A07)
└── *-*48-A01          ← CASCO / shell (plano *-*48.00) — NO es una cara
    ├── *-*48-A02      ← BASE
    ├── *-*48-A03      ← pared / SEGM
    ├── *-*48-A04      ← pared / SEGM
    ├── *-*48-A05      ← pared / SEGM
    └── *-*48-A06      ← pared / SEGM
         (+ A07+ = subarmados de pared / bridas / etc.)
```

La familia numerica usa los **ultimos 2 digitos**: `1246`→46 tanque, `1247`→47 tapa, `1248`→48 casco. El prefijo (`12`, etc.) puede variar; el rol lo dan `46/47/48`.

**Nota:** algunos STEP exportados vienen “planos” (pocos PRODUCT); el arbol completo se ve en el `.iam` de Inventor. El patron de planos PDF es el contrato estable.

## 2. Codigos de dibujo en 4. Planos

| Codigo PDF | Rol | Frecuencia |
|------------|-----|------------|
| `*-47.xx` | TOP_COVER | 14 |
| `*-48.xx` | CASCO_SHELL | 12 |
| `*-46.xx` | TANQUE_GENERAL | 11 |
| `*-51.xx` | DETALLE_O_SUB | 11 |
| `*-52.xx` | DETALLE_O_SUB | 10 |
| `*-60.xx` | DETALLE_O_SUB | 10 |
| `*-53.xx` | DETALLE_O_SUB | 9 |
| `*-54.xx` | DETALLE_O_SUB | 9 |
| `*-49.xx` | OTRO_O_DETALLE | 7 |
| `*-59.xx` | OTRO_O_DETALLE | 5 |
| `*-50.xx` | DETALLE_O_SUB | 5 |
| `*-58.xx` | OTRO_O_DETALLE | 1 |

## 3. Reglas para el generador (globalizar)

1. **Contenedor de cara** = IAM `A02`…`A06` (48) o `47-A*`, nunca `48-A01` ni `46-A01`.
2. **TOP** = familia rol `47`; **BASE** = `48-A02`; **SEGM** = `48-A03`…`A06`.
3. **Agrupar piezas** por nombre hasta `_` (`SP-852_1` = `SP-852_2`).
4. **Nesteo** `Pxx_1`/`Pxx_2`: excluir ambas mitades como 'pared madre'; origen = placa completa.
5. Piezas `SP-*`, `SF-*`, `L845.*`, `PT-*`, `FT-*` son accesorios tipicos OTC (no casco).
6. Planos `50`–`60` suelen ser detalles/subarmados; no redefinir TOP/BASE/SEGM.

## 4. Por tanque

### 1991 - OTC 62176 TANK

- Proyecto: `62176`
- Planos: si (32 archivos)
- Codigos PDF: 46.00, 47.00, 47.01, 48.00, 49.00, 51.00, 52.00, 53.00, 54.00, 59.00, 59.01, 60.00
- STEP: `62176-1246-A01.STEP` (83.42 MB, 11 products)
- STEP: `62176-1246-A01 LIMPIO Y MARCADO.stp` (6.88 MB, 0 products)

### 1992 - OTC 62177 TANK

- Proyecto: `62177`
- Planos: si (29 archivos)
- Codigos PDF: 46.00, 47.00, 47.01, 48.00, 49.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62177-1246-A01 Step 3D.STEP` (4.33 MB, 119 products)
- Contenedores: `62177-1248-A04=SEGM_WALL, 62177-1247-A04=TOP, 62177-1247-A02=TOP, 62177-1247-A05=TOP, 62177-1248-A01=CASCO_SHELL, 62177-1248-A05=SEGM_WALL, 62177-1246-A01=TANQUE, 62177-1248-A03=SEGM_WALL, 62177-1247-A01=TOP, 62177-1247-A03=TOP, 62177-1248-A02=BASE, 62177-1248-A06=SEGM_WALL`
- Roles IAM: `{'SUB_A_CASCO': 7, 'PIEZA_CASCO': 34, 'PIEZA_TOP': 10, 'OTRO': 31, 'SEGM_WALL': 4, 'TOP': 5, 'CASCO_SHELL': 1, 'TANQUE': 1, 'BASE': 1}`

### 2024 - OTC 62178 TANK

- Proyecto: `62178`
- Planos: si (28 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 49.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62178-1246-A01 Step 3D.STEP` (9.24 MB, 61 products)
- Contenedores: `62178-1248-A06=SEGM_WALL, 62178-1247-A01=TOP, 62178-1247-A03=TOP, 62178-1248-A03=SEGM_WALL, 62178-1246-A01=TANQUE, 62178-1248-A05=SEGM_WALL`
- Roles IAM: `{'OTRO': 16, 'PIEZA_CASCO': 21, 'SUB_A_CASCO': 3, 'SEGM_WALL': 3, 'TOP': 2, 'PIEZA_TOP': 6, 'TANQUE': 1}`

### 2028 - OTC 62140 TANK

- Proyecto: `62140`
- Planos: si (8 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 51.00, 52.00, 59.00, 60.00
- STEP: `62140-1248-A01.STEP` (3.04 MB, 87 products)
- Contenedores: `62140-1248-A02=BASE, 62140-1248-A03=SEGM_WALL, 62140-1248-A01=CASCO_SHELL, 62140-1248-A05=SEGM_WALL, 62140-1248-A04=SEGM_WALL`
- Roles IAM: `{'PIEZA_CASCO': 55, 'SUB_A_CASCO': 16, 'BASE': 1, 'SEGM_WALL': 3, 'CASCO_SHELL': 1}`

### 2081 - OTC 62147 TANK

- Proyecto: `62147`
- Planos: si (31 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 49.00, 50.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62147-1246-A01.STEP` (25.09 MB, 24 products)
- Roles IAM: `{'OTRO': 6, 'PIEZA_CASCO': 7, 'SUB_A_CASCO': 2, 'PIEZA_TOP': 3}`

### 2082 - OTC 62154 TANK

- Proyecto: `62154`
- Planos: si (32 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 49.00, 50.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62154-1246-A01.STEP` (24.26 MB, 28 products)
- Contenedores: `62154-1247-A01=TOP`
- Roles IAM: `{'OTRO': 7, 'TOP': 1, 'PIEZA_TOP': 4, 'PIEZA_CASCO': 10, 'SUB_A_CASCO': 2}`

### 2098 - OTC 62175 TANK

- Proyecto: `62175`
- Planos: si (29 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 49.00, 51.00, 52.00, 53.00, 54.00, 59.00, 60.00
- STEP: `62175-1246-A01 R1.STEP` (8.56 MB, 67 products)
- Contenedores: `62175-1247-A01=TOP, 62175-1248-A04=SEGM_WALL, 62175-1248-A02=BASE, 62175-1248-A06=SEGM_WALL, 62175-1248-A01=CASCO_SHELL`
- Roles IAM: `{'PIEZA_CASCO': 22, 'PIEZA_TOP': 2, 'OTRO': 14, 'SUB_A_CASCO': 7, 'TOP': 1, 'SEGM_WALL': 2, 'BASE': 1, 'CASCO_SHELL': 1}`

### 2099 - OTC 62155 TANK

- Proyecto: `62155`
- Planos: si (31 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 49.00, 50.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62154-1246-A01.STEP` (24.26 MB, 28 products)
- Contenedores: `62154-1247-A01=TOP`
- Roles IAM: `{'OTRO': 7, 'TOP': 1, 'PIEZA_TOP': 4, 'PIEZA_CASCO': 10, 'SUB_A_CASCO': 2}`

### 2208 - OTC 62200 TANK

- Proyecto: `62200`
- Planos: si (11 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 48.01, 50.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62200-1246-A01 Step 3D.STEP` (4.75 MB, 131 products)
- Contenedores: `62200-1247-A05=TOP, 62200-1247-A01=TOP, 62200-1247-A04=TOP, 62200-1247-A06=TOP, 62200-1247-A02=TOP, 62200-1246-A01=TANQUE, 62200-1248-A03=SEGM_WALL, 62200-1247-A07=TOP, 62200-1248-A04=SEGM_WALL, 62200-1248-A02=BASE, 62200-1248-A01=CASCO_SHELL, 62200-1248-A06=SEGM_WALL`
- Roles IAM: `{'OTRO': 32, 'PIEZA_CASCO': 43, 'PIEZA_TOP': 14, 'TOP': 6, 'SUB_A_CASCO': 12, 'TANQUE': 1, 'SEGM_WALL': 3, 'BASE': 1, 'CASCO_SHELL': 1}`

### 2211 - OTC 62201 TANK

- Proyecto: `62201`
- Planos: si (26 archivos)
- Codigos PDF: 46.00, 47.00, 47.01, 48.00, 50.00, 51.00, 52.00, 53.00, 54.00, 60.00
- STEP: `62201-1246-A01 R1.STEP` (8.94 MB, 88 products)
- Contenedores: `62201-1248-A05=SEGM_WALL, 62201-1248-A01=CASCO_SHELL, 62201-1247-A01=TOP, 62201-1248-A03=SEGM_WALL`
- Roles IAM: `{'PIEZA_CASCO': 29, 'OTRO': 19, 'PIEZA_TOP': 12, 'SUB_A_CASCO': 8, 'SEGM_WALL': 2, 'CASCO_SHELL': 1, 'TOP': 1}`

### 2275 - OTC 62248 TANK

- Proyecto: `62248`
- Planos: si (21 archivos)
- Codigos PDF: 46.00, 47.00, 48.00, 51.00, 58.00, 59.00

## 5. Implicacion

No hardcodear solo `62201-1248-A0x`. Usar regex `(?P<proj>\d+)-(?P<fam>\d+)-(?P<tipo>[AP])(?P<num>\d+)` y rol = ultimos 2 digitos de `fam` ∈ {46,47,48}.
