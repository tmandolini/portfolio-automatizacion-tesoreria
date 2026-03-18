# Liquidity Report - Generacion Automatica de Reportes de Liquidez

[Volver al Portfolio](../README.md)

---

## El Problema

Cada semana, Tesoreria debe completar un reporte de liquidez corporativo que consolida:
- **Saldos bancarios** de 8 cuentas en 3 monedas (ARS, USD, EUR)
- **Movimientos semanales** por categoria (Clientes, Proveedores, Sueldos, etc.)
- **Inversiones de corto plazo** (Balanz, Santander)
- **Resultados financieros** (ganancias/perdidas de inversiones)

El reporte tiene formato fijo y se entrega como archivo Excel estandarizado. Completarlo manualmente requeria abrir 3 archivos fuente diferentes, buscar la semana correcta, copiar valores, y convertir a miles.

---

## La Solucion

Dos modulos VBA que completan automaticamente el Liquidity Report extrayendo datos del Cash Flow y del libro de Inversiones.

### Flujo del Proceso

```
Cash Flow 2026 (saldos + movimientos)
        |
   [Buscar semana actual en Cash Flow]
   [Buscar ultimo Liquidity Report en carpeta]
        |
   [Completar saldos de 8 bancos en 3 monedas]
   [Completar movimientos semanales (5 categorias)]
        |
   [Guardar Liquidity Report]

Inversiones 2026 (portfolio)
        |
   [Buscar semana actual en Inversiones]
        |
   [Completar inversiones: Balanz ARS/USD + Santander ARS]
   [Completar resultado financiero (ganancia/perdida)]
        |
   [Guardar Liquidity Report]
```

### Modulos del Sistema

| Modulo | Lineas | Funcion |
|--------|--------|---------|
| **LiquidityReport** | 433 | Completa saldos bancarios y movimientos desde Cash Flow |
| **LiquidityInversiones** | 340 | Completa inversiones y resultado financiero |

---

## Desafios Tecnicos Resueltos

### 1. Deteccion dinamica del archivo correcto
El sistema busca automaticamente el Liquidity Report mas reciente en la carpeta, extrayendo el numero de semana del nombre del archivo (`Liquidity report Argentina_ wX-2026.xlsx`).

### 2. Busqueda dinamica de filas por nombre
En lugar de hardcodear posiciones, el sistema busca las filas por nombre de banco + moneda (ej: "Balanz" + "ARS"), lo que permite que el reporte cambie de estructura sin romper el codigo.

### 3. Conversion a miles
Todos los valores se dividen por 1.000 automaticamente, que es el formato requerido por el reporte corporativo.

### 4. Multi-moneda
El sistema maneja 3 monedas simultneamente:
- **ARS:** Nacion, Santander, Mercado Pago, Macro, Caja
- **USD:** Itau Nassau, Macro
- **EUR:** Itau Nassau

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 2 |
| Lineas de codigo | ~773 |
| Cuentas bancarias | 8 en 3 monedas |
| Archivos fuente | 3 (Cash Flow, Inversiones, Liquidity Report) |
| Formato de salida | Excel corporativo estandarizado |
| Tecnologias | VBA, manejo de Workbooks multiples |

---

## Resultado

- **Antes:** ~20 minutos semanales copiando datos entre 3 archivos
- **Despues:** Reporte completo en 1 clic, sin errores de transcripcion
- **Beneficio:** Reporte corporativo siempre consistente y auditable

---

[Volver al Portfolio](../README.md)
