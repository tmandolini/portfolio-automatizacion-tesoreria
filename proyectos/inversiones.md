# Inversiones - Tracking de Portfolio Multi-Plataforma

[Volver al Portfolio](../README.md)

---

## El Problema

La empresa mantiene inversiones en multiples plataformas:
- **Balanz**: FCI (BCAHA), Obligaciones Negociables (PN35O), USD MEP
- **Santander**: Super Ahorro Plus A, Super Ahorro $

Cada semana se necesita:
- Descargar el resumen de cuenta de Balanz (PDF exportado a TXT)
- Descargar movimientos de Balanz y Santander (Excel)
- Registrar los saldos actuales de cada posicion
- Calcular la ganancia/perdida semanal de cada inversion
- Importar las transacciones nuevas sin duplicar las existentes

Todo esto implicaba abrir multiples archivos, copiar datos manualmente, y calcular formulas a mano.

---

## La Solucion

Un sistema de 5 modulos VBA con un **orquestador central** que ejecuta todo el proceso semanal con un solo clic: importa saldos, importa movimientos de ambas plataformas, y calcula ganancia/perdida.

### Flujo del Proceso Semanal

```
1. Descargar archivos de Balanz y Santander (manual)
2. Ejecutar ActualizarMovimientos() (un clic)
        |
   [Importar saldos desde Balanz TXT]
   [Pedir saldos Santander via InputBox]
        |
   [Importar movimientos Balanz (Excel)]
   [Importar movimientos Santander (Excel)]
        |
   [Calcular Ganancia/Perdida semanal]
   [Generar formulas SUMIFS dinamicas]
        |
   [Resumen final con metricas]
```

### Modulos del Sistema

| Modulo | Lineas | Funcion |
|--------|--------|---------|
| **Orquestador** | 77 | Coordina los 3 pasos, modo silencioso, resumen final |
| **ImportarBalanz** | 465 | Lee TXT de Balanz, extrae USD MEP + saldos FCI + ON |
| **ImportarMovimientosBalanz** | 294 | Importa transacciones de Balanz, detecta duplicados |
| **ImportarMovimientosSantander** | 325 | Importa transacciones de Santander, detecta duplicados |
| **ActualizarGananciaPerdida** | 454 | Calcula P&L con formulas SUMIFS dinamicas |

---

## Desafios Tecnicos Resueltos

### 1. Parsing de texto semi-estructurado
El resumen de Balanz viene como TXT exportado desde PDF. El sistema extrae valores especificos (USD MEP, saldo BCAHA, saldo PN35O) parseando el texto con busqueda de patrones, manejando el formato numerico argentino (1.234.567,89).

### 2. Formulas SUMIFS dinamicas
En lugar de valores fijos, el calculo de ganancia/perdida genera formulas SUMIFS que filtran por rango de fechas y ticker. Esto permite que las formulas se recalculen si cambian los datos subyacentes:
```
G/P = Saldo Actual - Saldo Anterior + Rescates - Suscripciones
```

### 3. Deteccion de duplicados cross-plataforma
Cada plataforma usa una clave compuesta diferente para prevenir duplicados:
- **Balanz:** Concertacion|Ticker|Importe
- **Santander:** FechaSolicitud|NroComprobante|Importe

### 4. Busqueda inteligente de archivos
El sistema busca automaticamente el archivo mas reciente en la carpeta de Downloads, priorizando el que coincide con la fecha actual y fallback al mas reciente por fecha de modificacion.

---

## Posiciones Trackeadas

| Posicion | Plataforma | Moneda | Tipo |
|----------|-----------|--------|------|
| USD MEP | Balanz | USD | Tipo de cambio |
| Capital Ahorro FCI (BCAHA) | Balanz | ARS | Fondo Comun de Inversion |
| ON PAN AMERICAN (PN35O) | Balanz | ARS | Obligacion Negociable |
| Super Ahorro Plus A | Santander | ARS | Fondo mutuo |
| Super Ahorro $ | Santander | ARS | Fondo mutuo |

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 5 |
| Lineas de codigo | ~1.615 |
| Plataformas integradas | 2 (Balanz + Santander) |
| Posiciones trackeadas | 5 |
| Deteccion duplicados | Dictionary O(1) por plataforma |
| Formulas generadas | SUMIFS dinamicas con rangos de fecha |
| Tecnologias | VBA, Scripting.Dictionary, Regex, Arrays |

---

## Resultado

- **Antes:** Proceso manual semanal de ~30 minutos con riesgo de error en formulas
- **Despues:** Ejecucion automatizada en segundos, formulas dinamicas, historial completo de movimientos
- **Beneficio:** Vision unificada de todas las inversiones con P&L semanal automatico

---

[Volver al Portfolio](../README.md)
