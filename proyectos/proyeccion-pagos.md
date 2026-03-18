# Proyeccion de Pagos - Calendario de Pagos a Proveedores

[Volver al Portfolio](../README.md)

---

## El Problema

La empresa gestiona pagos a decenas de proveedores con diferentes condiciones de pago. El area de Tesoreria necesitaba:

- Saber cuanto se debe pagar cada semana y cada mes
- Identificar que facturas estan vencidas y cuales estan proximas a vencer
- Cruzar las facturas de ARCA (AFIP) con los pagos registrados en JDE (sistema contable)
- Excluir proveedores que se pagan por otro circuito
- Tener visibilidad del proximo lote de pagos con detalle

Todo esto se hacia con filtros manuales, BUSCARV, y mucha revision visual.

---

## La Solucion

Un sistema de 6 modulos VBA que procesa facturas de ARCA, las cruza con pagos de JDE, calcula fechas de vencimiento segun condiciones de pago por proveedor, y genera 3 calendarios automaticos.

### Flujo del Proceso

```
Reporte ARCA (CSV)          Resumen Pagos JDE (Excel)
        |                           |
   [Cargar bibliotecas]             |
   [Condiciones de pago]            |
   [Lista de exclusion]             |
        |                           |
   [Procesar facturas ARCA]         |
   [Mapear tipos: PV/PD/PA]        |
   [Calcular fecha vencimiento]     |
   [Excluir proveedores barrido]    |
        |                           |
        +--------+------------------+
                 |
        [Cruzar con JDE]
        [Marcar PAGADAS en verde]
                 |
        +--------+--------+--------+
        |        |        |        |
   [Proximo  [Calendario [Calendario
    Lote]     Semanal]    Mensual]
```

### Modulos del Sistema

| Modulo | Funcion |
|--------|---------|
| **Orquestador** | Macro principal, validaciones, limpieza, formateo con colores corporativos |
| **ARCA** | Procesa facturas de AFIP: mapeo de tipos, calculo de vencimientos, exclusiones |
| **PagosJDE** | Cruza con pagos de JDE: marca facturas pagadas, actualiza fechas de pago real |
| **Calendarios** | Genera 3 vistas: Proximo Lote, Calendario Semanal (5 semanas), Calendario Mensual (3 meses) |
| **General** | Funciones auxiliares compartidas |
| **Logs** | Sistema de auditoria |

---

## Desafios Tecnicos Resueltos

### 1. Mapeo inteligente de tipos de comprobante
ARCA usa codigos numericos (1, 2, 3, 6, 11, 12, 13) que el sistema traduce a codigos internos:
- Codigos 1, 11 -> PV (Factura)
- Codigos 2, 12 -> PD (Nota de Debito)
- Codigos 3, 13 -> PA (Nota de Credito)

Las Notas de Credito se registran con importe negativo automaticamente.

### 2. Calculo de vencimientos con condiciones por proveedor
Cada proveedor tiene dias de pago diferentes (30, 45, 60, 90 dias). El sistema:
- Busca el CUIT del proveedor en la biblioteca de condiciones
- Calcula: Fecha Vencimiento = Fecha Emision + Dias de Pago
- Los pagos siempre caen en Lunes (dia de pago de la empresa)

### 3. Cruce JDE por ultimos 6 digitos
El numero de factura en ARCA y JDE tiene formatos diferentes. El sistema matchea por los ultimos 6 digitos del numero + codigo de proveedor, tolerando las diferencias de formato.

### 4. Procesamiento incremental
El sistema no borra datos anteriores: agrega facturas nuevas, actualiza el estado de las existentes (marca como PAGADA si aparece en JDE), y regenera los calendarios.

### 5. Flagging visual de problemas
- **Rojo:** Facturas sin condicion de pago o tipo de comprobante no mapeado
- **Verde + Negrita:** Facturas pagadas
- **Naranja:** Proximo lote de pagos
- **Azul:** Pagos futuros

---

## Salidas Generadas

### 1. Proximo Lote
- Fecha del proximo lunes de pago
- Todas las facturas vencidas + las que vencen antes de ese lunes
- Dias de atraso por factura
- Total a pagar y cantidad de facturas

### 2. Calendario Semanal (5 semanas)
- Una columna por semana con fecha de lunes de pago
- Facturas vencidas incluidas en la primera semana
- Cantidad de facturas y total por semana
- Estados visuales: PROXIMO (naranja) y FUTURO (azul)

### 3. Calendario Mensual (3 meses)
- Seccion HISTORICO: facturas ya pagadas agrupadas por mes
- Seccion PROYECCION: facturas pendientes agrupadas por mes
- Mes actual con detalle, meses futuros agregados por proveedor

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 6 |
| Tipos de comprobante | 7 codigos ARCA -> 3 tipos internos |
| Calendarios generados | 3 (Lote, Semanal, Mensual) |
| Prevencion duplicados | Clave CUIT + TipoDoc + NroComprobante |
| Multi-moneda | Si (conversion USD -> ARS si aplica) |
| Tecnologias | VBA, Scripting.Dictionary, Arrays, CSV parsing |

---

## Resultado

- **Antes:** Horas de revision manual con filtros y BUSCARV para armar la proyeccion
- **Despues:** Proyeccion completa con 3 calendarios automaticos en segundos
- **Beneficio:** Visibilidad total de compromisos de pago, deteccion temprana de vencimientos, planificacion financiera precisa

---

[Volver al Portfolio](../README.md)
