# Dashboard Power BI - Visualizacion de Proyeccion de Pagos

[Volver al Portfolio](../README.md)

---

## El Problema

La proyeccion de pagos a proveedores genera datos valiosos, pero en formato Excel plano no es facil:
- Ver tendencias de pago a lo largo del tiempo
- Identificar rapidamente cuanto hay pendiente y cuanto esta vencido
- Saber cuales son los proveedores con mayor concentracion de deuda
- Presentar la informacion a gerencia de forma clara e interactiva

---

## La Solucion

Un dashboard interactivo en Power BI que toma los datos de la proyeccion de pagos y los transforma en visualizaciones ejecutivas, con identidad visual corporativa.

### Componentes del Proyecto

| Componente | Tecnologia | Funcion |
|-----------|-----------|---------|
| **Scripts de transformacion** | Power Query (M) | Limpieza y tipado de datos en Power BI |
| **Tema corporativo** | JSON | Paleta de colores y tipografia corporativa |
| **Dashboard** | Power BI | Visualizaciones interactivas |

### Visualizaciones del Dashboard

```
+------------------------------------------+
|  DASHBOARD PROYECCION DE PAGOS           |
+----------+----------+----------+---------+
| Total    | Monto    | Cantidad | % Venc. |
| Pendiente| Vencido  | Facturas | del Tot.|
| (KPI)    | (KPI)    | (KPI)   | (KPI)   |
+----------+----------+----------+---------+
|                     |                    |
|  Top 10 Proveedores |  Proyeccion       |
|  (Barras horiz.)    |  Semanal          |
|                     |  (Columnas)       |
|                     |                    |
+---------------------+--------------------+
|                                          |
|  Tabla de Facturas Pendientes            |
|  (interactiva, con filtros y busqueda)   |
|                                          |
+------------------------------------------+
```

---

## Pipeline de Datos

```
Proyeccion Pagos (Excel/VBA)
        |
   [Exportacion a CSV]
   - Detalle de facturas
   - Dias de atraso calculados
   - 4 archivos CSV
        |
   +----+----+----+----+
   |    |    |    |    |
   Detalle  Calendario  Facturas  Top 10
   Proveed. Semanal     Pendient. Proveed.
   (3.015   (semanal)   (301     (ranking)
   facturas)            pend.)
        |    |    |    |
   [Power Query - Transformacion]
   - Tipado de columnas
   - Formato de fechas
   - Limpieza de datos
        |
   [DAX - Medidas calculadas]
   - Total pendiente
   - Monto vencido
   - % vencido sobre total
        |
   [Dashboard Power BI]
```

### Archivos CSV Generados

| Archivo | Registros | Contenido |
|---------|-----------|-----------|
| **Detalle_Proveedores.csv** | 3.015 | Todas las facturas con estado |
| **Calendario_Semanal.csv** | Semanal | Fechas de pago y montos por semana |
| **Facturas_Pendientes.csv** | 301 | Facturas sin pagar con dias de atraso |
| **Top10_Proveedores.csv** | 10 | Ranking de proveedores por monto |

---

## Identidad Visual

El dashboard utiliza un tema personalizado con los colores corporativos de la empresa, aplicado a todos los elementos visuales (titulos, KPIs, graficos y estados).

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Fuentes de datos | 4 archivos CSV |
| Registros totales | +3.000 facturas |
| Visualizaciones | 4 (KPIs, barras, columnas, tabla interactiva) |
| Medidas DAX | Total pendiente, vencido, % vencido, conteo |
| Transformacion | Power Query (M) |
| Tema visual | JSON personalizado corporativo |

---

## Resultado

- **Antes:** Datos en Excel plano, dificil de presentar a gerencia
- **Despues:** Dashboard interactivo con filtros, busqueda y KPIs en tiempo real
- **Beneficio:** Visibilidad ejecutiva inmediata de la posicion de pagos, presentable a cualquier nivel de la organizacion

---

[Volver al Portfolio](../README.md)
