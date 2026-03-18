# Deuda Intercompany - Gestion de Cuentas por Pagar Multi-Moneda

[Volver al Portfolio](../README.md)

---

## El Problema

La empresa recibe facturas de proveedores internacionales en tres monedas diferentes (ARS, USD, EUR), clasificadas en dos tipos:
- **Mercaderia** (facturas tipo PL - Liquidacion)
- **Servicios** (facturas tipo PV - Factura de Venta, en USD o EUR)

El equipo necesitaba:
- Importar nuevas facturas desde un reporte del sistema contable
- Separar automaticamente por tipo y moneda
- Archivar las facturas ya pagadas sin perder el historial
- Calcular fechas de vencimiento segun condiciones de pago por proveedor
- Identificar rapidamente facturas que ya no aparecen en el sistema (posibles pagos)

---

## La Solucion

Un sistema modular de 5 modulos VBA que gestiona el ciclo completo de cuentas por pagar: importacion, clasificacion, calculo de vencimientos, archivado y validacion cruzada.

### Flujo del Proceso

```
Reporte del sistema contable
        |
   [Pegar en hoja "Pegar"]
        |
   [CopiarDatos() - Macro principal]
        |
   +----+----+----+
   |    |    |    |
   | Archivar facturas PAGADAS -> Historico
   |    |    |    |
   | Importar nuevas facturas
   | [Clasificar por tipo de documento y moneda]
   |    |    |    |
   | Validar cruce (marcar en rojo si no aparece en fuente)
   |    |    |    |
   | Calcular fechas de vencimiento
   |    |    |    |
   | Formatear hojas (fuente Calibri 10, colores corporativos)
```

### Modulos del Sistema

| Modulo | Funcion |
|--------|---------|
| **General** | Orquestador principal, constantes corporativas, validaciones, funciones compartidas |
| **Mercaderia** | Importa, archiva y valida facturas tipo PL |
| **ServiciosUSD** | Importa, archiva y valida facturas de servicios en USD |
| **ServiciosEUR** | Importa, archiva y valida facturas de servicios en EUR |
| **Logs** | Sistema de auditoria con registros por tipo (INFO/WARNING/ERROR) |

---

## Desafios Tecnicos Resueltos

### 1. Clasificacion automatica por tipo y moneda
El sistema lee el tipo de documento y la moneda de cada factura para enviarla automaticamente a la hoja correcta, separando mercaderia de servicios y agrupando por moneda.

### 2. Archivado inteligente
Las facturas marcadas como "PAGADA" se mueven automaticamente a la hoja Historico, preservando todos sus datos. Esto mantiene las hojas de trabajo limpias sin perder trazabilidad.

### 3. Validacion cruzada (deteccion de pagos)
Despues de importar los datos nuevos, el sistema compara cada factura existente contra la fuente. Las que **no aparecen** en el nuevo reporte se marcan en rojo, indicando que posiblemente fueron pagadas o eliminadas del sistema.

### 4. Condiciones de pago dinamicas
Cada proveedor tiene sus propios terminos de pago. El sistema busca en la tabla "Cond de Pago" y calcula:
```
Fecha Vencimiento = Fecha Factura + Dias de Pago del Proveedor
```

---

## Estructura de Hojas

| Hoja | Contenido |
|------|-----------|
| **Pegar** | Area de pegado para datos del sistema contable |
| **Mercaderia** | Facturas activas de mercaderia (14 columnas) |
| **Servicios USD** | Facturas activas de servicios en dolares (11 columnas) |
| **Servicios EUR** | Facturas activas de servicios en euros (11 columnas) |
| **Historico** | Archivo de facturas pagadas (todas las monedas) |
| **Cond de Pago** | Tabla de condiciones de pago por proveedor |
| **_Logs** | Auditoria oculta |

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 5 |
| Monedas soportadas | 3 (ARS, USD, EUR) |
| Tipos de documento | 2 |
| Logs por tipo | 3 (INFO azul, WARNING amarillo, ERROR rojo) |
| Prevencion duplicados | Verificacion antes de importar |
| Tecnologias | VBA, manejo de ListObjects, formateo condicional |

---

## Resultado

- **Antes:** Gestion manual de facturas en multiples monedas, riesgo de perder track de pagos
- **Despues:** Sistema unificado con archivado automatico, vencimientos calculados, y alertas visuales
- **Beneficio:** Control completo de cuentas por pagar internacionales en un solo lugar

---

[Volver al Portfolio](../README.md)
