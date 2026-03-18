# Tamara Mandolini - Portfolio de Automatizacion en Tesoreria

## Sobre mi

**Tamara Mandolini** | Responsable de Tesoreria en **empresa multinacional de consumo masivo**

Administradora de Empresas especializada en optimizacion de procesos financieros mediante automatizacion con VBA (Excel) y herramientas de Business Intelligence.

Mi enfoque: transformar tareas manuales y repetitivas del area de Tesoreria en sistemas automatizados que ahorran tiempo, eliminan errores y generan informacion en tiempo real para la toma de decisiones.

---

## Ecosistema de Automatizacion

Estos no son proyectos aislados: forman un **ecosistema integrado** donde los datos fluyen de un sistema a otro, cubriendo el ciclo completo de la operacion de Tesoreria.

```
                    OPERACION DIARIA
                         |
    Extractos Bancarios --+--> Cash Flow Diario --+
    (importacion y             (actualizacion      |
     clasificacion              automatica)        |
     de movimientos)                               +--> Liquidity Report
                                                   |    (reporte semanal
    Inversiones ----------------------------------+     corporativo)
    (tracking de portfolio
     multi-plataforma)

                    GESTION DE PAGOS
                         |
    Proyeccion de Pagos --+--> Dashboard Power BI
    (calendarios y              (visualizacion
     vencimientos)               ejecutiva)

    Conciliador ARCA           Deuda Intercompany
    (conciliacion fiscal       (cuentas por pagar
     automatica)                multi-moneda)
```

---

## Los Sistemas

### Extractos Bancarios
Importa y clasifica automaticamente los movimientos diarios de multiples bancos. Detecta duplicados, separa comisiones e impuestos, y permite clasificacion manual de casos ambiguos con dropdowns inteligentes. Es la puerta de entrada de datos al ecosistema.
**[Ver detalle](proyectos/extractos.md)**

### Cash Flow Diario
Con un clic, lee los extractos de 4 bancos, clasifica cada transaccion usando una biblioteca de mapeo, y actualiza el cash flow del dia. Determina la fecha correcta con un algoritmo de predominancia entre bancos.
**[Ver detalle](proyectos/cashflow.md)**

### Inversiones
Tracking semanal del portfolio de inversiones en multiples plataformas. Importa saldos y movimientos, detecta duplicados entre ejecuciones, y calcula ganancia/perdida semanal con formulas dinamicas.
**[Ver detalle](proyectos/inversiones.md)**

### Liquidity Report
Consolida datos de Cash Flow e Inversiones para generar automaticamente el reporte de liquidez semanal corporativo. Maneja 3 monedas y convierte valores al formato requerido.
**[Ver detalle](proyectos/liquidity.md)**

### Proyeccion de Pagos
Procesa facturas de ARCA, las cruza con pagos del sistema contable, calcula fechas de vencimiento por proveedor, y genera 3 calendarios automaticos: proximo lote, semanal y mensual.
**[Ver detalle](proyectos/proyeccion-pagos.md)**

### Conciliador ARCA
El sistema mas complejo del ecosistema (+2.600 lineas, 7 modulos). Concilia automaticamente facturas entre el organismo fiscal y el sistema contable interno. Incluye deteccion de duplicados en 3 niveles, matching inteligente, y envio de emails a proveedores.
**[Ver detalle](proyectos/conciliador.md)**

### Deuda Intercompany
Gestiona cuentas por pagar de proveedores internacionales en 3 monedas (ARS, USD, EUR). Clasifica por tipo y moneda, archiva facturas pagadas, y calcula vencimientos segun condiciones de pago.
**[Ver detalle](proyectos/reporte-ico.md)**

### Dashboard Power BI
Transforma los datos de la proyeccion de pagos en un dashboard interactivo con KPIs, ranking de proveedores, proyeccion semanal y tabla de facturas pendientes. Usa un tema visual con los colores corporativos.
**[Ver detalle](proyectos/powerbi-dashboard.md)**

---

## Stack Tecnologico

- **VBA (Excel)** - Desarrollo de macros y sistemas completos de automatizacion
- **Power BI** - Dashboards y visualizacion de datos financieros
- **Power Query (M)** - Transformacion y limpieza de datos
- **Git** - Control de versiones del codigo

## Metodologia de Trabajo

- Codigo modular con separacion clara de responsabilidades
- Sistema de logs y auditoria en cada proyecto
- Manejo robusto de errores con mensajes claros al usuario
- Optimizacion de rendimiento (arrays, diccionarios, procesamiento en lote)
- Estandares de nomenclatura y documentacion consistentes
- Prevencion de duplicados en todas las importaciones de datos

## Contacto

- **Email:** tamaramandolini@hotmail.com
- **LinkedIn:** [Tamara Mandolini](https://www.linkedin.com/in/tamaramandolini/)
- **GitHub:** [tmandolini](https://github.com/tmandolini)
