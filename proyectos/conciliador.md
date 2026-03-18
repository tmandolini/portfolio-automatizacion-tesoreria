# Conciliador ARCA - Conciliacion Automatica de Facturas

[Volver al Portfolio](../README.md)

---

## El Problema

Uno de los procesos mas criticos de Tesoreria: la **conciliacion de facturas** entre lo que reporta AFIP/ARCA (facturas recibidas segun el organismo fiscal) y lo que esta registrado en el sistema contable interno de Empresa (JDE).

Las discrepancias entre ambas fuentes pueden significar:
- Facturas registradas en AFIP que no estan en el sistema de la empresa
- Facturas en el sistema interno sin su correspondiente en AFIP
- Diferencias de importes entre ambos sistemas
- Proveedores sin CUIT cargado en el sistema

Este proceso se hacia manualmente con BUSCARV entre planillas enormes, tomando horas y con alto riesgo de error.

---

## La Solucion

El sistema mas complejo del portfolio: **7 modulos VBA, +2.600 lineas de codigo**. Automatiza la conciliacion completa incluyendo deteccion de duplicados, matching inteligente, gestion de pendientes, y envio de mails a proveedores.

### Arquitectura del Sistema

```
Datos ARCA (30+ columnas)     Datos Empresa (49 columnas)
        |                            |
   [ProcesarAFIP]              [ProcesarEmpresa]
   - Mapear tipos              - Buscar CUIT en BiblioCUIT
   - Aplicar tipo cambio       - Marcar "SIN CUIT" si falta
   - Excluir proveedores       - Filtrar solo PV/PD/PA
   - Detectar duplicados       - Excluir anuladas
        |                            |
        +------------+---------------+
                     |
              [PENDIENTES]
                     |
           [EjecutarConciliacion]
           - Matching por CUIT + Tipo + NroFactura
           - Tolerancia de $10 en importes
                     |
              +------+------+
              |             |
        [CONCILIADO]  [PENDIENTES]
         (matched)    (sin match)
                           |
                    [EnvioMailFinal]
                    - HTML con tabla de facturas
                    - Firma corporativa
                    - Tracking de envios
```

### Modulos del Sistema

| Modulo | Lineas | Responsabilidad |
|--------|--------|----------------|
| **Principal** | ~1.050 | Configuracion, validaciones, carga de bibliotecas, utilidades |
| **Conciliador** | ~520 | Algoritmo de matching, conciliacion manual, pendientes |
| **ARCA** | ~180 | Procesamiento de datos AFIP, mapeo de tipos de comprobante |
| **Empresa** | ~350 | Procesamiento de datos Empresa, busqueda de CUITs, actualizaciones |
| **Logs** | ~180 | Sistema de auditoria con acceso protegido por password |
| **Mails** | ~350 | Envio de emails HTML via Outlook con firma corporativa |
| **ThisWorkbook** | ~13 | Proteccion automatica al abrir el libro |

---

## Desafios Tecnicos Resueltos

### 1. Matching optimizado O(n) en lugar de O(n^2)
En lugar de comparar cada factura AFIP contra cada factura Empresa (lo que seria O(n^2)), el sistema construye un diccionario con clave `CUIT|TipoDoc|NroFact` y realiza busquedas O(1), reduciendo el tiempo de conciliacion de minutos a segundos.

### 2. Deteccion de duplicados en 3 niveles
Antes de agregar un registro a Pendientes, se verifica contra:
1. **Ejecucion actual** - duplicados dentro del mismo proceso
2. **Historial** - facturas ya conciliadas anteriormente
3. **Pendientes existentes** - facturas ya registradas como pendientes

Clave de duplicados: `ORIGEN|CUIT|TIPO|NROFACT|FECHA|IMPORTE`

### 3. Busqueda dinamica de columnas Empresa
El sistema Empresa exporta 49 columnas cuyas posiciones pueden variar. En lugar de hardcodear numeros de columna, el sistema busca por nombre de encabezado ("Tipo doc", "Descripcion numero proveedor", etc.), adaptandose automaticamente a cambios de formato.

### 4. Gestion de proveedores sin CUIT
Cuando un proveedor en Empresa no tiene CUIT cargado, el sistema:
- Lo marca como "SIN CUIT" en Pendientes
- Lo registra para revision
- Permite actualizar masivamente cuando se completa la BiblioCUIT

### 5. Emails HTML con firma corporativa
El modulo de mails:
- Agrupa facturas pendientes por proveedor
- Busca hasta 3 direcciones de email por proveedor en la BiblioCUIT
- Genera HTML con tabla formateada de facturas
- Exporta la firma desde una hoja del libro como imagen PNG
- Envia via Outlook y marca las filas enviadas con timestamp

### 6. Proteccion de datos
- Hojas protegidas con password al abrir el libro
- Estructura del libro protegida (no se pueden eliminar/renombrar hojas)
- Logs ocultos (xlVeryHidden) accesibles solo con password

---

## Hojas del Sistema

| Hoja | Funcion |
|------|---------|
| **PegarAFIP** | Area de pegado para datos de ARCA (30+ columnas) |
| **PegarEmpresa** | Area de pegado para datos de Empresa (49 columnas) |
| **Pendientes** | Facturas sin conciliar (14 columnas) |
| **Conciliado** | Facturas matcheadas (18 columnas) |
| **BiblioCUIT** | Base de datos de proveedores (NroProv, CUIT, emails) |
| **ProveedoresBarrer** | Lista de exclusion (CUITs a ignorar) |
| **Firmas** | Imagen de firma para emails |
| **_Logs** | Auditoria oculta |

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 7 |
| Lineas de codigo | +2.600 |
| Columnas fuente ARCA | 30+ |
| Columnas fuente Empresa | 49 |
| Tolerancia de matching | $10 ARS |
| Niveles de deteccion duplicados | 3 |
| Algoritmo de matching | Dictionary O(n) |
| Integracion email | Outlook COM via HTML |
| Tecnologias | VBA, Scripting.Dictionary, Outlook COM, HTML, FileSystemObject |

---

## Resultado

- **Antes:** Horas de trabajo manual con BUSCARV entre planillas de miles de filas, con alto riesgo de omitir facturas
- **Despues:** Conciliacion completa en segundos, con tracking de pendientes y notificacion automatica a proveedores
- **Beneficio:** Cumplimiento fiscal, control de proveedores, auditoria completa, reduccion drastica del riesgo operativo

---

[Volver al Portfolio](../README.md)
