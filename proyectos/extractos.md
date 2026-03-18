# Extractos Bancarios - Importacion y Clasificacion Automatica

[Volver al Portfolio](../README.md)

---

## El Problema

La empresa recibe diariamente extractos bancarios de multiples fuentes:
- **3 bancos via Interbanking** (Macro, Santander, Nacion) en formato Excel
- **Mercado Pago** en formato propio con comisiones e impuestos separados

El proceso manual implicaba:
- Abrir cada archivo, identificar de que banco es
- Copiar las transacciones a la hoja correcta
- Clasificar cada movimiento segun su tipo (proveedor, impuesto, transferencia, etc.)
- Calcular comisiones e IVA de Mercado Pago por separado
- Verificar que no haya duplicados con dias anteriores
- Mover los archivos procesados a otra carpeta

---

## La Solucion

Un sistema de 7 modulos VBA que automatiza todo el ciclo: importacion, clasificacion, seleccion manual de ambiguos, y actualizacion de saldos.

### Arquitectura del Sistema

```
Carpeta "Entrada" (archivos de bancos)
        |
   [Detectar tipo: Interbanking o Mercado Pago]
        |
   +--> Interbanking: Identificar banco por nro. cuenta
   |    [Mapear columnas -> hoja MACRO/SANTANDER/NACION]
   |
   +--> Mercado Pago: Separar en 4 lineas
        [Principal + Comision + IVA + Impuestos debito]
        |
   [Detectar duplicados via Dictionary O(1)]
        |
   [Clasificar contra Biblioteca de Codigos]
        |
   +--> Codigo unico: clasificacion automatica
   +--> Codigo ambiguo: dropdown para seleccion manual
        |
   [Bloquear dia anterior / Formatear]
        |
   [Mover archivos a "Procesados"]
```

### Modulos del Sistema

| Modulo | Responsabilidad |
|--------|----------------|
| **Orquestador** | Coordina todo el flujo: importar, clasificar, mover archivos |
| **Importador** | Lee archivos, detecta banco, mapea columnas, previene duplicados |
| **Clasificacion** | Busca codigos en biblioteca, asigna tipo/subtipo automaticamente |
| **ClasificacionMP** | Logica separada para Mercado Pago (deteccion de movimientos nuevos diferente) |
| **Selector** | Genera dropdowns con Data Validation para codigos ambiguos |
| **Logs** | Sistema de auditoria con registro de cada operacion |
| **ThisWorkbook** | Evento de hoja: cuando el usuario elige un Tipo, actualiza Subtipo automaticamente |

---

## Desafios Tecnicos Resueltos

### 1. Mercado Pago: una transaccion, cuatro lineas
Cada venta de Mercado Pago genera comisiones, IVA sobre la comision, e impuestos de debito. El sistema desglosa automaticamente una transaccion en hasta 4 lineas contables separadas, calculando el IVA al 21%.

### 2. Deteccion de duplicados de alto rendimiento
Usa `Scripting.Dictionary` con clave compuesta (fecha|comprobante|importe) para deteccion O(1). Esto permite procesar miles de transacciones sin degradacion de performance.

### 3. Clasificacion hibrida (automatica + manual)
Los codigos con una sola clasificacion posible se asignan automaticamente. Los codigos ambiguos (multiples tipos posibles) generan dropdowns filtrados para que el usuario elija. El sistema actualiza el subtipo automaticamente cuando se selecciona el tipo.

### 4. Separacion Interbanking vs Mercado Pago
La deteccion de movimientos nuevos es diferente en cada caso: Interbanking usa la columna de fecha vs banco; Mercado Pago usa la columna de saldo. Esta diferencia justifica modulos separados en lugar de condicionales complejos.

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Modulos VBA | 7 |
| Bancos soportados | 4 (Macro, Santander, Nacion, Mercado Pago) |
| Deteccion de duplicados | Dictionary O(1) con clave compuesta |
| Formato de fechas | Parsing custom DD/MM/YYYY (evita inversion MM/DD de VBA) |
| Manejo de archivos | Auto-deteccion, procesamiento y archivado con timestamp |
| Tecnologias | VBA, Scripting.Dictionary, Data Validation, Eventos de Worksheet |

---

## Resultado

- **Antes:** Proceso manual propenso a errores, especialmente en la clasificacion y el desglose de Mercado Pago
- **Despues:** Importacion y clasificacion automatica con un clic, dropdowns inteligentes para casos ambiguos
- **Volumen:** +1.000 transacciones diarias procesadas de forma consistente

---

[Volver al Portfolio](../README.md)
