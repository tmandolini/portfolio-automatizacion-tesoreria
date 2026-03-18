# Cash Flow Diario - Actualizacion Automatica

[Volver al Portfolio](../README.md)

---

## El Problema

Todos los dias, el area de Tesoreria necesita actualizar el Cash Flow con los movimientos reales de 4 bancos: **Macro, Santander, Nacion y Mercado Pago**. Este proceso implicaba:

- Abrir los extractos de cada banco
- Leer cada movimiento y clasificarlo manualmente en una categoria de cash flow
- Buscar la columna correcta del dia en la planilla
- Copiar los importes acumulados por categoria
- Actualizar los saldos bancarios
- Repetir esto todos los dias, con riesgo de errores humanos

**Tiempo estimado manual: ~45 minutos por dia.**

---

## La Solucion

Un sistema VBA que con **un solo clic** lee los 4 extractos bancarios, clasifica automaticamente cada transaccion usando una biblioteca de mapeo, acumula los importes por categoria, y escribe los resultados en la planilla de Cash Flow Diario.

### Arquitectura

```
Extracto Oficial (4 hojas de banco)
        |
   [Leer transacciones]
        |
   [Determinar fecha por predominancia]
        |
   [Cargar Biblioteca de Mapeo]
        |
   [Clasificar: TIPO + SUBTIPO -> Categoria Cash Flow]
        |
   [Acumular importes por categoria]
        |
   [Escribir en Cash Flow Diario + Saldos bancarios]
        |
   [Registrar en Log de auditoria]
```

### Componentes Principales

| Componente | Funcion |
|-----------|---------|
| **ActualizarCashFlow()** | Orquestador principal - coordina todo el proceso |
| **ObtenerFechaPorPredominancia()** | Algoritmo que determina la fecha correcta analizando frecuencias entre los 4 bancos |
| **CargarBibliotecaMapeo()** | Carga la tabla de mapeo (Tipo + Subtipo -> Categoria) en un diccionario para busquedas O(1) |
| **ProcesarHojaBanco()** | Procesa las transacciones de un banco individual |
| **EscribirEnCashFlow()** | Ubica la columna del dia y escribe los valores acumulados |
| **RegistrarLog()** | Registra cada ejecucion con timestamp, usuario y resultado |

---

## Desafios Tecnicos Resueltos

### 1. Determinacion inteligente de fecha
No siempre los 4 bancos reportan el mismo dia. El sistema usa un **algoritmo de predominancia**: cuenta la frecuencia de cada fecha en los 4 extractos y selecciona la mas comun, con desempate por fecha mas reciente.

### 2. Clasificacion automatica con wildcards
La biblioteca de mapeo soporta comodines (`*`) en el subtipo, permitiendo que una regla cubra multiples variantes de una transaccion sin necesidad de mapear cada caso individualmente.

### 3. Idempotencia
El proceso puede ejecutarse multiples veces para la misma fecha y siempre produce el mismo resultado, sin duplicar datos. Esto es critico para que el usuario pueda re-ejecutar sin miedo.

### 4. Performance
- Lectura masiva con arrays (no celda por celda)
- Diccionarios para busquedas O(1) en lugar de busquedas lineales
- Desactivacion de calculo automatico y actualizacion de pantalla durante la ejecucion
- Tiempo tipico: **2-5 segundos** para ~1.000 transacciones

---

## Datos Tecnicos

| Metrica | Valor |
|---------|-------|
| Lineas de codigo | ~800 |
| Bancos procesados | 4 (Macro, Santander, Nacion, Mercado Pago) |
| Categorias de Cash Flow | 33 fijas |
| Estructura de mapeo | Diccionario con clave INGRESO/EGRESO::TIPO::SUBTIPO |
| Tiempo de ejecucion | 2-5 segundos |
| Tecnologias | VBA, Scripting.Dictionary, Arrays |

---

## Resultado

- **Antes:** 45 minutos diarios de trabajo manual con riesgo de error
- **Despues:** 5 segundos con un clic, resultado auditado y consistente
- **Ahorro anual estimado:** +180 horas de trabajo manual

---

[Volver al Portfolio](../README.md)
