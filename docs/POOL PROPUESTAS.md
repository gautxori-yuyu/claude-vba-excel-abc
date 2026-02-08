# POOL DE PROPUESTAS DE MEJORA

> **VERSIÓN:** 1.1 (Corregida)  
> **FECHA:** 2026-01-13  
> **ROL:** Catálogo vivo de propuestas de mejora

---

## 🎯 PROPÓSITO

Este documento es un **catálogo vivo** de propuestas de mejora identificadas durante el análisis del código. 

**NO es un plan de trabajo** - es un **brainstorming estructurado**.

### Flujo de trabajo:

````
1. Propuesta identificada → AÑADIR aquí primero
   ↓
2. EVALUAR técnicamente (factibilidad, impacto, esfuerzo)
   ↓
3. PRIORIZAR según criterios definidos
   ↓
4. Propuesta aprobada → INCORPORAR a CLAUDE.md (refactor)
````

---

# INDICE GENERAL

1. Índice numerado de todas las propuestas
2. Para cada propuesta:
   - ID único
   - Nombre
   - Problema que resuelve
   - Patrón de diseño aplicado
   - Dónde se implementa
   - Prioridad
   - Estado (pendiente/en progreso/completado)
3. Referencias cruzadas a CLAUDE.md
4. EVALUACION de propuestas
---


## 📊 Criterios de Evaluación

Cada propuesta se evalúa en 5 dimensiones:

### 1. Factibilidad Técnica (0-5)

- **5:** Fácil de implementar, sin riesgos, bajo acoplamiento
- **4:** Requiere cambios en 1-2 clases
- **3:** Requiere cambios moderados en 3-5 clases
- **2:** Requiere refactorización significativa
- **1:** Muy complejo, alto riesgo, muchas dependencias

### 2. Impacto en Calidad (0-5)

- **5:** Resuelve problema crítico (God Object, Circular Dependency)
- **4:** Mejora arquitectural significativa
- **3:** Mejora notable pero no crítica
- **2:** Mejora menor
- **1:** Mejora marginal o cosmética

### 3. Esfuerzo (horas estimadas)

- **Bajo:** < 4 horas
- **Medio:** 4-16 horas
- **Alto:** > 16 horas

### 4. Dependencias

- **Ninguna:** Independiente, puede implementarse ahora
- **Baja:** Requiere 1 otra propuesta
- **Media:** Requiere 2-3 propuestas
- **Alta:** Requiere >3 propuestas o cambios mayores

### 5. Orden de Implementación Sugerido

- **Bottom-Up:** Infraestructura primero (Ej: ServiceManager)
- **Top-Down:** Fachada/UI primero (Ej: Ribbon)
- **Lateral:** Independiente, cualquier momento

---


## 🗂️ CATÁLOGO DE PROPUESTAS

> **NOTA:** Las propuestas a continuación son SOLO EJEMPLOS para ilustrar el formato.  
> Las propuestas reales deben ser añadidas durante el análisis del código.

---

### PROP-001: Interfaces de Escucha (Sink Interfaces)

**Estado:** 🟡 Pendiente evaluación

**Categoría:** Arquitectura / Eventos

**Problema que resuelve:**

EventCoordinator tiene `WithEvents` específico para cada servicio. Cada nuevo servicio requiere modificar EventCoordinator (violación OCP).
```vba
Private WithEvents mRibbonState As clsRibbonState
Private WithEvents mOpportunities As clsOpportunitiesMgr
Private WithEvents mChartManager As clsChartEventsManager
' ... 8 más
```

Cada nuevo servicio requiere:
1. Añadir variable `WithEvents` en EventCoordinator
2. Implementar manejadores específicos
3. **Viola OCP** (Open/Closed Principle)

**Solución propuesta:**

Implementar interfaz genérica de eventos:

````vba
' IEventPayload.cls (nueva interfaz)
Public Property Get EventType() As String
Public Property Get Source() As Object
Public Property Get Data() As Variant

' clsEventPayload.cls (implementación)
Private mEventType As String
Private mSource As Object
Private mData As Variant

' Los servicios disparan UN solo evento genérico:
Public Event OnAction(ByVal payload As IEventPayload)

' Ejemplo en clsOpportunitiesMgr:
Public Sub ChangeCurrOpportunity(index As Long)
	mCurrentIndex = index
	
	' Crear payload
	Dim payload As New clsEventPayload
	payload.EventType = "OpportunityChanged"
	Set payload.Source = Me
	payload.Data = index
	
	' Disparar evento genérico
	RaiseEvent OnAction(payload)
End Sub

' EventCoordinator escucha UN solo tipo de evento:
Private WithEvents mService As IService  ' Genérico

Private Sub mService_OnAction(ByVal payload As IEventPayload)
	Select Case payload.EventType
		Case "OpportunityChanged"
			HandleOpportunityChanged payload
		Case "ChartActivated"
			HandleChartActivated payload
		' ...
	End Select
End Sub
````

**Patrón de diseño:** Observer + Strategy + Command Pattern

**Ubicación de implementación:**

- **Crear:** `IEventPayload.cls`, `clsEventPayload.cls`
- **Modificar:** TODOS los servicios (añadir `Event OnAction`)
- **Modificar:** `clsEventCoordinator` (un solo `WithEvents` genérico)

**Evaluación:**

| Criterio | Valor | Justificación |
|----------|-------|---------------|
| **Factibilidad** | 3/5 | Requiere modificar TODOS los servicios (~10 clases) |
| **Impacto** | 4/5 | Reduce acoplamiento significativamente, facilita extensión |
| **Esfuerzo** | Alto (20h) | Tocar ~10 clases + EventCoordinator + Tests |
| **Dependencias** | Ninguna | Independiente, pero mejor después de infraestructura |
| **Orden sugerido** | Sprint 4 | Después de infraestructura base estable |

**Prioridad calculada:** MEDIA-ALTA

**Referencias:**
- Anti-patrón actual: [Ver CLAUDE.md (main) Sección 4.2]
- Arquitectura objetivo: [Ver CLAUDE.md (refactor) Sección X]

**Ventajas:**
- ✅ Nuevos servicios no requieren modificar EventCoordinator
- ✅ Facilita testing (mock de payload)
- ✅ Cumple OCP (Open/Closed Principle)

**Desventajas:**
- ⚠️ Pérdida de type safety (payload es genérico)
- ⚠️ Overhead de crear objetos payload
- ⚠️ Requiere refactorización de TODOS los servicios

**Decisión:** ⏳ PENDIENTE (requiere aprobación usuario)

**Implementado en:** (vacío hasta que se implemente)

---

### PROP-002: Lazy Loading Total en ServiceManager

**Estado:** ✅ Aprobada

**Categoría:** Infraestructura / Resiliencia

**Problema que resuelve:**

Si ocurre un reset de VBA (Ctrl+Break, error fatal, etc.):
1. Todas las variables de módulo/clase se pierden
2. `mServiceManager Is Nothing` → True
3. Toda la aplicación se cae
4. **Las UDFs dejan de funcionar** (problema crítico para usuario)

**Solución propuesta:**

ServiceManager no instancia nada en `Initialize`. Cada `Property Get` verifica si la instancia existe, si no → la crea.
````vba
' clsServiceManager

Public Property Get Configuration() As clsConfiguration
	' Lazy initialization
	If mConfiguration Is Nothing Then
		Set mConfiguration = New clsConfiguration
		mConfiguration.Initialize Me.AppContext
		LogInfo "ServiceManager", "Configuration lazy-loaded"
	End If
	Set Configuration = mConfiguration
End Property

' Repetir para TODOS los servicios
````

**Beneficio adicional:** Startup más rápido (no crea todo al abrir Excel)

**Patrón de diseño:** Lazy Initialization + Null Object Pattern

**Ubicación de implementación:**

- **Modificar:** `clsServiceManager` (~8 properties Get)
- **No tocar:** Clases de servicios (sin cambios)

**Evaluación:**

| Criterio | Valor | Justificación |
|----------|-------|---------------|
| **Factibilidad** | 5/5 | Cambio localizado, bajo riesgo, fácil de revertir |
| **Impacto** | 5/5 | Resiliencia crítica ante resets (problema real de usuarios) |
| **Esfuerzo** | Bajo (2h) | Solo modificar ServiceManager |
| **Dependencias** | Ninguna | Independiente |
| **Orden sugerido** | Sprint 1 | Infraestructura base - primero de todo |

**Prioridad calculada:** 🔴 **CRÍTICA**

**Decisión:** ✅ **APROBADA** - Implementar en Sprint 1 como prioridad máxima

**Implementado en:** (vacío hasta Sprint 1)

---

### PROP-003: EventCoordinator como Caja Negra

**Estado:** 🔴 Rechazada

**Categoría:** Optimización / Performance

**Problema que resuelve:**

EventCoordinator **podría** sobrecargarse con eventos irrelevantes:
- Cambios de celda individual (Worksheet_Change)
- Eventos de scroll
- Eventos de selección

**Solución propuesta:**

Filtrar eventos: solo los que afecten a "Oportunidades Comerciales" pasan al EventCoordinator.

**Evaluación:**

| Criterio | Valor | Justificación |
|----------|-------|---------------|
| **Factibilidad** | 4/5 | Requiere definir criterio de filtrado claro |
| **Impacto** | 2/5 | Beneficio marginal (no hay sobrecarga actual) |
| **Esfuerzo** | Medio (8h) | Implementar lógica de filtrado + tests |

**Prioridad calculada:** BAJA

**Razón de rechazo:** **Optimización prematura**
- No hay evidencia de problema de performance actual
- Añade complejidad sin beneficio medible
- **Revisar en futuro** si hay problemas reales

**Decisión:** ❌ **RECHAZADA**

**Alternativa sugerida:** Monitorear performance del EventCoordinator. Si se detecta problema → Reabrir propuesta.

---

<!-- REPETIR PLANTILLA PARA CADA PROPUESTA -->

### PROP-004: [Plantilla para nueva propuesta]

> Copiar esta plantilla para añadir nuevas propuestas

**Estado:** 🟡 Pendiente

**Categoría:** [Arquitectura / Infraestructura / Performance / UX / etc.]

**Problema que resuelve:**

[Describir problema actual en 2-3 párrafos]
- Síntoma 1
- Síntoma 2
- Consecuencia

**Solución propuesta:**

[Describir solución en detalle]

````vba
' Ejemplo de código si aplica
````

**Patrón de diseño:** [Nombre formal del patrón]

**Ubicación de implementación:**

- **Crear:** [Nuevos archivos]
- **Modificar:** [Archivos existentes]
- **Eliminar:** [Archivos obsoletos]

**Evaluación:**

| Criterio | Valor | Justificación |
|----------|-------|---------------|
| **Factibilidad** | X/5 | [Razón] |
| **Impacto** | X/5 | [Razón] |
| **Esfuerzo** | Bajo/Medio/Alto (Xh) | [Razón] |
| **Dependencias** | [Ninguna / Lista de propuestas] | [Razón] |
| **Orden sugerido** | Sprint X | [Razón] |

**Prioridad calculada:** [Crítica / Alta / Media / Baja]

**Referencias:**
- Anti-patrón actual: [Enlace a CLAUDE.md (main)]
- Arquitectura objetivo: [Enlace a CLAUDE.md (refactor)]

**Ventajas:**
- ✅ [Ventaja 1]
- ✅ [Ventaja 2]

**Desventajas:**
- ⚠️ [Desventaja 1]
- ⚠️ [Desventaja 2]

**Decisión:** ⏳ PENDIENTE

**Implementado en:** (vacío)

---


## 📊 MATRIZ DE PRIORIZACIÓN

> Actualizar después de cada evaluación de propuesta

| ID | Propuesta | Prioridad | Estado | Sprint | Esfuerzo | Dependencias |
|----|-----------|-----------|--------|--------|----------|--------------|
| PROP-002 | Lazy Loading | 🔴 Crítica | ✅ Aprobada | Sprint 1 | Bajo (2h) | Ninguna |
| PROP-001 | Sink Interfaces | 🟠 Media-Alta | 🟡 Pendiente | Sprint 4 | Alto (20h) | Ninguna |
| PROP-003 | Caja Negra | 🟢 Baja | 🔴 Rechazada | N/A | - | - |
| ... | ... | ... | ... | ... | ... | ... |

**Leyenda de Prioridades:**
- 🔴 Crítica: Implementar YA (Sprint 1)
- 🟠 Alta: Importante (Sprint 2-3)
- 🟡 Media: Deseable (Sprint 3-4)
- 🟢 Baja: Opcional (Backlog)

---

## 🔄 PROCESO DE GESTIÓN

### 1. Añadir Nueva Propuesta

````
PASOS:
1. Copiar plantilla PROP-XXX
2. Asignar siguiente ID (PROP-005, PROP-006, ...)
3. Completar TODOS los campos
   - Problema (qué resuelve)
   - Solución (cómo lo resuelve)
   - Evaluación (5 criterios)
4. Añadir a catálogo (sección anterior)
5. Añadir fila a Matriz de Priorización
6. Estado inicial: 🟡 Pendiente
7. **NO implementar aún** - esperar evaluación
````

---

### 2. Evaluar Propuesta

````
PASOS:
1. Revisar criterios técnicos (factibilidad, impacto, esfuerzo)
2. Calcular prioridad:
   - Factibilidad 4-5 + Impacto 4-5 + Esfuerzo Bajo → Crítica
   - Factibilidad 3-4 + Impacto 3-4 → Alta/Media
   - Impacto 1-2 → Baja
3. Identificar dependencias con otras propuestas
4. Asignar a sprint tentativo
5. Cambiar estado: 🟡 Pendiente → 🟢 Aprobada / 🔴 Rechazada
6. Actualizar Matriz de Priorización
````

---

### 3. Incorporar al Plan

````
PASOS (solo si estado = ✅ Aprobada):
1. Abrir CLAUDE.md (refactor)
2. Localizar sección del sprint correspondiente
3. Añadir propuesta con detalles:
   - Qué hacer
   - Cómo hacerlo
   - Tests de aceptación
4. Marcar en POOL_PROPUESTAS: "Incorporada al plan Sprint X"
5. Añadir enlace cruzado:
   POOL → CLAUDE.md (refactor)
````

---

### 4. Implementar Propuesta

````
PASOS (durante sprint):
1. Cambiar estado: ✅ Aprobada → 🟠 En progreso
2. Implementar según plan en CLAUDE.md (refactor)
3. Ejecutar tests de aceptación
4. Code review
5. Merge a rama refactor
````

---

### 5. Cerrar Propuesta Implementada

````
PASOS (al completar):
1. Cambiar estado: 🟠 En progreso → 🟢 Completada
2. Actualizar campo "Implementado en:":
   Ejemplo: "Sprint 1 - Commit abc123"
3. Actualizar Matriz de Priorización
4. (Opcional) Archivar moviendo a sección "Propuestas Completadas"
````

---

## 📚 REFERENCIAS

### Patrones de Diseño

- [Gang of Four - Design Patterns](https://refactoring.guru/design-patterns)
- [Martin Fowler - Refactoring](https://refactoring.com/)
- [Refactoring.Guru](https://refactoring.guru/)

### Anti-Patrones

- [SourceMaking - AntiPatterns](https://sourcemaking.com/antipatterns)
- [Code Smells](https://refactoring.guru/refactoring/smells)

### VBA Best Practices

- [RubberDuck VBA](https://rubberduckvba.com/)
- [Chip Pearson VBA](http://www.cpearson.com/excel/)
- [Excel VBA Best Practices (Microsoft)](https://docs.microsoft.com/en-us/office/vba/excel)

---

## 📝 CHANGELOG

| Fecha | Versión | Cambios | Autor |
|-------|---------|---------|-------|
| 2026-01-13 | 1.0 | Creación inicial | Humano + Claude |
| 2026-01-13 | 1.1 | Corrección encoding + formato | Claude |

---

**FIN DE POOL_PROPUESTAS.md v1.1**
