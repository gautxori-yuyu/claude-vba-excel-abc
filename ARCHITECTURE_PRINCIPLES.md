# Architecture Principles - VBA Excel Application

## Layer Model

```
┌──────────────────────────────────────────────┐
│  UI (Ribbon, Forms, Panels)                   │  ← Solo escucha eventos. Solo ejecuta órdenes.
├──────────────────────────────────────────────┤     No conoce nada abajo. Decide cómo "pintarse".
│  Domain Managers + Domain Events              │  ← Cambian estado. Después notifican (RaiseEvent).
├──────────────────────────────────────────────┤
│  Mediators (Infra + Domain)                   │  ← Orquestan. Conectan. Reenvían. NO emiten eventos.
├──────────────────────────────────────────────┤
│  Infrastructure (ExecutionContextMgr, FileMgr)│  ← Conoce Excel. Emite eventos semánticos.
├──────────────────────────────────────────────┤
│  Domain (Entities, State, Interfaces)         │  ← No conoce Workbook/Worksheet/Range/Application.
└──────────────────────────────────────────────┘
```

## Event Flow (Canonical)

```
[ Evento técnico Excel/FS ]
        ↓
[ ExecutionContextListener (ThisWorkbook) ]   ← Captura, reenvía sin semántica. Zero RaiseEvent.
        ↓ (llamada a método)
[ ExecutionContextMgr ]                       ← Traduce → actualiza State → RaiseEvent semántico
        ↓ (eventos semánticos de infraestructura)
[ Mediador de Infraestructura ]               ← Escucha. Decide a qué Mgr llamar. NO emite.
        ↓ (métodos)
[ Domain Managers ]                           ← Ejecutan lógica. Actualizan State. RaiseEvent dominio.
        ↓ (eventos de dominio)
[ UI / Ribbon / Forms ]                       ← Escucha eventos. Se invalida.
```

## Parallel: State Ownership

```
Mgr ──owns──► State          (composición: el Mgr crea y posee el State)
AppState ──refs──► all States (solo referencias para consultas coherentes)
```

## Roles y Responsabilidades

### State Classes (MEMORIA ESTRUCTURADA, no comportamiento)
- **Nunca** emiten eventos (RaiseEvent)
- **Nunca** tienen Property Let/Set públicos ni Friend (excepto métodos como Register/Clear/Reset llamados SOLO por su Mgr)
- **Nunca** toman decisiones ni validan
- Solo exponen Property Get (consultas de estado)
- Métodos permitidos: RegisterX, Clear, Snapshot, Reset (llamados por su Mgr)

### Manager Classes (ORDENES + ESTADO + EVENTOS)
- Ejecutan lógica imperativa (métodos, Property Let como equivalente a método)
- Actualizan su State (composición)
- Emiten eventos DESPUÉS de que el estado haya cambiado (notificación de hecho consumado)
- Un Manager solo emite eventos sobre estado que ÉL controla
- El flujo SIEMPRE es: Lógica → Actualiza State → RaiseEvent

### Mediator Classes (ORQUESTACIÓN PURA)
- Conectan publicadores con consumidores
- NO emiten eventos de negocio (ni de cualquier tipo)
- Conocen a las clases de infraestructura y de dominio
- Las clases de infraestructura NO conocen al mediador

### ApplicationState (RESULTADO, no origen)
- Consolida estado de todos los subsistemas (solo referencias, no posesión)
- Permite consultas coherentes de estado global
- Facilita logging y diagnósticos
- NO controla nada. NO da órdenes. NO emite eventos.
- Solo tiene Property Gets (fachada de lectura)

### UI Layer (clsRibbon, modCALLBACKS*)
- Capa más superior: nada la conoce
- Solo escucha eventos (decide cómo pintarse)
- Puede ejecutar órdenes (llamar métodos en Mgrs)
- Se invalida en respuesta a eventos, no por llamadas desde abajo
- La aplicación debe funcionar sin UI

## Event Types

1. **Técnicos**: Excel, FileSystem, COM, Host (WorkbookOpen, SheetActivate)
   - Capturados por ExecutionContextListener
   - No tienen semántica de dominio

2. **Semánticos de Infraestructura**: Traducción de los técnicos
   - Emitidos por ExecutionContextMgr (que es un Manager, no un Mediator)
   - Ejemplos: WorkbookSessionStarted, ActiveChartChanged, FileBecameReadOnly

3. **De Dominio**: Notificación de que algo ya ha ocurrido
   - Emitidos SOLO por Domain Managers
   - Ejemplos: OpportunityActivated, CollectionUpdated
   - Un evento es notificación de hecho consumado, NUNCA señal de inicio

## Interfaces (Puertos)

- El Dominio define contratos (interfaces): qué necesita, no cómo se obtiene
- La Infraestructura implementa: traduce tecnología → contrato
- Ejemplo: IOpportunitySource (dominio) ← ExcelOpportunitySource (infraestructura)
- El Dominio nunca referencia Workbook, Worksheet, Range directamente

## Invalidación del Ribbon (Patrón correcto)

```
INCORRECTO: mRibbon.InvalidateControl "btnGuardar"   ← Infraestructura sabe que hay Ribbon
CORRECTO:
    Infrastructure/Domain → RaiseEvent StateChanged(...)
    clsRibbon ← Sub mSource_StateChanged(...) → Me.Invalidate
```

## Property Get vs Property Let/Set

- **Property Get**: Vive en State o en Manager (como fachada). Solo consulta. No dispara eventos.
- **Property Let/Set en Manager**: Equivalente a método imperativo. Solo si el significado es "establecer el actual" y no hay lógica compleja.
- **Property Let/Set en State**: PROHIBIDO. Convierte el estado en anárquico.

## Who Listens to Events

- UI escucha TODOS los eventos (técnicos, semánticos, de dominio)
- Otros Mgrs escuchan eventos cuando es necesario (FileMgr → OpportunitiesMgr)
- La escucha se hace A TRAVÉS DE MEDIADORES, nunca directa entre Mgrs
- Estado → Mgr: NUNCA

## Estabilidad y Reset

- clsExecutionContext debe permitir: reenganchar eventos, reconstruir estado derivado, aislar el daño del reset
- Si Excel/VBA se resetea, solo se rompa infraestructura (NO dominio, NO UI, NO coordinadores)
- Eventos de ciclo de vida: ContextInvalidated(), ContextReinitialized()
- Solo clsExecutionContextMgr puede manejar reset/reenganche
