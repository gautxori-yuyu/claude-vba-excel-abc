# DOCUMENTO 3: PLAN DE MIGRACIÓN PASO A PASO

**Fecha:** 2026-01-22  
**Objetivo:** Guía ejecutable para refactorizar sin romper funcionalidad

---

## ESTRATEGIA GENERAL

### Principio: "Funciona en Cada Paso"
- Cada paso produce código **compilable**
- Cada paso mantiene **funcionalidad existente**
- Avanzamos por **capas**: desde base (infraestructura) hacia arriba (presentación)
- Probamos **después de cada paso**

### Orden de Migración
1. ✅ Crear estructura base (carpetas)
2. ✅ Renombrar archivos (Git mv)
3. ✅ Refactorizar Aplicación (clsApplication)
4. ✅ Refactorizar Infraestructura (Estado, DI, Adaptadores)
5. ✅ Refactorizar Servicios (eliminar WithEvents innecesarios)
6. ✅ Refactorizar Dominio y Datos (renombrado)
7. ✅ Refactorizar Presentación (callbacks)
8. ✅ Limpiar código muerto

---

## FASE 0: PREPARACIÓN

### 0.1 Backup y Branch
```bash
# Crear rama de refactorización
git checkout -b refactor/clean-architecture

# Asegurar que todo compila antes de empezar
# [Abrir VBA, compilar, verificar sin errores]
```

### 0.2 Crear Estructura de Carpetas

**Acción:** Crear carpetas en el proyecto VBA según arquitectura objetivo.

```
/src
├── /1-Application
├── /2-Infrastructure
│   ├── /DependencyInjection
│   ├── /State
│   ├── /ExcelIntegration
│   ├── /FileSystem
│   ├── /Configuration
│   └── /Logging
├── /3-Services
│   ├── /ChartManagement
│   ├── /FileManagement
│   ├── /RibbonManagement
│   └── /OpportunityManagement
├── /4-Domain
│   ├── /Entities
│   ├── /Templates
│   └── /ValueObjects
├── /5-Data
├── /6-Presentation
│   ├── /RibbonCallbacks
│   ├── /Forms
│   └── /Worksheets
├── /7-Macros
├── /8-UserDefinedFunctions
└── /9-Utilities
```

**Verificación:** Carpetas creadas en el VBA Project Explorer.

---

## FASE 1: RENOMBRADO DE ARCHIVOS

### 1.1 Clases Principales

**Objetivo:** Renombrar clases clave antes de modificar código.

| Paso | Antiguo | Nuevo | Carpeta Destino |
|------|---------|-------|-----------------|
| 1.1.1 | `clsAplicacion.cls` | `clsApplication.cls` | `/1-Application` |
| 1.1.2 | `clsApplicationContext.cls` | `clsApplicationState.cls` | `/2-Infrastructure/State` |
| 1.1.3 | `clsServiceManager.cls` | `clsServiceContainer.cls` | `/2-Infrastructure/DependencyInjection` |
| 1.1.4 | `clsExecutionContext.cls` | `clsExcelExecutionContext.cls` | `/2-Infrastructure/ExcelIntegration` |
| 1.1.5 | `clsChartEvents.cls` | `clsChartEventAdapter.cls` | `/2-Infrastructure/ExcelIntegration` |
| 1.1.6 | `clsChartEventsManager.cls` | `clsChartEventManager.cls` | `/3-Services/ChartManagement` |
| 1.1.7 | `clsRibbonEvents.cls` | `clsRibbonEventAdapter.cls` | `/2-Infrastructure/ExcelIntegration` |
| 1.1.8 | `clsRibbonUI.cls` | `clsRibbonManager.cls` | `/3-Services/RibbonManagement` |
| 1.1.9 | `clsFSWatcher.cls` | `clsFileSystemWatcher.cls` | `/2-Infrastructure/FileSystem` |
| 1.1.10 | `clsFSMonitoringCoord.cls` | `clsFileSystemMonitor.cls` | `/2-Infrastructure/FileSystem` |
| 1.1.11 | `clsOpportunitiesMgr.cls` | `clsOpportunityManager.cls` | `/3-Services/OpportunityManagement` |
| 1.1.12 | `clsConfiguration.cls` | `clsApplicationConfiguration.cls` | `/2-Infrastructure/Configuration` |

**Procedimiento por clase:**
1. Abrir VBA Editor
2. Exportar clase antigua: `File > Export File` → guardar en carpeta temporal
3. Editar archivo exportado:
   - Cambiar `Attribute VB_Name = "NombreAntiguo"` → `Attribute VB_Name = "NombreNuevo"`
   - Cambiar `'@Folder` según carpeta destino
4. Eliminar clase antigua del proyecto
5. Importar clase renombrada: `File > Import File`
6. **Compilar** (debe dar errores por referencias al nombre antiguo - es normal)

**Verificación:** Todas las clases renombradas importadas.

### 1.2 Actualizar Referencias en TODO el Código

**Herramienta:** Find & Replace en VBA Editor (Edit > Replace)

**Búsqueda EXACTA** (Match Case, Whole Word):

| Buscar | Reemplazar por | Ámbito |
|--------|---------------|--------|
| `clsAplicacion` | `clsApplication` | Entire Project |
| `clsApplicationContext` | `clsApplicationState` | Entire Project |
| `clsServiceManager` | `clsServiceContainer` | Entire Project |
| `clsExecutionContext` | `clsExcelExecutionContext` | Entire Project |
| `clsChartEvents` | `clsChartEventAdapter` | Entire Project |
| `clsChartEventsManager` | `clsChartEventManager` | Entire Project |
| `clsRibbonEvents` | `clsRibbonEventAdapter` | Entire Project |
| `clsRibbonUI` | `clsRibbonManager` | Entire Project |
| `clsFSWatcher` | `clsFileSystemWatcher` | Entire Project |
| `clsFSMonitoringCoord` | `clsFileSystemMonitor` | Entire Project |
| `clsOpportunitiesMgr` | `clsOpportunityManager` | Entire Project |
| `clsConfiguration` | `clsApplicationConfiguration` | Entire Project |

**Verificación:** `Debug > Compile VBAProject` → **SIN ERRORES**

### 1.3 Renombrar Variables Comunes

**En cada archivo que use estas variables:**

| Buscar (variable) | Reemplazar por | Contexto |
|-------------------|---------------|----------|
| `mServiceManager` | `_serviceContainer` | `clsApplication` |
| `mAppContext` | `_appState` | `clsApplication` |
| `ctx` (As clsExecutionContext) | `context` o `_excelContext` | Varios |
| `fw` (As clsFSWatcher) | `_watcher` | `clsFileSystemMonitor` |
| `m_xlApp` | `_app` | `clsExcelExecutionContext` |
| `mChart` | `_chart` | `clsChartEventAdapter` |

**Nota:** Hacer manualmente archivo por archivo (no global) para evitar colisiones.

**Verificación:** `Debug > Compile VBAProject` → **SIN ERRORES**

---

## FASE 2: REFACTORIZAR APLICACIÓN (clsApplication)

### 2.1 Eliminar WithEvents Innecesarios en clsApplication

**Archivo:** `clsApplication.cls`

**Estado actual (problemático):**
```vba
Private WithEvents mOpportunities As clsOpportunityManager
Private WithEvents mChartManager As clsChartEventManager
Private WithEvents mFSMonitoringCoord As clsFileSystemMonitor
Private WithEvents mRibbonState As clsRibbonState
Private WithEvents evRibbon As clsRibbonEventAdapter
Private WithEvents ctx As clsExcelExecutionContext
```

**Paso 2.1.1:** Eliminar `WithEvents` de las declaraciones
```vba
' ANTES
Private WithEvents mOpportunities As clsOpportunityManager

' DESPUÉS
Private _opportunityManager As clsOpportunityManager
```

**Cambios completos:**
```vba
' Declaraciones (parte superior de la clase)
Private _serviceContainer As clsServiceContainer
Private _appState As clsApplicationState
Private _excelContext As clsExcelExecutionContext
Private _chartEventManager As clsChartEventManager
Private _fileSystemMonitor As clsFileSystemMonitor
Private _ribbonManager As clsRibbonManager
Private _opportunityManager As clsOpportunityManager

' Estado (NO son servicios)
Private _ribbonState As clsRibbonState
Private _chartState As clsChartState
Private _fileState As clsFileState
```

**Paso 2.1.2:** Eliminar event handlers obsoletos

Buscar y ELIMINAR todos los procedimientos tipo:
```vba
Private Sub mOpportunities_currOpportunityChanged(...)
Private Sub mChartManager_ChartActivated(...)
Private Sub ctx_WorkbookOpened(...)
' etc
```

**Paso 2.1.3:** Reemplazar event handlers por callbacks directos

**ANTES (con WithEvents):**
```vba
Private Sub ctx_WorkbookOpened(ByVal wb As Workbook)
    ' Lógica aquí
End Sub
```

**DESPUÉS (callback directo):**
```vba
' Nuevo método público que el contexto llamará
Public Sub OnWorkbookOpened(ByVal wb As Workbook)
    ' Misma lógica aquí
End Sub
```

**Paso 2.1.4:** Modificar inicialización en `Initialize`

**ANTES:**
```vba
Set mOpportunities = New clsOpportunityManager
' Se suscribía automáticamente via WithEvents
```

**DESPUÉS:**
```vba
Set _opportunityManager = New clsOpportunityManager
_opportunityManager.Initialize _appState

' Para clsExcelExecutionContext, pasar callback handler
Set _excelContext = New clsExcelExecutionContext
_excelContext.Initialize Me  ' Pasa "Me" como handler
```

**Verificación:** `Debug > Compile` → **SIN ERRORES**

### 2.2 Refactorizar clsExcelExecutionContext

**Archivo:** `clsExcelExecutionContext.cls`

**Objetivo:** Eliminar `RaiseEvent`, usar callbacks directos.

**ANTES:**
```vba
Public Event WorkbookOpened(ByVal wb As Workbook)
Public Event WorkbookActivated(ByVal wb As Workbook)
' etc

Private Sub _app_WorkbookOpen(ByVal wb As Workbook)
    RaiseEvent WorkbookOpened(wb)
End Sub
```

**DESPUÉS:**
```vba
Private _callbackHandler As clsApplication

Public Sub Initialize(ByVal callbackHandler As clsApplication)
    Set _callbackHandler = callbackHandler
    Set _app = Application
End Sub

Private Sub _app_WorkbookOpen(ByVal wb As Workbook)
    If Not _callbackHandler Is Nothing Then
        _callbackHandler.OnWorkbookOpened wb
    End If
End Sub
```

**Cambios:**
1. Eliminar TODAS las declaraciones `Public Event ...`
2. Eliminar TODOS los `RaiseEvent ...`
3. Agregar campo `Private _callbackHandler As clsApplication`
4. Modificar `Initialize` para recibir el handler
5. Cambiar cada `RaiseEvent XXX` por `_callbackHandler.OnXXX`

**Métodos a crear en clsApplication:**
```vba
Public Sub OnWorkbookOpened(ByVal wb As Workbook)
    ' Código que estaba en el event handler
End Sub

Public Sub OnWorkbookActivated(ByVal wb As Workbook)
    ' ...
End Sub

' etc para cada evento
```

**Verificación:** `Debug > Compile` → **SIN ERRORES**

### 2.3 Refactorizar clsChartEventManager

**Archivo:** `clsChartEventManager.cls`

**Objetivo:** Eliminar eventos custom, usar métodos públicos.

**ANTES:**
```vba
Public Event ChartActivated(ByVal chart As Excel.Chart)
Public Event ChartDeactivated()

Public Sub OnChartActivated(...)
    RaiseEvent ChartActivated(...)
End Sub
```

**DESPUÉS:**
```vba
' Eliminar eventos
' Los métodos públicos ya existen, solo quitar RaiseEvent

Public Sub OnChartActivated(ByVal chart As Excel.Chart)
    ' Actualizar estado
    _appState.ChartState.ActiveChart = chart
    
    ' Llamada DIRECTA a otro servicio si necesario
    If Not _ribbonManager Is Nothing Then
        _ribbonManager.RefreshChartButtons
    End If
End Sub
```

**Cambios:**
1. Eliminar `Public Event ChartActivated`, etc
2. Eliminar `RaiseEvent` dentro de métodos
3. Inyectar dependencias necesarias (RibbonManager, AppState)

**En clsApplication.Initialize:**
```vba
' Inyectar dependencias
_chartEventManager.Initialize _appState, _ribbonManager
```

**Verificación:** `Debug > Compile` → **SIN ERRORES**

### 2.4 Refactorizar Otros Servicios

**Mismo patrón para:**
- `clsOpportunityManager` → eliminar eventos custom
- `clsRibbonEventAdapter` → eliminar eventos custom, usar callbacks
- `clsRibbonState` → eliminar `StateChanged`, usar property setters

**Verificación final FASE 2:** `Debug > Compile` → **SIN ERRORES**

---

## FASE 3: REFACTORIZAR INFRAESTRUCTURA

### 3.1 Separar Estado de Servicios

**Mover a carpetas correctas:**

| Archivo | Carpeta Actual | Carpeta Destino |
|---------|---------------|-----------------|
| `clsChartState.cls` | `2-Servicios.Excel.Charts` | `/2-Infrastructure/State` |
| `clsRibbonState.cls` | `2-Servicios.Excel.Ribbon` | `/2-Infrastructure/State` |
| `clsFileState.cls` | `2-Servicios.Archivos` | `/2-Infrastructure/State` |

**Actualizar @Folder:**
```vba
'@Folder "2-Infrastructure.State"
```

### 3.2 Mover Adaptadores de Eventos

| Archivo | Carpeta Destino |
|---------|-----------------|
| `clsChartEventAdapter.cls` | `/2-Infrastructure/ExcelIntegration` |
| `clsRibbonEventAdapter.cls` | `/2-Infrastructure/ExcelIntegration` |
| `clsExcelExecutionContext.cls` | `/2-Infrastructure/ExcelIntegration` |

**Actualizar @Folder:**
```vba
'@Folder "2-Infrastructure.ExcelIntegration"
```

### 3.3 Organizar FileSystem

| Archivo | Carpeta Destino |
|---------|-----------------|
| `clsFileSystemWatcher.cls` | `/2-Infrastructure/FileSystem` |
| `clsFileSystemMonitor.cls` | `/2-Infrastructure/FileSystem` |

**Actualizar @Folder:**
```vba
'@Folder "2-Infrastructure.FileSystem"
```

**Verificación FASE 3:** Todas las clases en carpetas correctas, compila sin errores.

---

## FASE 4: REFACTORIZAR SERVICIOS

### 4.1 Eliminar WithEvents en Servicios

**clsFileManager:**
```vba
' ANTES
Private WithEvents ctx As clsExcelExecutionContext

' DESPUÉS (sin WithEvents)
Private _excelContext As clsExcelExecutionContext

Public Sub Initialize(ByVal excelContext As clsExcelExecutionContext, ...)
    Set _excelContext = excelContext
    ' ...
End Sub
```

**clsOpportunityManager:**
```vba
' Mismo patrón: eliminar WithEvents, pasar por Initialize
```

### 4.2 Reorganizar por Subcarpetas

| Archivo | Subcarpeta |
|---------|-----------|
| `clsChartEventManager.cls` | `/3-Services/ChartManagement` |
| `clsFileManager.cls` | `/3-Services/FileManagement` |
| `clsRibbonManager.cls` | `/3-Services/RibbonManagement` |
| `clsOpportunityManager.cls` | `/3-Services/OpportunityManagement` |

**Actualizar @Folder** en cada archivo.

**Verificación FASE 4:** Compila sin errores.

---

## FASE 5: REFACTORIZAR DOMINIO Y DATOS

### 5.1 Renombrar Entidades

| Antiguo | Nuevo | Carpeta |
|---------|-------|---------|
| `clsOferta.cls` | `clsOffer.cls` | `/4-Domain/Entities` |
| `clsOfertaOtro.cls` | `clsOtherOffer.cls` | `/4-Domain/Entities` |
| `clsOpportunity.cls` | `clsOpportunity.cls` | `/4-Domain/Entities` (ya OK) |

**Actualizar referencias:**
- Find & Replace: `clsOferta` → `clsOffer`
- Find & Replace: `clsOfertaOtro` → `clsOtherOffer`

### 5.2 Mover Plantillas

| Archivo | Carpeta |
|---------|---------|
| `clsOpportunityOfferBudgetTpl.cls` | `/4-Domain/Templates` |
| `clsOpportunityOfferQuotationTpl.cls` | `/4-Domain/Templates` |

**Renombrar:**
- `...BudgetTpl` → `...BudgetTemplate`
- `...QuotationTpl` → `...QuotationTemplate`

### 5.3 Mover Repositorios

| Archivo | Nuevo Nombre | Carpeta |
|---------|--------------|---------|
| `clsOfertaRepository.cls` | `clsOfferRepository.cls` | `/5-Data` |
| `clsDBContext.cls` | `clsDBContext.cls` | `/5-Data` (ya OK) |

**Actualizar @Folder:**
```vba
'@Folder "5-Data"
```

**Verificación FASE 5:** Compila sin errores.

---

## FASE 6: REFACTORIZAR PRESENTACIÓN

### 6.1 Renombrar Callbacks

| Antiguo | Nuevo | Carpeta |
|---------|-------|---------|
| `modCALLBACKSRibbon.bas` | `modRibbonCallbacks.bas` | `/6-Presentation/RibbonCallbacks` |

**Actualizar @Folder:**
```vba
'@Folder "6-Presentation.RibbonCallbacks"
```

### 6.2 Mover Formularios

| Archivo | Carpeta |
|---------|---------|
| `frmConfiguracion.frm` | `/6-Presentation/Forms` |
| `frmImportExportMacros.frm` | `/6-Presentation/Forms` |
| `frmComparadorHojas.frm` | `/6-Presentation/Forms` |

### 6.3 Mover Worksheets

| Archivo | Carpeta |
|---------|---------|
| `wshUnidades.cls` | `/6-Presentation/Worksheets` |

**Verificación FASE 6:** Compila sin errores.

---

## FASE 7: ORGANIZAR MACROS Y UDFS

### 7.1 Renombrar Macros

| Patrón Antiguo | Patrón Nuevo |
|---------------|--------------|
| `modMACRO*.bas` | `modMacro*.bas` |

**Mover a `/7-Macros`**

### 7.2 Renombrar UDFs

| Patrón Antiguo | Patrón Nuevo |
|---------------|--------------|
| `UDFs_*.bas` | `modUdf*.bas` |

**Ejemplos:**
- `UDFs_Units.bas` → `modUdfUnits.bas`
- `UDFs_Utilids.bas` → `modUdfUtilities.bas`
- `UDFs_CGASING.bas` → `modUdfCgasing.bas`

**Mover a `/8-UserDefinedFunctions`**

**Verificación FASE 7:** Compila sin errores.

---

## FASE 8: UTILIDADES

### 8.1 Renombrar Utilidades

| Patrón Antiguo | Patrón Nuevo |
|---------------|--------------|
| `modUTILS*.bas` | `modUtils*.bas` |

**Mover a `/9-Utilities`**

### 8.2 Logging y Constantes

| Archivo | Nuevo Nombre | Carpeta |
|---------|--------------|---------|
| `mod_Logger.bas` | `mod_Logger.bas` (OK) | `/2-Infrastructure/Logging` |
| `mod_ConstantsGlobals.bas` | `mod_Constants.bas` | `/2-Infrastructure/Logging` |

**Verificación FASE 8:** Compila sin errores.

---

## FASE 9: LIMPIEZA FINAL

### 9.1 Identificar Código Muerto

**Buscar:**
- `clsEventDispatcher` → ¿Se usa? Si no, eliminar
- Módulos sin referencias
- Funciones no llamadas

**Herramienta:** VBA Code Analyzer o búsqueda manual.

### 9.2 Documentar

**Actualizar @Folder en TODOS los archivos** según nueva estructura.

**Agregar headers en clases principales:**
```vba
'@Folder "1-Application"
'@Description: Composition Root - Entry point and dependency wiring
```

---

## VERIFICACIÓN FINAL

### Checklist

- [ ] `Debug > Compile VBAProject` → **SIN ERRORES**
- [ ] Ejecutar aplicación desde `ThisWorkbook_Open` → **FUNCIONA**
- [ ] Probar funcionalidad clave:
  - [ ] Crear nueva oportunidad
  - [ ] Abrir gráfico (eventos)
  - [ ] Ribbon responde a comandos
  - [ ] FileSystem monitoring funciona
- [ ] Revisar que NO hay:
  - [ ] WithEvents innecesarios en `clsApplication`
  - [ ] Eventos custom 1-a-1
  - [ ] Referencias a nombres antiguos

### Commit

```bash
git add .
git commit -m "refactor: Clean Architecture - eliminate unnecessary events"
git push origin refactor/clean-architecture
```

---

## SIGUIENTE PASO

Ver **DOCUMENTO 4: SKILL PARA CLAUDE CODE** para instrucciones automatizadas.
