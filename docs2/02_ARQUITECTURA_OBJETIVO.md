# DOCUMENTO 2: ARQUITECTURA OBJETIVO

**Fecha:** 2026-01-22  
**Objetivo:** Definir estructura final limpia, mantenible y pragmática

---

## PRINCIPIOS ARQUITECTÓNICOS

### 1. Simplicidad sobre Elegancia
- ❌ NO usar patrones "porque son bonitos"
- ✅ SÍ usar patrones solo cuando resuelven problemas reales
- ✅ Llamadas directas > Eventos custom (si solo hay 1 suscriptor)

### 2. Separación de Responsabilidades (SoC)
- **Aplicación**: Inicialización y coordinación
- **Infraestructura**: Detalles técnicos (Excel, FileSystem, COM)
- **Servicios**: Lógica de aplicación (managers, coordinadores)
- **Dominio**: Entidades y reglas de negocio
- **Datos**: Acceso a persistencia

### 3. Eventos Solo Donde Tienen Sentido
- ✅ **Eventos COM** (WithEvents): Excel, Chart, FileSystem COM
- ✅ **Eventos de Dominio**: Cambios en oportunidades (1-a-N potencial)
- ❌ **Eventos 1-a-1**: Reemplazar por llamadas directas

### 4. Dependencias Claras
- La capa superior puede depender de la inferior
- La capa inferior NUNCA depende de la superior
- Sin referencias circulares

---

## ESTRUCTURA DE CARPETAS (Clean Architecture)

```
/src
├── /1-Application                    [Capa: Aplicación]
│   ├── clsApplication.cls                   ' Composition Root
│   └── ThisWorkbook.cls                     ' Entry point Excel
│
├── /2-Infrastructure                 [Capa: Infraestructura]
│   │
│   ├── /DependencyInjection
│   │   ├── clsServiceContainer.cls          ' DI Container (antes ServiceManager)
│   │   └── IService.cls                     ' Interface para servicios
│   │
│   ├── /State                        [Estado compartido]
│   │   ├── clsApplicationState.cls          ' Estado global (antes ApplicationContext)
│   │   ├── clsChartState.cls                ' Estado de gráficos
│   │   ├── clsRibbonState.cls               ' Estado del Ribbon
│   │   └── clsFileState.cls                 ' Estado de archivos
│   │
│   ├── /ExcelIntegration             [Integración Excel COM]
│   │   ├── clsExcelExecutionContext.cls     ' WithEvents Application (antes ExecutionContext)
│   │   ├── clsChartEventAdapter.cls         ' WithEvents Chart (antes ChartEvents)
│   │   └── clsRibbonEventAdapter.cls        ' WithEvents Ribbon (antes RibbonEvents)
│   │
│   ├── /FileSystem                   [Monitoreo sistema archivos]
│   │   ├── clsFileSystemWatcher.cls         ' WithEvents COM FolderWatcher (antes FSWatcher)
│   │   └── clsFileSystemMonitor.cls         ' Coordinador + eventos dominio (antes FSMonitoringCoord)
│   │
│   ├── /Configuration
│   │   ├── clsApplicationConfiguration.cls  ' Configuración (antes Configuration)
│   │   └── frmConfiguracion.frm             ' UI configuración
│   │
│   └── /Logging
│       ├── mod_Logger.bas                   ' Sistema de logging
│       └── mod_Constants.bas                ' Constantes globales
│
├── /3-Services                       [Capa: Servicios]
│   │
│   ├── /ChartManagement
│   │   └── clsChartEventManager.cls         ' Gestión eventos gráficos
│   │
│   ├── /FileManagement
│   │   └── clsFileManager.cls               ' Gestión de archivos
│   │
│   ├── /RibbonManagement
│   │   └── clsRibbonManager.cls             ' Gestión del Ribbon (antes RibbonUI)
│   │
│   └── /OpportunityManagement
│       └── clsOpportunityManager.cls        ' Gestión oportunidades (antes OpportunitiesMgr)
│
├── /4-Domain                         [Capa: Dominio]
│   │
│   ├── /Entities
│   │   ├── clsOpportunity.cls               ' Entidad oportunidad
│   │   ├── clsOffer.cls                     ' Entidad oferta (antes Oferta)
│   │   └── clsOtherOffer.cls                ' Oferta otro tipo (antes OfertaOtro)
│   │
│   ├── /Templates
│   │   ├── clsOpportunityBudgetTemplate.cls ' Plantilla presupuesto
│   │   └── clsOpportunityQuotationTemplate.cls ' Plantilla cotización
│   │
│   └── /ValueObjects
│       └── modOfferTypes.bas                ' Tipos de ofertas
│
├── /5-Data                           [Capa: Datos]
│   ├── clsOfferRepository.cls               ' Repositorio ofertas (antes OfertaRepository)
│   └── clsDBContext.cls                     ' Contexto de base datos
│
├── /6-Presentation                   [Capa: Presentación]
│   │
│   ├── /RibbonCallbacks
│   │   └── modRibbonCallbacks.bas           ' Callbacks del Ribbon
│   │
│   ├── /Forms
│   │   ├── frmImportExportMacros.frm
│   │   └── frmComparadorHojas.frm
│   │
│   └── /Worksheets
│       └── wshUnidades.cls
│
├── /7-Macros                         [Macros de usuario]
│   ├── modMacroAppLifecycle.bas
│   ├── modMacroGraphics.bas
│   ├── modMacroUnits.bas
│   ├── modMacroComparison.bas
│   └── (resto de macros...)
│
├── /8-UserDefinedFunctions           [UDFs]
│   ├── modUdfUnits.bas
│   ├── modUdfUtilities.bas
│   ├── modUdfCgasing.bas
│   ├── modUdfCoolprop.bas
│   └── (resto de UDFs...)
│
└── /9-Utilities                      [Utilidades]
    ├── modUtilsShellCmd.bas
    ├── modUtilsRefEditAPI.bas
    ├── modUtilsProcedureParsing.bas
    └── (resto de utilidades...)
```

---

## RESPONSABILIDADES POR CLASE (una línea)

### Capa 1: Aplicación
| Clase | Responsabilidad |
|-------|----------------|
| `clsApplication` | Composition Root: crea y conecta todos los componentes |
| `ThisWorkbook` | Entry point: captura Workbook_Open y delega a clsApplication |

### Capa 2: Infraestructura

#### 2.1 Dependency Injection
| Clase | Responsabilidad |
|-------|----------------|
| `clsServiceContainer` | DI Container: registra y resuelve servicios singleton |
| `IService` | Interface: contrato para servicios con ciclo de vida |

#### 2.2 State
| Clase | Responsabilidad |
|-------|----------------|
| `clsApplicationState` | Contenedor de estado global (ribbon, chart, file, execution) |
| `clsChartState` | Estado de gráficos activos y seleccionados |
| `clsRibbonState` | Estado del Ribbon (botones habilitados, etc) |
| `clsFileState` | Estado de archivos (abiertos, tracking) |

#### 2.3 Excel Integration
| Clase | Responsabilidad |
|-------|----------------|
| `clsExcelExecutionContext` | **WithEvents Application**: captura eventos Excel y ejecuta callbacks directos |
| `clsChartEventAdapter` | **WithEvents Chart**: captura eventos Chart y ejecuta callbacks directos |
| `clsRibbonEventAdapter` | **WithEvents Ribbon**: captura eventos Ribbon y ejecuta callbacks directos |

#### 2.4 FileSystem
| Clase | Responsabilidad |
|-------|----------------|
| `clsFileSystemWatcher` | **WithEvents FolderWatcher COM**: captura eventos FS y notifica monitor |
| `clsFileSystemMonitor` | Coordinador: interpreta eventos FS y **dispara eventos de dominio** |

#### 2.5 Configuration
| Clase | Responsabilidad |
|-------|----------------|
| `clsApplicationConfiguration` | Lee/escribe configuración desde Registry |
| `frmConfiguracion` | UI para editar configuración |

#### 2.6 Logging
| Módulo | Responsabilidad |
|--------|----------------|
| `mod_Logger` | Sistema de logging centralizado |
| `mod_Constants` | Constantes globales de aplicación |

### Capa 3: Servicios
| Clase | Responsabilidad |
|-------|----------------|
| `clsChartEventManager` | Procesa eventos de gráficos y actualiza estado |
| `clsFileManager` | Gestiona archivos (tracking, estados) |
| `clsRibbonManager` | Gestiona Ribbon (puntero IRibbonUI, invalidaciones) |
| `clsOpportunityManager` | Gestiona colección de oportunidades (CRUD, búsqueda) |

### Capa 4: Dominio
| Clase | Responsabilidad |
|-------|----------------|
| `clsOpportunity` | Entidad: representa una oportunidad comercial |
| `clsOffer` | Entidad: representa una oferta de compresor |
| `clsOtherOffer` | Entidad: representa oferta de otro tipo |
| `clsOpportunityBudgetTemplate` | Plantilla: genera presupuestos |
| `clsOpportunityQuotationTemplate` | Plantilla: genera cotizaciones |
| `modOfferTypes` | Value Objects: enumeraciones de tipos de ofertas |

### Capa 5: Datos
| Clase | Responsabilidad |
|-------|----------------|
| `clsOfferRepository` | Repositorio: CRUD de ofertas desde Excel/DB |
| `clsDBContext` | Contexto: conexión y queries a base de datos |

### Capa 6: Presentación
| Clase/Módulo | Responsabilidad |
|--------------|----------------|
| `modRibbonCallbacks` | Callbacks: recibe eventos Ribbon y llama a Application |
| `frm*` | Formularios de usuario |
| `wsh*` | Worksheets con código (eventos) |

---

## PATRONES DE COMUNICACIÓN

### 1. Excel COM → Infraestructura (WithEvents)

```vba
' clsExcelExecutionContext (WithEvents Application)
Private WithEvents _app As Application

Private Sub _app_WorkbookOpen(ByVal Wb As Workbook)
    ' Callback DIRECTO (sin RaiseEvent)
    If Not _callbackHandler Is Nothing Then
        _callbackHandler.OnWorkbookOpened Wb
    End If
End Sub

Public Sub Initialize(ByVal callbackHandler As clsApplication)
    Set _callbackHandler = callbackHandler
    Set _app = Application
End Sub
```

**Patrón:** WithEvents COM → Callback directo al handler (clsApplication)

### 2. FileSystem COM → Eventos de Dominio (WithEvents + RaiseEvent)

```vba
' clsFileSystemWatcher (WithEvents FolderWatcher)
Private WithEvents _watcher As FolderWatcher

Private Sub _watcher_SubfolderCreated(ByVal path As String)
    RaiseEvent SubfolderCreated(path)  ' OK: adaptador de eventos
End Sub

' clsFileSystemMonitor (suscriptor de clsFileSystemWatcher)
Private WithEvents _watcher As clsFileSystemWatcher

Private Sub _watcher_SubfolderCreated(ByVal path As String)
    ' Interpretar: ¿Es una oportunidad?
    If IsOpportunityFolder(path) Then
        RaiseEvent OpportunityCreated(path)  ' OK: evento de DOMINIO
    End If
End Sub
```

**Patrón:** WithEvents COM → Adaptador → Evento Dominio (1-a-N potencial)

### 3. Servicio → Servicio (Llamada Directa)

```vba
' clsChartEventManager
Public Sub OnChartActivated(ByVal chart As Excel.Chart)
    ' Actualizar estado (PULL)
    _appState.ChartState.ActiveChart = chart
    
    ' Notificar a otro servicio (LLAMADA DIRECTA)
    _ribbonManager.RefreshChartButtons
End Sub
```

**Patrón:** Llamada directa entre servicios (sin eventos custom)

### 4. Servicio → Estado (Pull - Get/Set)

```vba
' clsFileManager
Public Sub TrackFile(ByVal filePath As String)
    ' Acceso directo al estado
    _appState.FileState.AddTrackedFile filePath
End Sub
```

**Patrón:** Servicios acceden a estado mediante properties (Pull)

### 5. UI → Aplicación (Push - Comando)

```vba
' modRibbonCallbacks
Public Sub OnButtonNewOpportunity(ByVal control As IRibbonControl)
    ' Llamada directa a Application
    App.CreateNewOpportunity
End Sub
```

**Patrón:** Callback desde UI ejecuta comando en Application (Push)

---

## GRAFO DE DEPENDENCIAS (limpio)

```
[Presentación]
    modRibbonCallbacks
         ↓ (comando)
    clsApplication
         ↓ (crea/inyecta)
    ┌────────────────────────────────┐
    │  [Infraestructura]             │
    │  clsServiceContainer           │
    │  clsApplicationState           │
    │  clsExcelExecutionContext ✓    │ WithEvents Application
    │  clsChartEventAdapter ✓        │ WithEvents Chart
    │  clsFileSystemWatcher ✓        │ WithEvents COM
    │  clsFileSystemMonitor ✓        │ RaiseEvent (dominio)
    └────────────────────────────────┘
         ↓ (inyecta)
    [Servicios]
    clsChartEventManager
    clsFileManager
    clsRibbonManager
    clsOpportunityManager
         ↓ (usa)
    [Dominio]
    clsOpportunity
    clsOffer
         ↓ (persiste)
    [Datos]
    clsOfferRepository
    clsDBContext
```

**Reglas:**
- ✅ Dependencias fluyen hacia abajo
- ✅ Infraestructura no conoce Dominio
- ✅ Dominio no conoce Infraestructura (inversión via inyección)

---

## EVENTOS: DECISIÓN FINAL

### ✅ MANTENER WithEvents (eventos COM)
1. `clsExcelExecutionContext` → `Application`
2. `clsChartEventAdapter` → `Chart`
3. `clsRibbonEventAdapter` → `IRibbonUI` (si necesario)
4. `clsFileSystemWatcher` → `FolderWatcher` (COM)

### ✅ MANTENER RaiseEvent (eventos de dominio)
1. `clsFileSystemMonitor`:
   - `OpportunityCreated`
   - `OpportunityDeleted`
   - `OpportunityRenamed`
   - `TemplateCreated/Changed`
   - `GasFileCreated/Changed`
   - `MonitoringError/Reconnected/Failed`

2. `clsFileSystemWatcher`:
   - Eventos adaptados del COM (intermediario necesario)

### ❌ ELIMINAR RaiseEvent (reemplazar por llamadas directas)
1. `clsChartEventManager`: eventos → métodos públicos
2. `clsExcelExecutionContext`: eventos → callbacks directos
3. `clsOpportunityManager`: eventos → métodos públicos
4. `clsRibbonEventAdapter`: eventos → callbacks directos
5. `clsRibbonState`: `StateChanged` → property setters

### ❌ ELIMINAR WithEvents (reemplazar por referencias directas)
1. `clsApplication` → TODOS los WithEvents (6 eliminados)
2. `clsFileManager` → `ctx`
3. `clsOpportunityManager` → `ctx`
4. `clsFSMonitoringCoord` → `mFolderWatcher` (convertir a referencia directa)

---

## NOMENCLATURA ESTÁNDAR

### Reglas
1. **Todo en inglés** (sin excepciones)
2. **Sin abreviaturas** (salvo siglas conocidas: DB, UI, FS, ID)
3. **Prefijos claros**:
   - `cls` → Clases
   - `mod` → Módulos estándar
   - `modUdf` → User Defined Functions
   - `modMacro` → Macros de usuario
   - `modUtils` → Utilidades
   - `frm` → Formularios
   - `wsh` → Worksheets
4. **Variables privadas**: `_` underscore (no `m` ni `m_`)
5. **Parámetros/locales**: sin prefijo
6. **Constantes**: `UPPER_SNAKE_CASE`

### Tabla de Renombrado Completo

| ANTIGUO | NUEVO | Tipo |
|---------|-------|------|
| `clsAplicacion` | `clsApplication` | Clase |
| `clsApplicationContext` | `clsApplicationState` | Clase |
| `clsExecutionContext` | `clsExcelExecutionContext` | Clase |
| `clsChartEvents` | `clsChartEventAdapter` | Clase |
| `clsChartEventsManager` | `clsChartEventManager` | Clase |
| `clsRibbonEvents` | `clsRibbonEventAdapter` | Clase |
| `clsRibbonUI` | `clsRibbonManager` | Clase |
| `clsFSWatcher` | `clsFileSystemWatcher` | Clase |
| `clsFSMonitoringCoord` | `clsFileSystemMonitor` | Clase |
| `clsOpportunitiesMgr` | `clsOpportunityManager` | Clase |
| `clsOferta` | `clsOffer` | Clase |
| `clsOfertaOtro` | `clsOtherOffer` | Clase |
| `clsOfertaRepository` | `clsOfferRepository` | Clase |
| `clsServiceManager` | `clsServiceContainer` | Clase |
| `clsConfiguration` | `clsApplicationConfiguration` | Clase |
| `modCALLBACKSRibbon` | `modRibbonCallbacks` | Módulo |
| `modMACROAppLifecycle` | `modMacroAppLifecycle` | Módulo |
| `modMACRO*` | `modMacro*` | Módulos |
| `modAPP*` | `modApp*` | Módulos |
| `modUTILS*` | `modUtils*` | Módulos |
| `UDFs_*` | `modUdf*` | Módulos |
| `mod_Logger` | `mod_Logger` | Módulo (OK) |
| `mod_ConstantsGlobals` | `mod_Constants` | Módulo |

### Variables Comunes

| ANTIGUO | NUEVO |
|---------|-------|
| `ctx` | `context` o `excelContext` |
| `fw` | `watcher` |
| `m_xlApp` | `_app` |
| `mChart` | `_chart` |
| `oTextBox` | `_textBox` |
| `mServiceManager` | `_serviceContainer` |
| `mAppContext` | `_appState` |

---

## VERIFICACIÓN DE ARQUITECTURA

### Checklist por Capa

#### Capa 1: Aplicación
- [ ] `clsApplication` NO tiene WithEvents de servicios
- [ ] `clsApplication` solo crea y conecta componentes
- [ ] `ThisWorkbook` solo delega a `clsApplication`

#### Capa 2: Infraestructura
- [ ] Solo 4 WithEvents COM (Excel, Chart, Ribbon, FileSystem)
- [ ] Estado separado de servicios
- [ ] Adaptadores solo adaptan (no lógica de negocio)

#### Capa 3: Servicios
- [ ] Servicios NO tienen WithEvents entre ellos
- [ ] Servicios acceden a estado via Pull (properties)
- [ ] Servicios se llaman directamente (no eventos custom)

#### Capa 4: Dominio
- [ ] Entidades puras (sin dependencias de infraestructura)
- [ ] Lógica de negocio en las entidades

#### Capa 5: Datos
- [ ] Repositorios NO conocen servicios
- [ ] Solo acceso a datos

#### Capa 6: Presentación
- [ ] Callbacks llaman directamente a `clsApplication`
- [ ] Sin lógica de negocio en callbacks

---

## SIGUIENTE PASO

Ver **DOCUMENTO 3: PLAN DE MIGRACIÓN** para la estrategia paso a paso.
