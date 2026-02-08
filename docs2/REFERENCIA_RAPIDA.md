# TABLA DE REFERENCIA RÁPIDA

## EVENTOS: DECISIÓN FINAL

### ✅ MANTENER WithEvents (4 clases)

| Clase | WithEvents Variable | Tipo | Justificación |
|-------|-------------------|------|---------------|
| `clsExcelExecutionContext` | `_app` | `Application` | ✅ Eventos COM Excel (Workbook, Worksheet) |
| `clsChartEventAdapter` | `_chart` | `Chart` | ✅ Eventos COM Chart |
| `clsRibbonEventAdapter` | `_ribbon` | `IRibbonUI` | ✅ Eventos COM Ribbon (si usado) |
| `clsFileSystemWatcher` | `_watcher` | `FolderWatcher` | ✅ Eventos COM componente externo |

### ✅ MANTENER RaiseEvent (2 clases - eventos de dominio)

| Clase | Eventos | Justificación |
|-------|---------|---------------|
| `clsFileSystemMonitor` | `OpportunityCreated`<br>`OpportunityDeleted`<br>`OpportunityRenamed`<br>`TemplateCreated`<br>`GasFileCreated`<br>(etc) | ✅ Eventos de DOMINIO<br>✅ Potencial 1-a-N<br>✅ Separa infraestructura de dominio |
| `clsFileSystemWatcher` | `FileCreated`<br>`FileDeleted`<br>`SubfolderCreated`<br>(etc) | ✅ Adaptador de eventos COM<br>✅ Intermediario necesario |

### ❌ ELIMINAR WithEvents (6 en clsApplication + 2 en servicios)

| Clase | WithEvents a ELIMINAR | ❌ Por qué | ✅ Reemplazo |
|-------|----------------------|-----------|-------------|
| `clsApplication` | `mOpportunities` | Composition Root no escucha servicios | Llamada directa |
| `clsApplication` | `mChartManager` | Ídem | Llamada directa |
| `clsApplication` | `mFSMonitoringCoord` | Ídem | Llamada directa |
| `clsApplication` | `mRibbonState` | Ídem | Property setters |
| `clsApplication` | `evRibbon` | Ídem | Llamada directa |
| `clsApplication` | `ctx` | Ídem | Callback directo |
| `clsFileManager` | `ctx` | Servicio no escucha otro servicio | Inyección por Initialize |
| `clsOpportunityManager` | `ctx` | Ídem | Inyección por Initialize |

### ❌ ELIMINAR RaiseEvent (5 clases - eventos 1-a-1)

| Clase | Eventos a ELIMINAR | ❌ Por qué | ✅ Reemplazo |
|-------|-------------------|-----------|-------------|
| `clsChartEventManager` | `ChartActivated`<br>`ChartDeactivated`<br>`HojaConGraficosCambiada` | Solo 1 suscriptor (clsApplication) | Métodos públicos |
| `clsExcelExecutionContext` | `WorkbookOpened`<br>`WorkbookActivated`<br>`WorksheetActivated`<br>(etc - 8 eventos) | Solo 1 suscriptor | Callbacks directos |
| `clsOpportunityManager` | `currOpportunityChanged`<br>`OpportunityCollectionUpdate` | Solo 1 suscriptor | Métodos públicos |
| `clsRibbonEventAdapter` | `GenerarGraficosDesdeCurvasRto`<br>`InvertirEjes`<br>(etc - 6 eventos) | Solo 1 suscriptor | Callbacks directos |
| `clsRibbonState` | `StateChanged` | Solo 1 suscriptor | Property setter |

---

## RENOMBRADO COMPLETO

### Clases Principales

| ANTIGUO | NUEVO | Carpeta |
|---------|-------|---------|
| `clsAplicacion` | `clsApplication` | `/1-Application` |
| `clsApplicationContext` | `clsApplicationState` | `/2-Infrastructure/State` |
| `clsServiceManager` | `clsServiceContainer` | `/2-Infrastructure/DependencyInjection` |
| `clsExecutionContext` | `clsExcelExecutionContext` | `/2-Infrastructure/ExcelIntegration` |
| `clsChartEvents` | `clsChartEventAdapter` | `/2-Infrastructure/ExcelIntegration` |
| `clsChartEventsManager` | `clsChartEventManager` | `/3-Services/ChartManagement` |
| `clsRibbonEvents` | `clsRibbonEventAdapter` | `/2-Infrastructure/ExcelIntegration` |
| `clsRibbonUI` | `clsRibbonManager` | `/3-Services/RibbonManagement` |
| `clsFSWatcher` | `clsFileSystemWatcher` | `/2-Infrastructure/FileSystem` |
| `clsFSMonitoringCoord` | `clsFileSystemMonitor` | `/2-Infrastructure/FileSystem` |
| `clsOpportunitiesMgr` | `clsOpportunityManager` | `/3-Services/OpportunityManagement` |
| `clsOferta` | `clsOffer` | `/4-Domain/Entities` |
| `clsOfertaOtro` | `clsOtherOffer` | `/4-Domain/Entities` |
| `clsOfertaRepository` | `clsOfferRepository` | `/5-Data` |
| `clsConfiguration` | `clsApplicationConfiguration` | `/2-Infrastructure/Configuration` |
| `clsOpportunityOfferBudgetTpl` | `clsOpportunityBudgetTemplate` | `/4-Domain/Templates` |
| `clsOpportunityOfferQuotationTpl` | `clsOpportunityQuotationTemplate` | `/4-Domain/Templates` |

### Estado (separado de servicios)

| ANTIGUO | NUEVO | Carpeta |
|---------|-------|---------|
| `clsChartState` | `clsChartState` (OK) | `/2-Infrastructure/State` |
| `clsRibbonState` | `clsRibbonState` (OK) | `/2-Infrastructure/State` |
| `clsFileState` | `clsFileState` (OK) | `/2-Infrastructure/State` |

### Módulos

| ANTIGUO | NUEVO | Carpeta |
|---------|-------|---------|
| `modCALLBACKSRibbon` | `modRibbonCallbacks` | `/6-Presentation/RibbonCallbacks` |
| `modMACROAppLifecycle` | `modMacroAppLifecycle` | `/7-Macros` |
| `modMACRO*` | `modMacro*` | `/7-Macros` |
| `modAPP*` | `modApp*` | `/1-Application` o según función |
| `modUTILS*` | `modUtils*` | `/9-Utilities` |
| `UDFs_*` | `modUdf*` | `/8-UserDefinedFunctions` |
| `mod_ConstantsGlobals` | `mod_Constants` | `/2-Infrastructure/Logging` |

### Variables Comunes

| ANTIGUO | NUEVO | Uso |
|---------|-------|-----|
| `mServiceManager` | `_serviceContainer` | `clsApplication` |
| `mAppContext` | `_appState` | `clsApplication` |
| `ctx` | `context` o `_excelContext` | Varios |
| `fw` | `_watcher` | `clsFileSystemMonitor` |
| `m_xlApp` | `_app` | `clsExcelExecutionContext` |
| `mChart` | `_chart` | `clsChartEventAdapter` |
| `oTextBox` | `_textBox` | `CRefEdit` |
| `mOpportunities` | `_opportunityManager` | `clsApplication` |
| `mChartManager` | `_chartEventManager` | `clsApplication` |

---

## PATRONES DE COMUNICACIÓN

### Matriz de Patrones

| Desde | Hacia | Patrón | Ejemplo |
|-------|-------|--------|---------|
| **UI** | Application | Direct Call (Push) | `App.CreateNewOpportunity()` |
| **Application** | Service | Direct Call | `_chartEventManager.OnChartActivated(chart)` |
| **Service** | Service | Direct Call | `_ribbonManager.RefreshButtons()` |
| **Service** | State | Pull (Get/Set) | `_appState.ChartState.ActiveChart = chart` |
| **Service** | Repository | Pull (Query) | `_offerRepo.GetOfferById(id)` |
| **Excel COM** | Adapter | WithEvents | `Private WithEvents _app As Application` |
| **Adapter** | Application | Direct Callback | `_handler.OnWorkbookOpened(wb)` |
| **FS Monitor** | Subscribers | RaiseEvent (Domain) | `RaiseEvent OpportunityCreated(path)` |

### Código de Ejemplo por Patrón

#### Patrón 1: Eliminar WithEvents de clsApplication
```vba
' ANTES
Private WithEvents mOpportunities As clsOpportunityManager
Private Sub mOpportunities_currOpportunityChanged()
    ' Logic
End Sub

' DESPUÉS
Private _opportunityManager As clsOpportunityManager
Public Sub OnCurrentOpportunityChanged()
    ' Same logic
End Sub
' Service llama directamente cuando necesita
```

#### Patrón 2: COM Event → Direct Callback
```vba
' clsExcelExecutionContext
Private _callbackHandler As clsApplication
Public Sub Initialize(ByVal handler As clsApplication)
    Set _callbackHandler = handler
    Set _app = Application
End Sub

Private Sub _app_WorkbookOpen(ByVal wb As Workbook)
    _callbackHandler.OnWorkbookOpened wb
End Sub
```

#### Patrón 3: Service → Service Direct Call
```vba
' clsChartEventManager
Public Sub OnChartActivated(ByVal chart As Excel.Chart)
    _appState.ChartState.ActiveChart = chart
    _ribbonManager.RefreshChartButtons  ' Direct call
End Sub
```

#### Patrón 4: Service → State Pull
```vba
' clsFileManager
Public Sub TrackFile(ByVal path As String)
    _appState.FileState.AddTrackedFile path  ' Property access
End Sub
```

---

## FASES DE MIGRACIÓN (RESUMEN)

| Fase | Objetivo | Duración | Verificación |
|------|----------|----------|--------------|
| 0 | Preparación (backup, branch) | 15 min | Git branch creado |
| 1 | Renombrar archivos | 2-3 horas | Compila sin errores |
| 2 | Refactorizar Application | 3-4 horas | Compila, app inicia |
| 3 | Refactorizar Infraestructura | 1-2 horas | Compila |
| 4 | Refactorizar Servicios | 2-3 horas | Compila, funcionalidad OK |
| 5 | Refactorizar Dominio/Datos | 1-2 horas | Compila |
| 6 | Refactorizar Presentación | 1 hora | Compila, UI funciona |
| 7 | Organizar Macros/UDFs | 1 hora | Compila |
| 8 | Utilidades | 30 min | Compila |
| 9 | Limpieza final | 1 hora | Todo funciona |

**Total:** 8-12 horas (con Claude Code), 20-30 horas (manual)

---

## CHECKLIST RÁPIDO

### Antes de empezar
- [ ] Código actual compila
- [ ] Backup completo
- [ ] Git configurado
- [ ] He leído los 4 documentos

### Durante cada fase
- [ ] Cambios según plan
- [ ] Compila después de cada cambio
- [ ] Funcionalidad preservada

### Después de cada fase
- [ ] Debug > Compile → Sin errores
- [ ] Test funcionalidad básica
- [ ] Git commit

### Al terminar
- [ ] Solo 4 WithEvents (COM)
- [ ] Eventos custom solo dominio
- [ ] Todo en inglés
- [ ] Organización por capas
- [ ] Funcionalidad 100%

---

## ANTI-PATTERNS A EVITAR

❌ **NO HACER:**
1. WithEvents en clsApplication para servicios
2. RaiseEvent para comunicación 1-a-1
3. Servicios con WithEvents de otros servicios
4. Mezclar español e inglés
5. Usar notación húngara (m_, str, int)
6. Lógica de negocio en event handlers
7. Dependencias circulares

✅ **SÍ HACER:**
1. Llamadas directas entre servicios
2. Inyección de dependencias via Initialize
3. Estado accedido via Pull (properties)
4. Adaptadores delgados (sin lógica)
5. Nombres en inglés consistentes
6. Documentar con @Folder/@Description
7. Una responsabilidad por clase

---

## ARCHIVOS DE REFERENCIA

1. **Diagnóstico:** `01_ANALISIS_ARQUITECTONICO.md`
2. **Diseño:** `02_ARQUITECTURA_OBJETIVO.md`
3. **Ejecución:** `03_PLAN_MIGRACION.md`
4. **Skill:** `04_SKILL_CLAUDE_CODE.md`
5. **Resumen:** `README_REFACTORIZACION.md`
6. **Esta tabla:** `REFERENCIA_RAPIDA.md`
