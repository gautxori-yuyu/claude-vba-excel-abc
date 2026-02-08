# VBA CLEAN ARCHITECTURE REFACTORING SKILL

**Version:** 1.0  
**Date:** 2026-01-22  
**Purpose:** Guide Claude Code through systematic VBA refactoring

---

## PROJECT CONTEXT

### Current State
- **Language:** VBA (Visual Basic for Applications) in Excel
- **Total Files:** 66 (28 classes, 35 modules, 3 forms)
- **Main Problem:** Over-engineering with unnecessary events and WithEvents
- **Architecture:** Mixed layers, inconsistent naming, tight coupling

### Goal
Transform into clean, maintainable architecture:
- Remove unnecessary custom events (RaiseEvent)
- Remove unnecessary WithEvents between services
- Organize by layers (Application, Infrastructure, Services, Domain, Data, Presentation)
- Consistent English naming
- Direct calls between services (not events)

---

## ARCHITECTURAL PRINCIPLES

### 1. Events: When to Use

✅ **USE WithEvents ONLY for:**
- COM events from Excel/Office objects (`Application`, `Chart`, `Worksheet`)
- COM events from external components (`FolderWatcher`)
- TOTAL: Only 4 classes should have WithEvents

✅ **USE RaiseEvent ONLY for:**
- Domain events with potential 1-to-N subscribers
- Example: `OpportunityCreated`, `OpportunityDeleted`
- Adapters wrapping COM events

❌ **DO NOT USE WithEvents for:**
- Communication between custom VBA services
- `clsApplication` subscribing to services (Composition Root should NOT listen)
- Services subscribing to other services

❌ **DO NOT USE RaiseEvent for:**
- 1-to-1 communication (use direct method calls instead)
- Callbacks that only have 1 subscriber

### 2. Communication Patterns

| From → To | Pattern | Example |
|-----------|---------|---------|
| UI → Application | **Direct Call** | `App.CreateNewOpportunity()` |
| Service → Service | **Direct Call** | `chartManager.RefreshButtons()` |
| Service → State | **Pull (Get/Set)** | `appState.ChartState.ActiveChart = chart` |
| Excel COM → Adapter | **WithEvents** | `Private WithEvents _app As Application` |
| Adapter → Application | **Direct Callback** | `_handler.OnWorkbookOpened(wb)` |
| FileSystem Monitor → Subscribers | **RaiseEvent (Domain)** | `RaiseEvent OpportunityCreated(path)` |

### 3. Layer Dependencies

```
Presentation → Application → Services → Domain → Data
     ↓              ↓             ↓
Infrastructure ← ← ← (injected)
```

**Rules:**
- Top layers can depend on lower layers
- Lower layers NEVER depend on upper layers
- Infrastructure injected via Composition Root (clsApplication)

---

## NAMING CONVENTIONS

### Files
- **Classes:** `cls` prefix (e.g., `clsApplication.cls`)
- **Modules:** `mod` prefix:
  - Standard modules: `mod*.bas`
  - Macros: `modMacro*.bas`
  - UDFs: `modUdf*.bas`
  - Utils: `modUtils*.bas`
- **Forms:** `frm` prefix (e.g., `frmConfiguration.frm`)
- **Worksheets:** `wsh` prefix (e.g., `wshUnits.cls`)

### Variables
- **Private fields:** `_` underscore prefix (e.g., `_serviceContainer`)
  - DO NOT use `m` or `m_` (legacy Hungarian notation)
- **Parameters/Locals:** no prefix (e.g., `filePath`, `workbook`)
- **Constants:** `UPPER_SNAKE_CASE` (e.g., `MAX_RETRY_COUNT`)

### Language
- **Everything in English** (no Spanish)
- **No abbreviations** except known acronyms (DB, UI, FS, ID)
- **Full words:** `manager` not `mgr`, `context` not `ctx`

---

## FOLDER STRUCTURE

```
/1-Application
    clsApplication.cls
    ThisWorkbook.cls

/2-Infrastructure
    /DependencyInjection
        clsServiceContainer.cls
        IService.cls
    /State
        clsApplicationState.cls
        clsChartState.cls
        clsRibbonState.cls
        clsFileState.cls
    /ExcelIntegration
        clsExcelExecutionContext.cls
        clsChartEventAdapter.cls
        clsRibbonEventAdapter.cls
    /FileSystem
        clsFileSystemWatcher.cls
        clsFileSystemMonitor.cls
    /Configuration
        clsApplicationConfiguration.cls
        frmConfiguracion.frm
    /Logging
        mod_Logger.bas
        mod_Constants.bas

/3-Services
    /ChartManagement
        clsChartEventManager.cls
    /FileManagement
        clsFileManager.cls
    /RibbonManagement
        clsRibbonManager.cls
    /OpportunityManagement
        clsOpportunityManager.cls

/4-Domain
    /Entities
        clsOpportunity.cls
        clsOffer.cls
        clsOtherOffer.cls
    /Templates
        clsOpportunityBudgetTemplate.cls
        clsOpportunityQuotationTemplate.cls
    /ValueObjects
        modOfferTypes.bas

/5-Data
    clsOfferRepository.cls
    clsDBContext.cls

/6-Presentation
    /RibbonCallbacks
        modRibbonCallbacks.bas
    /Forms
        frm*.frm
    /Worksheets
        wsh*.cls

/7-Macros
    modMacro*.bas

/8-UserDefinedFunctions
    modUdf*.bas

/9-Utilities
    modUtils*.bas
```

---

## CLASS RESPONSIBILITIES (One Line Each)

| Class | Responsibility |
|-------|---------------|
| `clsApplication` | Composition Root: creates and wires all components (NO WithEvents) |
| `clsServiceContainer` | DI Container: registers and resolves singleton services |
| `clsApplicationState` | Aggregates all state (ribbon, chart, file, execution) |
| `clsExcelExecutionContext` | WithEvents Application: captures Excel events → direct callbacks |
| `clsChartEventAdapter` | WithEvents Chart: captures chart events → direct callbacks |
| `clsFileSystemWatcher` | WithEvents FolderWatcher: captures COM FS events → RaiseEvent |
| `clsFileSystemMonitor` | Coordinator: interprets FS events → RaiseEvent domain events |
| `clsChartEventManager` | Processes chart events, updates state, calls other services |
| `clsFileManager` | Manages file tracking and states |
| `clsRibbonManager` | Manages IRibbonUI pointer and invalidations |
| `clsOpportunityManager` | Manages opportunity collection (CRUD, search) |

---

## REFACTORING CHECKLIST (Per Step)

### Before Starting Any Step
- [ ] Current code compiles without errors
- [ ] Identify which files need changes
- [ ] Understand dependencies

### After Each Change
- [ ] File compiles (`Debug > Compile VBAProject`)
- [ ] No broken references
- [ ] Functionality preserved (if applicable)

### After Each Phase
- [ ] All affected files compile
- [ ] Test basic functionality
- [ ] Commit to Git

---

## ANTI-PATTERNS TO AVOID

### ❌ DO NOT
1. Add WithEvents to `clsApplication` for services
2. Use RaiseEvent for 1-to-1 communication
3. Have services subscribe to other services with WithEvents
4. Mix Spanish and English
5. Use Hungarian notation (`m`, `m_`, `str`, `int`)
6. Put business logic in event handlers
7. Create circular dependencies

### ✅ DO
1. Use direct method calls between services
2. Pass dependencies via Initialize() method
3. Access state via Pull (properties)
4. Keep adapters thin (no business logic)
5. Use English names consistently
6. Document with @Folder and @Description
7. One responsibility per class

---

## COMMON PATTERNS

### Pattern 1: Remove WithEvents from clsApplication

**BEFORE:**
```vba
Private WithEvents mOpportunities As clsOpportunityManager

Private Sub mOpportunities_currOpportunityChanged(...)
    ' Logic here
End Sub
```

**AFTER:**
```vba
Private _opportunityManager As clsOpportunityManager

' Instead of event handler, expose public method
Public Sub OnCurrentOpportunityChanged(...)
    ' Same logic here
End Sub

' Service calls directly when needed
' (in clsOpportunityManager)
Public Sub SetCurrentOpportunity(...)
    _currentOpportunity = value
    
    ' Direct call instead of RaiseEvent
    _application.OnCurrentOpportunityChanged _currentOpportunity
End Sub
```

### Pattern 2: COM Events → Direct Callback

**BEFORE:**
```vba
' In clsExcelExecutionContext
Public Event WorkbookOpened(ByVal wb As Workbook)

Private WithEvents _app As Application

Private Sub _app_WorkbookOpen(ByVal wb As Workbook)
    RaiseEvent WorkbookOpened(wb)
End Sub
```

**AFTER:**
```vba
' In clsExcelExecutionContext
Private WithEvents _app As Application
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

```vba
' In clsApplication
Public Sub OnWorkbookOpened(ByVal wb As Workbook)
    ' Logic that was in event handler
End Sub

Private Sub Initialize()
    Set _excelContext = New clsExcelExecutionContext
    _excelContext.Initialize Me  ' Pass self as handler
End Sub
```

### Pattern 3: Service → Service Communication

**BEFORE:**
```vba
' clsChartEventManager
Public Event ChartActivated(...)

Public Sub OnChartActivated(...)
    RaiseEvent ChartActivated(...)
End Sub
```

**AFTER:**
```vba
' clsChartEventManager
Private _ribbonManager As clsRibbonManager
Private _appState As clsApplicationState

Public Sub Initialize(appState As clsApplicationState, ribbonMgr As clsRibbonManager)
    Set _appState = appState
    Set _ribbonManager = ribbonMgr
End Sub

Public Sub OnChartActivated(ByVal chart As Excel.Chart)
    ' Update state (Pull)
    _appState.ChartState.ActiveChart = chart
    
    ' Direct call to other service
    _ribbonManager.RefreshChartButtons
End Sub
```

---

## RENAMING TABLE (OLD → NEW)

### Classes
| Old Name | New Name |
|----------|----------|
| `clsAplicacion` | `clsApplication` |
| `clsApplicationContext` | `clsApplicationState` |
| `clsServiceManager` | `clsServiceContainer` |
| `clsExecutionContext` | `clsExcelExecutionContext` |
| `clsChartEvents` | `clsChartEventAdapter` |
| `clsChartEventsManager` | `clsChartEventManager` |
| `clsRibbonEvents` | `clsRibbonEventAdapter` |
| `clsRibbonUI` | `clsRibbonManager` |
| `clsFSWatcher` | `clsFileSystemWatcher` |
| `clsFSMonitoringCoord` | `clsFileSystemMonitor` |
| `clsOpportunitiesMgr` | `clsOpportunityManager` |
| `clsOferta` | `clsOffer` |
| `clsOfertaOtro` | `clsOtherOffer` |
| `clsOfertaRepository` | `clsOfferRepository` |
| `clsConfiguration` | `clsApplicationConfiguration` |

### Modules
| Old Pattern | New Pattern |
|-------------|-------------|
| `modCALLBACKSRibbon` | `modRibbonCallbacks` |
| `modMACRO*` | `modMacro*` |
| `modAPP*` | `modApp*` |
| `modUTILS*` | `modUtils*` |
| `UDFs_*` | `modUdf*` |
| `mod_ConstantsGlobals` | `mod_Constants` |

### Variables
| Old | New |
|-----|-----|
| `ctx` | `context` or `excelContext` |
| `fw` | `watcher` |
| `m_xlApp` | `_app` |
| `mChart` | `_chart` |
| `mServiceManager` | `_serviceContainer` |
| `mAppContext` | `_appState` |

---

## STEP-BY-STEP EXECUTION GUIDE

### How to Use This Skill with Claude Code

1. **Read this entire skill first**
2. **Open the PLAN_MIGRACION.md document**
3. **Execute each phase sequentially**
4. **Verify compilation after EACH file change**
5. **Commit after EACH phase**

### Verification Command (VBA Editor)
```
Debug > Compile VBAProject
```
**Expected:** No errors

### Testing After Phases
- **After Phase 2:** Application initializes without errors
- **After Phase 5:** Domain operations work (create opportunity, etc)
- **After Phase 6:** UI callbacks work (Ribbon buttons)
- **After Phase 9:** Full application functional

---

## CRITICAL REMINDERS

1. **VBA is NOT .NET:** 
   - No automatic properties
   - No generics
   - Limited OOP features
   - Keep it simple

2. **WithEvents has limitations:**
   - Only for COM objects and VBA class modules
   - Cannot pass interfaces
   - Cannot be late-bound

3. **Compilation is mandatory:**
   - VBA errors are runtime by default
   - Always compile after changes
   - Fix all warnings

4. **Git workflow:**
   - Work in feature branch
   - Small, atomic commits
   - Meaningful commit messages
   - Test before merge

5. **When stuck:**
   - Check ANALISIS_ARQUITECTONICO.md for context
   - Check ARQUITECTURA_OBJETIVO.md for target
   - Verify against this skill's patterns
   - Ask for clarification (don't guess)

---

## SUCCESS CRITERIA

### Refactoring is complete when:

✅ **Architecture**
- [ ] Only 4 classes have WithEvents (Excel, Chart, Ribbon, FileSystem COM)
- [ ] clsApplication has NO WithEvents to services
- [ ] Services communicate via direct calls
- [ ] State accessed via Pull (properties)
- [ ] Domain events properly defined

✅ **Organization**
- [ ] All files in correct folders
- [ ] @Folder annotations updated
- [ ] Naming consistent (English, no abbreviations)

✅ **Quality**
- [ ] Debug > Compile → No errors
- [ ] No circular dependencies
- [ ] No code duplication
- [ ] All functionality works

✅ **Documentation**
- [ ] Each class has @Description
- [ ] Complex methods have comments
- [ ] README updated with new structure

---

## REFERENCES

- **Analysis:** `01_ANALISIS_ARQUITECTONICO.md`
- **Target:** `02_ARQUITECTURA_OBJETIVO.md`
- **Migration Plan:** `03_PLAN_MIGRACION.md`
- **This Skill:** `04_SKILL_CLAUDE_CODE.md`

---

**Last Updated:** 2026-01-22  
**Author:** Sergio (with Claude assistance)  
**License:** Internal use only
