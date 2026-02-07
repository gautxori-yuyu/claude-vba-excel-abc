# como deberia ser, en terminos generales, la arquitectura de la aplicacion
TODOS los cambios que vayamos haciendo tienen que llevarnos a ese diseño.
## CAPAS de la aplicación
Dominio: No conoce Workbook, Worksheet, Range, Application. Trabaja solo con abstracciones propias. Define qué necesita, no cómo se obtiene.
Infraestructura: Conoce Excel sin pudor. Implementa cómo se obtiene lo que el dominio pide. Traduce Excel → contratos del dominio.

### CAPA DE INFRAESTRUCTURA
"La infraestructura puede hablar en lenguaje de dominio, si implementa contratos definidos por el dominio". Esto quiere decir que:
- En infraestructura, además de clases como clsExecutionContext, clsFileMgr (FileSystemAdapter), aparecerán:
```
ExcelWorkbookAdapter (usa Workbook real)
├─ PdfFileAdapter
├─ FileSystemAdapter
└─ clsExecutionContext
```
y apareceran clases llamadas ExcelTechCalcSource, ExcelQuoteSource, ExcelOpportunitySource… (cada una haciendo referencia a distintos tipos de ficheros de excel, asociados a entidades (clases) de dominio), para FACILITAR LA INVERSION DE DEPENDENCIAS.
el nombre “ExcelOpportunitySource” expresa dos cosas distintas:
Parte del nombre: "Opportunity" - indica Contrato de dominio que implementa
Parte del nombre: "Excel" - indica Tecnología concreta usada

- Las clases de infraestructura NO deben conocer a la clase mediadora. La mediadora SÍ debe conocer a las clases de infraestructura.

### CAPA DE DOMINIO
- En DOMINIO NO pueden aparecer declaraciones como. "Public Sub RecalcularOferta(wb As Workbook) ...": NO debe haber referencias a entidades de Excel como Workbook
Ejemplo de uso legítimo, pero encapsulado: "Public Sub RecalcularOferta(ctx As IExcelSession)".

En dominio se declararan interfaces como:
' IOpportunitySource
Function LeerDatosOportunidad() As OpportunityDTO
Sub GuardarResultado(ByVal resultado As ResultadoDTO)
... y en infraestructura se implementaran:
' IExcelOpportunitySource
Private mWb As Workbook
IMplements IExcelOpportunitySource
Private Property Get IExcelOpportunitySource_LeerDatosOportunidad() As OpportunityDTO
IExcelOpportunitySource_LeerDatosOportunidad = Me.LeerDatosOportunidad
End Property
Function LeerDatosOportunidad() As OpportunityDTO
' aquí sí hay Worksheet, Range, etc.
End Function

De manera que luego el dominio, para usar excel, llama a la interfaz. El dominio define "el contrato":
+ qué es una Opportunity
+ qué operaciones necesita sobre ella
+ la interfaz (IOpportunitySource)
La infraestructura solo responde: “Ah, ¿esto es lo que entiendes por Opportunity? Vale, yo sé leerlo desde Excel.”. es decir, hace el mapeado tecnología (infraestructura) → contrato.

Todo ello, además, nos permitirá *gestionar el ciclo de vida DESDE LA INFRAESTRUCTURA*, y *garantizar desde ella la estabilidad de la aplicación*.

Asi que ya te haces a la idea: **VAMOS A TENER QUE IMPLEMENTAR UNAS CUANTAS INTERFACES** (aka "puertos"; pero tu cuando te dirijas a mi usa el termino anterior). Por dentro, las entidades de infraestructura que las implementan, usan Workbook, Worksheet, Range, fórmulas, tablas, lo que haga falta. Pero el dominio no sabe cómo.

## TIPOS DE EVENTOS. Escucha de eventos.
Eventos TÉCNICOS: Excel, FileSystem, COM, Host. Ejemplos: WorkbookOpen, SheetActivate. NO son eventos de dominio
Eventos SEMÁNTICOS DE INFRAESTRUCTURA, (traducción de los técnicos). Ejemplos:, WorkbookSessionStarted, ActiveChartChanged, FileBecameReadOnly. Estos SÍ los puede emitir el mediador
Eventos DE DOMINIO, Ejemplos: OpportunityActivated, QuoteCalculated, OfferInvalidated. Estos SOLO los emiten los Mgrs de dominio.

- QUIEN "ESCUCHA EVENTOS"? la INTERFAZ, y Otros Mgrs (si procede):
  ** FileMgr → OpportunitiesMgr
  ** ExecutionContext → ChartsMgr
  Pero nunca Estado → Mgr.
  Y esas clases NO HACEN UNA ESCUCHA DIRECTA!!: la hacen A TRAVES DE LOS MEDIADORES (aunque tenemos que aclarar COMO hace la escucha la interfaz, ...)

## CLASES
### clases mediadoras
- las clases mediadoras (Event Aggregator / Coordinator) NO emiten eventos "de negocio". Si pueden emitir otro tipo de eventos, más adelante se indica con qué fin. 
insisto: EN NINGUN MEDIADOR DEBE HABER UN RaiseEvent. El mediador: Orquesta, Conecta, Reenvía, o Coordina ... Pero NO es dueño de nada. Eso implica: conocer quién publica; conocer quién consume; gestionar el flujo.

### clases Manager
- la parte "Mgr" de los subsistemas (UI, Ribbon, Formularios, y a veces infraestructura (Excel abre un libro)) (las que NO son de estado, ni mediadoras), ORDENAN, siempre mediante métodos hacer algo. Los métodos on imperativos, van top → down o en transversal (entre clases de la misma capa) (eso ya lo sabias, verdad). el Mgr del subsistema DA ORDENES (ejecuta métodos), y cuando estos se terminan de ejecutar, CAMBIA ESTADO, y LEVANTA EVENTOS (RaiseEvent xx) para indicar que ha terminado.

INSISTO: Managers (y si existen, Servicios de dominio / infraestructura) CAMBIAN estado, y después notifican con eventos.

El que cambia el estado, SIEMPRE, emite el evento. El Mgr no debe levantar eventos sobre estado que no controla. El Mgr llama a su State y luego emite el evento. El mediador no emite eventos de negocio. NO notifica "cualquier Mgr": "El que notifica es el que entiende el significado".
  el flujo DEBE ser:
  [ Evento Excel / FS ]
  ↓
  [ Mediador ]
  ↓
  [ Mgr ]
  ↓
  [ Actualiza State ]
  ↓
  [ RaiseEvent semántico ]

## ESTADOS. Clases de estado.
- Se deben diferenciar (y lo estamos haciendo) ESTADOS DE NEGOCIO, de ESTADOS TECNICOS o de infraestructura. Y NO se deben mezclar.
  las clases State NO pueden implementar métodos para “ordenar”. Nunca.
  Las clases State: no toman decisiones; no validan; no lanzan eventos; no ordenan nada. Son memoria estructurada, no comportamiento.
  Si una State tiene métodos, solo pueden ser: RegisterX / Clear / Snapshot / Reset. Y solo llamados por su Mgr.
  Insisto: LAS CLASES STATE NO DEBEN LANZAR EVENTOS.
  clsApplicationState es resultado, no origen: Consolida estado, Permite consultas coherentes, Facilita UI, logging, decisiones. Pero no controla nada.
  Las clases State son snapshots, acumuladores, modelos de estado derivado. Si una clase State lanza eventos: 1. mezclas modelo con control. 2. rompes trazabilidad. 3. haces imposible reproducir el estado.
  Por tanto: si una ORDEN proviniente de una clase Mgr, puede cambiar un estado (=una variable de estado) por ejemplo de un subsistema de dominio... que al cambiar, conlleve una orden a un subsistema de infraestructura... : en principio, y salvo que este planteamiento contradiga alguno de los preceptos restantes, esa clase de estado, debería ordenar a su "clase padre, la Mgr", para que sea ella la que de la orden al subsistema.
- las responsabilidades separadas entre las clases de estado de subsistemas, y la de estado general, estan BIEN diseñadas. PERO las clases de estado de subsistemas NO TIENEN QUE EMITIR EVENTOS: clsApplicationState, OBSERVA, no gobierna: NO da ordenes ni emite eventos.

- CUALQUIER CAMBIO / proceso de cambio, en el sistema, DEBE partir de UNA SOLA clase Mgr: NO debe haber NINGUN CAMBIO (de estado) que, en función del evento / callback del que se parta, esté supeditado a distintas clases Mgr.

## CAPA DE INTERFAZ. Ribbon
- Si infraestructura conoce al Ribbon, se crean dependencias circulares lógicas y se pierde capacidad de evolución del sistema (aunque hoy “funcione”)
- LA CAPA DE INTERFAZ, UI, NO ANUNCIA NADA. Solo EJECUTA ORDENES (metodos). TODOS los componentes de la capa de interfaz, **INCLUIDO EL RIBBON**, están en la capa más superior: es la forma de garantizar que NO haya dependencias de ella. La aplicación DEBE poder ejecutarse, independientemente de la interfaz que eimplemente, o de si la implementa o NO!! (si desactivo el ribbon, NO deba haber NINGUNA ORDEN, en ninguna clase, dependiente del ribbon, que desestabilice la aplicación).

TODO el trabajo de la INTERFAZ, de la UI (Ribbon, Forms, Panels), DEPENDE DE EVENTOS: al estar en la capa más elevada, su trabajo se guía por ellos.

La interfaz conoce al mundo; el mundo no conoce a la interfaz. Nada que no sea UI llama a código de UI. Esto incluye: infraestructura, dominio, coordinadores, y event aggregators. La UI es "el consumidor final".
La UI solo escucha eventos y decide cómo pintarse.
  
### ¿cómo se invalida el Ribbon? :
- Incorrecto: "mRibbon.InvalidateControl "btnGuardar"". Esto está mal porque: infraestructura sabe que hay Ribbon, sabe que hay botones, sabe que hay UI. Eso no es su problema.
- Correcto (patrón limpio): Infraestructura levanta un evento semántico, La UI decide qué hacer con él:
  ' Infraestructura / Dominio:
  RaiseEvent OpportunityStateChanged(opId, newState)
  ' Interfaz (Ribbon):
  Private Sub mApp_OpportunityStateChanged(...)
  Me.Invalidate
  End Sub
  → El Ribbon se invalida solo.

### escucha de eventos
  * LA INTERFAZ PUEDE ESCUCHAR TODO TIPO DE EVENTOS, incluso los de nivel de infraestructura. Por ejemplo,
  ' Infraestructura:
  RaiseEvent WorkbookSessionStarted(path, readOnly)
  ' Interfaz:
  Private Sub mApp_WorkbookSessionStarted(...)
  Me.Invalidate
  End Sub

## METODOS PROPERTY GET / SET / LET
- Property Get:
  ** Dónde deben vivir: Clases State o Clases Manager (solo como fachada)
  ** Qué hacen: Exponen estado actual. NO disparan eventos. No ordenan nada.
  Ejemplo correcto:
  ' FileMgrState
  Public Property Get CurrentFileId() As String
  CurrentFileId = mCurrentFileId
  End Property
  ' O como fachada: FileMgr
  Public Property Get CurrentFileId() As String
  CurrentFileId = mState.CurrentFileId
  End Property
  → El Get no es una orden. Es una consulta.

- Property Let / Set: IMPORTANTE!!: NO DEBEN APARECER En clases State!!! Eso convierte el estado en anárquico: cualquiera puede mutarlo.
  SI PUEDEN APARECER En clases Manager, con condiciones: Un Let/Set en un Mgr es equivalente a un método imperativo.
  ' OpportunitiesMgr
  Public Property Let CurrentOpportunity(id As String)
  mState.SetCurrentOpportunity id
  RaiseEvent OpportunityActivated(id)
  End Property
  Conceptualmente, esto ES una orden, aunque sintácticamente sea un Property. No lo abuses. Úsalo solo si: el significado es “establecer el actual”; no hay lógica compleja; Si hay lógica → método explícito.


# IMPLEMENTACION del nuevo modelo arquitectonico
## DECISIONES que deben quedar reflejadas en la implementacion
- hay un error de planteamiento en mis propuestas anteriores, de cómo gestionar eventos. Un "sistema" (entendiendo por tal un conjunto "clase Mgr" + clase de estado) NO se genera eventos RaiseEvent para señalar que SE INICIA una acción, se registran cuando SE HA COMPLETADO la acción que lleva a un estado. Por tanto, NO debemos implementar NINGUN evento "RaiseEvent OpportunityChange", ni desde clases de estado, ni desde clases Mgr. Un evento es una notificación de que algo ya ha ocurrido.

- TODOS los raiseevent DEBEN llevarse a clases Mgr.

- LA UI, el Ribbon, DEBE "CAMBIAR DE POSICION". NO puede aparecer, para nada, en las capas de infraestructura o dominio!!!

Así, por ejemplo, ante un evento WorkbookOpened, el flujo sería el siguiente:

' Mediador
Private Sub mCtx_WorkbookOpened(path As String)
mFileMgr.RegisterWorkbook path
End Sub
' FileMgr
Public Sub RegisterWorkbook(path As String)
mState.RegisterWorkbook path
RaiseEvent WorkbookRegistered(mState.CurrentFileId)
End Sub

En consecuencia, tHISwORKBOOK y clsExecutionContext nos va a ayudar a controlar la recuperacion de la aplicacion ("Reset") en caso de errores en eventos, pérdidas de sink COM, etc. Actualmente se implementa PARCIALMENTE un proceso de recuperacion, en base a modMACROAppLifecycle, Thisworkbook,

## OBJETIVOS:

1. **Gestionar mejor el ESTADO DE LA APLICACION**:
  - *el ESTADO DE LA APLICACION debe quedar correctamente registrado en clsApplicationState*, con expresiones como "Public Property Get IsInitialized() As Boolean"
    + deben registrarse y detectarse resets, pej con variable Static en módulo estándar, o con contador de inicializaciones.
      clsExecutionContext DEBE permitir: reenganchar eventos, reconstruir estado derivado, aislar el daño del reset. Pero NO evita el reset.
      Además, la mediadora debe poder llamar a App() cuando detecta un reset. No tengo claro si solo la de infraestructura, creo que sí. Pero acláramelo tú.
      Y hay que perfeccionar el cache que hace clsExecutionContext, que en parte reconstruye estado derivado con esa caché, pero mezcla ese proceso con lógica de evento. y debe también aislar el daño del reset, Que si Excel/VBA se resetea, solo se rompa infraestructura, NO dominio, NO UI, NO coordinadores. Pare ello:
    + hay que añadir en el control de Estado de Aplicación algo como "Public Property Get IsContextValid() As Boolean",
    + y hay uque generar y hacer seguimiento en infraestructura de eventos como "Public Event ContextInvalidated()", "Public Event ContextReinitialized()". El mediador reaccionará a esos eventos, vuelve a inicializar, reengancha, y notifica al dominio. Eso SOLO puede hacerlo una clase como clsExecutionContext.

  - para conseguir lo anterior, *clsExecutionContext DEBE DESCOMPONERSE* en 3 clases:
    + una parte de su funcionalidad (escucha de eventos) DEBE ir a **un "Listener"** que implemente WithEvents Application y capture eventos técnicos. NO DEBE ser el mediador de infraestructura, conviene implementarlo en ThisWorkbook, CON Responsabilidad única: Capturar eventos técnicos del host Y Reenviarlos sin semántica.
    dE MODO QUE tHISwORKBOOK QUEDARÍA ALGO ASÍ (a añadir a lo que ya tiene):
    ```
    ' ThisWorkbook
    Private Sub Workbook_Open()
    ExecutionContextMgr.Initialize
    End Sub
    Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ExecutionContextMgr.Shutdown
    End Sub
    Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ExecutionContextMgr.OnSheetActivated Sh
    End Sub
    ' Y para gestionar Application: ojo a la diferencia entre mXLApp (Application) y mApp (clsApplication, accesible con App())
    Private WithEvents mXLApp As Application
    Private Sub Workbook_Open()
    Set mXLApp = Application
    End Sub
    Private Sub mXLApp_WorkbookOpen(ByVal Wb As Workbook)
    ExecutionContextMgr.OnWorkbookOpened Wb
    End Sub
    ```
    **Implementando el listener en Thisworkbook** (sin crear una clase adicional), se evita tener que hacer "rebinds" en muchos casos: los hace ese modulo automaticamente.

    + **clsExecutionContextState**: que determina ActiveWorkbook, ActiveChart, Flags técnicos. Sin eventos!
    + y **clsExecutionContextMgr**: Recibe eventos del Listener (para ello el Listener llama a sus métodos), Actualiza State, y Lanza eventos semánticos de infraestructura; y por último,... **_ maneja reenganche / reset _**.

    Como consecuencia de ello, el Mediador de infraestructura escucha eventos del ExecutionContextMgr, y Decide a qué Mgrs llamar.

    Mapa final
    ```
    Excel
    ↓ (eventos técnicos)
    ExecutionContextListener
    ↓
    ExecutionContextMgr
    ↓ (eventos semánticos de infraestructura)
    Infra Mediator
    ↓ (métodos)
    Domain Mgrs
    ↓ (eventos de dominio)
    UI / Ribbon / Forms

    Y en paralelo:

    Mgr ───► State
    AppState ───► all States
    ```

2. **Arquitectura correcta en VBA (pragmática, no dogmática)**
  - Thisworkbook. Responsabilidad única: enganchar eventos del host; reenviarlos sin semántica;
    Ejemplo:
    ```
    ' modExecutionContextListener
    Public WithEvents App As Application
    Public Sub Initialize()
    Set App = Application
    End Sub
    Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ExecutionContextMgr.OnWorkbookOpened Wb
    End Sub
    Nota clave: no RaiseEvent aquí; no lógica; solo reenviar
    ```
  - clsExecutionContextMgr. Responsabilidades: traducir eventos técnicos → semánticos de infraestructura; actualizar ExecutionContextState; emitir eventos de infraestructura
    Ejemplo:
    ```
    Public Event WorkbookSessionStarted(path As String, readOnly As Boolean)
    Public Sub OnWorkbookOpened(wb As Workbook)
    mState.RegisterWorkbook wb
    RaiseEvent WorkbookSessionStarted(wb.FullName, wb.ReadOnly)
    End Sub
    \*\* Mediador de infraestructura
    Responsabilidad: decidir qué Mgrs reaccionan
    Private Sub mExecCtx_WorkbookSessionStarted(path As String, ro As Boolean)
    mFileMgr.RegisterWorkbook path
    End Sub
    ```
3. **TODO EL ACCESO DESDE DOMINIO a Excel, se debe hacer a través de la capa de infraestructura**
- con llamadas como:
    "Public Sub RecalcularOferta(ctx As IExcelSession)", o
    "IOpportunityWorkbook
        .LeerDatosTecnicos()
        .EsOfertaValida()
        .GuardarResultado()"

4. **La clase clsFileXLS**
se encapsulará dentro de ExcelOpportunitySource / ExcelQuoteSource etc. esa clase proporcionará detalle técnico (Workbook, Path, estado, caché). Y ExcelOpportunitySource etc serán "adaptadores de dominio", que usan clsFileXLS e implementan IOpportunitySource, etc.
  Relación típica:
  ExcelOpportunitySource
  └── clsFileXLS
  └── Workbook
  El dominio solo ve IOpportunitySource.

5. **eventos**
sacar TODOS los RAISEEVENT de clsApplicationState, y ponerlos en METODOS de las clases Mgr que corresponda. Hay que determinar QUE EVENTOS DEBE EXPONER CADA MGR.

- NO PUEDE HABER "Private WithEvents ctx As clsExecutionContext" en la clase MEDIADORA DE DOMINIO!!, para evitarlo se puede seguir un patron como:
  Excel / Application
  │
  ▼
  clsExecutionContext (infraestructura pura)
  │
  ▼
  InfraEventCoordinator (mediadora infra)
  │
  ▼
  DomainEventBus / DomainCoordinator
  │
  ▼
  Clases de dominio
  , donde la mediadora de infraestructura se comunica con "el bus / la coordinadora de dominio", por TRADUCCION DE EVENTOS, algo así como:
  ' Evento técnico
  Private Sub mCtx_WorkbookOpened(ByVal wb As Workbook)
  ' Traducción semántica
  RaiseEvent WorkbookSessionStarted(wb.FullName,wb.ReadOnly )
  End Sub
  'El dominio escucha eso, no Excel
  Private WithEvents infraBus As clsInfraEventCoordinator
  Private Sub infraBus_WorkbookSessionStarted(...)
  ' actualizar caché
  End Sub
  Con ello grantizamos, por ejemplo, Un solo punto de captura de eventos de Excel.

6. **la clase clsFileManager**
- clsFileManager ES UNA CLASE DE INFRAESTRUCTURA. Conoce de Excel. pero NINGUNA CLASE DE DOMINIO debe conocer directamente de Excel, TODAS implementan funcionalidades sobre Excel mediante abstracciones relativas a oportunidades, ofertas, cálculos técnicos, etcétera.
  clsFileManager "escucha WorkbookOpen / WorkbookBeforeClose" (OJO!!: el mediador de infraestructura implementa esos callbacks, y esa clase "escucha" al recibir ordenes desde el mediador). Esa clase mantiene una caché de Workbook, y decide cuándo enganchar / desenganchar de ellos.

7. **el ribbon**
- el ribbon DEBE desaparecer del mediador de infraestructura, clsEventsMgrInfrastructure. DEBE diferenciarse una "capa de Interfaz" (Ribbon, formularios, panes), que escuche eventos que vienen del dominio / infraestructura, y ella misma decida cuándo invalidarse.

8. **proteccion contra resets**
- con todo ello, DEBEMOS conserguir la ESTABILIDAD de la aplicación: hay que conseguir evitar que Excel resetee el runtime sin avisar, o que los objetos COM (Workbook, Worksheet, Range) no sobrevivan, o que las referencias cruzadas se rompan silenciosamente.

## DECISIONES POR TOMAR.
Al margen de las decisiones que queden reflejadas en el resto del documento,
- debemos discutir un orden de PRIORIDAD, y una SECUENCIA, para llevar a efecto los objetivos anteriores: quiero que tú me hagas una propuesta para ello.

- Si ves alguna incongruencia en los anteriores preceptos o encomiendas, quiero que ME LA SEÑALES, para que la revisemos: soy humano, puedo equivocarme al escribir algo dos veces, de manera distinta.

## ADVERTENCIAS
NO deberia ser necesario decirte que "tu no eres un albañil, tu eres EL arquitecto.." y por tanto, todos estos errores que cometemos son responsabilidad tuya: tú estás para vigilar mis propuestas y para perfeccionarlas. CUALQUIERA de mis propuestas O comentarios debes recibirlos con juicio crítico y si no son correctos debes cuestionarlos y corregirlos antes de hacer nada basandote en ellos.
QUIERO QUE A PARTIR DE AHORA IMPLEMENTES TODOS LOS PROCESOS APLICANDO LOS ANTERIORES PRECEPTOS, ASI QUE MEMORIZALOS EN EL FICHERO DE SKILLS DE TU RAMA EN GITHUB, Y APLICALOS SIN EXCUSAS. Si en alguna implementación se tuvieran que vulnerar esos preceptos, tendrías que consultarmelo, para analizar el caso y justificarlo. Salvo que se justifique alguna corrección, MEMORIZA BIEN TODO LO ANTERIOR, y aplícate en aplicarlo.

Es posible que en el codigo que vas a revisar, veas que algunos cambios relativos a los comentarios anteriores ya están implementados.
QUiero que revises detenidamente el codigo que te facilito, y que hagas las propuestas necesarias para APLICAR los preceptos anteriores.

## PUNTOS POSIBLEMENTE YA IMPLEMENTADOS

Todos los puntos de esta sección han sido verificados y resueltos. Movidos a ##DONES.

## DONES

### Facade Property Gets en clsApplicationState (Fase F — commit e37520f)
Se eliminaron 5 Property Gets facade (CurrentOpportunity, CurrentFile, CurrentChart, RibbonMode, RibbonVisible) que no tenían callers externos. Cualquier entidad que necesita un subestado lo obtiene directamente del subsistema productor (ej: OpportunitiesMgr.State.CurrentOpportunity).

### IsRibbonTabVisible (Pre-E — verificado en commit 3bfa2a6)
La función IsRibbonTabVisible ya no existe en clsRibbonState. La visibilidad de la pestaña se controla directamente en modCALLBACKSRibbon mediante la consulta `App().AppState.IsOpportunityFileActive()`.

### Duplication AreSameFile (Pre-E — verificado en commit 3bfa2a6)
La función AreSameFile ya no existe duplicada en múltiples clases. La comparación de archivos se centraliza en la instancia que la necesita.

### Duplication AreSameOpportunity + double-check en Property Set (Pre-E + flag natural — commits 3bfa2a6, efae28c)
- AreSameOpportunity ya no existe duplicada en clsApplicationState y clsOpportunityState.
- clsApplicationState ya no tiene Property Set CurrentOpportunity (eliminado en Fase D/F: State no debe tener setters ni ordenar).
- El único guard contra asignación redundante vive en clsOpportunitiesMgr.CurrOpportunity (flag natural, commit efae28c): el Mgr compara antes de escribir al State y antes de emitir el evento. Una sola comprobación, en el lugar correcto.

### On* stubs y SyncOpportunityFromFile en clsApplicationState (Fase F — commit e37520f)
Los métodos Friend OnOpportunityChanged, OnFileChanged, OnOpportunityCollectionChanged y la Private SyncOpportunityFromFile solo logueaban sin ejercer coordinación real (AppState ya tiene referencias vivas a todos los subestados). Se eliminaron junto con las llamadas correspondientes en clsEventsMediatorDomain y el parámetro oApplicationState de su Initialize.

### Double-MsgBox en monitoreo (Fase F — commit 58b6fbe)
MonitoringError y MonitoringFailed mostraban MsgBox dos veces: una en clsFSMonitoringCoord (tras RaiseEvent) y otra en clsEventsMediatorDomain (en el handler síncrono). Se eliminaron los MsgBox del mediador. El mediador no es UI; los MsgBox originales en clsFSMonitoringCoord se mantienen como única notificación.

### Property Let Public en clsRibbonState (Fase F — commit c5d3cb0)
Property Let Modo y Visible cambiados de Public a Friend. Todos sus callers son clsRibbon (su Mgr). Per preceptos: State no debe exponer setters públicos.

### Corrupción encoding en clsRibbonState (Fase F — commit c5d3cb0)
@Description tenía "lógico" con bytes U+FFFD (ef bf bd, codificación UTF-8 del replacement character) en lugar de \xf3 (ó en ISO-8859-1). Corregido.

### clsExecutionContext descompuesto (Fase E — commit e80be1f)
clsExecutionContext (289 líneas) eliminado e íntegramente redistribuido en:
- Listener en ThisWorkbook (WithEvents Application, 5 handlers, cero lógica)
- clsExecutionContextState (estado cacheado puro, Friend setters)
- clsExecutionContextMgr (recibe Listener, actualiza State, lanza eventos semánticos, DetectChart)

### Bug residual Fase E: mCtx en ChartManager.Initialize (Fase F — commit e37520f)
`mChartManager.Initialize mCtx` en clsApplication referenciaba la variable eliminada. Corregido a `mCtxState`.

### Llamadas inadecuadas al Ribbon desde Dispatcher (Fase G — commit e6cea5e)
clsEventDispatcher llamaba directamente a App.ribbon.InvalidarControl y App.ribbon.InvalidarRibbon en 3 lugares (btnNuevaOp, btnOpRefresh, SetRibbonSelectionIndex). Eliminadas estas llamadas. El Ribbon ahora se invalida automáticamente via eventos que clsRibbon escucha con WithEvents de los Managers (OpportunitiesMgr, FileMgr, etc.).

### Detección de reset VBA (Fase G — commit e6cea5e)
Añadida detección de reset de VBA mediante variable Static en modMACROAppLifecycle:
- `DetectVBAResetOccurred()`: Detecta si ocurrió un reset desde la última llamada
- `InitializationCount`: Contador de inicializaciones (>1 indica que hubo resets)
- `LastInitializationTime`: Timestamp de la última inicialización

### Validación de contexto COM (Fase G — commit e6cea5e)
clsExecutionContextMgr ahora tiene:
- `IsContextValid()`: Verifica que las referencias COM estén vivas (no zombis)
- `RefreshContextState()`: Detecta y actualiza el contexto actual desde Excel
- Handlers de ContextInvalidated/ContextReinitialized en clsEventsMgrInfrastructure

### IsInitialized en clsApplicationState (Fase G — commit e6cea5e)
Añadidas propiedades para diagnosticar estado de inicialización:
- `IsInitialized`: True si todos los subestados críticos fueron inyectados
- `GetInitializationStatus()`: String con estado de cada subestado

### Interfaz IOpportunity de dominio (Fase H — commit 6acb541)
IOpportunity.cls implementada como interfaz de dominio que define el contrato para trabajar con oportunidades. Métodos: OpportunityId, BasePath, DisplayName, IsValid, CanGenerateQuote, ReadTechnicalData.

### Adaptador ExcelOpportunitySource (Fase H — commit 6acb541)
ExcelOpportunitySource.cls creado como adaptador de infraestructura que implementa IOpportunity usando Excel. Encapsula clsFileXLS y traduce sus operaciones al contrato de dominio. El dominio solo ve IOpportunity; la infraestructura sabe hablar Excel.

### Detección de reset integrada en bootstrap (Feature 1 — commit 5001caf)
ThisWorkbook.Workbook_Open ahora llama a DetectVBAResetOccurred() al inicio y fuerza reinicialización del contexto si detecta reset. Loguea número de inicialización para diagnóstico.

### Flujo de reinicialización completo en dominio (Feature 1 — commit 5001caf, actualizado Feature 2)
- clsEventsMediatorDomain escucha InfraContextInvalidated/InfraContextReinitialized via WithEvents mInfraMediador (eventos traducidos)
- Durante InfraContextInvalidated: pausa FSMonitoringCoord
- Durante InfraContextReinitialized: reanuda FSMonitoringCoord y refresca oportunidades
- clsFSMonitoringCoord: nuevos métodos PausarMonitoreo/ReanudarMonitoreo y propiedad IsPaused

### clsOpportunity implementa IOpportunity (Feature 1 — commit 5001caf)
La entidad de dominio clsOpportunity ahora implementa la interfaz IOpportunity directamente, exponiendo: OpportunityId, BasePath, DisplayName, IsValid, CanGenerateQuote, ReadTechnicalData.

### CurrentOpportunitySource en clsOpportunitiesMgr (Feature 1 — commit 5001caf)
Nueva propiedad `CurrentOpportunitySource As IOpportunity` que permite al código de dominio trabajar con la oportunidad actual usando la interfaz abstracta en lugar de la implementación concreta.

### Interfaz IOferta de dominio (Feature 1 — commit 5001caf)
IOferta.cls implementada como interfaz de dominio para ofertas. Métodos: OfertaId, NumeroOferta, FechaOferta, IsValid, IsDirty, IsNew, GetDatosGenerales, OtrosCount.

### Traducción de eventos infra→dominio (Feature 2)
Corregida la violación del Objetivo 5: el mediador de dominio ya NO escucha directamente a clases de infraestructura.
- clsEventsMgrInfrastructure ahora emite eventos semánticos traducidos:
  - `InfraContextInvalidated` (traducción de mCtxMgr_ContextInvalidated)
  - `InfraContextReinitialized` (traducción de mCtxMgr_ContextReinitialized)
  - `ActiveFileSessionChanged` (traducción de mFileMgr_ActiveFileChanged)
- clsEventsMediatorDomain ahora escucha `WithEvents mInfraMediador` en lugar de escuchar directamente a mCtxMgr y mFileMgr
- El flujo correcto es: Excel → ExecutionContextMgr → InfraMediador (traduce) → DomainMediador
- clsApplication.cls actualizado para pasar mEvMgrInfrastructure al inicializar el mediador de dominio

### clsFileManager reclasificado como infraestructura (Feature 2)
@Folder cambiado de "4-Servicios.Archivos" a "2-Infraestructura" según Objetivo 6 de REFERENCE_NOTES.md.

### Violación pendiente documentada: uso de clsFileXLS en domain mediator (Feature 2)
En mOpportunities_currOpportunityChanged se usa clsFileXLS directamente. Documentado con TODO explicando que requiere decisión de diseño cuando se tenga la especificación de dominio. Opciones: delegarlo a servicio de infraestructura o que el mediador de infraestructura exponga método semántico.