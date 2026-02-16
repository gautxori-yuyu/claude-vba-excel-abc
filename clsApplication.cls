VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' gestion de la aplicacion 'gestor de oportunidades ABC'
'@Folder "1-Aplicacion"
Option Explicit

Private Const MODULE_NAME As String = "clsApplication"

Private mCtxMgr As clsExecutionContextMgr

Private mEventDispatcher As clsEventDispatcher
Private mEvMediatorInfrastructure As clsEventsMediatorInfrastructure

Private mFileMgr As clsFileManager
Private mOpportunityProvider As FileSystemOpportunityProvider
Private mCalcProvider As FileSystemTechnicalCalcProvider
Private mQuoteProvider As FileSystemEconomicQuoteProvider
Private mOpportunities As clsOpportunitiesMgr
Private mFSMonitoringCoord As clsFSMonitoringCoord
Private mConfiguration As clsConfiguration
                '
Private mEvMediatorDomain As clsEventsMediatorDomain

Private mRibbon As clsRibbon

' Estado centralizado de la aplicaciï¿½n
Private mApplicationState As clsApplicationState

Public Property Get ChartManager() As clsChartMgr
    If Not mCtxMgr Is Nothing Then Set ChartManager = mCtxMgr.ChartMgr
End Property

Public Property Get ExecutionContextMgr() As clsExecutionContextMgr
    Set ExecutionContextMgr = mCtxMgr
End Property

Public Static Property Get ribbon() As clsRibbon
    Set ribbon = mRibbon
End Property

Public Property Get Dispatcher() As clsEventDispatcher
    Set Dispatcher = mEventDispatcher
End Property

Public Property Get FileMgr() As clsFileManager
    Set FileMgr = mFileMgr
End Property
                
Public Property Get FSMonitoringCoord() As clsFSMonitoringCoord
    Set FSMonitoringCoord = mFSMonitoringCoord
End Property

Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr
    Set OpportunitiesMgr = mOpportunities
End Property

Public Property Get CalcProvider() As ITechnicalCalcProvider
    Set CalcProvider = mCalcProvider
End Property

'@Description: Obtiene el proveedor de valoraciones economicas
Public Property Get QuoteProvider() As IEconomicQuoteProvider
    Set QuoteProvider = mQuoteProvider
End Property

'@Description: Obtiene la configuracion de la aplicacion
Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

'@Description: Obtiene el gestor de estado centralizado de la aplicaciï¿½n
Public Property Get State() As clsApplicationState
    Set State = mApplicationState
End Property

' -------------------------------------------------------------
' Class_Initialize: SOLO inicializaciï¿½n mï¿½nima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[Class_Initialize] Iniciando aplicacion"
    ' fixme: revisar que el orden de instanciaciï¿½n de cada componente es EL CORRECTO!!
    ' (se ha intentado respetar el orden Infraestructura -> dominio -> UI -> Aplicaciï¿½n;
    ' y dentro de cada capa, primero los managers (que crean ellos mismos los State), luego
    ' los servicios y la configuracion de la capa; y finalmente mediadores de eventos, y coordinadores

    Set mApplicationState = New clsApplicationState
    
    InitializeInfrastructure
    
    InitializeDomain

    InitializeUI

    LogInfo MODULE_NAME, "[Class_Initialize] inicializaciï¿½n completada"
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Class_Initialize]"
    Err.Raise Err.Number, MODULE_NAME & "[Class_Initialize]", _
              "Error inicializando aplicaciï¿½n: " & Err.Description
End Sub
Private Sub InitializeInfrastructure()
    On Error GoTo ErrHandler
    ' --- CORE ---
    Set mCtxMgr = New clsExecutionContextMgr
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    ' clsChartMgr nace como composicion interna de clsExecutionContextMgr
    mApplicationState.Initialize CtxState:=mCtxMgr.State, ChartState:=mCtxMgr.ChartMgr.State

    Set mFileMgr = New clsFileManager
    mFileMgr.Init mCtxMgr.Application.Workbooks
    mApplicationState.Initialize FileState:=mFileMgr.State

    ' GESTOR DE EVENTOS DE INFRAESTRUCTURA (adapter pattern)
    Set mEvMediatorInfrastructure = New clsEventsMediatorInfrastructure
    Call mEvMediatorInfrastructure.Initialize(mFileMgr, mCtxMgr)

    ' --- CONFIGURACION ---
    Set mConfiguration = New clsConfiguration

    ' --- MONITOREO ---
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    mFSMonitoringCoord.IniciarMonitoreo mConfiguration.oDicFoldersToWatch

    ' --- PROVIDERS (implementaciones de interfaces de dominio) ---
    Set mOpportunityProvider = New FileSystemOpportunityProvider
    mOpportunityProvider.Initialize mFileMgr, mConfiguration, mFSMonitoringCoord

    Set mCalcProvider = New FileSystemTechnicalCalcProvider
    mCalcProvider.Initialize mFileMgr

    Set mQuoteProvider = New FileSystemEconomicQuoteProvider
    mQuoteProvider.Initialize mFileMgr

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[InitializeInfrastructure]"
    Err.Raise Err.Number, MODULE_NAME & "[InitializeInfrastructure]", _
              Err.Description
End Sub
Private Sub InitializeDomain()
    On Error GoTo ErrHandler
    ' DOMINIO - Crear manager de oportunidades (crea su estado internamente)
    ' Se pasa tanto la interfaz (para llamadas de dominio) como la clase concreta
    ' (para WithEvents - limitacion VBA: WithEvents no funciona con interfaces)
    Set mOpportunities = New clsOpportunitiesMgr
    mOpportunities.Initialize mOpportunityProvider, mOpportunityProvider
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    mApplicationState.Initialize OpportunityState:=mOpportunities.State

    ' MEDIADOR DE EVENTOS DE DOMINIO
    ' Escucha eventos SEMANTICOS del mediador de infraestructura y de FSMonitoringCoord
    Set mEvMediatorDomain = New clsEventsMediatorDomain
    Call mEvMediatorDomain.Initialize(mCtxMgr.State, _
                                    mEvMediatorInfrastructure, _
                                    mFSMonitoringCoord, _
                                    mOpportunities)
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[InitializeDomain]"
    Err.Raise Err.Number, MODULE_NAME & "[InitializeDomain]", _
              Err.Description
End Sub
Private Sub InitializeUI()
    On Error GoTo ErrHandler
    ' INTERFAZ - Crear Ribbon (crea RibbonState internamente)
    Set mEventDispatcher = New clsEventDispatcher
    
    Set mRibbon = New clsRibbon
    mRibbon.Initialize mOpportunities, mFileMgr, mCtxMgr.ChartMgr, mCtxMgr

    ' Recuperacion automatica tras reset de VBA: si el IRibbonUI sigue vivo en Excel4, reutilizarlo
    Dim recoveredRibbon As IRibbonUI
    Set recoveredRibbon = RecoverRibbonFromExcelNames()
    If Not recoveredRibbon Is Nothing Then
        mRibbon.Init recoveredRibbon
        LogInfo MODULE_NAME, "[InitializeUI] Ribbon recuperado automaticamente de sesion anterior"
    End If
    Set recoveredRibbon = Nothing

    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    mApplicationState.Initialize RibbonState:=mRibbon.State

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[InitializeUI]"
    Err.Raise Err.Number, MODULE_NAME & "[InitializeUI]", _
              Err.Description
End Sub
' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Terminate] Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza continï¿½a aunque haya errores

    ' ORDEN INVERSO A LA INSTANCIACION:
    ' Infraestructura: CtxMgr -> FileMgr -> EvMediatorInfra -> Configuration -> FSMonitoringCoord
    '                  -> OpportunityProvider -> CalcProvider -> QuoteProvider
    ' Dominio: Opportunities -> EvMediatorDomain
    ' UI: Ribbon -> EventDispatcher

    DisposeUI
    
    DisposeDomain
    
    DisposeInfrastructure

    ' 12. Liberar estado centralizado de aplicacion
    Set mApplicationState = Nothing

    LogInfo MODULE_NAME, "[Terminate] Limpieza completada"
End Sub
Private Sub DisposeInfrastructure()
    ' Liberar proveedores (orden inverso a instanciacion)
    Set mQuoteProvider = Nothing
    Set mCalcProvider = Nothing
    Set mOpportunityProvider = Nothing

    ' Apagado explicito del monitor de archivos
    If Not mFSMonitoringCoord Is Nothing Then
        mFSMonitoringCoord.Shutdown
        Set mFSMonitoringCoord = Nothing
    End If

    ' Liberar configuracion
    Set mConfiguration = Nothing

    ' Liberar mediador de eventos de infraestructura
    Set mEvMediatorInfrastructure = Nothing

    ' Liberar gestor de archivos
    Set mFileMgr = Nothing

    ' clsChartMgr liberado internamente por clsExecutionContextMgr (composicion)
    ' Liberar contexto de ejecucion (primero en instanciarse, ultimo en liberarse)
    Set mCtxMgr = Nothing
End Sub
Private Sub DisposeDomain()
    ' Liberar mediador de eventos de dominio (ultimo en instanciarse)
    Set mEvMediatorDomain = Nothing

    ' Liberar gestor de oportunidades
    Set mOpportunities = Nothing
End Sub
Private Sub DisposeUI()
    ' Liberar Ribbon (con apagado explicito de eventos)
    If Not mRibbon Is Nothing Then
        mRibbon.StopEvents
        Set mRibbon = Nothing
    End If

    ' Liberar dispatcher de eventos
    Set mEventDispatcher = Nothing
End Sub


