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

Private mChartManager As clsChartMgr
Private mCtxMgr As clsExecutionContextMgr

Private mEventDispatcher As clsEventDispatcher
Private mEvMediatorInfrastructure As clsEventsMediatorInfrastructure

Private mFileMgr As clsFileManager
Private mOpportunityProvider As FileSystemOpportunityProvider
Private mOpportunities As clsOpportunitiesMgr
Private mFSMonitoringCoord As clsFSMonitoringCoord
Private mConfiguration As clsConfiguration
                '
Private mEvMediatorDomain As clsEventsMediatorDomain

Private mRibbon As clsRibbon

' Estado centralizado de la aplicación
Private mApplicationState As clsApplicationState

Public Property Get ChartManager() As clsChartMgr
    Set ChartManager = mChartManager
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

'@Description: Obtiene la configuracion de la aplicacion
Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

'@Description: Obtiene el gestor de estado centralizado de la aplicación
Public Property Get AppState() As clsApplicationState
    Set AppState = mApplicationState
End Property

' -------------------------------------------------------------
' Class_Initialize: SOLO inicialización mínima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[Class_Initialize] Iniciando aplicacion"
    ' fixme: revisar que el orden de instanciación de cada componente es EL CORRECTO!!
    ' (se ha intentado respetar el orden Infraestructura -> dominio -> UI -> Aplicación;
    ' y dentro de cada capa, primero los managers (que crean ellos mismos los State), luego
    ' los servicios y la configuracion de la capa; y finalmente mediadores de eventos, y coordinadores

    ' INFRAESTRUCTURA
    Set mCtxMgr = New clsExecutionContextMgr
    ' FIXME: Es mejor UNIFICAR, o HACER COMPOSICION, de clsChartMgr, dentro de clsExecutionContext (ChartMgr PERTENECE al exec context,
    ' mas que depender de él; actualmente parece que la relación está invertida, está asi para minimizar los cambios al resto del codigo...)
'    Set mChartManager = New clsChartMgr
'    mChartManager.Initialize mCtxMgr.State
'    Set mEventDispatcher = New clsEventDispatcher
'
'    Set mFileMgr = New clsFileManager
'    mFileMgr.Init mCtxMgr.Application.Workbooks
'
'    ' GESTOR DE EVENTOS DE INFRAESTRUCTURA
'    Set mEvMediatorInfrastructure = New clsEventsMediatorInfrastructure
'    Call mEvMediatorInfrastructure.Initialize(mFileMgr, mChartManager, mCtxMgr)

    ' CONFIGURACION (servicio de infraestructura)
'    Set mConfiguration = New clsConfiguration

    ' SERVICIOS DE MONITOREO
'    Set mFSMonitoringCoord = New clsFSMonitoringCoord
'    mFSMonitoringCoord.IniciarMonitoreo mConfiguration.oDicFoldersToWatch

    ' PROVIDER DE OPORTUNIDADES (infraestructura)
'    Set mOpportunityProvider = New FileSystemOpportunityProvider
'    mOpportunityProvider.Initialize mFileMgr, mConfiguration

    ' DOMINIO - Crear managers (ellos crean sus estados internamente)
    'Set mOpportunities = New clsOpportunitiesMgr
    'mOpportunities.Initialize mOpportunityProvider

    ' INTERFAZ - Crear Ribbon (crea RibbonState internamente)
    Set mRibbon = New clsRibbon
    mRibbon.Initialize mOpportunities, mFileMgr, mChartManager, mCtxMgr
    
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
'    Set mApplicationState = New clsApplicationState
'    mApplicationState.Initialize RibbonState:=mRibbon.State, _
'                                 OpportunityState:=mOpportunities.State, _
'                                 FileState:=mFileMgr.State, _
'                                 ChartState:=mChartManager.State, _
'                                 CtxState:=mCtxMgr.State
    
    ' MEDIADOR DE EVENTOS DE DOMINIO
    ' Escucha eventos SEMANTICOS del mediador de infraestructura (no directamente de mCtxMgr)
'    Set mEvMediatorDomain = New clsEventsMediatorDomain
'    Call mEvMediatorDomain.Initialize(mCtxMgr.State, _
'                                    mEvMediatorInfrastructure, _
'                                    mFSMonitoringCoord, _
'                                    mOpportunities, _
'                                    mFileMgr)

    LogInfo MODULE_NAME, "[Class_Initialize] inicialización completada"
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Class_Initialize]"
    Err.Raise Err.Number, MODULE_NAME & "[Class_Initialize]", _
              "Error inicializando aplicación: " & Err.Description
End Sub

' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Terminate] Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza continúa aunque haya errores

    ' ORDEN INVERSO A LA INSTANCIACION:
    ' Instanciacion: CtxMgr -> ChartMgr -> EventDispatcher -> FileMgr -> Configuration
    '                -> FSMonitoringCoord -> EvMediatorInfrastructure -> OpportunityProvider
    '                -> Opportunities -> Ribbon -> ApplicationState -> EvMediatorDomain
    ' Liberacion (inverso): EvMediatorDomain -> ApplicationState -> Ribbon -> Opportunities
    '                       -> OpportunityProvider -> EvMediatorInfrastructure -> FSMonitoringCoord
    '                       -> Configuration -> FileMgr -> EventDispatcher -> ChartMgr -> CtxMgr

    ' 1. Liberar mediador de eventos de dominio (ultimo en instanciarse)
    Set mEvMediatorDomain = Nothing

    ' 2. Liberar estado centralizado de aplicacion
    Set mApplicationState = Nothing

    ' 3. Liberar Ribbon (con apagado explicito de eventos)
    If Not mRibbon Is Nothing Then
        mRibbon.StopEvents
        Set mRibbon = Nothing
    End If

    ' 4. Liberar gestor de oportunidades
    Set mOpportunities = Nothing

    ' 5. Liberar proveedor de oportunidades
    Set mOpportunityProvider = Nothing

    ' 6. Apagado explícito del monitor de archivos
    If Not mFSMonitoringCoord Is Nothing Then
        mFSMonitoringCoord.Shutdown
        Set mFSMonitoringCoord = Nothing
    End If

    ' 7. Liberar configuracion
    Set mConfiguration = Nothing

    ' 8. Liberar mediador de eventos de infraestructura
    Set mEvMediatorInfrastructure = Nothing

    ' 9. Liberar gestor de archivos
    Set mFileMgr = Nothing

    ' 10. Liberar dispatcher de eventos
    Set mEventDispatcher = Nothing

    ' 11. Liberar gestor de graficos
    Set mChartManager = Nothing

    ' 12. Liberar contexto de ejecucion (primero en instanciarse, ultimo en liberarse)
    Set mCtxMgr = Nothing

    LogInfo MODULE_NAME, "[Terminate] Limpieza completada"
End Sub

