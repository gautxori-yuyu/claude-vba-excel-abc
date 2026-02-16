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
Private mOpportunities As clsOpportunitiesMgr
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1
Private WithEvents mConfiguration As clsConfiguration
Attribute mConfiguration.VB_VarHelpID = -1
                '
Private mEvMediatorDomain As clsEventsMediatorDomain

Private mRibbon As clsRibbon

' Estado centralizado de la aplicación
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

'@Description: Obtiene la configuracion de la aplicacion
Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

'@Description: Obtiene el gestor de estado centralizado de la aplicación
Public Property Get State() As clsApplicationState
    Set State = mApplicationState
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

    Set mApplicationState = New clsApplicationState
    
    InitializeInfrastructure
    
    InitializeDomain

    InitializeUI

    LogInfo MODULE_NAME, "[Class_Initialize] inicialización completada"
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Class_Initialize]"
    Err.Raise Err.Number, MODULE_NAME & "[Class_Initialize]", _
              "Error inicializando aplicación: " & Err.Description
End Sub
Private Sub InitializeInfrastructure()
    On Error GoTo ErrHandler
    ' INFRAESTRUCTURA
    Set mCtxMgr = New clsExecutionContextMgr
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    ' clsChartMgr nace como composicion interna de clsExecutionContextMgr
    mApplicationState.Initialize CtxState:=mCtxMgr.State, ChartState:=mCtxMgr.ChartMgr.State

    Set mFileMgr = New clsFileManager
    mFileMgr.Init mCtxMgr.Application.Workbooks
    mApplicationState.Initialize FileState:=mFileMgr.State

    ' GESTOR DE EVENTOS DE INFRAESTRUCTURA
    Set mEvMediatorInfrastructure = New clsEventsMediatorInfrastructure
    Call mEvMediatorInfrastructure.Initialize(mFileMgr, mCtxMgr)
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[InitializeInfrastructure]"
    Err.Raise Err.Number, MODULE_NAME & "[InitializeInfrastructure]", _
              Err.Description
End Sub
Private Sub InitializeDomain()
    On Error GoTo ErrHandler
    ' CONFIGURACION (servicio de infraestructura)
    Set mConfiguration = New clsConfiguration

    ' SERVICIOS DE MONITOREO
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    mFSMonitoringCoord.IniciarMonitoreo mConfiguration.oDicFoldersToWatch

    ' PROVIDER DE OPORTUNIDADES (infraestructura)
    Set mOpportunityProvider = New FileSystemOpportunityProvider
    mOpportunityProvider.Initialize mFileMgr, mConfiguration

    ' PROVIDER DE CALCULOS TECNICOS (infraestructura)
    Set mCalcProvider = New FileSystemTechnicalCalcProvider
    mCalcProvider.Initialize mFileMgr

    ' DOMINIO - Crear manager de oportunidades (crea su estado internamente)
    Set mOpportunities = New clsOpportunitiesMgr
    mOpportunities.Initialize mOpportunityProvider
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    mApplicationState.Initialize OpportunityState:=mOpportunities.State

    ' MEDIADOR DE EVENTOS DE DOMINIO
    ' Escucha eventos SEMANTICOS del mediador de infraestructura (no directamente de mCtxMgr)
    Set mEvMediatorDomain = New clsEventsMediatorDomain
    Call mEvMediatorDomain.Initialize(mCtxMgr.State, _
                                    mEvMediatorInfrastructure, _
                                    mFSMonitoringCoord, _
                                    mOpportunities, _
                                    mFileMgr)
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

    On Error Resume Next  ' Asegurar que la limpieza continúa aunque haya errores

    ' ORDEN INVERSO A LA INSTANCIACION:
    ' Instanciacion: CtxMgr -> ChartMgr -> EventDispatcher -> FileMgr -> Configuration
    '                -> FSMonitoringCoord -> EvMediatorInfrastructure -> OpportunityProvider
    '                -> Opportunities -> Ribbon -> ApplicationState -> EvMediatorDomain
    ' Liberacion (inverso): EvMediatorDomain -> ApplicationState -> Ribbon -> Opportunities
    '                       -> OpportunityProvider -> EvMediatorInfrastructure -> FSMonitoringCoord
    '                       -> Configuration -> FileMgr -> EventDispatcher -> ChartMgr -> CtxMgr

    DisposeUI
    
    DisposeDomain
    
    DisposeInfrastructure

    ' 12. Liberar estado centralizado de aplicacion
    Set mApplicationState = Nothing

    LogInfo MODULE_NAME, "[Terminate] Limpieza completada"
End Sub
Private Sub DisposeInfrastructure()
    ' Liberar mediador de eventos de infraestructura
    Set mEvMediatorInfrastructure = Nothing

    ' Liberar gestor de archivos
    Set mFileMgr = Nothing

    ' clsChartMgr liberado internamente por clsExecutionContextMgr (composicion)
    ' Liberar contexto de ejecucion (primero en instanciarse, ultimo en liberarse)
    Set mCtxMgr = Nothing
End Sub
Private Sub DisposeDomain()
    ' 1. Liberar mediador de eventos de dominio (ultimo en instanciarse)
    Set mEvMediatorDomain = Nothing

    ' 3. Liberar gestor de oportunidades
    Set mOpportunities = Nothing

    ' 4. Liberar proveedores
    Set mOpportunityProvider = Nothing
    Set mCalcProvider = Nothing

    ' 5. Apagado explícito del monitor de archivos
    If Not mFSMonitoringCoord Is Nothing Then
        mFSMonitoringCoord.Shutdown
        Set mFSMonitoringCoord = Nothing
    End If

    ' 6. Liberar configuracion
    Set mConfiguration = Nothing
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

' ==================================================================
' EVENTOS DE CONFIGURACION
' ==================================================================

' ==================================================================
' EVENTOS DE MONITORIZACION DE ARCHIVOS (actualizacion de mIsReady)
' ==================================================================

'@Description: Error de monitorizacion: puede indicar que la carpeta no esta accesible.
'              Actualiza el flag de disponibilidad del provider.
Private Sub mFSMonitoringCoord_MonitoringError(ByVal folder As String, ByVal errorMessage As String)
    On Error Resume Next
    LogWarning MODULE_NAME, "[mFSMonitoringCoord_MonitoringError] Error en: " & folder & " -> RefreshStorageAvailability"
    If Not mOpportunityProvider Is Nothing Then
        mOpportunityProvider.RefreshStorageAvailability
    End If
    On Error GoTo 0
End Sub

'@Description: Monitorizacion fallida definitivamente.
'              Marca el almacenamiento como no disponible.
Private Sub mFSMonitoringCoord_MonitoringFailed(ByVal folder As String, ByVal reason As String)
    On Error Resume Next
    LogError MODULE_NAME, "[mFSMonitoringCoord_MonitoringFailed] Fallo en: " & folder & " -> RefreshStorageAvailability"
    If Not mOpportunityProvider Is Nothing Then
        mOpportunityProvider.RefreshStorageAvailability
    End If
    On Error GoTo 0
End Sub

'@Description: Monitorizacion reconectada. Actualiza disponibilidad y recarga oportunidades.
Private Sub mFSMonitoringCoord_MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
    On Error Resume Next
    LogInfo MODULE_NAME, "[mFSMonitoringCoord_MonitoringReconnected] Reconectado: " & folder & " tras " & attempts & " intentos"
    If Not mOpportunityProvider Is Nothing Then
        mOpportunityProvider.RefreshStorageAvailability
    End If
    If Not mOpportunities Is Nothing Then
        mOpportunities.RefreshOpportunities
    End If
    On Error GoTo 0
End Sub

' ==================================================================
' EVENTOS DE CONFIGURACION
' ==================================================================

'@Description: Se dispara cuando cambia la ruta de oportunidades en la configuracion.
'              Actualiza el flag de disponibilidad y recarga la lista de oportunidades.
Private Sub mConfiguration_OpportunitiesPathChanged(ByVal newPath As String)
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[mConfiguration_OpportunitiesPathChanged] Nueva ruta: " & newPath

    ' Actualizar flag de disponibilidad del provider (re-consulta el FS con la nueva ruta)
    If Not mOpportunityProvider Is Nothing Then
        mOpportunityProvider.RefreshStorageAvailability
    End If

    ' Recargar la lista de oportunidades desde la nueva ruta
    If Not mOpportunities Is Nothing Then
        mOpportunities.RefreshOpportunities
    End If

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[mConfiguration_OpportunitiesPathChanged]"
End Sub

