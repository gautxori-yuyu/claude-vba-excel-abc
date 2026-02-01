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
'@Folder "5-Aplicac (Coord)"
Option Explicit

Private Const MODULE_NAME As String = "clsApplication"

Private mChartManager As clsChartMgr
Private mCtx As clsExecutionContext
Attribute mCtx.VB_VarHelpID = -1

' Nuevo contexto refactorizado
Private mCtxMgr As clsExecutionContextMgr

Private mEventDispatcher As clsEventDispatcher
Private mEvMgrInfrastructure As clsEventsMgrInfrastructure

Private mFileMgr As clsFileManager
Private mOpportunities As clsOpportunitiesMgr
Private mFSMonitoringCoord As clsFSMonitoringCoord
                '
Private mEvMediatorDomain As clsEventsMediatorDomain

Private mRibbon As clsRibbon

' Estado centralizado de la aplicaci�n
Private mApplicationState As clsApplicationState

' Nuevo gestor de contexto de ejecuci�n
Private mExecutionContextMgr As clsExecutionContextMgr

Public Property Get ChartManager() As clsChartMgr
    Set ChartManager = mChartManager
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = mCtx
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

'@Description: Obtiene el gestor de estado centralizado de la aplicaci�n
Public Property Get AppState() As clsApplicationState
    Set AppState = mApplicationState
End Property

' -------------------------------------------------------------
' Class_Initialize: SOLO inicializaci�n m�nima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[Class_Initialize] - Iniciando aplicacion"

    ' INFRAESTRUCTURA
    Set mCtx = New clsExecutionContext
    mCtx.Initialize
    ' TODO: creo que es mejor UNIFICARLO, o HACER COMPOSICION, de clsChartMgr, dentro de clsExecutionContext (ChartMgr PERTENECE al exec context,
    ' mas que depender de �l; actualmente parece que la relaci�n est� invertida, est� asi para minimizar los cambios al resto del codigo...)
    Set mChartManager = New clsChartMgr
    mChartManager.Initialize mCtx
    Set mEventDispatcher = New clsEventDispatcher

    Set mFileMgr = New clsFileManager
    mFileMgr.Init mCtx.Application.Workbooks

    ' DOMINIO - Crear managers (ellos crean sus estados internamente)
    Set mOpportunities = New clsOpportunitiesMgr
    ' Inyectar dependencia de FileManager y arrancar
    Set mOpportunities.FileManager = mFileMgr
    mOpportunities.Initialize

    ' SERVICIOS DE MONITOREO
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    mFSMonitoringCoord.IniciarMonitoreo mOpportunities.Conf.oDicFoldersToWatch

    ' INTERFAZ - Crear Ribbon (crea RibbonState internamente)
    Set mRibbon = New clsRibbon
    mRibbon.Initialize mOpportunities, mFileMgr, mChartManager, mCtx, mExecutionContextMgr
    
    ' ESTADO CENTRALIZADO - Inyectar referencias a subestados (Dependency Injection)
    Set mApplicationState = New clsApplicationState
    mApplicationState.Initialize RibbonState:=mRibbon.State, _
                                 OpportunityState:=mOpportunities.State, _
                                 FileState:=mFileMgr.State, _
                                 ChartState:=mChartManager.State
    ' GESTOR DE EVENTOS DE INFRAESTRUCTURA
    Set mEvMgrInfrastructure = New clsEventsMgrInfrastructure
    Call mEvMgrInfrastructure.Initialize(mFileMgr, mChartManager, mExecutionContextMgr)

    ' MEDIADOR DE EVENTOS DE DOMINIO - Inyectar ApplicationState
    Set mEvMediatorDomain = New clsEventsMediatorDomain
    Call mEvMediatorDomain.Initialize(mCtx, _
                                    mExecutionContextMgr, _
                                    mFSMonitoringCoord, _
                                    mOpportunities, _
                                    mFileMgr, _
                                    mApplicationState)

    LogInfo MODULE_NAME, "[Class_Initialize] - inicializaci�n completada"
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Class_Initialize]"
    Err.Raise Err.Number, MODULE_NAME & "[Class_Initialize]", _
              "Error inicializando aplicaci�n: " & Err.Description
End Sub

' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Terminate] - Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza contin�a aunque haya errores
    
    ' 2. Liberar Ribbon events
    If Not mRibbon Is Nothing Then
        mRibbon.StopEvents
        Set mRibbon = Nothing
    End If

    Set mEvMediatorDomain = Nothing

    Set mFSMonitoringCoord = Nothing

    Set mOpportunities = Nothing

    ' 3. Liberar managers
    Set mFileMgr = Nothing

    Set mEvMgrInfrastructure = Nothing

    Set mEventDispatcher = Nothing
    
    Set mChartManager = Nothing

    ' Liberar nuevo gestor de contexto
    Set mExecutionContextMgr = Nothing

    ' Liberar estado centralizado
    Set mApplicationState = Nothing

    LogInfo MODULE_NAME, "[Terminate] - Limpieza completada"
End Sub


