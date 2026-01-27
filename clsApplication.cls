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
Attribute mChartManager.VB_VarHelpID = -1

Private mEventDispatcher As clsEventDispatcher

Private mEvMgrInfrastructure As clsEventsMgrInfrastructure

Private mFileMgr As clsFileManager

Private mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1

Private mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1
                '
Private mEvMediatorDomain As clsEventsMediatorDomain

Private mRibbon As clsRibbon

Public Property Get ChartManager() As clsChartMgr
    Set ChartManager = mChartManager
End Property

Public Static Property Get Ribbon() As clsRibbon
    Set Ribbon = mRibbon
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

' -------------------------------------------------------------
' Class_Initialize: SOLO inicializaci?n m?nima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[Class_Initialize] - Iniciando aplicacion"
    
    ' INFRAESTRUCTURA
    Set mChartManager = New clsChartMgr
    
    Set mEventDispatcher = New clsEventDispatcher
    
    Set mEvMgrInfrastructure = New clsEventsMgrInfrastructure
    Call mEvMgrInfrastructure.Initialize(mRibbon, mChartManager)
    
    ' DOMINIO
    Set mFileMgr = New clsFileManager

    Set mOpportunities = New clsOpportunitiesMgr
    ' Inyectar dependencia de FileManager
    Set mOpportunities.FileManager = mFileMgr

    ' SERVICIOS DE NIVEL DE DOMINIO (crean eventos y generan logica "de negocio")
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    mFSMonitoringCoord.IniciarMonitoreo mOpportunities.Conf.oDicFoldersToWatch
    
    Set mEvMediatorDomain = New clsEventsMediatorDomain
    Call mEvMediatorDomain.Initialize(mFSMonitoringCoord, mOpportunities, mRibbon, mFileMgr)
    
    ' INTERFAZ
    Set mRibbon = New clsRibbon
    
    LogInfo MODULE_NAME, "[Class_Initialize] - inicializaci?n completada"
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Class_Initialize]"
    Err.Raise Err.Number, MODULE_NAME & "[Class_Initialize]", _
              "Error inicializando aplicaci?n: " & Err.Description
End Sub

' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Terminate] - Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza contin?a aunque haya errores
    
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

    LogInfo MODULE_NAME, "[Terminate] - Limpieza completada"
End Sub

