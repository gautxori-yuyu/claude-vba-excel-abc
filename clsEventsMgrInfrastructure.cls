VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMgrInfrastructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura"
'@ModuleDescription "Gestor centralizado de eventos (orquestador)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMgrInfrastructure"

Private WithEvents mChartManager As clsChartMgr
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mCtx As clsExecutionContext
Attribute mCtx.VB_VarHelpID = -1

Private mFileMgr As clsFileManager

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub
Public Sub Initialize(oFileMgr As clsFileManager, oChartManager As clsChartMgr, oCtx As clsExecutionContext)
    On Error GoTo ErrHandler
    
    Set mChartManager = oChartManager
    
    Set mCtx = oCtx
    
    Set mFileMgr = oFileMgr
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", _
              "Error inicializando aplicación: " & Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler
    
    Set mChartManager = Nothing
    
    Set mCtx = Nothing
    
    Set mFileMgr = Nothing
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicación: " & Err.Description
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Eventos de ExecutionContext
' -------------------------------------------------------------

'@Description: Se dispara cuando se activa un workbook
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: wb As Workbook

Private Sub mCtx_WorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookActivated] " & wb.Name
    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtx_SheetActivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetActivated] " & sh.Name & "'"
    mChartManager.WatchSheet sh
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetActivated] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtx_SheetDeactivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetDeactivated] '" & sh.Name & "'"
    mChartManager.StopWatching
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de tracking de ficheros (desde clsFileMgr)
' -------------------------------------------------------------

'@Description: Callback cuando se abre un workbook
Private Sub mCtx_WorkbookOpened(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookOpened] " & wb.Name
    mFileMgr.GetOrTrackWorkbook wb
End Sub

'@Description: Callback cuando se cierra un workbook
Private Sub mCtx_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtx_WorkbookBeforeClose] " & wb.Name
    mFileMgr.UntrackWorkbook wb
End Sub

' -------------------------------------------------------------
' Eventos de cambio de seleccion de charts (desde ChartManager)
' -------------------------------------------------------------

Private Sub mChartManager_ChartActivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartMgr ChartActivated] - Gráfico activado: " & cht.Name
End Sub

Private Sub mChartManager_ChartDeactivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartMgr ChartActivated] - Gráfico DESactivado: " & cht.Name
End Sub
