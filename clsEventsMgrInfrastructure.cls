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

Private mChartManager As clsChartMgr
Private WithEvents mCtxMgr As clsExecutionContextMgr
Attribute mCtxMgr.VB_VarHelpID = -1

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
Public Sub Initialize(oFileMgr As clsFileManager, oChartManager As clsChartMgr, oCtxMgr As clsExecutionContextMgr)
    On Error GoTo ErrHandler
    
    Set mChartManager = oChartManager
    
    Set mCtxMgr = oCtxMgr
    
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
    
    Set mCtxMgr = Nothing
    
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

Private Sub mCtxMgr_WorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookActivated] " & wb.Name
    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtxMgr_SheetActivated(ByVal sh As Object)
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

Private Sub mCtxMgr_SheetDeactivated(ByVal sh As Object)
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
Private Sub mCtxMgr_WorkbookOpened(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookOpened] " & wb.Name
    mFileMgr.GetOrTrackWorkbook wb
End Sub

'@Description: Callback cuando se cierra un workbook
Private Sub mCtxMgr_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtx_WorkbookBeforeClose] " & wb.Name
    mFileMgr.UntrackWorkbook wb
End Sub
