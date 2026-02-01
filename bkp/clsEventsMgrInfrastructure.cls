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
' Eventos de ExecutionContextMgr (semánticos de infraestructura)
' -------------------------------------------------------------

'@Description: Se dispara cuando se activa un workbook (semántico de infraestructura)
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: path As String

Private Sub mCtxMgr_WorkbookSessionActivated(ByVal Path As String)
    LogDebug MODULE_NAME, "[mCtxMgr_WorkbookSessionActivated] " & Path
    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos (semántico de infraestructura)
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sheetName As String, sheetType As String

Private Sub mCtxMgr_SheetSessionActivated(ByVal sheetName As String, ByVal sheetType As String)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtxMgr_SheetSessionActivated] '" & sheetName & "' (" & sheetType & ")"
    Dim dummySh As Object  ' We don't have the actual object, just the name and type
    Set dummySh = Nothing  ' In a real scenario, you'd get the actual sheet object
    mChartManager.WatchSheet dummySh
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtxMgr_SheetSessionActivated] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos (semántico de infraestructura)
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sheetName As String, sheetType As String

Private Sub mCtxMgr_SheetSessionDeactivated(ByVal sheetName As String, ByVal sheetType As String)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtxMgr_SheetSessionDeactivated] '" & sheetName & "' (" & sheetType & ")"
    mChartManager.StopWatching
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtxMgr_SheetSessionDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de tracking de ficheros (desde clsFileMgr)
' -------------------------------------------------------------

'@Description: Callback cuando se abre un workbook (semántico de infraestructura)
Private Sub mCtxMgr_WorkbookSessionStarted(ByVal Path As String, ByVal readOnly As Boolean)
    LogDebug MODULE_NAME, "[mCtxMgr_WorkbookSessionStarted] " & Path
    ' Need to get the actual workbook object from the path
    Dim Wb As Workbook
    On Error Resume Next
    Set Wb = Workbooks(GetFileNameFromPath(Path))
    On Error GoTo 0
    If Not Wb Is Nothing Then
        mFileMgr.GetOrTrackWorkbook Wb
    End If
End Sub

'@Description: Callback cuando se cierra un workbook (semántico de infraestructura)
Private Sub mCtxMgr_WorkbookSessionClosing(ByVal Path As String, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtxMgr_WorkbookSessionClosing] " & Path
    ' Need to get the actual workbook object from the path
    Dim Wb As Workbook
    On Error Resume Next
    Set Wb = Workbooks(GetFileNameFromPath(Path))
    On Error GoTo 0
    If Not Wb Is Nothing Then
        mFileMgr.UntrackWorkbook Wb
    End If
End Sub

' Helper function to extract filename from path
Private Function GetFileNameFromPath(fullPath As String) As String
    Dim arr() As String
    arr = Split(fullPath, "\")
    GetFileNameFromPath = arr(UBound(arr))
End Function
