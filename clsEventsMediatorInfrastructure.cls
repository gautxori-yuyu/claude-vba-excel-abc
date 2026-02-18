VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMediatorInfrastructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.2-Mediadores"
'@ModuleDescription "Mediador de eventos de infraestructura: traduce eventos de Excel a acciones sobre FileManager"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMediatorInfrastructure"

' ==================================================================
' REFERENCIAS INTERNAS
' ==================================================================
Private WithEvents mCtxMgr As clsExecutionContextMgr
Attribute mCtxMgr.VB_VarHelpID = -1

Private mFileMgr As clsFileManager

' ==================================================================
' INICIALIZACIÓN
' ==================================================================
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

Public Sub Initialize(oFileMgr As clsFileManager, oCtxMgr As clsExecutionContextMgr)
    On Error GoTo ErrHandler
    Set mCtxMgr = oCtxMgr
    Set mFileMgr = oFileMgr
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", _
              "Error inicializando mediador: " & Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler
    Set mCtxMgr = Nothing
    Set mFileMgr = Nothing
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
End Sub

' ==================================================================
' EVENTOS DE ExecutionContext
' ==================================================================

'@Description: Workbook activado -> tracking + sincronización con FileManager
Private Sub mCtxMgr_WorkbookActivated(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookActivated] " & Wb.Name & " -> GetOrTrackWorkbook + SetActiveFile"
    If Not mFileMgr Is Nothing Then
        Dim f As clsFileXLS
        Set f = mFileMgr.GetOrTrackWorkbook(Wb)
        mFileMgr.SetActiveFile f
    End If
End Sub

'@Description: Hoja activada (logging; chart watch delegado a clsExecutionContextMgr)
Private Sub mCtxMgr_SheetActivated(ByVal Sh As Object)
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[mCtx_SheetActivated] " & Sh.Name
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetActivated] Error", , Err.Description
End Sub

Private Sub mCtxMgr_SheetDeactivated(ByVal Sh As Object)
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[mCtx_SheetDeactivated] '" & Sh.Name & "'"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetDeactivated] Error", , Err.Description
End Sub

'@Description: Workbook abierto (solo notificación; tracking real en WorkbookActivated)
Private Sub mCtxMgr_WorkbookOpened(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookOpened] " & Wb.Name & " (solo notificación; tracking en WorkbookActivated)"
End Sub

'@Description: Workbook cerrándose -> untrack
Private Sub mCtxMgr_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtx_WorkbookBeforeClose] " & Wb.Name
    If Not mFileMgr Is Nothing Then
        mFileMgr.UntrackWorkbook Wb
    End If
End Sub

'@Description: Contexto invalidado (log solamente; sin propagación de evento)
Private Sub mCtxMgr_ContextInvalidated()
    On Error Resume Next
    LogWarning MODULE_NAME, "[mCtxMgr_ContextInvalidated] Contexto invalidado"
    On Error GoTo 0
End Sub

'@Description: Contexto reinicializado -> re-trackear workbook activo
Private Sub mCtxMgr_ContextReinitialized()
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[mCtxMgr_ContextReinitialized] Re-enganchando"

    Dim Wb As Workbook
    Set Wb = Excel.Application.ActiveWorkbook
    If Not Wb Is Nothing Then
        mFileMgr.GetOrTrackWorkbook Wb
    End If

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtxMgr_ContextReinitialized] Error durante reconexión", , Err.Description
End Sub
