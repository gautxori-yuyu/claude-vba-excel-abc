VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Manager del contexto de ejecución Excel. Recibe eventos del Listener (ThisWorkbook),
'              actualiza State, emite eventos semánticos de infraestructura.
'@Category: Manager
'@Folder "2-Infraestructura"
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextMgr"

' ==========================================
' EVENTOS SEMÓNTICOS DE INFRAESTRUCTURA
' ==========================================
Public Event WorkbookOpened(ByVal wb As Workbook)
Public Event WorkbookActivated(ByVal wb As Workbook)
Public Event WorkbookBeforeClose(ByVal wb As Workbook, ByRef Cancel As Boolean)
Public Event WorksheetActivated(ByVal ws As Worksheet)
Public Event WorksheetDeactivated(ByVal ws As Worksheet)
Public Event SheetActivated(ByVal sh As Object)
Public Event SheetDeactivated(ByVal sh As Object)
Public Event ContextInvalidated()
Public Event ContextReinitialized()

Private mState As clsExecutionContextState

' ==========================================
' INICIALIZACIÓN
' ==========================================

Public Sub Initialize(oState As clsExecutionContextState)
    Set mState = oState
    mState.SetApplication Excel.Application
    LogInfo MODULE_NAME, "[Initialize] - Contexto inicializado"
End Sub

'@Description: Expone el State para lectura por consumidores
Public Property Get State() As clsExecutionContextState
    Set State = mState
End Property

Private Sub Class_Terminate()
    Set mState = Nothing
End Sub

' ==========================================
' MÉTODOS DEL LISTENER (llamados por ThisWorkbook)
' ==========================================

Public Sub OnWorkbookOpened(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookOpened] " & wb.Name
    RaiseEvent WorkbookOpened(wb)
End Sub

Public Sub OnWorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookActivated] " & wb.Name
    mState.UpdateWorkbook wb
    RaiseEvent WorkbookActivated(wb)
End Sub

Public Sub OnWorkbookBeforeClose(ByVal wb As Workbook, ByRef Cancel As Boolean)
    LogDebug MODULE_NAME, "[OnWorkbookBeforeClose] " & wb.Name
    RaiseEvent WorkbookBeforeClose(wb, Cancel)
End Sub

Public Sub OnSheetActivated(ByVal sh As Object)
    LogDebug MODULE_NAME, "[OnSheetActivated] " & TypeName(sh) & " '" & sh.Name & "'"
    ' Invalidar cache de chart al cambiar de hoja
    mState.InvalidateChart
    If TypeOf sh Is Worksheet Then
        mState.UpdateWorksheet sh
        RaiseEvent WorksheetActivated(sh)
    End If
    ' Detectar chart en la nueva hoja
    DetectChart
    RaiseEvent SheetActivated(sh)
End Sub

Public Sub OnSheetDeactivated(ByVal sh As Object)
    LogDebug MODULE_NAME, "[OnSheetDeactivated] " & sh.Name
    mState.InvalidateChart
    If TypeOf sh Is Worksheet Then
        RaiseEvent WorksheetDeactivated(sh)
    End If
    RaiseEvent SheetDeactivated(sh)
End Sub

' ==========================================
' SHUTDOWN Y RESET
' ==========================================

'@Description: Shutdown del contexto (llamado por ThisWorkbook en BeforeClose)
Public Sub Shutdown()
    LogInfo MODULE_NAME, "[Shutdown]"
    mState.InvalidateContext
    RaiseEvent ContextInvalidated()
End Sub

'@Description: Reinicializa el contexto tras un reset de VBA/Excel.
'@Note: Este método será llamado por el mediador cuando detecte que el contexto
'       se ha invalidado y necesita ser reconstruido.
'TODO: Implementar reenganche completo de eventos tras reset
Public Sub ReInitializeContext()
    mState.ValidateContext
    mState.SetApplication Excel.Application
    RaiseEvent ContextReinitialized()
    LogInfo MODULE_NAME, "[ReInitializeContext] - Contexto reinicializado tras reset"
End Sub

' ==========================================
' ELECCIÓN DE CHART (lógica privada del Mgr)
' ==========================================

Private Sub DetectChart()
    On Error GoTo ErrHandler
    Dim Ch As Chart

    If mState.Application Is Nothing Then Exit Sub

    ' Caso 1: ActiveChart (ChartSheet o Chart activado en hoja)
    Set Ch = mState.Application.ActiveChart
    If Not Ch Is Nothing Then
        mState.UpdateChart Ch, "ActiveChart", TypeName(Ch.Parent)
        Exit Sub
    End If

    ' Caso 2: Selection es ChartObject o subcomponente
    Dim sel As Object
    Set sel = mState.Selection
    If sel Is Nothing Then Exit Sub

    Select Case TypeName(sel)
        Case "ChartObject"
            Set Ch = sel.Chart
        Case "ChartArea", "PlotArea", "Legend", "Axis", "Series"
            ' Subir hasta encontrar el Chart
            Dim obj As Object: Set obj = sel
            Do While Not (TypeOf obj Is Chart) And Not (obj Is Nothing)
                On Error Resume Next
                Set obj = obj.Parent
                On Error GoTo 0
                If obj Is Nothing Then Exit Do
            Loop
            If TypeOf obj Is Chart Then Set Ch = obj
    End Select

    ' Si encontramos un Chart, actualizarlo en State
    If Not Ch Is Nothing Then
        Dim parentType As String
        If TypeOf Ch.Parent Is Worksheet Then
            parentType = "Worksheet"
        ElseIf TypeOf Ch.Parent Is Chart Then
            parentType = "ChartSheet"
        Else
            parentType = TypeName(Ch.Parent)
        End If
        mState.UpdateChart Ch, TypeName(sel), parentType
    End If

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[DetectChart]"
End Sub

