VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Manager del contexto de ejecuciÃ³n Excel. Recibe eventos del Listener (ThisWorkbook),
'              actualiza State, emite eventos semÃ³nticos de infraestructura.
'@Category: Manager
'@Folder "2-Infraestructura"
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextMgr"

' ==========================================
' EVENTOS SEMÃ³NTICOS DE INFRAESTRUCTURA
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
' INICIALIZACIÃ³N
' ==========================================

Public Sub Initialize(oState As clsExecutionContextState)
    Set mState = oState
    RefreshContextState
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
' MÃ³TODOS DEL LISTENER (llamados por ThisWorkbook)
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
'@Note: Llamado por ThisWorkbook cuando detecta que el contexto se ha invalidado.
'       Reconstruye State desde el estado real de Excel y notifica a consumidores.
Public Sub ReInitializeContext()
    mState.InvalidateContext
    RaiseEvent ContextInvalidated()
    RefreshContextState
    mState.ValidateContext
    RaiseEvent ContextReinitialized()
    LogInfo MODULE_NAME, "[ReInitializeContext] - Contexto reinicializado tras reset"
End Sub

'@Description: Verifica si el contexto completo es valido (referencias COM vivas)
'@Returns: True si todas las referencias criticas estan vivas y responden
'@Note: Usa accesos reales a propiedades para verificar que los objetos no son zombis
Public Function IsContextValid() As Boolean
    On Error GoTo InvalidContext

    ' 1. Verificar que tenemos State
    If mState Is Nothing Then
        LogWarning MODULE_NAME, "[IsContextValid] mState Is Nothing"
        GoTo InvalidContext
    End If

    ' 2. Verificar que Application existe y responde
    If mState.Application Is Nothing Then
        LogWarning MODULE_NAME, "[IsContextValid] Application Is Nothing"
        GoTo InvalidContext
    End If

    ' 3. Verificar que Application responde (no es referencia zombi)
    Dim testName As String
    testName = mState.Application.Name
    If Len(testName) = 0 Then
        LogWarning MODULE_NAME, "[IsContextValid] Application.Name vacio"
        GoTo InvalidContext
    End If

    ' 4. Verificar que podemos acceder a Workbooks
    Dim testCount As Long
    testCount = mState.Application.Workbooks.Count

    ' 5. Verificar flag de validez del State
    If Not mState.IsContextValid Then
        LogWarning MODULE_NAME, "[IsContextValid] State.IsContextValid = False"
        GoTo InvalidContext
    End If

    IsContextValid = True
    Exit Function

InvalidContext:
    IsContextValid = False
    LogWarning MODULE_NAME, "[IsContextValid] Contexto invalido detectado"
End Function

'@Description: Detecta el contexto actual de Excel (workbook, worksheet, chart)
'              y actualiza el State. Se usa en Initialize (bootstrap) y en
'              ReInitializeContext (tras reset).
Private Sub RefreshContextState()
    mState.SetApplication Excel.Application

    On Error Resume Next
    Dim wb As Workbook
    Set wb = Excel.Application.ActiveWorkbook
    On Error GoTo 0

    If Not wb Is Nothing Then
        mState.UpdateWorkbook wb

        On Error Resume Next
        Dim ws As Worksheet
        Set ws = wb.ActiveSheet   ' Falla si es Chart sheet
        On Error GoTo 0

        If Not ws Is Nothing Then
            mState.UpdateWorksheet ws
        End If

        DetectChart
    End If
End Sub

' ==========================================
' ELECCIÃ³N DE CHART (lÃ³gica privada del Mgr)
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

