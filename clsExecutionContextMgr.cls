VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Manager del contexto de ejecuciï¿½n Excel. Recibe eventos del Listener (ThisWorkbook),
'              actualiza State, emite eventos semï¿½nticos de infraestructura.
'@Category: Manager
'@Folder "2-Infraestructura"
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextMgr"

' ==========================================
' EVENTOS SEMï¿½NTICOS DE INFRAESTRUCTURA
' ==========================================
Public Event WorkbookOpened(ByVal Wb As Workbook)
Public Event WorkbookActivated(ByVal Wb As Workbook)
Public Event WorkbookBeforeClose(ByVal Wb As Workbook, ByRef Cancel As Boolean)
Public Event WorksheetActivated(ByVal ws As Worksheet)
Public Event WorksheetDeactivated(ByVal ws As Worksheet)
Public Event SheetActivated(ByVal Sh As Object)
Public Event SheetDeactivated(ByVal Sh As Object)
Public Event ContextInvalidated()
Public Event ContextReinitialized()

Private mState As clsExecutionContextState
Private mChartMgr As clsChartMgr

' ==========================================
' INICIALIZACIï¿½N
' ==========================================

Private Sub Class_Initialize()
    Set mState = New clsExecutionContextState
    RefreshContextState
    ' clsChartMgr es composicion: nace y muere con el contexto de ejecucion
    Set mChartMgr = New clsChartMgr
    mChartMgr.Initialize mState
    LogInfo MODULE_NAME, "[Initialize] Contexto inicializado"
End Sub

Private Sub Class_Terminate()
    If Not mChartMgr Is Nothing Then
        mChartMgr.StopWatching
        Set mChartMgr = Nothing
    End If
    Set mState = Nothing
End Sub

'@Description: Expone el State para lectura por consumidores
Public Property Get State() As clsExecutionContextState
    Set State = mState
End Property

' ==========================================
' Mï¿½TODOS DE ESTADO DE CONVENIENCIA
' ==========================================
'@Description: Expone el State para lectura por consumidores
Public Property Get ActiveWorkbook() As Workbook
    Set ActiveWorkbook = mState.Workbook
End Property

Public Property Get ActiveWorksheet() As Worksheet
    Set ActiveWorksheet = mState.Worksheet
End Property

Public Property Get Selection() As Object
    Set Selection = mState.Selection
End Property

Public Property Get Application() As Application
    Set Application = mState.Application
End Property

Public Property Get ActiveChart() As Chart
    Set ActiveChart = mState.Chart
End Property

'@Description: Expone el gestor de graficos (composicion interna del contexto)
Public Property Get ChartMgr() As clsChartMgr
    Set ChartMgr = mChartMgr
End Property
' ==========================================
' Mï¿½TODOS DEL LISTENER (llamados por ThisWorkbook)
' ==========================================

Public Sub OnWorkbookOpened(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookOpened] " & Wb.Name
    RaiseEvent WorkbookOpened(Wb)
End Sub

Public Sub OnWorkbookActivated(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[OnWorkbookActivated] " & Wb.Name
    mState.UpdateWorkbook Wb
    mChartMgr.RefreshCurrentSheet
    RaiseEvent WorkbookActivated(Wb)
End Sub

Public Sub OnWorkbookBeforeClose(ByVal Wb As Workbook, ByRef Cancel As Boolean)
    LogDebug MODULE_NAME, "[OnWorkbookBeforeClose] " & Wb.Name
    RaiseEvent WorkbookBeforeClose(Wb, Cancel)
End Sub

Public Sub OnSheetActivated(ByVal Sh As Object)
    LogDebug MODULE_NAME, "[OnSheetActivated] " & TypeName(Sh) & " '" & Sh.Name & "'"
    ' Invalidar cache de chart al cambiar de hoja
    mState.InvalidateChart
    If TypeOf Sh Is Worksheet Then
        mState.UpdateWorksheet Sh
        RaiseEvent WorksheetActivated(Sh)
    End If
    ' Detectar chart en la nueva hoja y arrancar vigilancia de eventos
    DetectChart
    mChartMgr.WatchSheet Sh
    RaiseEvent SheetActivated(Sh)
End Sub

Public Sub OnSheetDeactivated(ByVal Sh As Object)
    LogDebug MODULE_NAME, "[OnSheetDeactivated] " & Sh.Name
    mState.InvalidateChart
    mChartMgr.StopWatching
    If TypeOf Sh Is Worksheet Then
        RaiseEvent WorksheetDeactivated(Sh)
    End If
    RaiseEvent SheetDeactivated(Sh)
End Sub

' ==========================================
' SHUTDOWN Y RESET
' ==========================================

'@Description: Shutdown del contexto (llamado por ThisWorkbook en BeforeClose)
Public Sub Shutdown()
    LogInfo MODULE_NAME, "[Shutdown]"
    mChartMgr.StopWatching
    mState.InvalidateContext
    RaiseEvent ContextInvalidated
End Sub

'@Description: Reinicializa el contexto tras un reset de VBA/Excel.
'@Note: Llamado por ThisWorkbook cuando detecta que el contexto se ha invalidado.
'       Reconstruye State desde el estado real de Excel y notifica a consumidores.
Public Sub ReInitializeContext()
    mState.InvalidateContext
    RaiseEvent ContextInvalidated
    RefreshContextState
    mState.ValidateContext
    mChartMgr.RefreshCurrentSheet
    RaiseEvent ContextReinitialized
    LogInfo MODULE_NAME, "[ReInitializeContext] Contexto reinicializado tras reset"
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
    Dim Wb As Workbook
    Set Wb = Excel.Application.ActiveWorkbook
    On Error GoTo 0

    If Not Wb Is Nothing Then
        mState.UpdateWorkbook Wb

        On Error Resume Next
        Dim ws As Worksheet
        Set ws = Wb.ActiveSheet   ' Falla si es Chart sheet
        On Error GoTo 0

        If Not ws Is Nothing Then
            mState.UpdateWorksheet ws
        End If

        DetectChart
    End If
End Sub

' ==========================================
' ELECCIï¿½N DE CHART (lï¿½gica privada del Mgr)
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

