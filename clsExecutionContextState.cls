VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado cached del contexto de ejecuciÃ³n Excel (Workbook, Worksheet, Chart)
'@Category: State
'@Folder "2-Infraestructura.Estado"
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextState"

' ==========================================
' ESTADO CACHED
' ==========================================
Private m_Workbook As Workbook
Private m_Worksheet As Worksheet
Private m_xlApp As Application          ' Referencia para lecturas directas (Selection)

Private Type T_CachedChartInfo
    Chart As Chart
    SourceType As String                ' "ActiveChart", "ChartObject", etc.
    ParentType As String                ' "ChartSheet", "Worksheet"
    IsValid As Boolean
End Type
Private m_chartInfo As T_CachedChartInfo

' Flag de validez del contexto (protecciÃ³n contra reset)
Private m_IsContextValid As Boolean

' ==========================================
' INICIALIZACIÃN
' ==========================================

Friend Sub SetApplication(oApp As Application)
    Set m_xlApp = oApp
    m_IsContextValid = True
End Sub

Private Sub Class_Terminate()
    Set m_xlApp = Nothing
    Set m_Workbook = Nothing
    Set m_Worksheet = Nothing
    m_chartInfo.IsValid = False
    Set m_chartInfo.Chart = Nothing
End Sub

' ==========================================
' CONSULTAS PÃBLICAS (Property Get)
' ==========================================

Public Property Get Workbook() As Workbook
    On Error GoTo SafeExit
    Set Workbook = m_Workbook
    Exit Property
SafeExit:
    Set Workbook = Nothing
End Property

Public Property Get Worksheet() As Worksheet
    On Error GoTo SafeExit
    Set Worksheet = m_Worksheet
    Exit Property
SafeExit:
    Set Worksheet = Nothing
End Property

Public Property Get Selection() As Object
    On Error GoTo SafeExit
    If m_xlApp Is Nothing Then Exit Property
    Set Selection = m_xlApp.Selection
    Exit Property
SafeExit:
    Set Selection = Nothing
End Property

Public Property Get Application() As Application
    Set Application = m_xlApp
End Property

Public Property Get Chart() As Chart
    On Error GoTo SafeExit
    ' Si el cache es vÃ¡lido, devolver valor cached
    If m_chartInfo.IsValid Then
        Set Chart = m_chartInfo.Chart
        Exit Property
    End If
    ' Fallback de seguridad: ActiveChart directo (no cachea - el Mgr gestiona el cache)
    If Not m_xlApp Is Nothing Then
        If Not m_xlApp.ActiveChart Is Nothing Then
            Set Chart = m_xlApp.ActiveChart
            Exit Property
        End If
    End If
SafeExit:
    Set Chart = Nothing
End Property

Public Property Get IsContextValid() As Boolean
    IsContextValid = m_IsContextValid
End Property

' ==========================================
' CONSULTAS DE CONVENIENCIA
' ==========================================

Public Property Get HasWorkbook() As Boolean
    HasWorkbook = Not (Me.Workbook Is Nothing)
End Property

Public Property Get HasWorksheet() As Boolean
    HasWorksheet = Not (Me.Worksheet Is Nothing)
End Property

Public Property Get HasSelection() As Boolean
    HasSelection = Not (Me.Selection Is Nothing)
End Property

Public Property Get HasChart() As Boolean
    HasChart = Not (Me.Chart Is Nothing)
End Property

Public Function GetSelectedRange() As Range
    On Error GoTo SafeExit
    Dim sel As Object: Set sel = Me.Selection
    If TypeName(sel) = "Range" Then
        Set GetSelectedRange = sel
    End If
SafeExit:
End Function

' ==========================================
' ACTUALIZADORES (Friend - llamados SOLO por clsExecutionContextMgr)
' ==========================================

Friend Sub UpdateWorkbook(wb As Workbook)
    Set m_Workbook = wb
End Sub

Friend Sub UpdateWorksheet(ws As Worksheet)
    Set m_Worksheet = ws
End Sub

Friend Sub UpdateChart(ch As Chart, ByVal sourceType As String, ByVal parentType As String)
    Set m_chartInfo.Chart = ch
    m_chartInfo.SourceType = sourceType
    m_chartInfo.ParentType = parentType
    m_chartInfo.IsValid = Not (ch Is Nothing)
End Sub

Friend Sub InvalidateChart()
    m_chartInfo.IsValid = False
    Set m_chartInfo.Chart = Nothing
End Sub

Friend Sub InvalidateContext()
    m_IsContextValid = False
    Set m_Workbook = Nothing
    Set m_Worksheet = Nothing
    m_chartInfo.IsValid = False
    Set m_chartInfo.Chart = Nothing
End Sub

Friend Sub ValidateContext()
    m_IsContextValid = True
End Sub

' ==========================================
' DIAGNÃSTICO
' ==========================================

Public Function Diagnostics() As String
    Dim s As String
    s = "Contexto actual:" & vbCrLf
    s = s & "  Workbook: " & IIf(Me.HasWorkbook, Me.Workbook.Name, "(ninguno)") & vbCrLf
    s = s & "  Worksheet: " & IIf(Me.HasWorksheet, Me.Worksheet.Name, "(ninguna)") & vbCrLf
    s = s & "  Selection: " & IIf(Me.HasSelection, TypeName(Me.Selection), "(nada)") & vbCrLf
    If Me.HasChart Then
        s = s & "  Chart: Si (" & m_chartInfo.SourceType & " en " & m_chartInfo.ParentType & ")" & vbCrLf
    Else
        s = s & "  Chart: No" & vbCrLf
    End If
    s = s & "  ContextValid: " & CStr(m_IsContextValid) & vbCrLf
    Diagnostics = s
End Function

