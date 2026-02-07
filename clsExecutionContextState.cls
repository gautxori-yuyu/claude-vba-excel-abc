VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado cached del contexto de ejecución Excel (Workbook, Worksheet, Chart)
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
    sourceType As String                ' "ActiveChart", "ChartObject", etc.
    parentType As String                ' "ChartSheet", "Worksheet"
    IsValid As Boolean
End Type
Private m_chartInfo As T_CachedChartInfo

' Flag de validez del contexto (protección contra reset)
Private m_IsContextValid As Boolean

' ==========================================
' INICIALIZACIÓN
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
' CONSULTAS PÜBLICAS (Property Get)
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
    ' Si el cache es válido, devolver valor cached
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

Friend Sub UpdateWorkbook(Wb As Workbook)
    Set m_Workbook = Wb
End Sub

Friend Sub UpdateWorksheet(ws As Worksheet)
    Set m_Worksheet = ws
End Sub

Friend Sub UpdateChart(Ch As Chart, ByVal sourceType As String, ByVal parentType As String)
    Set m_chartInfo.Chart = Ch
    m_chartInfo.sourceType = sourceType
    m_chartInfo.parentType = parentType
    m_chartInfo.IsValid = Not (Ch Is Nothing)
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
' DIAGNÓSTICO
' ==========================================

Public Function Diagnostics() As String
    Dim s As String
    s = "Contexto actual:" & vbCrLf
    s = s & "  Workbook: " & IIf(Me.HasWorkbook, Me.Workbook.Name, "(ninguno)") & vbCrLf
    s = s & "  Worksheet: " & IIf(Me.HasWorksheet, Me.Worksheet.Name, "(ninguna)") & vbCrLf
    s = s & "  Selection: " & IIf(Me.HasSelection, TypeName(Me.Selection), "(nada)") & vbCrLf
    If Me.HasChart Then
        s = s & "  Chart: Si (" & m_chartInfo.sourceType & " en " & m_chartInfo.parentType & ")" & vbCrLf
    Else
        s = s & "  Chart: No" & vbCrLf
    End If
    s = s & "  ContextValid: " & CStr(m_IsContextValid) & vbCrLf
    Diagnostics = s
End Function

