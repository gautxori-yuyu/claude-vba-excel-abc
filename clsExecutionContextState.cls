VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContextState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.Estado"
'@Description: Estado t�cnico de contexto de ejecuci�n (Workbook, Worksheet, Chart, Selection)
'@Note: Solo memoria estructurada, no comportamiento. No emite eventos.
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContextState"

' ==========================================
' CACHE DE ESTADO T�CNICO
' ==========================================
Private m_lastWorkbookObjKey As Double   ' Cache coherente (solo si sigue vigente)
Private m_lastWorksheetObjKey As Double
Private m_lastChartObjKey As Double
Private m_lastSelectionObjKey As Double

' Structure to hold chart info - using individual properties since UDTs can't be returned from public functions
Private m_cachedChartInfo_IsValid As Boolean
Private m_cachedChartInfo_Chart As Chart
Private m_cachedChartInfo_SourceType As String
Private m_cachedChartInfo_ParentType As String

' ==========================================
' ESTADO T�CNICO
' ==========================================
Private m_activeWorkbook As Workbook
Private m_activeWorksheet As Worksheet
Private m_activeSelection As Object
Private m_activeChart As Chart

' ==========================================
' INICIALIZACI�N
' ==========================================
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"

    ' Inicializar estructura de chart info
    m_cachedChartInfo_IsValid = False
    Set m_cachedChartInfo_Chart = Nothing
    m_cachedChartInfo_SourceType = ""
    m_cachedChartInfo_ParentType = ""
End Sub

Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
    ClearState
End Sub

' ==========================================
' M�TODOS DE MANIPULACI�N DE ESTADO (llamados solo por el Manager)
' ==========================================
Friend Sub SetActiveWorkbook(wb As Workbook)
    Set m_activeWorkbook = wb
    If Not wb Is Nothing Then
        m_lastWorkbookObjKey = ObjPtr(wb)
    End If
End Sub

Friend Sub SetActiveWorksheet(ws As Worksheet)
    Set m_activeWorksheet = ws
    If Not ws Is Nothing Then
        m_lastWorksheetObjKey = ObjPtr(ws)
    End If
End Sub

Friend Sub SetActiveSelection(sel As Object)
    Set m_activeSelection = sel
    If Not sel Is Nothing Then
        m_lastSelectionObjKey = ObjPtr(sel)
    End If
End Sub

Friend Sub SetActiveChart(cht As Chart)
    Set m_activeChart = cht
    If Not cht Is Nothing Then
        m_lastChartObjKey = ObjPtr(cht)
    End If

    ' Actualizar cache de chart info
    If Not cht Is Nothing Then
        Set m_cachedChartInfo_Chart = cht
        m_cachedChartInfo_SourceType = "ActiveChart"
        If TypeOf cht.Parent Is Worksheet Then
            m_cachedChartInfo_ParentType = "Worksheet"
        ElseIf TypeOf cht.Parent Is Chart Then
            m_cachedChartInfo_ParentType = "ChartSheet"
        Else
            m_cachedChartInfo_ParentType = TypeName(cht.Parent)
        End If
        m_cachedChartInfo_IsValid = True
    Else
        m_cachedChartInfo_IsValid = False
        Set m_cachedChartInfo_Chart = Nothing
    End If
End Sub

' ==========================================
' M�TODOS DE LIMPIEZA
' ==========================================
Friend Sub ClearState()
    Set m_activeWorkbook = Nothing
    Set m_activeWorksheet = Nothing
    Set m_activeSelection = Nothing
    Set m_activeChart = Nothing
    
    m_lastWorkbookObjKey = 0
    m_lastWorksheetObjKey = 0
    m_lastChartObjKey = 0
    m_lastSelectionObjKey = 0
    
    With m_cachedChartInfo
        .IsValid = False
        Set .Chart = Nothing
        .SourceType = ""
        .ParentType = ""
    End With
End Sub

' ==========================================
' PROPIEDADES P�BLICAS - Consulta de estado (solo lectura)
' ==========================================
Public Property Get ActiveWorkbook() As Workbook
    Set ActiveWorkbook = m_activeWorkbook
End Property

Public Property Get ActiveWorksheet() As Worksheet
    Set ActiveWorksheet = m_activeWorksheet
End Property

Public Property Get ActiveSelection() As Object
    Set ActiveSelection = m_activeSelection
End Property

Public Property Get ActiveChart() As Chart
    Set ActiveChart = m_activeChart
End Property

' ==========================================
' M�TODOS DE CONSULTA DE ESTADO
' ==========================================
Public Function HasWorkbook() As Boolean
    HasWorkbook = Not (m_activeWorkbook Is Nothing)
End Function

Public Function HasWorksheet() As Boolean
    HasWorksheet = Not (m_activeWorksheet Is Nothing)
End Function

Public Function HasSelection() As Boolean
    HasSelection = Not (m_activeSelection Is Nothing)
End Function

Public Function HasChart() As Boolean
    HasChart = Not (m_activeChart Is Nothing)
End Function

' Properties to access chart info individually since we can't return a UDT from a public function
Public Property Get CachedChartInfo_IsValid() As Boolean
    CachedChartInfo_IsValid = m_cachedChartInfo_IsValid
End Property

Public Property Get CachedChartInfo_Chart() As Chart
    Set CachedChartInfo_Chart = m_cachedChartInfo_Chart
End Property

Public Property Get CachedChartInfo_SourceType() As String
    CachedChartInfo_SourceType = m_cachedChartInfo_SourceType
End Property

Public Property Get CachedChartInfo_ParentType() As String
    CachedChartInfo_ParentType = m_cachedChartInfo_ParentType
End Property

Public Function GetLastWorkbookObjKey() As Double
    GetLastWorkbookObjKey = m_lastWorkbookObjKey
End Function

Public Function GetLastWorksheetObjKey() As Double
    GetLastWorksheetObjKey = m_lastWorksheetObjKey
End Function

Public Function GetLastChartObjKey() As Double
    GetLastChartObjKey = m_lastChartObjKey
End Function

Public Function GetLastSelectionObjKey() As Double
    GetLastSelectionObjKey = m_lastSelectionObjKey
End Function

' ==========================================
' M�TODOS DE VALIDACI�N DE OBJETOS
' ==========================================
Public Function IsWorkbookValid(wb As Workbook) As Boolean
    If wb Is Nothing Then
        IsWorkbookValid = False
        Exit Function
    End If
    
    If m_activeWorkbook Is Nothing Then
        IsWorkbookValid = False
        Exit Function
    End If
    
    ' Comparar por ObjPtr para verificar que es el mismo objeto
    IsWorkbookValid = (ObjPtr(wb) = ObjPtr(m_activeWorkbook))
End Function

Public Function IsWorksheetValid(ws As Worksheet) As Boolean
    If ws Is Nothing Then
        IsWorksheetValid = False
        Exit Function
    End If
    
    If m_activeWorksheet Is Nothing Then
        IsWorksheetValid = False
        Exit Function
    End If
    
    ' Comparar por ObjPtr para verificar que es el mismo objeto
    IsWorksheetValid = (ObjPtr(ws) = ObjPtr(m_activeWorksheet))
End Function

' ==========================================
' DIAGN�STICO
' ==========================================
Public Function Diagnostics() As String
    Dim s As String
    s = "Estado t�cnico de contexto:" & vbCrLf
    
    s = s & "  ActiveWorkbook: " & IIf(HasWorkbook(), m_activeWorkbook.Name, "(ninguno)") & vbCrLf
    s = s & "  ActiveWorksheet: " & IIf(HasWorksheet(), m_activeWorksheet.Name, "(ninguna)") & vbCrLf
    s = s & "  ActiveSelection: " & IIf(HasSelection(), TypeName(m_activeSelection), "(nada)") & vbCrLf
    
    If HasChart() Then
        s = s & "  ActiveChart: Sí (" & CachedChartInfo_SourceType & " en " & CachedChartInfo_ParentType & ")" & vbCrLf
    Else
        s = s & "  ActiveChart: No" & vbCrLf
    End If
    
    Diagnostics = s
End Function