VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutionContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3-Aplicac (Coord)"
'@Description: Provee acceso seguro y explícito al contexto de ejecución (Workbook, Worksheet, Chart, Selection)
'@Note: Diseñado para reemplazar ActiveWorkbook/ActiveSheet/ActiveChart/Selection con acceso seguro y consistente.
'       Basado en la lógica real de EsValidoInvertirEjes y clsChartEvents.
Option Explicit

Public Event WorkbookChanged(ByVal wb As Workbook)
Public Event WorksheetChanged(ByVal ws As Worksheet)
Public Event SelectionChanged(ByVal sel As Object)

Private m_app As Application
Private m_lastWorkbookObjKey As Double   ' Cache coherente (solo si sigue vigente)
Private m_lastWorksheetObjKey As Double
Private m_lastChartObjKey As Double
Private m_lastSelectionObjKey As Double

Private Type T_CachedChartInfo
    IsValid As Boolean
    Chart As Chart                      ' Chart activo (ActiveChart o desde ChartObject/Selection)
    SourceType As String                ' "ActiveChart", "ChartObject", "ChartArea", etc.
    ParentType As String                ' "ChartSheet", "Worksheet"
End Type
Private m_cachedChartInfo As T_CachedChartInfo

' ---------------------------------------------------------------------------
' INICIALIZACIÓN
' ---------------------------------------------------------------------------
Public Sub Initialize()
    Set m_app = Application
    ' No cacheamos nada aquí ? lazy evaluation segura
End Sub

' ---------------------------------------------------------------------------
' PROPIEDADES PÚBLICAS — seguras y con fallback explícito
' ---------------------------------------------------------------------------
Public Property Get Workbook() As Workbook
    On Error GoTo SafeExit
    If m_app.ActiveWorkbook Is Nothing Then Exit Property
    If ObjPtr(m_app.ActiveWorkbook) = m_lastWorkbookObjKey Then
        Set Workbook = m_app.ActiveWorkbook
        Exit Property
    End If
    ' Refrescar cache
    m_lastWorkbookObjKey = ObjPtr(m_app.ActiveWorkbook)
    Set Workbook = m_app.ActiveWorkbook
    Exit Property
SafeExit:
    Set Workbook = Nothing
End Property

Public Property Get Worksheet() As Worksheet
    On Error GoTo SafeExit
    If TypeOf m_app.ActiveSheet Is Worksheet Then
        If ObjPtr(m_app.ActiveSheet) = m_lastWorksheetObjKey Then
            Set Worksheet = m_app.ActiveSheet
            Exit Property
        End If
        m_lastWorksheetObjKey = ObjPtr(m_app.ActiveSheet)
        Set Worksheet = m_app.ActiveSheet
    End If
    Exit Property
SafeExit:
    Set Worksheet = Nothing
End Property

Public Property Get Selection() As Object
    On Error GoTo SafeExit
    Set Selection = m_app.Selection
    Exit Property
SafeExit:
    Set Selection = Nothing
End Property

Public Property Get Application() As Application
    Set Application = m_app
End Property

' ---------------------------------------------------------------------------
' CHART — LA PARTE CRÍTICA (reescrita para tu lógica exacta)
' ---------------------------------------------------------------------------
Public Property Get Chart() As Chart
    On Error GoTo SafeExit
    
    ' Si ya está cacheado y válido, devolverlo
    If m_cachedChartInfo.IsValid Then
        Set Chart = m_cachedChartInfo.Chart
        Exit Property
    End If

    ' === LÓGICA IDÉNTICA A TU EsValidoInvertirEjes ===
    Dim Ch As Chart
    Dim sel As Object
    Dim ws As Worksheet

    ' 1. Caso 1: ActiveChart (ChartSheet o Chart activado en hoja)
    Set Ch = m_app.ActiveChart
    If Not Ch Is Nothing Then
        m_cachedChartInfo.Chart = Ch
        m_cachedChartInfo.SourceType = "ActiveChart"
        m_cachedChartInfo.ParentType = TypeName(Ch.Parent)
        m_cachedChartInfo.IsValid = True
        Set Chart = Ch
        Exit Property
    End If

    ' 2. Caso 2: Selection es ChartObject, ChartArea, PlotArea, etc.
    Set sel = Me.Selection
    If Not sel Is Nothing Then
        Select Case TypeName(sel)
            Case "ChartObject"
                Set Ch = sel.Chart
            Case "ChartArea", "PlotArea", "Legend", "Axis", "Series"
                ' Subir hasta encontrar el Chart (como haces tú)
                Dim obj As Object: Set obj = sel
                Do While Not (TypeOf obj Is Chart) And Not (obj Is Nothing)
                    On Error Resume Next
                    Set obj = obj.Parent
                    On Error GoTo 0
                    If obj Is Nothing Then Exit Do
                Loop
                If TypeOf obj Is Chart Then Set Ch = obj
            Case Else
                ' No es un gráfico ? salir
        End Select
    End If

    ' Si encontramos un Chart, cachearlo
    If Not Ch Is Nothing Then
        m_cachedChartInfo.Chart = Ch
        m_cachedChartInfo.SourceType = TypeName(sel)
        If TypeOf Ch.Parent Is Worksheet Then
            m_cachedChartInfo.ParentType = "Worksheet"
        ElseIf TypeOf Ch.Parent Is Chart Then
            m_cachedChartInfo.ParentType = "ChartSheet"
        Else
            m_cachedChartInfo.ParentType = TypeName(Ch.Parent)
        End If
        m_cachedChartInfo.IsValid = True
        Set Chart = Ch
        Exit Property
    End If

    ' No hay gráfico activo
    m_cachedChartInfo.IsValid = False
    Set Chart = Nothing
    Exit Property
SafeExit:
    m_cachedChartInfo.IsValid = False
    Set Chart = Nothing
End Property

' ---------------------------------------------------------------------------
' MÉTODOS DE CONVENIENCIA — idénticos a los tuyos, pero seguros
' ---------------------------------------------------------------------------
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
    ' Nada ? Nothing por defecto
End Function

' ---------------------------------------------------------------------------
' DIAGNÓSTICO — clave para depuración
' ---------------------------------------------------------------------------
Public Function Diagnostics() As String
    Dim s As String
    s = "Contexto actual:" & vbCrLf
    
    s = s & "  Workbook: " & IIf(Me.HasWorkbook, Me.Workbook.Name, "(ninguno)") & vbCrLf
    s = s & "  Worksheet: " & IIf(Me.HasWorksheet, Me.Worksheet.Name, "(ninguna)") & vbCrLf
    s = s & "  Selection: " & IIf(Me.HasSelection, TypeName(Me.Selection), "(nada)") & vbCrLf
    
    If Me.HasChart Then
        With m_cachedChartInfo
            s = s & "  Chart: Sí (" & .SourceType & " en " & .ParentType & ")" & vbCrLf
        End With
    Else
        s = s & "  Chart: No" & vbCrLf
    End If
    
    ' Diagnóstico adicional para depuración de errores de contexto
    On Error Resume Next
    s = s & "  App.ActiveChart: " & IIf(m_app.ActiveChart Is Nothing, "Nothing", "Sí") & vbCrLf
    s = s & "  App.Selection: " & TypeName(m_app.Selection) & vbCrLf
    On Error GoTo 0
    
    Diagnostics = s
End Function

' ---------------------------------------------------------------------------
' LIMPIEZA
' ---------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set m_app = Nothing
    m_cachedChartInfo.IsValid = False
    Set m_cachedChartInfo.Chart = Nothing
End Sub
