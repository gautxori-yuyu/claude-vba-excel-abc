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
'@Note: Centraliza la suscripción a eventos de Application y propaga cambios a otras clases.
'       Diseñado para reemplazar ActiveWorkbook/ActiveSheet/ActiveChart/Selection con acceso seguro.
Option Explicit

Private Const MODULE_NAME As String = "clsExecutionContext"

' ==========================================
' EVENTOS PUBLICOS - Otras clases pueden suscribirse
' ==========================================
Public Event WorkbookOpened(ByVal wb As Workbook)
Public Event WorkbookActivated(ByVal wb As Workbook)
Public Event WorkbookBeforeClose(ByVal wb As Workbook, ByRef Cancel As Boolean)
Public Event WorksheetActivated(ByVal ws As Worksheet)
Public Event WorksheetDeactivated(ByVal ws As Worksheet)
Public Event SheetActivated(ByVal sh As Object)
Public Event SheetDeactivated(ByVal sh As Object)
Public Event SelectionChanged(ByVal sel As Object)

' ==========================================
' SUSCRIPCION A EVENTOS DE APPLICATION
' ==========================================
Private WithEvents m_xlApp As Application
Attribute m_xlApp.VB_VarHelpID = -1

' ==========================================
' CACHE DE ESTADO
' ==========================================
Private m_lastWorkbookObjKey As Double
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
' INICIALIZACION
' ---------------------------------------------------------------------------
Public Sub Initialize()
    Set m_xlApp = Application
    LogInfo MODULE_NAME, "[Initialize] - Suscrito a eventos de Application"
End Sub

' ---------------------------------------------------------------------------
' EVENTOS DE APPLICATION - Capturados y propagados
' ---------------------------------------------------------------------------

Private Sub m_xlApp_WorkbookOpen(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[xlApp_WorkbookOpen] " & wb.Name
    RaiseEvent WorkbookOpened(wb)
End Sub

Private Sub m_xlApp_WorkbookActivate(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[xlApp_WorkbookActivate] " & wb.Name
    m_lastWorkbookObjKey = ObjPtr(wb)
    RaiseEvent WorkbookActivated(wb)
End Sub

Private Sub m_xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[xlApp_WorkbookBeforeClose] " & wb.Name
    RaiseEvent WorkbookBeforeClose(wb, Cancel)
End Sub

Private Sub m_xlApp_SheetActivate(ByVal sh As Object)
    LogDebug MODULE_NAME, "[xlApp_SheetActivate] " & TypeName(sh) & " '" & sh.Name & "'"
    If TypeOf sh Is Worksheet Then
        m_lastWorksheetObjKey = ObjPtr(sh)
        RaiseEvent WorksheetActivated(sh)
    End If
    RaiseEvent SheetActivated(sh)
End Sub

Private Sub m_xlApp_SheetDeactivate(ByVal sh As Object)
    LogDebug MODULE_NAME, "[xlApp_SheetDeactivate] " & sh.Name
    If TypeOf sh Is Worksheet Then
        RaiseEvent WorksheetDeactivated(sh)
    End If
    RaiseEvent SheetDeactivated(sh)
End Sub

' ---------------------------------------------------------------------------
' PROPIEDADES PUBLICAS - seguras y con fallback explicito
' ---------------------------------------------------------------------------
Public Property Get Workbook() As Workbook
    On Error GoTo SafeExit
    If m_xlApp Is Nothing Then Exit Property
    If m_xlApp.ActiveWorkbook Is Nothing Then Exit Property
    If ObjPtr(m_xlApp.ActiveWorkbook) = m_lastWorkbookObjKey Then
        Set Workbook = m_xlApp.ActiveWorkbook
        Exit Property
    End If
    ' Refrescar cache
    m_lastWorkbookObjKey = ObjPtr(m_xlApp.ActiveWorkbook)
    Set Workbook = m_xlApp.ActiveWorkbook
    Exit Property
SafeExit:
    Set Workbook = Nothing
End Property

Public Property Get Worksheet() As Worksheet
    On Error GoTo SafeExit
    If m_xlApp Is Nothing Then Exit Property
    If TypeOf m_xlApp.ActiveSheet Is Worksheet Then
        If ObjPtr(m_xlApp.ActiveSheet) = m_lastWorksheetObjKey Then
            Set Worksheet = m_xlApp.ActiveSheet
            Exit Property
        End If
        m_lastWorksheetObjKey = ObjPtr(m_xlApp.ActiveSheet)
        Set Worksheet = m_xlApp.ActiveSheet
    End If
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

' ---------------------------------------------------------------------------
' CHART - LA PARTE CRITICA (logica para detectar chart activo)
' ---------------------------------------------------------------------------
Public Property Get Chart() As Chart
    On Error GoTo SafeExit

    If m_xlApp Is Nothing Then Exit Property

    ' Si ya esta cacheado y valido, devolverlo
    If m_cachedChartInfo.IsValid Then
        Set Chart = m_cachedChartInfo.Chart
        Exit Property
    End If

    Dim Ch As Chart
    Dim sel As Object

    ' 1. Caso 1: ActiveChart (ChartSheet o Chart activado en hoja)
    Set Ch = m_xlApp.ActiveChart
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
                ' Subir hasta encontrar el Chart
                Dim obj As Object: Set obj = sel
                Do While Not (TypeOf obj Is Chart) And Not (obj Is Nothing)
                    On Error Resume Next
                    Set obj = obj.Parent
                    On Error GoTo 0
                    If obj Is Nothing Then Exit Do
                Loop
                If TypeOf obj Is Chart Then Set Ch = obj
            Case Else
                ' No es un grafico - salir
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

    ' No hay grafico activo
    m_cachedChartInfo.IsValid = False
    Set Chart = Nothing
    Exit Property
SafeExit:
    m_cachedChartInfo.IsValid = False
    Set Chart = Nothing
End Property

' ---------------------------------------------------------------------------
' METODOS DE CONVENIENCIA - seguros
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
' DIAGNOSTICO - clave para depuracion
' ---------------------------------------------------------------------------
Public Function Diagnostics() As String
    Dim s As String
    s = "Contexto actual:" & vbCrLf

    s = s & "  Workbook: " & IIf(Me.HasWorkbook, Me.Workbook.Name, "(ninguno)") & vbCrLf
    s = s & "  Worksheet: " & IIf(Me.HasWorksheet, Me.Worksheet.Name, "(ninguna)") & vbCrLf
    s = s & "  Selection: " & IIf(Me.HasSelection, TypeName(Me.Selection), "(nada)") & vbCrLf

    If Me.HasChart Then
        With m_cachedChartInfo
            s = s & "  Chart: Si (" & .SourceType & " en " & .ParentType & ")" & vbCrLf
        End With
    Else
        s = s & "  Chart: No" & vbCrLf
    End If

    ' Diagnostico adicional
    On Error Resume Next
    If Not m_xlApp Is Nothing Then
        s = s & "  App.ActiveChart: " & IIf(m_xlApp.ActiveChart Is Nothing, "Nothing", "Si") & vbCrLf
        s = s & "  App.Selection: " & TypeName(m_xlApp.Selection) & vbCrLf
    End If
    On Error GoTo 0

    Diagnostics = s
End Function

' ---------------------------------------------------------------------------
' LIMPIEZA
' ---------------------------------------------------------------------------
Private Sub Class_Terminate()
    LogDebug MODULE_NAME, "[Class_Terminate]"
    Set m_xlApp = Nothing
    m_cachedChartInfo.IsValid = False
    Set m_cachedChartInfo.Chart = Nothing
End Sub
