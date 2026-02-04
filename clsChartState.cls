VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado relacionado con el grÃ¡fico (Chart) activo
'@Category: State
'@Folder "5-Aplicac (Coord).Estado"
Option Explicit

Private Const MODULE_NAME As String = "clsChartState"

' ==========================================
' ESTADO DE CHART
' ==========================================
Private mCurrentChart As Chart

' ==========================================
' INICIALIZACIÃN
' ==========================================

Private Sub Class_Initialize()
    Set mCurrentChart = Nothing
End Sub

' ==========================================
' PROPIEDADES PÃBLICAS
' ==========================================

'@Description: Obtiene o establece el Chart activo actual
Public Property Get CurrentChart() As Chart
    Set CurrentChart = mCurrentChart
End Property

Friend Property Set CurrentChart(Ch As Chart)
    Set mCurrentChart = Ch
End Property

' ==========================================
' CONSULTAS DE ESTADO
' ==========================================

'@Description: Indica si hay un grÃ¡fico actualmente activo/seleccionado
Public Function HasCurrentChart() As Boolean
    HasCurrentChart = Not (mCurrentChart Is Nothing)
End Function

'@Description: Obtiene el nombre del chart actual
'@Returns: String con el nombre del chart o "(ninguno)"
Public Function GetCurrentChartName() As String
    If mCurrentChart Is Nothing Then
        GetCurrentChartName = "(ninguno)"
    Else
        On Error Resume Next
        GetCurrentChartName = mCurrentChart.Name
        If Err.Number <> 0 Then
            GetCurrentChartName = "(error al obtener nombre)"
        End If
        On Error GoTo 0
    End If
End Function

