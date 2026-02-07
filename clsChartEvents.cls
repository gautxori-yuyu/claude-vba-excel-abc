VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MissingAnnotationArgument
'@Folder "2-Infraestructura.Excel.Charts"
'@ModuleDescription "Evento ligero por grafico -- notifica a su gestor"
Option Explicit

Private Const MODULE_NAME As String = "clsChartEvents"

'--------------------------------------------------------------
' Miembros
'--------------------------------------------------------------
Private WithEvents mChart As Chart
Attribute mChart.VB_VarHelpID = -1
Private mChartMgr As clsChartMgr

'--------------------------------------------------------------
' Propiedad para asociar al gestor
'--------------------------------------------------------------
'@Description: Asigna el gestor padre al que notificar los eventos de activacion/desactivacion
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: p As clsChartMgr
Public Property Set ParentManager(oEvMgrChart As clsChartMgr)
    Set mChartMgr = oEvMgrChart
    LogInfo MODULE_NAME, "[ParentManager] " & IIf(oEvMgrChart Is Nothing, "DESASIGNADO", "asignado")
End Property

'--------------------------------------------------------------
' Asignacion segura de mChart (con diagnostico)
'--------------------------------------------------------------
'@Description: Vincula el receptor de eventos con un grafico de Excel
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Public Property Set Chart(oChart As Chart)
    
    On Error GoTo ErrHandler
    
    If Not mChart Is Nothing Then
        LogWarning MODULE_NAME, "[Chart Set] Reasignando mChart"
        Set mChart = Nothing
    End If
    Set mChart = oChart
    If Not oChart Is Nothing Then
        LogInfo MODULE_NAME, " - mChart asignado: " & TypeName(oChart.Parent) & "!" & oChart.Name & " @ " & ObjPtr(oChart)
    Else
        LogError MODULE_NAME, " - mChart liberado"
    End If

    Exit Property
ErrHandler:
    LogCurrentError MODULE_NAME, " [Chart]"
    Err.Raise Err.Number, MODULE_NAME & " [Chart]", Err.Description
End Property

'--------------------------------------------------------------
' Eventos de Chart
'--------------------------------------------------------------
Private Sub mChart_Activate()
    
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[clsChartEvents.Activate] " & TypeName(mChart.Parent) & "!" & mChart.Name
    If mChartMgr Is Nothing Then
        LogError MODULE_NAME, "[clsChartEvents.Activate] ParentManager es Nothing"
    Else
        mChartMgr.NotifyChartActivated mChart
    End If

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, " [clsChartEvents.Activate]"
    Err.Raise Err.Number, MODULE_NAME & " [clsChartEvents.Activate]", Err.Description
End Sub

Private Sub mChart_Deactivate()
    
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[clsChartEvents.Deactivate] " & TypeName(mChart.Parent) & "!" & mChart.Name
    If mChartMgr Is Nothing Then
        LogError MODULE_NAME, "[clsChartEvents.Deactivate] ParentManager es Nothing"
    Else
        mChartMgr.NotifyChartDeactivated mChart
    End If

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, " [clsChartEvents.Deactivate]"
    Err.Raise Err.Number, MODULE_NAME & " [clsChartEvents.Deactivate]", Err.Description
End Sub


