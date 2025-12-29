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
'@Folder "4-Tipos de ficheros.gas_vbnet.Graficos Rendimiento"
'@ModuleDescription "Evento ligero por grafico -- notifica a su gestor"
Option Explicit
'--------------------------------------------------------------
' Miembros
'--------------------------------------------------------------
'@MemberAttribute VB_VarHelpID, -1
Public WithEvents mChart As chart
Attribute mChart.VB_VarHelpID = -1
Private p_ParentChartEventsMgr As clsChartEventsManager

'--------------------------------------------------------------
' Propiedad para asociar al gestor
'--------------------------------------------------------------
'@Description: Asigna el gestor padre al que notificar los eventos de activacion/desactivacion
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: p As clsChartEventsManager
Public Property Set ParentManager(oChartEventsManager As clsChartEventsManager)
    Set p_ParentChartEventsMgr = oChartEventsManager
    Debug.Print "[clsChartEvents] [i] ParentManager " & _
                IIf(oChartEventsManager Is Nothing, "DESASIGNADO", "asignado")
End Property

'--------------------------------------------------------------
' Asignacion segura de mChart (con diagnostico)
'--------------------------------------------------------------
'@Description: Vincula el receptor de eventos con un grafico de Excel
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Public Property Set chart(oChart As chart)
    
    On Error GoTo ErrHandler
    
    If Not mChart Is Nothing Then
        Debug.Print "[clsChartEvents] [!] Reasignando mChart"
        Set mChart = Nothing
    End If
    Set mChart = oChart
    If Not oChart Is Nothing Then
        Debug.Print "[clsChartEvents] [OK] mChart asignado: " & TypeName(oChart.Parent) & "!" & oChart.Name & " @ " & ObjPtr(oChart)
    Else
        Debug.Print "[clsChartEvents] [ERR] mChart liberado"
    End If

    Exit Property
ErrHandler:
    Debug.Print "[clsChartEvents Set chart] error: " & Err.Description
End Property

'--------------------------------------------------------------
' Eventos de Chart
'--------------------------------------------------------------
Private Sub mChart_Activate()
    
    On Error GoTo ErrHandler
    
    Debug.Print "[clsChartEvents.Activate] [i] " & TypeName(mChart.Parent) & "!" & mChart.Name
    If p_ParentChartEventsMgr Is Nothing Then
        Debug.Print "[clsChartEvents.Activate] [ERR] ParentManager es Nothing"
    Else
        p_ParentChartEventsMgr.NotifyChartActivated mChart
    End If

    Exit Sub
ErrHandler:
    Debug.Print "[clsChartEvents mChart_Activate] error: " & Err.Description
End Sub

Private Sub mChart_Deactivate()
    
    On Error GoTo ErrHandler
    
    Debug.Print "[clsChartEvents.Deactivate] [i] " & TypeName(mChart.Parent) & "!" & mChart.Name
    If p_ParentChartEventsMgr Is Nothing Then
        Debug.Print "[clsChartEvents.Deactivate] [ERR] ParentManager es Nothing"
    Else
        p_ParentChartEventsMgr.NotifyChartDeactivated mChart
    End If

    Exit Sub
ErrHandler:
    Debug.Print "[clsChartEvents mChart_Deactivate] error: " & Err.Description
End Sub


