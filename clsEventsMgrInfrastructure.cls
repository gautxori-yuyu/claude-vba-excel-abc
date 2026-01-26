VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMgrInfrastructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura"
'@ModuleDescription "Gestor centralizado de eventos (orquestador)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMgrInfrastructure"

Private WithEvents mXlApp As Application
Attribute mXlApp.VB_VarHelpID = -1
Private WithEvents mChartManager As clsChartMgr
Attribute mChartManager.VB_VarHelpID = -1

Private mRibbon As clsRibbon

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    Set mXlApp = Application
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

Public Sub Initialize(oRibbon As clsRibbon, oChartManager As clsChartMgr)
    On Error GoTo ErrHandler
    
    Set mRibbon = oRibbon
    
    Set mChartManager = oChartManager
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", _
              "Error inicializando aplicación: " & Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler
    
    Set mChartManager = Nothing
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicación: " & Err.Description
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Eventos de Application
' -------------------------------------------------------------

'@Description: Se dispara cuando se activa un workbook
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: wb As Workbook

Private Sub mXlApp_WorkbookActivate(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[mXlApp_WorkbookActivated] " & Wb.Name
    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mXlApp_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mXlApp_SheetActivated] " & Sh.Name & "'"
    mChartManager.WatchSheet Sh
    mRibbon.InvalidarControl "btnGenerarGraficos"
    mRibbon.InvalidarControl "btnCGASING"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mXlApp_SheetActivated] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mXlApp_SheetDeactivate(ByVal Sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mXlApp_SheetDeactivated] '" & Sh.Name & "'"
    mChartManager.StopWatching
    mRibbon.InvalidarControl "btnGenerarGraficos"
    mRibbon.InvalidarControl "btnCGASING"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mXlApp_SheetDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de cambio de seleccion de charts (desde ChartManager)
' -------------------------------------------------------------

Private Sub mChartManager_ChartActivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartMgr ChartActivated] - Gráfico activado: " & cht.Name
    ' Aquí decides:
    ' - ¿Invalidar ribbon?
    ' - ¿Cargar datos del gráfico?
    ' - ¿Actualizar estado de la oportunidad?
    mRibbon.InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartManager_ChartDeactivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartMgr ChartActivated] - Gráfico DESactivado: " & cht.Name
    mRibbon.InvalidarControl "btnInvertirSeries"
End Sub
