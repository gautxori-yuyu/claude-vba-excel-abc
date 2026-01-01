VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartEventsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MissingAnnotationArgument
'@Folder "4-Tipos de ficheros.gas_vbnet.Graficos Rendimiento"
'@ModuleDescription "Gestor centralizado de eventos de gráficos (orquestador)"
Option Explicit

'--------------------------------------------------------------
' Eventos públicos: notifican cambios de contexto gráfico
'--------------------------------------------------------------
'@Description: Se dispara cuando el usuario activa un grafico
'@Category: Gráficos
'@ArgumentDescriptions: chart As Chart
Public Event ChartActivated(Chart As Chart)
'@Description: Se dispara cuando el usuario desactiva un grafico
'@Category: Gráficos
'@ArgumentDescriptions: chart As Chart
Public Event ChartDeactivated(Chart As Chart)

Public Event HojaConGraficosCambiada(sePuedeInvertir As Boolean)
'--------------------------------------------------------------
' Variables privadas
'--------------------------------------------------------------
Private mActiveCharts As Collection              ' Colección de instancias de clsChartEvents activas
Private mWatchingSheet As Object                 ' Hoja actualmente observada

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mActiveCharts = New Collection
    Set mWatchingSheet = Nothing
    Debug.Print "[clsChartEventsManager Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Call StopWatching
    Debug.Print "[clsChartEventsManager Class_Terminate]"
End Sub

'@Description: Inicia la vigilancia de graficos en una hoja (normal o de grafico)
'@Category: Gráficos
'@Scope: argumentos
'@ArgumentDescriptions: sh As Object
Public Sub WatchSheet(sh As Object)
    Dim oChartEvents As clsChartEvents
    
    On Error GoTo ErrHandler
    
    Debug.Print "[clsChartEventsManager WatchSheet] - Empezando a observar hoja: " & sh.Name
    
    ' Detener observación anterior (libera eventos)
    ' Call StopWatching
    
    Set mWatchingSheet = sh
    
    ' Caso 1: Hoja de gráfico (Chart Sheet)
    If TypeName(sh) = "Chart" Then
        Set oChartEvents = New clsChartEvents
        Set oChartEvents.mChart = sh
        Set oChartEvents.ParentManager = Me
        mActiveCharts.Add oChartEvents, CStr(ObjPtr(sh))
        Exit Sub
    End If
    
    ' Caso 2: Hoja con gráficos embebidos
    If sh.ChartObjects.Count > 0 Then
        Dim chtObj As ChartObject
        For Each chtObj In sh.ChartObjects
            Set oChartEvents = New clsChartEvents
            Set oChartEvents.mChart = chtObj.Chart
            Set oChartEvents.ParentManager = Me
            mActiveCharts.Add oChartEvents, CStr(ObjPtr(chtObj.Chart))
        Next chtObj
    End If
    
    Debug.Print "[clsChartEventsManager WatchSheet] - gráficos en observación: " & mActiveCharts.Count
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsChartEventsManager WatchSheet] ERROR: " & Err.Description
    Err.Raise Err.Number, "clsChartEventsManager.WatchSheet", _
              "Error al configurar la vigilancia de selección de gráficos en hoja: " & Err.Description
End Sub

'--------------------------------------------------------------
' Detener toda observación
'--------------------------------------------------------------
'@Description: Detiene la vigilancia de todos los graficos actualmente observados
'@Category: Gráficos
'@Scope: interna, auxiliar
Public Sub StopWatching()
    On Error GoTo ErrHandler
    Dim oChartEvents As clsChartEvents
    For Each oChartEvents In mActiveCharts
        Set oChartEvents.mChart = Nothing
    Next oChartEvents
    Set mActiveCharts = New Collection
    Set mWatchingSheet = Nothing
    Debug.Print "[clsChartEventsManager StopWatching] - Eventos de gráficos desactivados"

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsChartEventsManager.StopWatching", _
              "Error al detener la vigilancia de selección de gráficos en hoja: " & Err.Description
End Sub

Public Sub RefreshCurrentSheet()
    On Error GoTo ErrHandler
    
    If Not Application.ActiveSheet Is Nothing Then
        Debug.Print "[clsChartEventsManager RefreshCurrentSheet] Re-inicializando vigilancia"
        Call StopWatching
        Call WatchSheet(Application.ActiveSheet)
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsChartEventsManager RefreshCurrentSheet] ERROR: " & Err.Description
End Sub

'--------------------------------------------------------------
' Métodos internos: llamados por clsChartEvents
'--------------------------------------------------------------
'@Description: Notifica la activacion de un grafico al gestor central
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Friend Sub NotifyChartActivated(cht As Chart)
    RaiseEvent ChartActivated(cht)
End Sub

'@Description: Notifica la desactivacion de un grafico al gestor central
'@Category: Gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Friend Sub NotifyChartDeactivated(cht As Chart)
    RaiseEvent ChartDeactivated(cht)
End Sub


