VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule VariableNotAssigned
'@Folder "6-DOMINIO-Oportunidades y compresores"
'@ModuleDescription "Gestor de eventos de graficos (ChartEvents) -- notifica a su coordinador"
Option Explicit

Private Const MODULE_NAME As String = "clsChartMgr"

' ==========================================
' EVENTOS
' ==========================================
Public Event ChartActivated(Chart As Chart)
Public Event ChartDeactivated(Chart As Chart)

' ==========================================
' VARIABLES MIEMBRO
' ==========================================
Private mCtx As clsExecutionContext

' Coleccion de receptores de eventos de graficos activos
Private mChartsInActiveSheet As Collection

' Estado de chart (composicion - Manager posee el estado)
Private mChartState As clsChartState

' ==========================================
' INICIALIZACION
' ==========================================
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"

    ' Crear estado de Chart (composicion - Manager posee el estado)
    Set mChartState = New clsChartState

    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Call StopWatching

    ' Liberar estado
    Set mChartState = Nothing

    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'@Description: Expone el estado de Chart (para inyeccion en ApplicationState)
Public Property Get State() As clsChartState
    Set State = mChartState
End Property

Public Sub Initialize(oCtx As clsExecutionContext)
    Set mCtx = oCtx
    ' No cacheamos nada aqi ? lazy evaluation segura
    LogInfo MODULE_NAME, "[Initialize] - Suscrito a eventos de Application"
End Sub

'--------------------------------------------------------------
' @Description: Comprueba si el grfico activo es vlido para invertir ejes
'--------------------------------------------------------------
' TODO: unificar esta funcion con lo de clsExecutioContext: alli la funcion Chart DEVUELVE CASI TODO LO QUE HACE ESTA FUNCION,
' queda cambiar el Select Case TypeName(mCtx.Selection) usando la informacion cacheada alli (m_cachedChartInfo) (y sigo pensando que
' clsExecutionContext "ESTA DE SOBRA", lo unico que aporta es el "Cacheado"... y pendiente de incoporarle funcionalidad para,
' cuando se pierde el contexto, **reenganchar eventos, **reconstruir estado derivado, y **aislar el dao del reset)

Public Function EsValidoInvertirEjes() As Boolean
    On Error GoTo SafeExit

    ' 1. Validación segura de contexto mnimo
    If mCtx.Workbook Is Nothing Then Exit Function
    If mCtx.Selection Is Nothing Then Exit Function
    If mCtx.Chart Is Nothing Then Exit Function

    ' 3. Caso de hoja de grfico activa (ChartSheet)
    If TypeName(Application.ActiveSheet) = "Chart" Then
        EsValidoInvertirEjes = True
        GoTo CheckSeries
    End If

    ' 4. Caso de grfico incrustado en hoja de clculo (ChartObject)
    If TypeName(mCtx.Selection) = "ChartObject" Then
        EsValidoInvertirEjes = True
        GoTo CheckSeries
    End If

    ' 5. Caso de rea del grfico activa (ChartArea, Series, Legend, etc.)
    Select Case TypeName(mCtx.Selection)
        Case "ChartArea", "PlotArea", "Legend", "Axis", "Series", "Gridlines", "Trendline"
            EsValidoInvertirEjes = True
            GoTo CheckSeries
    End Select

    ' Ninguno de los casos anteriores
    Exit Function

CheckSeries:
    ' 6. Validación adicional: que el grfico tenga series
    On Error Resume Next
    EsValidoInvertirEjes = (mCtx.Chart.SeriesCollection.Count > 0)
    On Error GoTo 0

SafeExit:
    ' Nada ? False por defecto
End Function

'--------------------------------------------------------------
' @Description: Invierte los ejes X/Y de un grafico activo
'--------------------------------------------------------------
'@Note: Requiere que el grafico este activo y tenga series
'@Returns: True si la operacion fue exitosa, False en caso contrario
Public Function InvertirEjes() As Boolean
    On Error GoTo ErrHandler

    ' Validar contexto
    If Not EsValidoInvertirEjes() Then
        LogWarning MODULE_NAME, "[InvertirEjes] - Contexto no vlido para invertir ejes"
        Exit Function
    End If

    ' Invertir ejes
    mCtx.Chart.Select
    ActiveChart.PlotBy = xlRows  ' Cambia entre xlColumns y xlRows

    ' Registrar xito
    InvertirEjes = True
    LogInfo MODULE_NAME, "[InvertirEjes] - Ejes invertidos correctamente"

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[InvertirEjes]"
    InvertirEjes = False
End Function

'--------------------------------------------------------------
' Activar vigilancia de graficos en una hoja
'--------------------------------------------------------------
'@Description: Inicia la vigilancia de eventos de graficos en una hoja determinada
'@Category: gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: sh As Object
Public Sub WatchSheet(Sh As Object)
    On Error GoTo ErrHandler

    ' Detener vigilancia previa
    Call StopWatching

    ' Crear coleccion de receptores de eventos
    Set mChartsInActiveSheet = New Collection

    ' Verificar que sh es un objeto que puede tener ChartObjects
    If TypeOf Sh Is Worksheet Then
        ' Iterar sobre objetos grafico en la hoja
        Dim chObj As ChartObject
        For Each chObj In Sh.ChartObjects
            Dim oChartEvents As clsChartEvents
            Set oChartEvents = New clsChartEvents
            Set oChartEvents.ParentManager = Me
            Set oChartEvents.Chart = chObj.Chart
            mChartsInActiveSheet.Add oChartEvents
        Next chObj
    End If

    LogDebug MODULE_NAME, "[WatchSheet] - gráficos en observación: " & mChartsInActiveSheet.Count

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[WatchSheet] Error", , Err.Description
    Err.Raise Err.Number, "clsChartMgr.WatchSheet", _
              "Error al configurar la vigilancia de selección de gráficos en hoja: " & Err.Description
End Sub

'--------------------------------------------------------------
' Detener toda vigilancia
'--------------------------------------------------------------
'@Description: Detiene la vigilancia de todos los graficos actualmente observados
'@Category: gráficos
'@Scope: interna, auxiliar
Public Sub StopWatching()
    On Error GoTo ErrHandler
    
    ' Check if the collection exists before iterating
    If Not mChartsInActiveSheet Is Nothing Then
        Dim oChartEvents As clsChartEvents
        For Each oChartEvents In mChartsInActiveSheet
            Set oChartEvents.Chart = Nothing
        Next oChartEvents
    End If
    
    Set mChartsInActiveSheet = New Collection
    LogDebug MODULE_NAME, "[StopWatching] - Eventos de gráficos desactivados"

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[StopWatching]"
    Err.Raise Err.Number, MODULE_NAME & " [StopWatching] - Error al detener la vigilancia de selección de gráficos en hoja", Err.Description
End Sub

Public Sub RefreshCurrentSheet()
    On Error GoTo ErrHandler

    If Not Application.ActiveSheet Is Nothing Then
        LogDebug MODULE_NAME, "[RefreshCurrentSheet] Re-inicializando vigilancia"
        Call StopWatching
        Call WatchSheet(Application.ActiveSheet)
    End If

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[RefreshCurrentSheet]"
End Sub

'--------------------------------------------------------------
' Mtodos internos: llamados por clsChartEvents
'--------------------------------------------------------------
'@Description: Notifica la activacion de un grafico al gestor central
'@Category: gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Friend Sub NotifyChartActivated(cht As Chart)
    ' Actualizar estado primero
    Set mChartState.CurrentChart = cht
    ' Luego notificar el evento
    RaiseEvent ChartActivated(cht)
End Sub

'@Description: Notifica la desactivacion de un grafico al gestor central
'@Category: gráficos
'@Scope: interna, auxiliar
'@ArgumentDescriptions: cht As Chart
Friend Sub NotifyChartDeactivated(cht As Chart)
    ' Actualizar estado primero
    Set mChartState.CurrentChart = Nothing
    ' Luego notificar el evento
    RaiseEvent ChartDeactivated(cht)
End Sub


