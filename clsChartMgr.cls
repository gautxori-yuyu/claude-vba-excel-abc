VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MissingAnnotationArgument
'@Folder "2-Infraestructura.Excel.Charts"
'@ModuleDescription "Gestor centralizado de eventos de gráficos (orquestador)"
Option Explicit

Private Const MODULE_NAME As String = "clsChartMgr"

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

'--------------------------------------------------------------
' Variables privadas
'--------------------------------------------------------------
Private mChartsInActiveSheet As Collection              ' Colección de instancias de clsChartEvents activas
Private mChartState As clsChartState              ' Estado de Chart activo (composición)
Private mCtxState As clsExecutionContextState

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mChartsInActiveSheet = New Collection

    ' Crear estado de Chart (composición - Manager posee el estado)
    Set mChartState = New clsChartState

    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Call StopWatching

    ' Liberar estado
    Set mChartState = Nothing

    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'@Description: Expone el estado de Chart (para inyección en ApplicationState)
Public Property Get State() As clsChartState
    Set State = mChartState
End Property

Public Sub Initialize(oCtxState As clsExecutionContextState)
    Set mCtxState = oCtxState
    ' No cacheamos nada aquí -> lazy evaluation segura
    LogInfo MODULE_NAME, "[Initialize] Suscrito a eventos de Application"
End Sub

'--------------------------------------------------------------
' @Description: Comprueba si el gráfico activo es válido para invertir ejes
'--------------------------------------------------------------
' TODO: unificar esta funcion con lo de clsExecutioContext: alli la funcion Chart DEVUELVE CASI TODO LO QUE HACE ESTA FUNCION,
' queda cambiar el Select Case TypeName(mCtxState.Selection) usando la informacion cacheada alli (m_cachedChartInfo) (y sigo pensando que
' clsExecutionContext "ESTA DE SOBRA", lo unico que aporta es el "Cacheado"... y pendiente de incoporarle funcionalidad para,
' cuando se pierde el contexto, **reenganchar eventos, **reconstruir estado derivado, y **aislar el daño del reset)

Public Function EsValidoInvertirEjes() As Boolean
    On Error GoTo SafeExit
    
    ' 1. Validación segura de contexto mínimo
    If mCtxState.Workbook Is Nothing Then Exit Function
    If mCtxState.Selection Is Nothing Then Exit Function
    If mCtxState.Chart Is Nothing Then Exit Function
    
    ' 3. Caso de hoja de gráfico activa (ChartSheet)
    If TypeName(Application.ActiveSheet) = "Chart" Then
        EsValidoInvertirEjes = True
        GoTo CheckSeries
    End If
    
    ' 4. Caso de gráfico incrustado en hoja de cálculo (ChartObject)
    Select Case TypeName(mCtxState.Selection)
        Case "ChartObject", "ChartArea"
            ' No tengo nada claro que en este tipo de objetos se puedan invertir los ejes; en todo caso, NO son los que yo creo
            EsValidoInvertirEjes = True
        Case "DrawingObjects", "Picture", "Shape", "GroupObject", "OLEObject", "TextBox"
            ' Explícitamente NO es un gráfico
            EsValidoInvertirEjes = False
            Exit Function
        Case Else
            ' Otras selecciones no válidas
            EsValidoInvertirEjes = False
            Exit Function
    End Select
    ' HASTA AQUI SE REEMPLAZARIA CON EL Chart de clsExecutionContext; A PARTIR DE AQUI "EMPIEZA" EsValidoInvertirEjes
CheckSeries:
    ' 5. Comprobar al menos una serie en cada eje (tu lógica exacta)
    Dim tienePrimario As Boolean, tieneSecundario As Boolean
    Dim s As Series
    For Each s In ActiveChart.SeriesCollection
        If s.AxisGroup = xlPrimary Then
            tienePrimario = True
        ElseIf s.AxisGroup = xlSecondary Then
            tieneSecundario = True
        End If
    Next s
    EsValidoInvertirEjes = tienePrimario And tieneSecundario
    Exit Function

SafeExit:
    EsValidoInvertirEjes = False
    ' Opcional: loguear solo en modo depuración:
    LogCurrentError MODULE_NAME, "[EsValidoInvertirEjes]"
End Function

'--------------------------------------------------------------
'
'--------------------------------------------------------------
'@Description: Inicia la vigilancia de graficos en una hoja (normal o de grafico)
'@Category: Gráficos
'@Scope: argumentos
'@ArgumentDescriptions: sh As Object
Public Sub WatchSheet(Sh As Object)
    Dim oChartEvents As clsChartEvents
    
    On Error GoTo ErrHandler
    
    LogInfo MODULE_NAME, "[WatchSheet] Empezando a observar hoja: " & Sh.Name
    
    ' Detener observación anterior (libera eventos)
    ' Call StopWatching
    
    ' Caso 1: Hoja de gráfico (Chart Sheet)
    If TypeName(Sh) = "Chart" Then
        Set oChartEvents = New clsChartEvents
        Set oChartEvents.Chart = Sh
        Set oChartEvents.ParentManager = Me
        mChartsInActiveSheet.Add oChartEvents, CStr(ObjPtr(Sh))
        Exit Sub
    End If
    
    ' Caso 2: Hoja con gráficos embebidos
    If Sh.ChartObjects.Count > 0 Then
        Dim chtObj As ChartObject
        For Each chtObj In Sh.ChartObjects
            Set oChartEvents = New clsChartEvents
            Set oChartEvents.Chart = chtObj.Chart
            Set oChartEvents.ParentManager = Me
            mChartsInActiveSheet.Add oChartEvents, CStr(ObjPtr(chtObj.Chart))
        Next chtObj
    End If
    
    LogDebug MODULE_NAME, "[WatchSheet] gráficos en observación: " & mChartsInActiveSheet.Count
    
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[WatchSheet] Error", , Err.Description
    Err.Raise Err.Number, "clsChartMgr.WatchSheet", _
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
    For Each oChartEvents In mChartsInActiveSheet
        Set oChartEvents.Chart = Nothing
    Next oChartEvents
    Set mChartsInActiveSheet = New Collection
    LogDebug MODULE_NAME, "[StopWatching] Eventos de gráficos desactivados"

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[StopWatching]"
    Err.Raise Err.Number, MODULE_NAME & " [StopWatching] Error al detener la vigilancia de selección de gráficos en hoja", Err.Description
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


