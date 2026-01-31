VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Gestor centralizado del estado actual de la aplicación
'@Category: State
'@Folder "5-Aplicac (Coord).Estado"
Option Explicit

Private Const MODULE_NAME As String = "clsApplicationState"

' ==========================================
' SUBCLASES DE ESTADO (Composición)
' ==========================================
Private mRibbonState As clsRibbonState
Private mOpportunityState As clsOpportunityState
Private mChartState As clsChartState
Private mFileState As clsFileState

' Referencia al Ribbon para invalidación
Private mRibbon As clsRibbon

' ==========================================
' EVENTOS
' ==========================================

'@Description: Se dispara cuando cambia la oportunidad actual
Public Event CurrentOpportunityChanged(op As clsOpportunity)

'@Description: Se dispara cuando cambia el archivo Excel activo
Public Event CurrentFileChanged(f As clsFileXLS)

'@Description: Se dispara cuando cambia el Chart activo
Public Event CurrentChartChanged(Ch As Chart)

'@Description: Se dispara cuando cambia el modo del Ribbon
Public Event RibbonModeChanged(mode As eRibbonMode)

'@Description: Se dispara cuando cambia el contexto general (invalida todo el Ribbon)
Public Event ApplicationContextChanged()

' ==========================================
' INICIALIZACIÓN
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    ' NO crear subestados aquí - se inyectan mediante Initialize()
End Sub

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate]"

    Set mRibbonState = Nothing
    Set mOpportunityState = Nothing
    Set mChartState = Nothing
    Set mFileState = Nothing
    Set mRibbon = Nothing
End Sub

'@Description: Inicializa el coordinador con referencias a subestados (Dependency Injection)
'@Note: Todos los parámetros son opcionales (Nothing es un estado válido)
'@Param ribbonState: Instancia de clsRibbonState (creada y poseída por clsRibbon)
'@Param opportunityState: Instancia de clsOpportunityState (creada y poseída por clsOpportunitiesMgr)
'@Param fileState: Instancia de clsFileState (creada y poseída por clsFileManager)
'@Param chartState: Instancia de clsChartState (creada y poseída por clsChartMgr)
Public Sub Initialize(Optional RibbonState As clsRibbonState = Nothing, _
                      Optional OpportunityState As clsOpportunityState = Nothing, _
                      Optional FileState As clsFileState = Nothing, _
                      Optional ChartState As clsChartState = Nothing)
    On Error GoTo ErrHandler

    ' Inyectar referencias a subestados
    Set mRibbonState = RibbonState
    Set mOpportunityState = OpportunityState
    Set mFileState = FileState
    Set mChartState = ChartState

    LogInfo MODULE_NAME, "[Initialize] Referencias inyectadas: " & _
            "Ribbon=" & CStr(Not (RibbonState Is Nothing)) & ", " & _
            "Opportunity=" & CStr(Not (OpportunityState Is Nothing)) & ", " & _
            "File=" & CStr(Not (FileState Is Nothing)) & ", " & _
            "Chart=" & CStr(Not (ChartState Is Nothing))

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
End Sub

' ==========================================
' CONFIGURACIÓN
' ==========================================

'@Description: Inyecta la referencia al Ribbon para invalidación de controles
'@Param ribbon: Instancia de clsRibbon
Friend Sub SetRibbon(ribbon As clsRibbon)
    Set mRibbon = ribbon
End Sub

' ==========================================
' ACCESO A SUBCLASES (para inyección en componentes)
' ==========================================

'@Description: Estado del Ribbon (para inyectar en clsRibbon)
'Public Property Get RibbonState() As clsRibbonState
'    Set RibbonState = mRibbonState
'End Property
'
''@Description: Estado de oportunidades (para inyectar en clsOpportunitiesMgr u otros)
'Public Property Get OpportunityState() As clsOpportunityState
'    Set OpportunityState = mOpportunityState
'End Property
'
''@Description: Estado de charts (para inyectar en clsChartMgr u otros)
'Public Property Get ChartState() As clsChartState
'    Set ChartState = mChartState
'End Property
'
''@Description: Estado de archivo activo (para inyectar en clsFileManager u otros)
'Public Property Get FileState() As clsFileState
'    Set FileState = mFileState
'End Property

' ==========================================
' PROPIEDADES FACADE - Acceso conveniente al estado
' ==========================================

'@Description: Oportunidad actual (delegado a OpportunityState)
Public Property Get CurrentOpportunity() As clsOpportunity
    Set CurrentOpportunity = mOpportunityState.CurrentOpportunity
End Property

Public Property Set CurrentOpportunity(op As clsOpportunity)
    'TODO: Pendiente de depurar, no tengo la certeza de que la secuencia de activación de estado esté correcta
    Stop ' De momento no se usa; tiene utilidad para sistemas externos a la aplicación
    If Not op.Equals(mOpportunityState.CurrentOpportunity) Then
        Set mOpportunityState.CurrentOpportunity = op

        LogInfo MODULE_NAME, "[CurrentOpportunity] Cambiada a: " & mOpportunityState.GetCurrentOpportunityLabel()

        ' Sincronizar otros estados si es necesario
        SyncRibbonAfterOpportunityChange

        ' Invalidar controles específicos relacionados con oportunidades
        InvalidateOpportunityControls

        ' Disparar eventos
        RaiseEvent CurrentOpportunityChanged(op)
    End If
End Property

'@Description: Archivo Excel activo actual (delegado a FileState)
Public Property Get CurrentFile() As clsFileXLS
    Set CurrentFile = mFileState.CurrentFile
End Property

Public Property Set CurrentFile(f As clsFileXLS)
    'TODO: Pendiente de depurar, no tengo la certeza de que la secuencia de activación de estado esté correcta
    Stop ' De momento no se usa; tiene utilidad para sistemas externos a la aplicación
    If Not f.Equals(mFileState.CurrentFile) Then
        Set mFileState.CurrentFile = f

        LogInfo MODULE_NAME, "[CurrentFile] Cambiado a: " & mFileState.GetCurrentFileName()

        ' Sincronizar oportunidad desde archivo si aplica
        SyncOpportunityFromFile f

        ' Invalidar controles específicos relacionados con archivos
        InvalidateFileControls

        ' Disparar eventos
        RaiseEvent CurrentFileChanged(f)
    End If
End Property

'@Description: Chart activo actual (delegado a ChartState)
Public Property Get CurrentChart() As Chart
    Set CurrentChart = mChartState.CurrentChart
End Property

Public Property Set CurrentChart(Ch As Chart)
    'TODO: Pendiente de depurar, no tengo la certeza de que la secuencia de activación de estado esté correcta
    Stop ' De momento no se usa; tiene utilidad para sistemas externos a la aplicación
    If Not AreSameChart(mChartState.CurrentChart, Ch) Then
        Set mChartState.CurrentChart = Ch

        LogInfo MODULE_NAME, "[CurrentChart] Cambiado a: " & mChartState.GetCurrentChartName()

        ' Invalidar controles específicos relacionados con charts
        InvalidateChartControls

        ' Disparar eventos
        RaiseEvent CurrentChartChanged(Ch)
    End If
End Property

'@Description: Modo del Ribbon (delegado a RibbonState)
Public Property Get RibbonMode() As eRibbonMode
    RibbonMode = mRibbonState.Modo
End Property

Public Property Let RibbonMode(mode As eRibbonMode)
    'TODO: Pendiente de depurar, no tengo la certeza de que la secuencia de activación de estado esté correcta
    Stop ' De momento no se usa; tiene utilidad para sistemas externos a la aplicación
    If mRibbonState.Modo <> mode Then
        mRibbonState.Modo = mode

        LogInfo MODULE_NAME, "[RibbonMode] Cambiado a: " & mode

        ' Cuando cambia el modo del Ribbon, invalidar todo
        InvalidateRibbonFull

        ' Disparar eventos
        RaiseEvent RibbonModeChanged(mode)
        RaiseEvent ApplicationContextChanged
    End If
End Property

'@Description: Visibilidad del Ribbon (delegado a RibbonState)
Public Property Get RibbonVisible() As Boolean
    RibbonVisible = mRibbonState.Visible
End Property

Public Property Let RibbonVisible(value As Boolean)
    'TODO: Pendiente de depurar, no tengo la certeza de que la secuencia de activación de estado esté correcta
    Stop ' De momento no se usa; tiene utilidad para sistemas externos a la aplicación
    If mRibbonState.Visible <> value Then
        mRibbonState.Visible = value
        LogInfo MODULE_NAME, "[RibbonVisible] Cambiado a: " & value
    End If
End Property

' ==========================================
' CONSULTAS DE CONTEXTO
' ==========================================

'@Description: Indica si hay un archivo de oportunidad activo
'@Returns: True si hay archivo activo Y es de oportunidad Y hay oportunidad seleccionada
Public Function IsOpportunityFileActive() As Boolean
    IsOpportunityFileActive = mFileState.isOpportunityFile And mOpportunityState.HasCurrentOpportunity()
End Function

'@Description: Indica si hay un chart activo
Public Function IsChartActive() As Boolean
    Stop ' De momento no se usa; tiene utilidad ¿para gestores de eventos?/¿para sistemas externos a la aplicación?
    IsChartActive = mChartState.HasCurrentChart()
End Function

'@Description: Indica si hay un archivo activo
Public Function IsFileActive() As Boolean
    Stop ' De momento no se usa; tiene utilidad ¿para gestores de eventos?/¿para sistemas externos a la aplicación?
    IsFileActive = mFileState.HasCurrentFile()
End Function

'@Description: Obtiene diagnóstico completo del contexto actual
Public Function GetContextDiagnostics() As String
    Dim info As String
    info = "=== CONTEXTO DE APLICACIÓN ===" & vbCrLf
    info = info & "Oportunidad: " & mOpportunityState.GetCurrentOpportunityLabel() & vbCrLf
    info = info & "Archivo: " & mFileState.GetCurrentFileName() & vbCrLf
    info = info & "  - Es oportunidad: " & CStr(mFileState.isOpportunityFile) & vbCrLf
    info = info & "Chart: " & mChartState.GetCurrentChartName() & vbCrLf
    info = info & "Ribbon Mode: " & mRibbonState.Description() & vbCrLf
    info = info & "Ribbon Visible: " & CStr(mRibbonState.Visible) & vbCrLf
    GetContextDiagnostics = info
End Function

' ==========================================
' INVALIDACIÓN DE CONTROLES DEL RIBBON
' ==========================================

'@Description: Invalida controles relacionados con oportunidades
Private Sub InvalidateOpportunityControls()
    If mRibbon Is Nothing Then Exit Sub

    ' Controles que dependen de la oportunidad actual
    mRibbon.InvalidarControl "btnGenerarGraficos"
    mRibbon.InvalidarControl "dropdownOportunidades"
    mRibbon.InvalidarControl "btnCreaOportunidad"
    mRibbon.InvalidarControl "grpOportunidades"

    LogDebug MODULE_NAME, "[InvalidateOpportunityControls] Controles de oportunidad invalidados"
End Sub

'@Description: Invalida controles relacionados con archivos
Private Sub InvalidateFileControls()
    If mRibbon Is Nothing Then Exit Sub

    ' Controles que dependen del archivo activo
    mRibbon.InvalidarControl "btnGenerarGraficos"
    mRibbon.InvalidarControl "btnExportarPDF"
    mRibbon.InvalidarControl "btnLimpiarFormato"
    mRibbon.InvalidarControl "btnAjustarFilas"

    LogDebug MODULE_NAME, "[InvalidateFileControls] Controles de archivo invalidados"
End Sub

'@Description: Invalida controles relacionados con charts
Private Sub InvalidateChartControls()
    If mRibbon Is Nothing Then Exit Sub

    ' Controles que dependen del chart activo
    mRibbon.InvalidarControl "btnInvertirSeries"
    mRibbon.InvalidarControl "btnFormatearGrafico"

    LogDebug MODULE_NAME, "[InvalidateChartControls] Controles de chart invalidados"
End Sub

'@Description: Invalida todo el Ribbon
Private Sub InvalidateRibbonFull()
    If mRibbon Is Nothing Then Exit Sub

    mRibbon.InvalidarRibbon

    LogDebug MODULE_NAME, "[InvalidateRibbonFull] Ribbon completo invalidado"
End Sub

'@Description: Invalida todo el contexto de la aplicación (método público para forzar invalidación completa)
Public Sub InvalidateAll()
    InvalidateRibbonFull
    RaiseEvent ApplicationContextChanged
End Sub

' ==========================================
' SINCRONIZACIÓN ENTRE SUBCLASES (Coordinación privada)
' ==========================================

'@Description: Sincroniza estado del Ribbon cuando cambia la oportunidad
Private Sub SyncRibbonAfterOpportunityChange()
    ' Si el modo es OpportunityOnly, actualizar visibilidad según si hay oportunidad
    If mRibbonState.Modo = Ribbon_OpportunityOnly Then
        If mOpportunityState.HasCurrentOpportunity() Then
            mRibbonState.Visible = True
        Else
            mRibbonState.Visible = False
        End If
    End If
End Sub

'@Description: Sincroniza oportunidad actual desde el archivo activo (si corresponde)
Private Sub SyncOpportunityFromFile(f As clsFileXLS)
    ' Si el archivo es de oportunidad, aquí se podría actualizar CurrentOpportunity
    ' basándose en el path del archivo
    ' Esta lógica requiere acceso a OpportunitiesMgr, por lo que se delega
    ' al mediador de eventos (clsEventsMediatorDomain)

    ' Por ahora, solo loguear
    If Not f Is Nothing Then
        If mFileState.isOpportunityFile Then
            LogDebug MODULE_NAME, "[SyncOpportunityFromFile] Archivo de oportunidad detectado: " & f.Name
        End If
    End If
End Sub

' ==========================================
' FUNCIONES AUXILIARES DE COMPARACIÓN
' ==========================================

'@Description: Compara dos charts para ver si son el mismo
Private Function AreSameChart(ch1 As Chart, ch2 As Chart) As Boolean
    On Error Resume Next

    If ch1 Is Nothing And ch2 Is Nothing Then
        AreSameChart = True
    ElseIf ch1 Is Nothing Or ch2 Is Nothing Then
        AreSameChart = False
    Else
        ' Comparar por puntero de objeto
        AreSameChart = (ObjPtr(ch1) = ObjPtr(ch2))
    End If

    On Error GoTo 0
End Function

' ==========================================
' COORDINACIÓN DE EVENTOS (llamados por Mediadores)
' ==========================================

'@Description: Coordinación cuando cambia la oportunidad actual (llamado desde EventsMediatorDomain)
'@Param op: Nueva oportunidad seleccionada
'@Param index: Índice en la colección
Friend Sub OnOpportunityChanged(op As clsOpportunity, Index As Long)
    On Error Resume Next

    LogDebug MODULE_NAME, "[OnOpportunityChanged] op=" & IIf(op Is Nothing, "(Nothing)", op.Label) & _
                         ", index=" & Index

    ' Actualizar índice en OpportunityState si hace falta
    If mOpportunityState.CurrentIndex <> Index Then
        mOpportunityState.CurrentIndex = Index
    End If

    ' Sincronizar ribbon si hace falta
    SyncRibbonAfterOpportunityChange

    ' Invalidar controles relacionados con oportunidades
    InvalidateOpportunityControls

    On Error GoTo 0
End Sub

'@Description: Coordinación cuando cambia el archivo activo (llamado desde EventsMediatorDomain)
'@Param f: Nuevo archivo activo
Friend Sub OnFileChanged(f As clsFileXLS)
    On Error Resume Next

    Dim fileName As String
    If f Is Nothing Then
        fileName = "(ninguno)"
    Else
        fileName = f.Name
        If Err.Number <> 0 Then fileName = "(error)"
    End If

    LogDebug MODULE_NAME, "[OnFileChanged] file=" & fileName

    ' Sincronizar oportunidad desde archivo si aplica
    SyncOpportunityFromFile f

    ' Invalidar controles relacionados con archivos
    InvalidateFileControls

    On Error GoTo 0
End Sub

'@Description: Coordinación cuando cambia la colección de oportunidades (llamado desde EventsMediatorDomain)
'@Param changeType: Tipo de cambio (agregado, eliminado, renombrado)
Friend Sub OnOpportunityCollectionChanged(changeType As String)
    On Error Resume Next

    LogDebug MODULE_NAME, "[OnOpportunityCollectionChanged] changeType=" & changeType

    ' Cuando cambia la colección, invalidar el dropdown de oportunidades
    If Not mRibbon Is Nothing Then
        mRibbon.InvalidarControl "ddlOportunidades"
        mRibbon.InvalidarControl "dropdownOportunidades"
    End If

    On Error GoTo 0
End Sub


