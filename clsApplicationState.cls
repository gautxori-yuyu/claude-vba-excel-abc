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

'@Description: Archivo Excel activo actual (delegado a FileState)
Public Property Get CurrentFile() As clsFileXLS
    Set CurrentFile = mFileState.CurrentFile
End Property

'@Description: Chart activo actual (delegado a ChartState)
Public Property Get CurrentChart() As Chart
    Set CurrentChart = mChartState.CurrentChart
End Property

'@Description: Modo del Ribbon (delegado a RibbonState)
Public Property Get RibbonMode() As eRibbonMode
    RibbonMode = mRibbonState.Modo
End Property

'@Description: Visibilidad del Ribbon (delegado a RibbonState)
Public Property Get RibbonVisible() As Boolean
    RibbonVisible = mRibbonState.Visible
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
' SINCRONIZACIÓN ENTRE SUBCLASES (Coordinación privada)
' ==========================================

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

    On Error GoTo 0
End Sub

'@Description: Coordinación cuando cambia la colección de oportunidades (llamado desde EventsMediatorDomain)
'@Param changeType: Tipo de cambio (agregado, eliminado, renombrado)
Friend Sub OnOpportunityCollectionChanged(changeType As String)
    On Error Resume Next

    LogDebug MODULE_NAME, "[OnOpportunityCollectionChanged] changeType=" & changeType


    On Error GoTo 0
End Sub


