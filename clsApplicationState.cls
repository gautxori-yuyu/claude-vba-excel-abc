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
Private mCtxState As clsExecutionContextState


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
    Set mCtxState = Nothing
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
                      Optional ChartState As clsChartState = Nothing, _
                      Optional CtxState As clsExecutionContextState = Nothing)
    On Error GoTo ErrHandler

    ' Inyectar referencias a subestados
    Set mRibbonState = RibbonState
    Set mOpportunityState = OpportunityState
    Set mFileState = FileState
    Set mChartState = ChartState
    Set mCtxState = CtxState

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
' CONSULTAS DE CONTEXTO
' ==========================================

'@Description: Indica si hay un archivo de oportunidad activo
'@Returns: True si hay archivo activo Y es de oportunidad Y hay oportunidad seleccionada
Public Function IsOpportunityFileActive() As Boolean
    IsOpportunityFileActive = mFileState.isOpportunityFile And mOpportunityState.HasCurrentOpportunity()
End Function

'@Description: Indica si el contexto de ejecucion Excel es valido
Public Property Get IsContextValid() As Boolean
    If mCtxState Is Nothing Then
        IsContextValid = False
    Else
        IsContextValid = mCtxState.IsContextValid
    End If
End Property

'@Description: Indica si la aplicacion esta completamente inicializada
'@Returns: True si todos los subestados criticos han sido inyectados
'@Note: Util para detectar estados parciales tras un reset de VBA
Public Property Get IsInitialized() As Boolean
    ' Verificar que todos los subestados criticos estan presentes
    IsInitialized = Not (mCtxState Is Nothing) And _
                    Not (mRibbonState Is Nothing) And _
                    Not (mOpportunityState Is Nothing) And _
                    Not (mFileState Is Nothing)
End Property

'@Description: Obtiene informacion de diagnostico sobre el estado de inicializacion
'@Returns: String con el estado de cada subestado
Public Function GetInitializationStatus() As String
    Dim status As String
    status = "Subestados: "
    status = status & "Ctx=" & CStr(Not (mCtxState Is Nothing)) & ", "
    status = status & "Ribbon=" & CStr(Not (mRibbonState Is Nothing)) & ", "
    status = status & "Opp=" & CStr(Not (mOpportunityState Is Nothing)) & ", "
    status = status & "File=" & CStr(Not (mFileState Is Nothing)) & ", "
    status = status & "Chart=" & CStr(Not (mChartState Is Nothing))
    GetInitializationStatus = status
End Function

