VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMediatorDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "5-Aplicac (Coord)"
'@ModuleDescription "Gestor de eventos (mediador entre nivel de infraestructura y de dominio)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMediatorDomain"

Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1

Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1

' NUEVO: Escuchar eventos de subestados
Private WithEvents mOpportunityState As clsOpportunityState
Attribute mOpportunityState.VB_VarHelpID = -1

Private mRibbon As clsRibbon
Private mFileMgr As clsFileManager
Private mApplicationState As clsApplicationState

'--------------------------------------------------------------
' Inicializaci�n
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'TODO: faltaria crear clases para otras funciones que estan en macros, y pasarlas como parametros...
Public Sub Initialize(oFSMonitoringCoord As clsFSMonitoringCoord, _
                oOpportunities As clsOpportunitiesMgr, _
                oRibbon As clsRibbon, _
                oFileMgr As clsFileManager, _
                oApplicationState As clsApplicationState)
    On Error GoTo ErrHandler

    Set mRibbon = oRibbon

    Set mFileMgr = oFileMgr

    Set mFSMonitoringCoord = oFSMonitoringCoord

    Set mOpportunities = oOpportunities

    ' NUEVO: Guardar referencia a ApplicationState y escuchar eventos de subestados
    Set mApplicationState = oApplicationState
    Set mOpportunityState = oOpportunities.State

    LogInfo MODULE_NAME, "[Initialize] Mediador configurado con ApplicationState y subestados"

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler

    ' Liberar referencias WithEvents primero
    Set mOpportunityState = Nothing
    Set mOpportunities = Nothing
    Set mFSMonitoringCoord = Nothing

    ' Luego referencias normales
    Set mApplicationState = Nothing
    Set mFileMgr = Nothing
    Set mRibbon = Nothing

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicacion: " & Err.Description
End Sub
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Evento: cambios detectados en oportunidades (desde OpportunitiesMgr)
' -------------------------------------------------------------
Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr OpportunityCollectionUpdate] - cambios detectados: " & cambios
    
    If Not mRibbon Is Nothing Then
        'mRibbon.InvalidarRibbon
        mRibbon.InvalidarControl "ddlOportunidades"
        LogDebug MODULE_NAME, "[mOpportunities_OpportunityCollectionUpdate] - Ribbon invalidado"
    End If
End Sub
'TODO: Extraer la l�gica de clsFileXLS A una funci�n (En principio se mantendr� en Celesa aplicaci�n pero es posible que se lleve a otra clase)
Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr currOpportunityChanged] - Oportunidad seleccionada: " & Path
    
    Dim f As clsFileXLS
    Set f = mFileMgr.ActiveExcelFile
    If Not f Is Nothing Then
        If f.DebeMoverseAOp(Path, mOpportunities.BaseFolder) Then
            f.MoverAOp Path, mOpportunities.BaseFolder
        End If
    End If
    ' (Aqu� no tiene sentido invalidar el ribbon, esto no tiene nada que ver con ribbon, SOLO L�GICA DE NEGOCIO)
End Sub

'--------------------------------------------------------------
' NUEVO: Evento de cambio de oportunidad desde OpportunityState
'--------------------------------------------------------------
Private Sub mOpportunityState_CurrentOpportunityChanged(op As clsOpportunity, index As Long)
    LogInfo MODULE_NAME, "[callback: OpportunityState CurrentOpportunityChanged] - Index=" & index

    ' Llamar a ApplicationState para coordinar estados
    If Not mApplicationState Is Nothing Then
        mApplicationState.OnOpportunityChanged op, index
    End If

    ' NOTA: La lógica de mover archivo se mantiene en handler viejo por compatibilidad
End Sub

' -------------------------------------------------------------
' Eventos de mFSMonitoringCoord (cambios en sistema de archivos)
' clsApplication recibe estos eventos y delega a OpportunitiesMgr
' -------------------------------------------------------------

Private Sub mFSMonitoringCoord_OpportunityCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_OpportunityCreated] " & subfolderName
    mOpportunities.ProcesarCambiosEnOportunidades subfolderName
End Sub

Private Sub mFSMonitoringCoord_OpportunityDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_OpportunityDeleted] " & subfolderName
    mOpportunities.ProcesarCambiosEnOportunidades subfolderName
End Sub

Private Sub mFSMonitoringCoord_OpportunityRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_OpportunityRenamed] " & oldName & " -> " & newName
    mOpportunities.ProcesarCambiosEnOportunidades oldName
    mOpportunities.ProcesarCambiosEnOportunidades newName
End Sub

Private Sub mFSMonitoringCoord_OpportunityItemDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_OpportunityItemDeleted] " & fileName
    mOpportunities.ProcesarCambiosEnItemsOportunidad fileName
End Sub

Private Sub mFSMonitoringCoord_OpportunityItemRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_OpportunityItemRenamed] " & oldName & " -> " & newName
    mOpportunities.ProcesarCambiosEnItemsOportunidad newName
End Sub

Private Sub mFSMonitoringCoord_TemplateCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_TemplateCreated] " & fileName
    ' TODO: Implementar logica de actualizacion de plantillas si es necesario
End Sub

Private Sub mFSMonitoringCoord_TemplateChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_TemplateChanged] " & fileName
    ' TODO: Notificar cambio en plantilla si es necesario
End Sub

Private Sub mFSMonitoringCoord_GasFileCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_GasFileCreated] " & fileName
    ' TODO: Implementar logica de sincronizacion Gas si es necesario
End Sub

Private Sub mFSMonitoringCoord_GasFileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_GasFileChanged] " & fileName
    ' TODO: Actualizar referencias Gas si es necesario
End Sub

Private Sub mFSMonitoringCoord_MonitoringError(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[mFSMonitoringCoord_MonitoringError] " & folder & ": " & errorMessage

    ' Mostrar al usuario solo errores criticos
    If InStr(1, errorMessage, "Limite", vbTextCompare) > 0 Or _
       InStr(1, errorMessage, "no encontrada", vbTextCompare) > 0 Then
        MsgBox "Error en monitoreo de carpeta:" & vbCrLf & _
               folder & vbCrLf & vbCrLf & _
               errorMessage, vbExclamation, "FolderWatcher"
    End If
End Sub

Private Sub mFSMonitoringCoord_MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
    LogInfo MODULE_NAME, "[mFSMonitoringCoord_MonitoringReconnected] " & folder & " tras " & attempts & " intentos"
End Sub

Private Sub mFSMonitoringCoord_MonitoringFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[mFSMonitoringCoord_MonitoringFailed] " & folder & ": " & reason

    MsgBox "El monitoreo de la carpeta ha fallado permanentemente:" & vbCrLf & _
           folder & vbCrLf & vbCrLf & _
           "Razon: " & reason & vbCrLf & vbCrLf & _
           "Es posible que la carpeta no este accesible o haya problemas de red.", _
           vbExclamation, "FolderWatcher - Reconexion Fallida"
End Sub

