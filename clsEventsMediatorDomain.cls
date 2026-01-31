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

Private WithEvents mFileMgr As clsFileManager
Attribute mFileMgr.VB_VarHelpID = -1

Private mCtx As clsExecutionContext
Private mCtxMgr As clsExecutionContextMgr
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
Public Sub Initialize(oCtx As clsExecutionContext, _
                oCtxMgr As clsExecutionContextMgr, _
                oFSMonitoringCoord As clsFSMonitoringCoord, _
                oOpportunities As clsOpportunitiesMgr, _
                oFileMgr As clsFileManager, _
                oApplicationState As clsApplicationState)
    On Error GoTo ErrHandler

    Set mCtx = oCtx
    Set mCtxMgr = oCtxMgr

    Set mFileMgr = oFileMgr

    Set mFSMonitoringCoord = oFSMonitoringCoord

    Set mOpportunities = oOpportunities

    ' Guardar referencia a ApplicationState
    Set mApplicationState = oApplicationState

    LogInfo MODULE_NAME, "[Initialize] Mediador configurado con ApplicationState y subestados"

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler

    ' Liberar referencias WithEvents primero
    Set mOpportunities = Nothing
    
    Set mFSMonitoringCoord = Nothing

    ' Luego referencias normales
    Set mApplicationState = Nothing
    Set mFileMgr = Nothing
    Set mCtx = Nothing
    
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicaci�n: " & Err.Description
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Evento: cambios detectados en oportunidades (desde OpportunitiesMgr)
' -------------------------------------------------------------
Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr OpportunityCollectionUpdate] - cambios detectados: " & cambios

    ' Llamar a ApplicationState para coordinar estados
    If Not mApplicationState Is Nothing Then
        mApplicationState.OnOpportunityCollectionChanged cambios
    End If
End Sub
'TODO: Extraer la l�gica de clsFileXLS A una funci�n
Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr currOpportunityChanged] - Oportunidad seleccionada: " & Path
    
    Dim f As clsFileXLS
    Set f = mFileMgr.GetOrTrackWorkbook(mCtx.Workbook)
    If Not f Is Nothing Then
        If f.DebeMoverseAOp(Path, mOpportunities.BaseFolder) Then
            f.MoverAOp Path, mOpportunities.BaseFolder
        End If
    End If
    ' Coordinar estado en ApplicationState
    If Not mApplicationState Is Nothing Then
        Dim op As clsOpportunity
        Set op = mOpportunities.State.CurrentOpportunity
        mApplicationState.OnOpportunityChanged op, Index
    End If
End Sub


'--------------------------------------------------------------
' Evento de cambio de archivo activo (desde FileManager)
'--------------------------------------------------------------
Private Sub mFileMgr_ActiveFileChanged(f As clsFileXLS)
    Dim fileName As String
    If f Is Nothing Then
        fileName = "(ninguno)"
    Else
        On Error Resume Next
        fileName = f.Name
        If Err.Number <> 0 Then fileName = "(error)"
        On Error GoTo 0
    End If

    LogInfo MODULE_NAME, "[callback: FileManager ActiveFileChanged] - File=" & fileName & _
                         ", IsOpportunity=" & mFileMgr.State.isOpportunityFile

    ' Llamar a ApplicationState para coordinar estados
    If Not mApplicationState Is Nothing Then
        mApplicationState.OnFileChanged f
    End If
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


