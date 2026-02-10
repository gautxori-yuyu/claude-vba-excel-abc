VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMediatorDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3-Dominio.2-2-Mediadores"
'@ModuleDescription "Gestor de eventos (mediador entre nivel de infraestructura y de dominio)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMediatorDomain"

' ==================================================================
' SUSCRIPCIONES A EVENTOS
' ==================================================================
' El mediador de dominio escucha:
' - mInfraMediador: eventos SEMANTICOS traducidos de infraestructura
' - mFSMonitoringCoord: eventos de sistema de archivos (ya son semanticos)
' - mOpportunities: eventos de dominio (Manager de oportunidades)
'
' NO escucha directamente a clases de infraestructura pura (mCtxMgr, mFileMgr)

Private WithEvents mInfraMediador As clsEventsMediatorInfrastructure
Attribute mInfraMediador.VB_VarHelpID = -1

Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1

Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1

Private mCtxState As clsExecutionContextState
Private mFileMgr As clsFileManager   ' Referencia sin WithEvents (solo para consultas)

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

'TODO: faltaria crear clases para otras funciones que estan en macros, y pasarlas como parametros...
Public Sub Initialize(oCtxState As clsExecutionContextState, _
                oInfraMediador As clsEventsMediatorInfrastructure, _
                oFSMonitoringCoord As clsFSMonitoringCoord, _
                oOpportunities As clsOpportunitiesMgr, _
                oFileMgr As clsFileManager)
    On Error GoTo ErrHandler

    Set mCtxState = oCtxState

    ' Escuchar eventos SEMANTICOS del mediador de infraestructura
    Set mInfraMediador = oInfraMediador

    ' Referencia a FileMgr solo para consultas (no WithEvents)
    Set mFileMgr = oFileMgr

    Set mFSMonitoringCoord = oFSMonitoringCoord

    Set mOpportunities = oOpportunities

    LogInfo MODULE_NAME, "[Initialize] Mediador configurado (escucha via InfraMediador)"

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
    Set mInfraMediador = Nothing

    ' Luego referencias normales
    Set mFileMgr = Nothing
    Set mCtxState = Nothing

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicacion: " & Err.Description
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Evento: cambios detectados en oportunidades (desde OpportunitiesMgr)
' -------------------------------------------------------------
Private Sub mOpportunities_OpportunityCollectionUpdated(ByVal reason As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr OpportunityCollectionUpdated] razon: " & reason
End Sub

'TODO: VIOLACION ARQUITECTONICA - El mediador de dominio no deberia usar clsFileXLS directamente.
'      Esta logica de "mover archivo a carpeta de oportunidad" deberia:
'      1. Delegarse a un servicio de infraestructura (FileLocationService?), o
'      2. El mediador de infraestructura deberia exponer un evento/metodo semantico
'      Pendiente de decision cuando se tenga la especificacion de dominio.
Private Sub mOpportunities_CurrentOpportunityChanged(ByVal opportunityName As String, ByVal opportunityPath As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr CurrentOpportunityChanged] Oportunidad seleccionada: " & opportunityPath

    ' TEMPORAL: Uso directo de clsFileXLS (ver TODO arriba)
    Dim f As clsFileXLS
    Set f = mFileMgr.GetOrTrackWorkbook(mCtxState.Workbook)
    If Not f Is Nothing Then
        If f.DebeMoverseAOp(opportunityPath, mOpportunities.BaseFolder) Then
            f.MoverAOp opportunityPath, mOpportunities.BaseFolder
        End If
    End If
End Sub


' -------------------------------------------------------------
' Eventos SEMANTICOS del mediador de infraestructura
' (traducidos de eventos tecnicos - el dominio no conoce Excel/COM)
' -------------------------------------------------------------

'@Description: Evento semantico: el archivo activo ha cambiado
Private Sub mInfraMediador_ActiveFileSessionChanged(ByVal FilePath As String, ByVal isOpportunityFile As Boolean)
    Dim fileName As String
    If Len(FilePath) = 0 Then
        fileName = "(ninguno)"
    Else
        ' Extraer solo el nombre del archivo del path
        Dim pos As Long
        pos = InStrRev(FilePath, "\")
        If pos > 0 Then
            fileName = Mid$(FilePath, pos + 1)
        Else
            fileName = FilePath
        End If
    End If

    LogInfo MODULE_NAME, "[InfraMediador_ActiveFileSessionChanged] File=" & fileName & _
                         ", IsOpportunity=" & isOpportunityFile
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
End Sub

Private Sub mFSMonitoringCoord_MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
    LogDebug MODULE_NAME, "[mFSMonitoringCoord_MonitoringReconnected] " & folder & " tras " & attempts & " intentos"
End Sub

Private Sub mFSMonitoringCoord_MonitoringFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[mFSMonitoringCoord_MonitoringFailed] " & folder & ": " & reason
End Sub

'@Description: Evento semantico: el contexto de infraestructura se ha invalidado
'@Note: Permite al dominio prepararse para la invalidacion sin conocer detalles tecnicos
Private Sub mInfraMediador_InfraContextInvalidated()
    LogWarning MODULE_NAME, "[InfraMediador_InfraContextInvalidated] Contexto invalidado, notificando a dominio"

    ' Pausar monitoreo de FS si esta activo (puede tener referencias invalidas)
    On Error Resume Next
    If Not mFSMonitoringCoord Is Nothing Then
        mFSMonitoringCoord.PausarMonitoreo
    End If
    On Error GoTo 0
End Sub

'@Description: Evento semantico: el contexto de infraestructura se ha reinicializado
'@Note: Permite al dominio reenganchar sus dependencias y refrescar estado
Private Sub mInfraMediador_InfraContextReinitialized()
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[InfraMediador_InfraContextReinitialized] Reconectando dominio"

    ' Reanudar monitoreo de FS
    If Not mFSMonitoringCoord Is Nothing Then
        mFSMonitoringCoord.ReanudarMonitoreo
    End If

    ' Refrescar coleccion de oportunidades por si hubo cambios durante el reset
    If Not mOpportunities Is Nothing Then
        mOpportunities.RefreshOpportunities
    End If

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[InfraMediador_InfraContextReinitialized] Error durante reconexion de dominio", , Err.Description
End Sub
