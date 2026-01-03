VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSMonitoringCoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "3-Aplicac (Coord)"
Option Explicit

Private Const MODULE_NAME As String = "clsFSMonitoringCoord"

Private WithEvents mFolderWatcher As clsFSWatcher
Attribute mFolderWatcher.VB_VarHelpID = -1

' Rutas almacenadas localmente (recibidas en IniciarMonitoreo)
Private m_rutaOportunidades As String
Private m_rutaPlantillas As String
Private m_rutaGasVBNet As String

' -------------------------------------------------------------
' EVENTOS PUBLICOS - clsAplicacion los escucha y delega
' -------------------------------------------------------------

' Eventos de oportunidades (subcarpetas)
Public Event OpportunityCreated(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event OpportunityDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event OpportunityRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)

' Eventos de items dentro de oportunidades (archivos)
Public Event OpportunityItemDeleted(ByVal folder As String, ByVal fileName As String)
Public Event OpportunityItemRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)

' Eventos de plantillas
Public Event TemplateCreated(ByVal folder As String, ByVal fileName As String)
Public Event TemplateChanged(ByVal folder As String, ByVal fileName As String)

' Eventos de Gas VBNet
Public Event GasFileCreated(ByVal folder As String, ByVal fileName As String)
Public Event GasFileChanged(ByVal folder As String, ByVal fileName As String)

' Eventos de error/reconexion
Public Event MonitoringError(ByVal folder As String, ByVal errorMessage As String)
Public Event MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
Public Event MonitoringFailed(ByVal folder As String, ByVal reason As String)

' -------------------------------------------------------------
' Property: Acceso al FolderWatcher (para debugging)
' -------------------------------------------------------------
Public Property Get FolderWatcher() As clsFSWatcher
    Set FolderWatcher = mFolderWatcher
End Property

Private Sub Class_Initialize()
    Set mFolderWatcher = New clsFSWatcher
End Sub

Private Sub Class_Terminate()
    If Not mFolderWatcher Is Nothing Then
        LogInfo MODULE_NAME, "[Terminate] - Liberando FolderWatcher COM"
        mFolderWatcher.Dispose
        Set mFolderWatcher = Nothing
    End If
End Sub

' -------------------------------------------------------------
' IniciarMonitoreo: RECIBE EL DICCIONARIO DE CARPETAS COMO PARAMETRO
' @param oDicFolders: Dictionary con claves CFG_RUTA_* y valores = rutas
' -------------------------------------------------------------
Friend Sub IniciarMonitoreo(ByVal oDicFolders As Object)
    LogInfo MODULE_NAME, "[IniciarMonitoreo]"

    On Error GoTo ErrHandler

    If oDicFolders Is Nothing Or oDicFolders.Count = 0 Then
        LogWarning MODULE_NAME, "[IniciarMonitoreo] No hay carpetas configuradas para monitorear"
        Exit Sub
    End If

    ' Almacenar rutas localmente para usar en callbacks
    If oDicFolders.Exists(CFG_RUTA_OPORTUNIDADES) Then
        m_rutaOportunidades = oDicFolders(CFG_RUTA_OPORTUNIDADES)
    End If
    If oDicFolders.Exists(CFG_RUTA_PLANTILLAS) Then
        m_rutaPlantillas = oDicFolders(CFG_RUTA_PLANTILLAS)
    End If
    If oDicFolders.Exists(CFG_RUTA_GAS_VBNET) Then
        m_rutaGasVBNet = oDicFolders(CFG_RUTA_GAS_VBNET)
    End If

    ' Monitorear cada carpeta configurada
    Dim key As Variant
    For Each key In oDicFolders.Keys
        Dim rutaCarpeta As String
        rutaCarpeta = oDicFolders(key)

        If RutaExiste(rutaCarpeta) Then
            Select Case key
            Case CFG_RUTA_OPORTUNIDADES
                ' Monitoreo de subcarpetas para oportunidades
                mFolderWatcher.IniciarMonitoreo _
                    folderPath:=rutaCarpeta, _
                    includeSubdirs:=False, _
                    filterPattern:="*", _
                    eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
                    inactivityMinutes:=10, _
                    foldersOnly:=True
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Oportunidades monitoreadas: " & rutaCarpeta

            Case CFG_RUTA_PLANTILLAS
                ' Monitoreo de archivos para plantillas
                mFolderWatcher.IniciarMonitoreo _
                    folderPath:=rutaCarpeta, _
                    includeSubdirs:=False, _
                    filterPattern:="*.xls*", _
                    eventsToWatch:=Array("Created", "Changed"), _
                    inactivityMinutes:=10, _
                    foldersOnly:=False
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Plantillas monitoreadas: " & rutaCarpeta

            Case CFG_RUTA_GAS_VBNET
                ' Monitoreo de archivos para Gas VBNet
                mFolderWatcher.IniciarMonitoreo _
                    folderPath:=rutaCarpeta, _
                    includeSubdirs:=True, _
                    filterPattern:="*.*", _
                    eventsToWatch:=Array("Created", "Changed"), _
                    inactivityMinutes:=10, _
                    foldersOnly:=False
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Gas monitoreado: " & rutaCarpeta

            Case Else
                ' Monitoreo generico para otras carpetas
                mFolderWatcher.IniciarMonitoreo _
                    folderPath:=rutaCarpeta, _
                    includeSubdirs:=False, _
                    filterPattern:="*.*", _
                    eventsToWatch:=Array("Created", "Changed"), _
                    inactivityMinutes:=10, _
                    foldersOnly:=False
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Monitoreo generico: " & rutaCarpeta
            End Select
        Else
            LogWarning MODULE_NAME, "[IniciarMonitoreo] ADVERTENCIA: Ruta no existe: " & rutaCarpeta
        End If
    Next key

    LogInfo MODULE_NAME, "[IniciarMonitoreo] Monitoreo iniciado en " & oDicFolders.Count & " carpetas"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[IniciarMonitoreo] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de cambios en el sistema de archivos (desde FolderWatch)
' NOTA: Todos los callbacks disparan RaiseEvent, NO llaman a App()
' -------------------------------------------------------------

' Eventos de subcarpetas (oportunidades)
Private Sub mFolderWatcher_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[SubfolderCreated] " & subfolderName & " en " & parentFolder

    If EsRutaOportunidades(parentFolder) Then
        LogInfo MODULE_NAME, "[SubfolderCreated] Nueva oportunidad: " & subfolderName
        RaiseEvent OpportunityCreated(parentFolder, subfolderName)
    End If
End Sub

Private Sub mFolderWatcher_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[SubfolderDeleted] " & subfolderName & " en " & parentFolder

    If EsRutaOportunidades(parentFolder) Then
        LogInfo MODULE_NAME, "[SubfolderDeleted] Oportunidad eliminada: " & subfolderName
        RaiseEvent OpportunityDeleted(parentFolder, subfolderName)
    End If
End Sub

Private Sub mFolderWatcher_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[SubfolderRenamed] " & oldName & " -> " & newName & " en " & parentFolder

    If EsRutaOportunidades(parentFolder) Then
        LogInfo MODULE_NAME, "[SubfolderRenamed] Oportunidad renombrada: " & oldName & " -> " & newName
        RaiseEvent OpportunityRenamed(parentFolder, oldName, newName)
    End If
End Sub

' Eventos de archivos
Private Sub mFolderWatcher_FileCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileCreated] " & fileName & " en " & folder

    If EsRutaPlantillas(folder) Then
        LogInfo MODULE_NAME, "[FileCreated] Plantilla nueva: " & fileName
        RaiseEvent TemplateCreated(folder, fileName)
    ElseIf EsRutaGasVBNet(folder) Then
        LogInfo MODULE_NAME, "[FileCreated] Archivo Gas nuevo: " & fileName
        RaiseEvent GasFileCreated(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileDeleted] " & fileName & " en " & folder

    If EsRutaOportunidades(folder) Then
        LogInfo MODULE_NAME, "[FileDeleted] Item de oportunidad eliminado: " & fileName
        RaiseEvent OpportunityItemDeleted(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileChanged] " & fileName & " en " & folder

    If EsRutaPlantillas(folder) Then
        LogInfo MODULE_NAME, "[FileChanged] Plantilla modificada: " & fileName
        RaiseEvent TemplateChanged(folder, fileName)
    ElseIf EsRutaGasVBNet(folder) Then
        LogInfo MODULE_NAME, "[FileChanged] Archivo Gas modificado: " & fileName
        RaiseEvent GasFileChanged(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[FileRenamed] " & oldName & " -> " & newName & " en " & folder

    If EsRutaOportunidades(folder) Then
        LogInfo MODULE_NAME, "[FileRenamed] Item de oportunidad renombrado: " & oldName & " -> " & newName
        RaiseEvent OpportunityItemRenamed(folder, oldName, newName)
    End If
End Sub

' Eventos de estado del watcher
Private Sub mFolderWatcher_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat - solo para debug detallado si es necesario
    ' LogDebug MODULE_NAME, "[Heartbeat] " & folder & " - " & lastUpdate
End Sub

Private Sub mFolderWatcher_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[ErrorOccurred] ERROR en " & folder & ": " & errorMessage
    RaiseEvent MonitoringError(folder, errorMessage)
End Sub

Private Sub mFolderWatcher_WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
    LogDebug MODULE_NAME, "[WatcherReconnected] " & folder & " reconectado tras " & attempts & " intento(s)"
    If attempts > 1 Then
        LogWarning MODULE_NAME, "Monitoreo restaurado en " & folder & " tras " & attempts & " intentos"
    End If
    RaiseEvent MonitoringReconnected(folder, attempts)
End Sub

Private Sub mFolderWatcher_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[WatcherReconnectionFailed] " & folder & ": " & reason
    RaiseEvent MonitoringFailed(folder, reason)
End Sub

' -------------------------------------------------------------
' Funciones helper privadas para comparar rutas
' -------------------------------------------------------------

Private Function EsRutaOportunidades(ByVal folder As String) As Boolean
    If Len(m_rutaOportunidades) = 0 Then Exit Function
    EsRutaOportunidades = (InStr(1, folder, m_rutaOportunidades, vbTextCompare) > 0)
End Function

Private Function EsRutaPlantillas(ByVal folder As String) As Boolean
    If Len(m_rutaPlantillas) = 0 Then Exit Function
    EsRutaPlantillas = (InStr(1, folder, m_rutaPlantillas, vbTextCompare) > 0)
End Function

Private Function EsRutaGasVBNet(ByVal folder As String) As Boolean
    If Len(m_rutaGasVBNet) = 0 Then Exit Function
    EsRutaGasVBNet = (InStr(1, folder, m_rutaGasVBNet, vbTextCompare) > 0)
End Function
