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

Public Event OpportunityCreated(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event OpportunityDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event OpportunityRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)

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
        mFolderWatcher.Dispose  ' Usar el nuevo metodo Dispose
        Set mFolderWatcher = Nothing
    End If
End Sub

' -------------------------------------------------------------
' IniciarMonitoreo: CONFIGURA EL FOLDER WATCHER
' -------------------------------------------------------------
Friend Sub IniciarMonitoreo()
    LogInfo MODULE_NAME, "[IniciarMonitoreo]"

    On Error GoTo ErrHandler

    ' Obtener carpetas a monitorear desde configuración via App()
    Dim oDicFolders As Object
    Set oDicFolders = App().Configuration.oDicFoldersToWatch

    If oDicFolders Is Nothing Or oDicFolders.Count = 0 Then
        LogWarning MODULE_NAME, "[IniciarMonitoreo] No hay carpetas configuradas para monitorear"
        Exit Sub
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
                ' Monitoreo genérico para otras carpetas
                mFolderWatcher.IniciarMonitoreo _
                    folderPath:=rutaCarpeta, _
                    includeSubdirs:=False, _
                    filterPattern:="*.*", _
                    eventsToWatch:=Array("Created", "Changed"), _
                    inactivityMinutes:=10, _
                    foldersOnly:=False
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Monitoreo genérico: " & rutaCarpeta
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
' -------------------------------------------------------------

' Eventos de subcarpetas (reemplazan los eventos de archivos para oportunidades)
Private Sub mFolderWatcher_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch SubfolderCreated] " & subfolderName & " en " & parentFolder

    ' Determinar tipo de carpeta padre
    If InStr(1, parentFolder, App().Configuration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Es una nueva oportunidad - delegar a OpportunitiesMgr
        App().OpportunitiesMgr.ProcesarCambiosEnOportunidades subfolderName
        LogInfo MODULE_NAME, "[callback: SubfolderCreated] Nueva oportunidad: " & subfolderName
    End If
End Sub

Private Sub mFolderWatcher_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch SubfolderDeleted] " & subfolderName & " en " & parentFolder

    If InStr(1, parentFolder, App().Configuration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Oportunidad eliminada - delegar a OpportunitiesMgr
        App().OpportunitiesMgr.ProcesarCambiosEnOportunidades subfolderName
        LogInfo MODULE_NAME, "[callback: SubfolderDeleted] Oportunidad eliminada: " & subfolderName
    End If
End Sub

Private Sub mFolderWatcher_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch SubfolderRenamed] " & oldName & " -> " & newName & " en " & parentFolder

    If InStr(1, parentFolder, App().Configuration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Oportunidad renombrada - delegar a OpportunitiesMgr
        App().OpportunitiesMgr.ProcesarCambiosEnOportunidades oldName
        App().OpportunitiesMgr.ProcesarCambiosEnOportunidades newName
        LogInfo MODULE_NAME, "[callback: SubfolderRenamed] Oportunidad renombrada: " & oldName & " -> " & newName
    End If
End Sub

' Los eventos de archivos individuales se mantienen para plantillas y Gas
Private Sub mFolderWatcher_FileCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch FileCreated] " & fileName & " en " & folder

    ' Para plantillas o archivos Gas
    If InStr(1, folder, App().Configuration.RutaPlantillas, vbTextCompare) > 0 Then
        LogInfo MODULE_NAME, "[callback: FileCreated] Plantilla nueva: " & fileName
        ' TODO: Implementar lógica de actualización de plantillas
    ElseIf InStr(1, folder, App().Configuration.RutaGasVBNet, vbTextCompare) > 0 Then
        LogInfo MODULE_NAME, "[callback: FileCreated] Archivo Gas nuevo: " & fileName
        ' TODO: Implementar lógica de sincronización Gas
    End If
End Sub

Private Sub mFolderWatcher_FileDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch FileDeleted] " & fileName & " en " & folder

    If InStr(1, folder, App().Configuration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Es un archivo de oportunidad eliminado - delegar a OpportunitiesMgr
        App().OpportunitiesMgr.ProcesarCambiosEnItemsOportunidad fileName
    End If
End Sub

Private Sub mFolderWatcher_FileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch FileChanged] " & fileName & " en " & folder

    ' Detectar cambios en plantillas o archivos Gas
    If InStr(1, folder, App().Configuration.RutaPlantillas, vbTextCompare) > 0 Then
        LogInfo MODULE_NAME, "[callback: FileChanged] Plantilla modificada: " & fileName
        ' TODO: Notificar cambio en plantilla
    ElseIf InStr(1, folder, App().Configuration.RutaGasVBNet, vbTextCompare) > 0 Then
        LogInfo MODULE_NAME, "[callback: FileChanged] Archivo Gas modificado: " & fileName
        ' TODO: Actualizar referencias
    End If
End Sub

Private Sub mFolderWatcher_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch FileRenamed] " & oldName & " -> " & newName & " en " & folder

    If InStr(1, folder, App().Configuration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Archivo de oportunidad renombrado - delegar a OpportunitiesMgr
        App().OpportunitiesMgr.ProcesarCambiosEnItemsOportunidad newName
    End If
End Sub

' Los eventos Heartbeat y ErrorOccurred se mantienen igual
Private Sub mFolderWatcher_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat - solo para debug detallado si es necesario
    ' LogDebug MODULE_NAME, "[callback: clsFolderWatch Heartbeat] " & folder & " - " & lastUpdate
End Sub

Private Sub mFolderWatcher_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[callback: clsFolderWatch ErrorOccurred] ERROR en " & folder & ": " & errorMessage
    
    ' Mostrar al usuario solo errores cr�ticos
    If InStr(1, errorMessage, "L�mite", vbTextCompare) > 0 Or _
       InStr(1, errorMessage, "no encontrada", vbTextCompare) > 0 Then
        MsgBox "Error en monitoreo de carpeta:" & vbCrLf & _
               folder & vbCrLf & vbCrLf & _
               errorMessage, vbExclamation, "FolderWatcher"
    End If
End Sub

' Eventos de reconexion automatica del FolderWatcher
Private Sub mFolderWatcher_WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
    LogDebug MODULE_NAME, "[callback: clsFolderWatch WatcherReconnected] " & folder & " reconectado tras " & attempts & " intento(s)"
    ' Opcional: notificar al usuario solo si hubo multiples intentos
    If attempts > 1 Then
        LogWarning MODULE_NAME, "Monitoreo restaurado en " & folder & " tras " & attempts & " intentos"
    End If
End Sub

Private Sub mFolderWatcher_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[callback: clsFolderWatch WatcherReconnectionFailed] " & folder & ": " & reason

    ' Notificar al usuario que el monitoreo fallo permanentemente
    MsgBox "El monitoreo de la carpeta ha fallado permanentemente:" & vbCrLf & _
           folder & vbCrLf & vbCrLf & _
           "Razon: " & reason & vbCrLf & vbCrLf & _
           "Es posible que la carpeta no este accesible o haya problemas de red.", _
           vbExclamation, "FolderWatcher - Reconexion Fallida"
End Sub

