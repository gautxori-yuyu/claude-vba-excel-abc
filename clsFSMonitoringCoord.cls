VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSMonitoringCoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.4-Servicios.FSWatcher"
Option Explicit

Private Const MODULE_NAME As String = "clsFSMonitoringCoord"

Private WithEvents mFolderWatcher As clsFSWatcher
Attribute mFolderWatcher.VB_VarHelpID = -1

' Clasificador de rutas (sabe a que tipo semantico pertenece cada carpeta)
Private mFolderMgr As clsFolderMgr

Private mDisposed As Boolean ' Flag para limpieza explícita

' -------------------------------------------------------------
' EVENTOS PUBLICOS - clsApplication los escucha y delega
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

' Eventos de error/reconexión del monitoreo
Public Event MonitoringError(ByVal folder As String, ByVal errorMessage As String)
Public Event MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
Public Event MonitoringFailed(ByVal folder As String, ByVal reason As String)

' -------------------------------------------------------------
' Properties
' -------------------------------------------------------------
'@Description: Acceso al FolderWatcher (para diagnostico y tests)
Public Property Get FolderWatcher() As clsFSWatcher
    Set FolderWatcher = mFolderWatcher
End Property

Private Sub Class_Initialize()
    Set mFolderWatcher = New clsFSWatcher
    Set mFolderMgr = New clsFolderMgr
    mDisposed = False
End Sub

Private Sub Class_Terminate()
    ' Red de seguridad por si Shutdown no fue llamado explícitamente
    LogInfo MODULE_NAME, "[Terminate] Limpieza de seguridad iniciada"
    Shutdown
End Sub

'@Description: Libera explícitamente los recursos del monitor de archivos.
Public Sub Shutdown()
    If mDisposed Then Exit Sub

    On Error Resume Next
    LogInfo MODULE_NAME, "[Shutdown] Liberando FolderWatcher COM explícitamente"

    If Not mFolderWatcher Is Nothing Then
        mFolderWatcher.Dispose
        Set mFolderWatcher = Nothing
    End If

    Set mFolderMgr = Nothing

    mDisposed = True
    LogDebug MODULE_NAME, "[Shutdown] Limpieza explícita completada"
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

    ' Cargar el registro de rutas en el clasificador
    mFolderMgr.Initialize oDicFolders

    ' Monitorear cada carpeta configurada usando las funciones helper
    Dim key As Variant
    For Each key In oDicFolders.Keys
        Dim rutaCarpeta As String
        rutaCarpeta = oDicFolders(key)

        If RutaExiste(rutaCarpeta) Then
            Select Case key
            Case CFG_RUTA_OPORTUNIDADES
                ' Usar función especializada para oportunidades
                ConfigurarMonitoreoOportunidades rutaCarpeta
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Oportunidades monitoreadas: " & rutaCarpeta

            Case CFG_RUTA_PLANTILLAS
                ' Usar función especializada para plantillas
                ConfigurarMonitoreoPlantillas rutaCarpeta
                LogInfo MODULE_NAME, "[IniciarMonitoreo] Plantillas monitoreadas: " & rutaCarpeta

            Case CFG_RUTA_GAS_VBNET
                ' Usar función especializada para Gas
                ConfigurarMonitoreoGasVBNet rutaCarpeta
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
' NOTA: Todos los callbacks disparan RaiseEvent, NO llaman a App()
' -------------------------------------------------------------

' Eventos de subcarpetas (reemplazan los eventos de archivos para oportunidades)
Private Sub mFolderWatcher_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)

    LogDebug MODULE_NAME, "[SubfolderCreated] " & subfolderName & " en " & parentFolder

    ' Determinar tipo de carpeta padre
    If mFolderMgr.ClassifyPath(parentFolder) = ftOpportunities Then
        LogInfo MODULE_NAME, "[SubfolderCreated] Nueva oportunidad: " & subfolderName
        RaiseEvent OpportunityCreated(parentFolder, subfolderName)
    End If
End Sub

Private Sub mFolderWatcher_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)

    LogDebug MODULE_NAME, "[SubfolderDeleted] " & subfolderName & " en " & parentFolder

    If mFolderMgr.ClassifyPath(parentFolder) = ftOpportunities Then
        LogInfo MODULE_NAME, "[SubfolderDeleted] Oportunidad eliminada: " & subfolderName
        RaiseEvent OpportunityDeleted(parentFolder, subfolderName)
    End If
End Sub

Private Sub mFolderWatcher_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)

    LogDebug MODULE_NAME, "[SubfolderRenamed] " & oldName & " -> " & newName & " en " & parentFolder

    If mFolderMgr.ClassifyPath(parentFolder) = ftOpportunities Then
        LogInfo MODULE_NAME, "[SubfolderRenamed] Oportunidad renombrada: " & oldName & " -> " & newName
        RaiseEvent OpportunityRenamed(parentFolder, oldName, newName)
    End If
End Sub

' Los eventos de archivos individuales se mantienen para plantillas y Gas
Private Sub mFolderWatcher_FileCreated(ByVal folder As String, ByVal fileName As String)

    LogDebug MODULE_NAME, "[FileCreated] " & fileName & " en " & folder

    ' Para plantillas o archivos Gas - implementar según necesidad
    Select Case mFolderMgr.ClassifyPath(folder)
    Case ftTemplates
        LogInfo MODULE_NAME, "[FileCreated] Plantilla nueva: " & fileName
        ' TODO: Implementar lógica de actualización de plantillas
        RaiseEvent TemplateCreated(folder, fileName)
    Case ftGasVBNet
        LogInfo MODULE_NAME, "[FileCreated] Archivo Gas nuevo: " & fileName
        ' TODO: Implementar lógica de sincronización Gas
        RaiseEvent GasFileCreated(folder, fileName)
    End Select
End Sub

Private Sub mFolderWatcher_FileDeleted(ByVal folder As String, ByVal fileName As String)

    LogDebug MODULE_NAME, "[FileDeleted] " & fileName & " en " & folder

    If mFolderMgr.ClassifyPath(folder) = ftOpportunities Then
        LogInfo MODULE_NAME, "[FileDeleted] Item de oportunidad eliminado: " & fileName
        RaiseEvent OpportunityItemDeleted(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileChanged(ByVal folder As String, ByVal fileName As String)

    LogDebug MODULE_NAME, "[FileChanged] " & fileName & " en " & folder

    ' Detectar cambios en plantillas o archivos Gas
    Select Case mFolderMgr.ClassifyPath(folder)
    Case ftTemplates
        LogInfo MODULE_NAME, "[FileChanged] Plantilla modificada: " & fileName
        ' TODO: Notificar cambio en plantilla
        RaiseEvent TemplateChanged(folder, fileName)
    Case ftGasVBNet
        LogInfo MODULE_NAME, "[FileChanged] Archivo Gas modificado: " & fileName
        ' TODO: Actualizar referencias
        RaiseEvent GasFileChanged(folder, fileName)
    End Select
End Sub

Private Sub mFolderWatcher_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)

    LogDebug MODULE_NAME, "[FileRenamed] " & oldName & " -> " & newName & " en " & folder

    If mFolderMgr.ClassifyPath(folder) = ftOpportunities Then
        LogInfo MODULE_NAME, "[FileRenamed] Item de oportunidad renombrado: " & oldName & " -> " & newName
        RaiseEvent OpportunityItemRenamed(folder, oldName, newName)
    End If
End Sub

' Los eventos Heartbeat y ErrorOccurred se mantienen igual
Private Sub mFolderWatcher_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat - solo para debug detallado si es necesario
    ' LogDebug MODULE_NAME, "[Heartbeat] " & folder & " - " & lastUpdate
End Sub

Private Sub mFolderWatcher_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[ErrorOccurred] ERROR en " & folder & ": " & errorMessage
    RaiseEvent MonitoringError(folder, errorMessage)
End Sub

Private Sub mFolderWatcher_WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
    If attempts > 1 Then
        LogWarning MODULE_NAME, "Monitoreo restaurado en " & folder & " tras " & attempts & " intentos"
    End If
    RaiseEvent MonitoringReconnected(folder, attempts)
End Sub

Private Sub mFolderWatcher_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[WatcherReconnectionFailed] " & folder & ": " & reason
    RaiseEvent MonitoringFailed(folder, reason)
End Sub

' =====================================================
' FUNCIONES DE CONFIGURACIÓN RÁPIDA
' =====================================================

'@Description: Configura monitoreo de subcarpetas de oportunidades
'@Scope: Friend (solo clsApplication)
Public Sub ConfigurarMonitoreoOportunidades(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoOportunidades] Ruta de oportunidades no existe: " & rutaBase
        Exit Sub
    End If

    ' Monitorear SOLO subcarpetas (no archivos individuales)
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=IIf(IsNetworkPath(rutaBase), 15, 5), _
        foldersOnly:=True

    ' Configurar filtro para solo carpetas
    mFolderWatcher.ConfigurarFiltroSoloCarpetas rutaBase

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoOportunidades] configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarMonitoreoOportunidades]"
End Sub

'@Description: Configura monitoreo de archivos de plantillas
Public Sub ConfigurarMonitoreoPlantillas(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoPlantillas] Ruta de plantillas no existe: " & rutaBase
        Exit Sub
    End If

    ' Monitorear cambios en archivos Excel
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=True, _
        filterPattern:="*.xlsx;*.xlsm", _
        eventsToWatch:=Array("Changed", "Created"), _
        inactivityMinutes:=10, _
        foldersOnly:=False

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoPlantillas] configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarMonitoreoPlantillas]"
End Sub

'@Description: Configura monitoreo de archivos Gas (C-GAS-ING)
Public Sub ConfigurarMonitoreoGasVBNet(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoGasVBNet] Ruta no existe: " & rutaBase
        Exit Sub
    End If

    ' Monitorear cambios en archivos Excel
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaBase, _
        includeSubdirs:=True, _
        filterPattern:="*.xlsx;*.xlsm", _
        eventsToWatch:=Array("Changed"), _
        inactivityMinutes:=10, _
        foldersOnly:=False

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoGasVBNet] configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarMonitoreoGasVBNet]"
End Sub

' =====================================================
' FUNCIONES DE CONSULTA Y DIAGNÓSTICO
' =====================================================

'@Description: Muestra estadísticas de monitoreo en MessageBox
Public Sub VerEstadisticasMonitoreo()
    If mFolderWatcher Is Nothing Then
        ShowTaskDialogError "FolderWatcher", "No disponible", "FolderWatcher no está activo"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' Obtener carpetas activas
    Dim carpetas As Variant
    carpetas = mFolderWatcher.CarpetasActivas()

    If Not IsArray(carpetas) Then
        ShowTaskDialogError "Estadísticas", "Sin datos", "No hay carpetas monitoreadas"
        Exit Sub
    End If

    If UBound(carpetas) < LBound(carpetas) Then
        ShowTaskDialogError "Estadísticas", "Sin datos", "No hay carpetas monitoreadas"
        Exit Sub
    End If

    ' Mostrar estadísticas de cada carpeta
    Dim msg As String
    msg = "CARPETAS MONITOREADAS:" & vbCrLf & vbCrLf

    Dim i As Long
    For i = LBound(carpetas) To UBound(carpetas)
        Dim stats As Variant
        stats = mFolderWatcher.ObtenerEstadisticas(CStr(carpetas(i)))

        If IsArray(stats) And UBound(stats) >= 8 Then
            msg = msg & "[i] " & stats(0) & vbCrLf
            msg = msg & "   Total eventos: " & stats(1) & vbCrLf
            msg = msg & "   Creados: " & stats(2) & " | "
            msg = msg & "Eliminados: " & stats(3) & " | "
            msg = msg & "Modificados: " & stats(4) & vbCrLf
            msg = msg & "   Renombrados: " & stats(5) & " | "
            msg = msg & "Errores: " & stats(8) & vbCrLf
            msg = msg & "   Eventos/hora: " & Format(stats(7), "0.0") & vbCrLf
            msg = msg & "   Última actividad: " & Format(stats(6), "dd/mm/yyyy hh:mm:ss") & vbCrLf
            msg = msg & vbCrLf
        End If
    Next i

    ShowTaskDialogError "Estadísticas FolderWatcher", "Monitoreo activo", msg
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[VerEstadisticasMonitoreo]"
    Err.Raise Err.Number, MODULE_NAME & " [VerEstadisticasMonitoreo]", Err.Description
End Sub

'@Description: Genera hoja de Excel con historial de eventos
Public Sub VerHistorialMonitoreo(Optional ByVal lastMinutes As Long = 60)
    If mFolderWatcher Is Nothing Then
        ShowTaskDialogError "FolderWatcher", "No disponible", "FolderWatcher no está activo"
        Exit Sub
    End If

    On Error GoTo ErrHandler

    ' Obtener historial
    Dim history As Variant
    history = mFolderWatcher.ObtenerHistorial(lastMinutes, "")

    If Not IsArray(history) Then
        ShowTaskDialogError "Historial", "Sin datos", "No hay eventos en el historial"
        Exit Sub
    End If

    If UBound(history, 1) < 0 Then
        ShowTaskDialogError "Historial", "Sin datos", "No hay eventos en el historial"
        Exit Sub
    End If

    ' Crear hoja con historial
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = "FW_Historial_" & Format(Now, "hhmmss")

    ' Encabezados
    With ws
        .Range("A1:E1").Value = Array("Timestamp", "Evento", "Carpeta", "Archivo", "Tamaño (bytes)")
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)

        ' Datos
        Dim rowCount As Long
        rowCount = UBound(history, 1) - LBound(history, 1) + 1
        .Range("A2").Resize(rowCount, 5).Value = history

        ' Formato
        .Columns("A:A").NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Columns("E:E").NumberFormat = "#,##0"
        .Columns("A:E").AutoFit
        .Columns("C:C").ColumnWidth = 50
        .Columns("D:D").ColumnWidth = 30
    End With

    ShowTaskDialogError "Historial", "Generado correctamente", "Historial generado: " & rowCount & " eventos"
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[VerHistorialMonitoreo]"
    Err.Raise Err.Number, MODULE_NAME & " [VerHistorialMonitoreo]", Err.Description
End Sub

'@Description: Limpia el historial de eventos
Public Sub LimpiarHistorialMonitoreo()
    If mFolderWatcher Is Nothing Then
        ShowTaskDialogError "FolderWatcher", "No disponible", "FolderWatcher no está activo"
        Exit Sub
    End If

    If ShowTaskDialogYesNo("Limpiar historial", "¿Deseas limpiar el historial?", "Se eliminarán todos los eventos registrados.") = vbYes Then
        mFolderWatcher.LimpiarHistorial
        ShowTaskDialogError "Historial", "Operación completada", "Historial limpiado correctamente"
    End If
End Sub

'@Description: Muestra información de configuración del watcher
Public Sub VerConfiguracionWatcher()
    If mFolderWatcher Is Nothing Then
        ShowTaskDialogError "FolderWatcher", "No disponible", "FolderWatcher no está activo"
        Exit Sub
    End If

    Dim config As String
    config = mFolderWatcher.Configuracion

    Dim carpetas As Variant
    carpetas = mFolderWatcher.CarpetasActivas()

    Dim msg As String
    msg = "CONFIGURACIÓN DEL FOLDERWATCHER" & vbCrLf & vbCrLf
    msg = msg & config & vbCrLf & vbCrLf
    msg = msg & "CARPETAS ACTIVAS:" & vbCrLf

    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            msg = msg & "  -> " & carpetas(i) & vbCrLf
        Next i
    Else
        msg = msg & "  (ninguna)" & vbCrLf
    End If

    ShowTaskDialogError "Configuración FolderWatcher", "Estado actual", msg
End Sub
