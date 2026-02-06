VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSMonitoringCoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "5-Aplicac (Coord)"
Option Explicit

Private Const MODULE_NAME As String = "clsFSMonitoringCoord"

Private WithEvents mFolderWatcher As clsFSWatcher
Attribute mFolderWatcher.VB_VarHelpID = -1

' Rutas almacenadas localmente (recibidas en IniciarMonitoreo)
Private m_rutaOportunidades As String
Private m_rutaPlantillas As String
Private m_rutaGasVBNet As String

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
    If EsRutaOportunidades(parentFolder) Then
        ' Es una nueva oportunidad - delegar a OpportunitiesMgr
        LogInfo MODULE_NAME, "[SubfolderCreated] Nueva oportunidad: " & subfolderName
        RaiseEvent OpportunityCreated(parentFolder, subfolderName)
    End If
End Sub

Private Sub mFolderWatcher_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[SubfolderDeleted] " & subfolderName & " en " & parentFolder

    If EsRutaOportunidades(parentFolder) Then
        ' Oportunidad eliminada - delegar a OpportunitiesMgr
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

' Los eventos de archivos individuales se mantienen para plantillas y Gas
Private Sub mFolderWatcher_FileCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileCreated] " & fileName & " en " & folder
    
    ' Para plantillas o archivos Gas - implementar según necesidad
    If EsRutaPlantillas(folder) Then
        LogInfo MODULE_NAME, "[FileCreated] Plantilla nueva: " & fileName
        ' TODO: Implementar lógica de actualización de plantillas
        RaiseEvent TemplateCreated(folder, fileName)
    ElseIf EsRutaGasVBNet(folder) Then
        LogInfo MODULE_NAME, "[FileCreated] Archivo Gas nuevo: " & fileName
        ' TODO: Implementar lógica de sincronización Gas
         RaiseEvent GasFileCreated(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileDeleted] " & fileName & " en " & folder
    ' Implementar según necesidad
    ' Determinar tipo de carpeta
    If EsRutaOportunidades(folder) Then
        LogInfo MODULE_NAME, "[FileDeleted] Item de oportunidad eliminado: " & fileName
        RaiseEvent OpportunityItemDeleted(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[FileChanged] " & fileName & " en " & folder
    
    ' Detectar cambios en plantillas o archivos Gas
    If EsRutaPlantillas(folder) Then
        LogInfo MODULE_NAME, "[FileChanged] Plantilla modificada: " & fileName
        ' TODO: Notificar cambio en plantilla
        RaiseEvent TemplateChanged(folder, fileName)
    ElseIf EsRutaGasVBNet(folder) Then
        LogInfo MODULE_NAME, "[FileChanged] Archivo Gas modificado: " & fileName
        ' TODO: Actualizar referencias
        RaiseEvent GasFileChanged(folder, fileName)
    End If
End Sub

Private Sub mFolderWatcher_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[FileRenamed] " & oldName & " -> " & newName & " en " & folder
    ' Implementar según necesidad
    ' Determinar tipo de carpeta
    If EsRutaOportunidades(folder) Then
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
    ' Mostrar al usuario solo errores críticos
    If InStr(1, errorMessage, "Límite", vbTextCompare) > 0 Or _
       InStr(1, errorMessage, "no encontrada", vbTextCompare) > 0 Then
        MsgBox "Error en monitoreo de carpeta:" & vbCrLf & _
               folder & vbCrLf & vbCrLf & _
               errorMessage, vbExclamation, "FolderWatcher"
    End If
End Sub

' Eventos de reconexion automatica del FolderWatcher
Private Sub mFolderWatcher_WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
    LogDebug MODULE_NAME, "[WatcherReconnected] " & folder & " reconectado tras " & attempts & " intento(s)"
    ' Opcional: notificar al usuario solo si hubo multiples intentos
    If attempts > 1 Then
        LogWarning MODULE_NAME, "Monitoreo restaurado en " & folder & " tras " & attempts & " intentos"
    End If
    RaiseEvent MonitoringReconnected(folder, attempts)
End Sub

Private Sub mFolderWatcher_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[WatcherReconnectionFailed] " & folder & ": " & reason
    RaiseEvent MonitoringFailed(folder, reason)

    ' Notificar al usuario que el monitoreo fallo permanentemente
    MsgBox "El monitoreo de la carpeta ha fallado permanentemente:" & vbCrLf & _
           folder & vbCrLf & vbCrLf & _
           "Razon: " & reason & vbCrLf & vbCrLf & _
           "Es posible que la carpeta no este accesible o haya problemas de red.", _
           vbExclamation, "FolderWatcher - Reconexion Fallida"
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

' =====================================================
' FUNCIONES DE CONFIGURACIÓN RÁPIDA
' =====================================================

'@Description: Configura monitoreo de subcarpetas de oportunidades
'@Scope: Friend (solo clsApplication)
Public Sub ConfigurarMonitoreoOportunidades(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoOportunidades] - Ruta de oportunidades no existe: " & rutaBase
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

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoOportunidades] - configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[ConfigurarMonitoreoOportunidades] - Error", Err.Number, Err.Description
End Sub

'@Description: Configura monitoreo de archivos de plantillas
Public Sub ConfigurarMonitoreoPlantillas(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - Ruta de plantillas no existe: " & rutaBase
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

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - Error", Err.Number, Err.Description
End Sub

'@Description: Configura monitoreo de archivos Gas (C-GAS-ING)
Public Sub ConfigurarMonitoreoGasVBNet(ByVal rutaBase As String)
    On Error GoTo ErrHandler

    If Not RutaExiste(rutaBase) Then
        LogWarning MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - Ruta no existe: " & rutaBase
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

    LogInfo MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - configurado en: " & rutaBase
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[ConfigurarMonitoreoPlantillas] - Error", Err.Number, Err.Description
End Sub

' =====================================================
' FUNCIONES DE CONSULTA Y DIAGNÓSTICO
' =====================================================

'@Description: Muestra estadísticas de monitoreo en MessageBox
Public Sub VerEstadisticasMonitoreo()
    If mFolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Obtener carpetas activas
    Dim carpetas As Variant
    carpetas = mFolderWatcher.CarpetasActivas()
    
    If Not IsArray(carpetas) Then
        MsgBox "No hay carpetas monitoreadas", vbInformation
        Exit Sub
    End If
    
    If UBound(carpetas) < LBound(carpetas) Then
        MsgBox "No hay carpetas monitoreadas", vbInformation
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
    
    MsgBox msg, vbInformation, "Estadísticas FolderWatcher"
    Exit Sub
    
ErrHandler:
    LogCurrentError MODULE_NAME, "[VerEstadisticasMonitoreo]"
    Err.Raise Err.Number, MODULE_NAME & " [VerEstadisticasMonitoreo]", Err.Description
End Sub

'@Description: Genera hoja de Excel con historial de eventos
Public Sub VerHistorialMonitoreo(Optional ByVal lastMinutes As Long = 60)
    If mFolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    
    ' Obtener historial
    Dim history As Variant
    history = mFolderWatcher.ObtenerHistorial(lastMinutes, "")
    
    If Not IsArray(history) Then
        MsgBox "No hay eventos en el historial", vbInformation
        Exit Sub
    End If
    
    If UBound(history, 1) < 0 Then
        MsgBox "No hay eventos en el historial", vbInformation
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
    
    MsgBox "Historial generado: " & rowCount & " eventos", vbInformation
    Exit Sub
    
ErrHandler:
    LogCurrentError MODULE_NAME, "[VerHistorialMonitoreo]"
    Err.Raise Err.Number, MODULE_NAME & " [VerHistorialMonitoreo]", Err.Description
End Sub

'@Description: Limpia el historial de eventos
Public Sub LimpiarHistorialMonitoreo()
    If mFolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("¿Deseas limpiar el historial de eventos?", vbQuestion + vbYesNo) = vbYes Then
        mFolderWatcher.LimpiarHistorial
        MsgBox "Historial limpiado", vbInformation
    End If
End Sub

'@Description: Muestra información de configuración del watcher
Public Sub VerConfiguracionWatcher()
    If mFolderWatcher Is Nothing Then
        MsgBox "FolderWatcher no está activo", vbExclamation
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
    
    MsgBox msg, vbInformation, "Configuración FolderWatcher"
End Sub

' =====================================================
' FUNCIONES DE TEST Y DEBUGGING
' =====================================================

'@Description: Test de monitoreo de subcarpetas
Sub TestMonitoreoSubcarpetas()
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFolderWatcher"
    
    ' Crear carpeta de test si no existe
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Configurar monitoreo
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=5, _
        foldersOnly:=True
    
    mFolderWatcher.ConfigurarFiltroSoloCarpetas rutaTest
    
    MsgBox "Monitoreo de test iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Ahora crea, renombra o elimina subcarpetas para probar.", vbInformation
End Sub

'@Description: Test de filtros de tamaño
Sub TestFiltroTamaño()
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFiltros"
    
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Monitorear solo archivos mayores a 1 KB
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*.*", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    mFolderWatcher.ConfigurarFiltroTamaño rutaTest, 1024, 999999999
    
    MsgBox "Monitoreo con filtro de tamaño iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Solo detectará archivos mayores a 1 KB", vbInformation
End Sub

'@Description: Test de detección de rutas de red
Sub TestDeteccionRutasRed()
    Dim rutasPrueba As Variant
    rutasPrueba = Array( _
                  "C:\Windows", _
                  "Z:\Compartido\Docs", _
                  "\\servidor\publico", _
                  "\\192.168.1.100\files" _
                  )
    
    Dim ruta As Variant
    For Each ruta In rutasPrueba
        Debug.Print ruta & " -> " & IIf(IsNetworkPath(CStr(ruta)), "RED", "LOCAL")
    Next
    
    MsgBox "Ver resultados en Ventana Inmediato (Ctrl+G)", vbInformation
End Sub

'@Description: Test de acción automática de mover
Sub TestAccionMover()
    Dim rutaOrigen As String, rutaDestino As String
    rutaOrigen = Environ("TEMP") & "\TestOrigen"
    rutaDestino = Environ("TEMP") & "\TestDestino"
    
    ' Crear carpetas si no existen
    If Not RutaExiste(rutaOrigen) Then MkDir rutaOrigen
    If Not RutaExiste(rutaDestino) Then MkDir rutaDestino
    
    ' Configurar monitoreo con acción de mover
    mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaOrigen, _
        includeSubdirs:=False, _
        filterPattern:="*.txt", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    mFolderWatcher.ConfigurarAccionMover rutaOrigen, rutaDestino
    
    MsgBox "Test de acción automática iniciado:" & vbCrLf & _
           "Origen: " & rutaOrigen & vbCrLf & _
           "Destino: " & rutaDestino & vbCrLf & vbCrLf & _
           "Los archivos .txt se moverán automáticamente", vbInformation
End Sub

'@Description: Test completo del sistema
Sub TestSistemaCompleto()
    Debug.Print "==== TEST COMPLETO FOLDERWATCHER ===="
    Debug.Print "Configuración: " & mFolderWatcher.Configuracion
    
    Dim carpetas As Variant
    carpetas = mFolderWatcher.CarpetasActivas()
    
    Debug.Print "Carpetas monitoreadas: " & IIf(IsArray(carpetas), UBound(carpetas) + 1, 0)
    
    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            Debug.Print "  " & i & ": " & carpetas(i)
            
            Dim stats As Variant
            stats = mFolderWatcher.ObtenerEstadisticas(CStr(carpetas(i)))
            
            If IsArray(stats) Then
                Debug.Print "     Eventos: " & stats(1) & " | Última actividad: " & stats(6)
            End If
        Next i
    End If
    
    Debug.Print "==== FIN TEST ===="
End Sub


