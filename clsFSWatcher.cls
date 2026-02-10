VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================
' WRAPPER VBA PARA COMPONENTE COM FolderWatcher
' =====================================================
' Este módulo envuelve el componente COM FolderWatcherCOM.dll
' escrito en VB.NET (.NET Framework 4.0).
'
' CARGA DEL COM SIN REGISTRO:
' -------------------------------------------------------------------------------------
' El COM se carga usando Activation Context API.
' Esto permite usar el COM sin registrarlo en Windows (sin permisos admin).
' Los archivos necesarios (DLL y manifest) deben estar en la carpeta de AddIns.
'
' GESTIÓN DE RECURSOS:
' -------------------------------------------------------------------------------------
' El COM implementa correctamente IDisposable (ver FolderWatcher.vb):
' - Dispose() llama a EnableRaisingEvents = False antes de liberar cada watcher
' - Finalize() actúa como safety net si Dispose() no se llama
' - GC.SuppressFinalize(Me) evita doble limpieza
'
' IMPORTANTE: Desde VBA, llamar siempre a StopWatching o usar el método
' Dispose() de esta clase ANTES de hacer Set = Nothing, para garantizar
' que el COM libere correctamente los FileSystemWatcher internos.
'
' NOTA SOBRE RECURSOS COM:
' El componente FolderWatcherCOM.dll (.NET Framework 4.0) usa FileSystemWatcher
' que puede quedarse residente si no se libera correctamente.
' Ver clsFolderWatch.cls para documentacion detallada y recomendaciones
' para el codigo VB.NET del COM.
'
' SOLUCION IMPLEMENTADA:
' - clsFolderWatch.Dispose() libera todos los recursos
' - clsApplication.Terminate() llama a Dispose antes de Set = Nothing
' =====================================================

'@Folder "2-Infraestructura.4-Servicios.FSWatcher"
Option Explicit

Private Const MODULE_NAME As String = "clsFSWatcher"

' =====================================================
' ENUMERACIONES (deben coincidir con FolderWatcher.vb)
' =====================================================

Public Enum FileFilterType
    None = 0
    FileSize = 1
    fileDate = 2
    FileAttributes = 4
    foldersOnly = 8
    FilesOnly = 16
End Enum

Public Enum AutoActionType
    None = 0
    MoveFile = 1
    CopyFile = 2
    Archive = 3
    Delete = 4
    LogOnly = 5
End Enum

Public Enum DateCompareMode
    CreatedAfter = 1
    CreatedBefore = 2
    ModifiedAfter = 3
    ModifiedBefore = 4
End Enum

' =====================================================
' WRAPPER VBA PARA COMPONENTE COM
' =====================================================

Private WithEvents fw As FolderWatcher
Attribute fw.VB_VarHelpID = -1

Private mInitialized As Boolean
Private mDisposed As Boolean                     ' Flag para evitar doble limpieza
Private mWatchedFolders As Object                ' Dictionary de carpetas monitoreadas

' =====================================================
' EVENTOS PÚBLICOS (se propagan a clsApplication)
' =====================================================

Public Event FileCreated(ByVal folder As String, ByVal fileName As String)
Public Event FileDeleted(ByVal folder As String, ByVal fileName As String)
Public Event FileChanged(ByVal folder As String, ByVal fileName As String)
Public Event FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
Public Event Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
Public Event ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)

' Nuevos eventos para subcarpetas
Public Event SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
Public Event SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)

' Eventos de reconexión automática para carpetas de red
Public Event WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
Public Event WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)

' =====================================================
' INICIALIZACIÓN Y LIMPIEZA
' =====================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize] Inicializando wrapper VBA para FolderWatcher COM"

    On Error GoTo ErrHandler

    mDisposed = False
    Set mWatchedFolders = CreateObject("Scripting.Dictionary")

    If Not CrearCOMConActivationContext() Then
        ' Si falla, intentar con CreateObject normal (COM registrado)
        LogWarning MODULE_NAME, "Activation Context falló, intentando con COM registrado"
        On Error GoTo ErrHandlerCOM
        Set fw = CreateObject("FolderWatcher.Monitor")
    End If

    If fw Is Nothing Then
        LogError MODULE_NAME, "[Class_Initialize] No se pudo crear el objeto COM FolderWatcher"
        mInitialized = False
        Exit Sub
    End If

    mInitialized = True
    LogDebug MODULE_NAME, "[Class_Initialize] Componente COM conectado: " & fw.GetConfiguration()
    Exit Sub

ErrHandlerCOM:
    LogError MODULE_NAME, "[Class_Initialize] Error creando COM (ni Activation Context ni registro)", Err.Number, Err.Description
    mInitialized = False
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[Class_Initialize] Error en inicialización COM", Err.Number, Err.Description
    mInitialized = False
End Sub

'@Description: Intenta crear el COM usando Activation Context (sin registro)
Private Function CrearCOMConActivationContext() As Boolean
    On Error GoTo ErrHandler

    ' Verificar si los archivos existen
    If Not ComprobarArchivosCOM() Then
        LogWarning MODULE_NAME, "[CrearCOMConActivationContext] Archivos COM no encontrados en: " & Application.UserLibraryPath
        CrearCOMConActivationContext = False
        Exit Function
    End If

    ' Intentar crear el COM usando Activation Context (sin registro)
    ' Inicializar Activation Context
    With VBA.CreateObject("Microsoft.Windows.ActCtx")
        .Manifest = ObtenerRutaManifest 'ObtenerRutaCOM '
        Set fw = .CreateObject("FolderWatcher.Monitor")
    End With

    If fw Is Nothing Then
        CrearCOMConActivationContext = False
    Else
        LogDebug MODULE_NAME, "[CrearCOMConActivationContext] COM creado exitosamente con Activation Context"
        CrearCOMConActivationContext = True
    End If

    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[CrearCOMConActivationContext]"
    CrearCOMConActivationContext = False
End Function

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "[Class_Terminate] Iniciando limpieza"
    Dispose
End Sub

'@Description: Limpia explicitamente todos los recursos del FolderWatcher
'@Note: Llamar este metodo ANTES de Set obj = Nothing para garantizar limpieza
Public Sub Dispose()
    On Error Resume Next  ' Asegurar que la limpieza continua aunque haya errores

    ' Evitar doble limpieza
    If mDisposed Then
        LogDebug MODULE_NAME, "[Dispose] Ya fue ejecutado, omitiendo"
        Exit Sub
    End If

    LogDebug MODULE_NAME, "[Dispose] Iniciando limpieza de recursos"

    ' 1. Detener todos los monitoreos activos
    If Not mWatchedFolders Is Nothing Then
        Dim carpeta As Variant
        Dim countFolders As Long
        countFolders = mWatchedFolders.Count

        For Each carpeta In mWatchedFolders.Keys
            If mInitialized And Not fw Is Nothing Then
                LogDebug MODULE_NAME, "[Dispose] Deteniendo monitoreo de: " & CStr(carpeta)
                fw.StopWatching CStr(carpeta)
            End If
        Next carpeta

        mWatchedFolders.RemoveAll
        LogDebug MODULE_NAME, "[Dispose] " & countFolders & " carpetas liberadas"
    End If

    ' 2. Liberar referencia al COM
    If Not fw Is Nothing Then
        LogDebug MODULE_NAME, "[Dispose] Liberando referencia COM"
        Set fw = Nothing
    End If

    ' 3. Liberar diccionario
    If Not mWatchedFolders Is Nothing Then
        Set mWatchedFolders = Nothing
    End If

    mInitialized = False
    mDisposed = True

    LogInfo MODULE_NAME, "[Dispose] Limpieza completada"
End Sub

'@Description: Indica si el objeto ha sido liberado
Public Property Get IsDisposed() As Boolean
    IsDisposed = mDisposed
End Property

'@Description: Indica si el COM esta inicializado correctamente
Public Property Get IsInitialized() As Boolean
    IsInitialized = mInitialized And Not mDisposed
End Property

' =====================================================
' MÉTODOS PRINCIPALES
' =====================================================

'@Description: Inicia monitoreo de una carpeta
'@Scope: público
'@ArgumentDescriptions:
'   folderPath: ruta completa de la carpeta a monitorear
'   includeSubdirs: monitorear subcarpetas recursivamente
'   filterPattern: patrones de archivo separados por ; (ej: "*.xlsx;*.xlsm")
'   eventsToWatch: array de eventos a monitorear ("Created","Deleted","Changed","Renamed")
'   inactivityMinutes: tiempo de inactividad antes de reinicio automático
'   foldersOnly: TRUE para monitorear solo cambios en subcarpetas, no archivos
Public Sub IniciarMonitoreo(ByVal folderPath As String, _
                            Optional ByVal includeSubdirs As Boolean = True, _
                            Optional ByVal filterPattern As String = "*.*", _
                            Optional ByVal eventsToWatch As Variant = Empty, _
                            Optional ByVal inactivityMinutes As Double = -1, _
                            Optional ByVal foldersOnly As Boolean = False)

    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then
        Err.Raise vbObjectError + 600, MODULE_NAME, "Componente COM no inicializado o ya liberado"
    End If

    ' Normalizar ruta
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    ' Verificar que la carpeta existe
    If Not RutaExiste(folderPath) Then
        LogWarning MODULE_NAME, "[IniciarMonitoreo] Ruta no existe, omitiendo monitoreo: " & folderPath
        Exit Sub
    End If

    ' Preparar eventos a monitorear
    Dim eventos As Variant
    If IsEmpty(eventsToWatch) Or Not IsArray(eventsToWatch) Then
        eventos = Array("Created", "Deleted", "Renamed")
    Else
        eventos = eventsToWatch
    End If

    ' Llamar al componente COM
    fw.WatchFolder folderPath, includeSubdirs, filterPattern, eventos, inactivityMinutes, foldersOnly

    ' Registrar en diccionario local
    If Not mWatchedFolders.Exists(folderPath) Then
        mWatchedFolders.Add folderPath, Now
    End If

    LogDebug MODULE_NAME, "[IniciarMonitoreo] Monitoreo iniciado: " & folderPath & IIf(foldersOnly, " (solo subcarpetas)", "")

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[IniciarMonitoreo] Error al iniciar monitoreo de: " & folderPath, Err.Number, Err.Description
    RaiseEvent ErrorOccurred(folderPath, Err.Description)
End Sub

'@Description: Detiene el monitoreo de una carpeta específica
Public Sub DetenerMonitoreo(ByVal folderPath As String)
    On Error Resume Next

    If mInitialized And Not mDisposed And Not fw Is Nothing Then
        fw.StopWatching folderPath
    End If

    If Not mWatchedFolders Is Nothing Then
        If mWatchedFolders.Exists(folderPath) Then
            mWatchedFolders.Remove folderPath
        End If
    End If

    LogInfo MODULE_NAME, "[DetenerMonitoreo] Monitoreo detenido: " & folderPath
End Sub

' =====================================================
' CONFIGURACIÓN DE FILTROS
' =====================================================

'@Description: Configura filtro por tamaño de archivo
Public Sub ConfigurarFiltroTamaño(ByVal folderPath As String, _
                                  ByVal minSize As Long, _
                                  ByVal maxSize As Long)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetFilter folderPath, FileFilterType.FileSize, minSize, maxSize
    LogDebug MODULE_NAME, "[ConfigurarFiltroTamaño] Filtro de tamaño: " & folderPath & " [" & minSize & "-" & maxSize & "]"

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ConfigurarFiltroTamaño] Error: " & folderPath, Err.Number, Err.Description
End Sub

'@Description: Configura filtro por fecha de archivo
Public Sub ConfigurarFiltroFecha(ByVal folderPath As String, _
                                 ByVal compareDate As Date, _
                                 ByVal dateMode As DateCompareMode)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetFilter folderPath, FileFilterType.fileDate, 0, 0, compareDate, dateMode
    Dim strFilt
    Select Case dateMode
        Case CreatedAfter: strFilt = "creac. > " & compareDate
        Case CreatedBefore: strFilt = "modif. < " & compareDate
        Case ModifiedAfter: strFilt = "creac. > " & compareDate
        Case ModifiedBefore: strFilt = "modif. < " & compareDate
    End Select

    LogDebug MODULE_NAME, "[ConfigurarFiltroFecha] Filtro de fecha configurado, " & strFilt & " en: " & folderPath

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ConfigurarFiltroFecha] Error: " & folderPath, Err.Number, Err.Description
End Sub

'@Description: Configura filtro para monitorear solo carpetas
Public Sub ConfigurarFiltroSoloCarpetas(ByVal folderPath As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetFilter folderPath, FileFilterType.foldersOnly
    LogDebug MODULE_NAME, "[ConfigurarFiltroSoloCarpetas] Filtro solo carpetas en: " & folderPath

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ConfigurarFiltroSoloCarpetas] Error: " & folderPath, Err.Number, Err.Description
End Sub

'@Description: Limpia todos los filtros de una carpeta
Public Sub LimpiarFiltros(ByVal folderPath As String)
    On Error Resume Next
    If mInitialized And Not mDisposed Then fw.ClearFilter folderPath
End Sub

' =====================================================
' CONFIGURACIÓN DE ACCIONES AUTOMÁTICAS
' =====================================================

'@Description: Configura acción de mover archivos automáticamente
Public Sub ConfigurarAccionMover(ByVal folderPath As String, _
                                 ByVal targetFolder As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetAutoAction folderPath, AutoActionType.MoveFile, targetFolder
    LogDebug MODULE_NAME, "[ConfigurarAccionMover] Accion mover configurada: " & folderPath & " -> " & targetFolder

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarAccionMover]"
End Sub

'@Description: Configura acción de copiar archivos automáticamente
Public Sub ConfigurarAccionCopiar(ByVal folderPath As String, _
                                  ByVal targetFolder As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetAutoAction folderPath, AutoActionType.CopyFile, targetFolder
    LogDebug MODULE_NAME, "[ConfigurarAccionCopiar] Accion copiar configurada: " & folderPath & " -> " & targetFolder

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarAccionCopiar]"
End Sub

'@Description: Configura acción de archivar con timestamp
Public Sub ConfigurarAccionArchivar(ByVal folderPath As String, _
                                    ByVal archiveFolder As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetAutoAction folderPath, AutoActionType.Archive, archiveFolder
    LogDebug MODULE_NAME, "[ConfigurarAccionArchivar] Accion archivar configurada: " & folderPath & " -> " & archiveFolder

    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[ConfigurarAccionArchivar]"
End Sub

'@Description: Limpia acciones automáticas de una carpeta
Public Sub LimpiarAcciones(ByVal folderPath As String)
    On Error Resume Next
    If mInitialized And Not mDisposed Then fw.ClearAutoAction folderPath
End Sub

' =====================================================
' PROPIEDADES Y CONSULTAS
' =====================================================

Public Property Get Configuracion() As String
    On Error Resume Next
    If mInitialized And Not mDisposed Then
        Configuracion = fw.GetConfiguration()
    Else
        Configuracion = "(no inicializado o liberado)"
    End If
End Property

Public Property Get CarpetasMonitoreadas() As Long
    If mWatchedFolders Is Nothing Then
        CarpetasMonitoreadas = 0
    Else
        CarpetasMonitoreadas = mWatchedFolders.Count
    End If
End Property

Public Function CarpetasActivas() As Variant
    On Error Resume Next
    If mInitialized And Not mDisposed Then
        CarpetasActivas = fw.ActiveFolders
    Else
        CarpetasActivas = Array()
    End If
End Function

Public Function ObtenerHistorial(Optional ByVal lastMinutes As Long = 0, _
                                 Optional ByVal eventType As String = "") As Variant
    On Error Resume Next
    If mInitialized And Not mDisposed Then
        ObtenerHistorial = fw.GetEventHistory(lastMinutes, eventType)
    Else
        ObtenerHistorial = Array()
    End If
End Function

Public Function ObtenerEstadisticas(ByVal folderPath As String) As Variant
    On Error Resume Next
    If mInitialized And Not mDisposed Then
        ObtenerEstadisticas = fw.GetStatistics(folderPath)
    Else
        ObtenerEstadisticas = Array()
    End If
End Function

Public Sub LimpiarHistorial()
    On Error Resume Next
    If mInitialized And Not mDisposed Then fw.ClearHistory
End Sub

' =====================================================
' MANEJADORES DE EVENTOS COM (propagan a VBA)
' =====================================================

Private Sub fw_FileCreated(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor FileCreated] " & fileName & " en " & folder
    RaiseEvent FileCreated(folder, fileName)
End Sub

Private Sub fw_FileDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor FileDeleted] " & fileName & " en " & folder
    RaiseEvent FileDeleted(folder, fileName)
End Sub

Private Sub fw_FileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor FileChanged] " & fileName & " en " & folder
    RaiseEvent FileChanged(folder, fileName)
End Sub

Private Sub fw_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor FileRenamed] " & oldName & " -> " & newName & " en " & folder
    RaiseEvent FileRenamed(folder, oldName, newName)
End Sub

Private Sub fw_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat es muy frecuente, no logear para evitar ruido
    RaiseEvent Heartbeat(folder, lastUpdate)
End Sub

Private Sub fw_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "[callback: FolderWatcher.Monitor ErrorOccurred] Error COM en " & folder & ": " & errorMessage
    RaiseEvent ErrorOccurred(folder, errorMessage)
End Sub

' Manejadores para eventos de subcarpetas
Private Sub fw_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor SubfolderCreated] " & subfolderName & " en " & parentFolder
    RaiseEvent SubfolderCreated(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor SubfolderDeleted] " & subfolderName & " en " & parentFolder
    RaiseEvent SubfolderDeleted(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor SubfolderRenamed] " & oldName & " -> " & newName & " en " & parentFolder
    RaiseEvent SubfolderRenamed(parentFolder, oldName, newName)
End Sub

' Manejadores para eventos de reconexión automática
Private Sub fw_WatcherReconnected(ByVal folder As String, ByVal attempts As Long)
    'LogDebug MODULE_NAME, "[callback: FolderWatcher.Monitor WatcherReconnected] " & folder & " (tras " & attempts & " intento(s))"
    RaiseEvent WatcherReconnected(folder, attempts)
End Sub

Private Sub fw_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[callback: FolderWatcher.Monitor WatcherReconnectionFailed] " & folder & " - " & reason
    RaiseEvent WatcherReconnectionFailed(folder, reason)
End Sub

' =====================================================
' FUNCIONES DE UTILIDAD
' =====================================================

'@Description: Obtiene la ruta donde debería estar el COM (carpeta AddIns)
Private Function ObtenerRutaCOM() As String
    ObtenerRutaCOM = Application.UserLibraryPath & FOLDERWATCHERCOM_NAME
End Function

'@Description: Obtiene la ruta donde debería estar el manifest (carpeta AddIns)
Private Function ObtenerRutaManifest() As String
    ObtenerRutaManifest = Application.UserLibraryPath & FOLDERWATCHERCOM_NAME & ".manifest"
End Function

'@Description: Verifica si los archivos del COM están presentes
Private Function ComprobarArchivosCOM() As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ComprobarArchivosCOM = fso.FileExists(ObtenerRutaCOM()) And _
                           fso.FileExists(ObtenerRutaManifest())

    Set fso = Nothing
End Function


' =====================================================
' FUNCIONES DE TEST Y DEBUGGING
' =====================================================

'@Description: Test de monitoreo de subcarpetas
Public Sub TestMonitoreoSubcarpetas()
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFolderWatcher"
    
    ' Crear carpeta de test si no existe
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Configurar monitoreo
    IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*", _
        eventsToWatch:=Array("Created", "Deleted", "Renamed"), _
        inactivityMinutes:=5, _
        foldersOnly:=True
    
    ConfigurarFiltroSoloCarpetas rutaTest
    
    MsgBox "Monitoreo de test iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Ahora crea, renombra o elimina subcarpetas para probar.", vbInformation
End Sub

'@Description: Test de filtros de tamaño
Public Sub TestFiltroTamaño()
    Dim rutaTest As String
    rutaTest = Environ("TEMP") & "\TestFiltros"
    
    If Not RutaExiste(rutaTest) Then
        MkDir rutaTest
    End If
    
    ' Monitorear solo archivos mayores a 1 KB
    IniciarMonitoreo _
        folderPath:=rutaTest, _
        includeSubdirs:=False, _
        filterPattern:="*.*", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    ConfigurarFiltroTamaño rutaTest, 1024, 999999999
    
    MsgBox "Monitoreo con filtro de tamaño iniciado en:" & vbCrLf & rutaTest & vbCrLf & vbCrLf & _
           "Solo detectará archivos mayores a 1 KB", vbInformation
End Sub

'@Description: Test de acción automática de mover
Public Sub TestAccionMover()
    Dim rutaOrigen As String, rutaDestino As String
    rutaOrigen = Environ("TEMP") & "\TestOrigen"
    rutaDestino = Environ("TEMP") & "\TestDestino"
    
    ' Crear carpetas si no existen
    If Not RutaExiste(rutaOrigen) Then MkDir rutaOrigen
    If Not RutaExiste(rutaDestino) Then MkDir rutaDestino
    
    ' Configurar monitoreo con acción de mover
    IniciarMonitoreo _
        folderPath:=rutaOrigen, _
        includeSubdirs:=False, _
        filterPattern:="*.txt", _
        eventsToWatch:=Array("Created"), _
        inactivityMinutes:=5
    
    ConfigurarAccionMover rutaOrigen, rutaDestino
    
    MsgBox "Test de acción automática iniciado:" & vbCrLf & _
           "Origen: " & rutaOrigen & vbCrLf & _
           "Destino: " & rutaDestino & vbCrLf & vbCrLf & _
           "Los archivos .txt se moverán automáticamente", vbInformation
End Sub

'@Description: Test completo del sistema
Public Sub TestSistemaCompleto()
    Debug.Print "==== TEST COMPLETO FOLDERWATCHER ===="
    Debug.Print "Configuración: " & Configuracion
    
    Dim carpetas As Variant
    carpetas = CarpetasActivas()
    
    Debug.Print "Carpetas monitoreadas: " & IIf(IsArray(carpetas), UBound(carpetas) + 1, 0)
    
    If IsArray(carpetas) Then
        Dim i As Long
        For i = LBound(carpetas) To UBound(carpetas)
            Debug.Print "  " & i & ": " & carpetas(i)
            
            Dim stats As Variant
            stats = ObtenerEstadisticas(CStr(carpetas(i)))
            
            If IsArray(stats) Then
                Debug.Print "     Eventos: " & stats(1) & " | Última actividad: " & stats(6)
            End If
        Next i
    End If
    
    Debug.Print "==== FIN TEST ===="
End Sub


