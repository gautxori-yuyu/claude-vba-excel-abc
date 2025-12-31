VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolderWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "5-FolderWatcher"
Option Explicit

' =====================================================
' WRAPPER VBA PARA COMPONENTE COM FolderWatcher
' =====================================================
' Este módulo envuelve el componente COM FolderWatcherCOM.dll
' escrito en VB.NET (.NET Framework 4.0).
'
' CARGA DEL COM SIN REGISTRO:
' -------------------------------------------------------------------------------------
' El COM se carga usando Activation Context API (ver modActivationContext.bas).
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
' =====================================================

' =====================================================
' ENUMERACIONES (deben coincidir con FolderWatcher.vb)
' =====================================================

Public Enum FileFilterType
    None = 0
    FileSize = 1
    FileDate = 2
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

Private Const MODULE_NAME As String = "clsFolderWatch"

' =====================================================
' EVENTOS PÚBLICOS (se propagan a clsAplicacion)
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

' =====================================================
' INICIALIZACIÓN Y LIMPIEZA
' =====================================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "Inicializando wrapper VBA para FolderWatcher COM"

    On Error GoTo ErrHandler

    mDisposed = False
    Set mWatchedFolders = CreateObject("Scripting.Dictionary")

    ' Intentar crear el COM usando Activation Context (sin registro)
    If Not CrearCOMConActivationContext() Then
        ' Si falla, intentar con CreateObject normal (COM registrado)
        LogWarning MODULE_NAME, "Activation Context falló, intentando con COM registrado"
        On Error GoTo ErrHandlerCOM
        Set fw = CreateObject("FolderWatcher.Monitor")
    End If

    If fw Is Nothing Then
        LogError MODULE_NAME, "No se pudo crear el objeto COM FolderWatcher"
        mInitialized = False
        Exit Sub
    End If

    mInitialized = True
    LogInfo MODULE_NAME, "Componente COM conectado: " & fw.GetConfiguration()
    Exit Sub

ErrHandlerCOM:
    LogError MODULE_NAME, "Error creando COM (ni Activation Context ni registro)", Err.Number, Err.Description
    mInitialized = False
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error en inicialización COM", Err.Number, Err.Description
    mInitialized = False
End Sub

'@Description: Intenta crear el COM usando Activation Context (sin registro)
Private Function CrearCOMConActivationContext() As Boolean
    On Error GoTo ErrHandler

    Dim rutaManifest As String
    rutaManifest = ObtenerRutaManifest()

    ' Verificar si los archivos existen
    If Not ComprobarArchivosCOM() Then
        LogWarning MODULE_NAME, "Archivos COM no encontrados en: " & Application.UserLibraryPath
        CrearCOMConActivationContext = False
        Exit Function
    End If

    ' Inicializar Activation Context
    If Not InicializarActivationContext(rutaManifest) Then
        LogWarning MODULE_NAME, "No se pudo inicializar Activation Context"
        CrearCOMConActivationContext = False
        Exit Function
    End If

    ' Activar contexto y crear objeto
    If Not ActivarContexto() Then
        LogWarning MODULE_NAME, "No se pudo activar contexto"
        CrearCOMConActivationContext = False
        Exit Function
    End If

    ' Crear el objeto COM mientras el contexto está activo
    Set fw = CreateObject("FolderWatcher.Monitor")

    ' Desactivar contexto (el objeto ya está creado)
    DesactivarContexto

    If fw Is Nothing Then
        CrearCOMConActivationContext = False
    Else
        LogInfo MODULE_NAME, "COM creado con Activation Context"
        CrearCOMConActivationContext = True
    End If

    Exit Function

ErrHandler:
    DesactivarContexto
    LogError MODULE_NAME, "Error en CrearCOMConActivationContext", Err.Number, Err.Description
    CrearCOMConActivationContext = False
End Function

Private Sub Class_Terminate()
    LogInfo MODULE_NAME, "Class_Terminate - Iniciando limpieza"
    Dispose
End Sub

'@Description: Limpia explicitamente todos los recursos del FolderWatcher
'@Note: Llamar este metodo ANTES de Set obj = Nothing para garantizar limpieza
Public Sub Dispose()
    On Error Resume Next  ' Asegurar que la limpieza continua aunque haya errores

    ' Evitar doble limpieza
    If mDisposed Then
        LogDebug MODULE_NAME, "Dispose: Ya fue ejecutado, omitiendo"
        Exit Sub
    End If

    LogInfo MODULE_NAME, "Dispose: Iniciando limpieza de recursos"

    ' 1. Detener todos los monitoreos activos
    If Not mWatchedFolders Is Nothing Then
        Dim carpeta As Variant
        Dim countFolders As Long
        countFolders = mWatchedFolders.Count

        For Each carpeta In mWatchedFolders.Keys
            If mInitialized And Not fw Is Nothing Then
                LogDebug MODULE_NAME, "Dispose: Deteniendo monitoreo de: " & CStr(carpeta)
                fw.StopWatching CStr(carpeta)
            End If
        Next carpeta

        mWatchedFolders.RemoveAll
        LogInfo MODULE_NAME, "Dispose: " & countFolders & " carpetas liberadas"
    End If

    ' 2. Liberar referencia al COM
    If Not fw Is Nothing Then
        LogDebug MODULE_NAME, "Dispose: Liberando referencia COM"
        Set fw = Nothing
    End If

    ' 3. Liberar diccionario
    If Not mWatchedFolders Is Nothing Then
        Set mWatchedFolders = Nothing
    End If

    mInitialized = False
    mDisposed = True

    LogInfo MODULE_NAME, "Dispose: Limpieza completada"

    On Error GoTo 0
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
        LogWarning MODULE_NAME, "Ruta no existe, omitiendo monitoreo: " & folderPath
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

    LogInfo MODULE_NAME, "Monitoreo iniciado: " & folderPath & IIf(foldersOnly, " (solo subcarpetas)", "")

    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Error al iniciar monitoreo de: " & folderPath, Err.Number, Err.Description
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

    LogInfo MODULE_NAME, "Monitoreo detenido: " & folderPath
End Sub

'@Description: Detiene todos los monitoreos activos
'@Note: Este metodo es llamado por Dispose, usar Dispose para limpieza completa
Public Sub DetenerTodo()
    On Error Resume Next

    If mDisposed Then
        LogDebug MODULE_NAME, "DetenerTodo: Objeto ya liberado"
        Exit Sub
    End If

    If Not mWatchedFolders Is Nothing Then
        Dim carpeta As Variant
        Dim countStopped As Long
        countStopped = 0

        For Each carpeta In mWatchedFolders.Keys
            If mInitialized And Not fw Is Nothing Then
                fw.StopWatching CStr(carpeta)
                countStopped = countStopped + 1
            End If
        Next carpeta

        mWatchedFolders.RemoveAll
        LogInfo MODULE_NAME, "DetenerTodo: " & countStopped & " monitoreos detenidos"
    End If
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
    LogDebug MODULE_NAME, "Filtro de tamaño configurado: " & folderPath & " [" & minSize & "-" & maxSize & "]"

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando filtro tamaño: " & folderPath, Err.Number, Err.Description
End Sub

'@Description: Configura filtro por fecha de archivo
Public Sub ConfigurarFiltroFecha(ByVal folderPath As String, _
                                 ByVal compareDate As Date, _
                                 ByVal dateMode As DateCompareMode)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetFilter folderPath, FileFilterType.FileDate, 0, 0, compareDate, dateMode
    LogDebug MODULE_NAME, "Filtro de fecha configurado: " & folderPath

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando filtro fecha: " & folderPath, Err.Number, Err.Description
End Sub

'@Description: Configura filtro para monitorear solo carpetas
Public Sub ConfigurarFiltroSoloCarpetas(ByVal folderPath As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetFilter folderPath, FileFilterType.foldersOnly
    LogDebug MODULE_NAME, "Filtro solo carpetas configurado: " & folderPath

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando filtro carpetas: " & folderPath, Err.Number, Err.Description
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
    LogDebug MODULE_NAME, "Accion mover configurada: " & folderPath & " -> " & targetFolder

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando accion mover", Err.Number, Err.Description
End Sub

'@Description: Configura acción de copiar archivos automáticamente
Public Sub ConfigurarAccionCopiar(ByVal folderPath As String, _
                                  ByVal targetFolder As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetAutoAction folderPath, AutoActionType.CopyFile, targetFolder
    LogDebug MODULE_NAME, "Accion copiar configurada: " & folderPath & " -> " & targetFolder

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando accion copiar", Err.Number, Err.Description
End Sub

'@Description: Configura acción de archivar con timestamp
Public Sub ConfigurarAccionArchivar(ByVal folderPath As String, _
                                    ByVal archiveFolder As String)
    On Error GoTo ErrHandler

    If Not mInitialized Or mDisposed Then Exit Sub

    fw.SetAutoAction folderPath, AutoActionType.Archive, archiveFolder
    LogDebug MODULE_NAME, "Accion archivar configurada: " & folderPath & " -> " & archiveFolder

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Error configurando accion archivar", Err.Number, Err.Description
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
    LogDebug MODULE_NAME, "FileCreated: " & fileName & " en " & folder
    RaiseEvent FileCreated(folder, fileName)
End Sub

Private Sub fw_FileDeleted(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "FileDeleted: " & fileName & " en " & folder
    RaiseEvent FileDeleted(folder, fileName)
End Sub

Private Sub fw_FileChanged(ByVal folder As String, ByVal fileName As String)
    LogDebug MODULE_NAME, "FileChanged: " & fileName & " en " & folder
    RaiseEvent FileChanged(folder, fileName)
End Sub

Private Sub fw_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "FileRenamed: " & oldName & " -> " & newName & " en " & folder
    RaiseEvent FileRenamed(folder, oldName, newName)
End Sub

Private Sub fw_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat es muy frecuente, no logear para evitar ruido
    RaiseEvent Heartbeat(folder, lastUpdate)
End Sub

Private Sub fw_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    LogError MODULE_NAME, "Error COM en " & folder & ": " & errorMessage
    RaiseEvent ErrorOccurred(folder, errorMessage)
End Sub

' Manejadores para eventos de subcarpetas
Private Sub fw_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "SubfolderCreated: " & subfolderName & " en " & parentFolder
    RaiseEvent SubfolderCreated(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    LogDebug MODULE_NAME, "SubfolderDeleted: " & subfolderName & " en " & parentFolder
    RaiseEvent SubfolderDeleted(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    LogDebug MODULE_NAME, "SubfolderRenamed: " & oldName & " -> " & newName & " en " & parentFolder
    RaiseEvent SubfolderRenamed(parentFolder, oldName, newName)
End Sub


