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
Private mWatchedFolders As Object                ' Dictionary de carpetas monitoreadas

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
' INICIALIZACIÓN
' =====================================================

Private Sub Class_Initialize()
    Debug.Print "[clsFolderWatch] Inicializando wrapper VBA"
    
    On Error GoTo ErrHandler
    
    Set fw = New FolderWatcher
    Set mWatchedFolders = CreateObject("Scripting.Dictionary")
    mInitialized = True
    
    Debug.Print "[clsFolderWatch] Componente COM conectado: " & fw.GetConfiguration()
    Exit Sub
    
ErrHandler:
    Debug.Print "[clsFolderWatch] ERROR en inicialización: " & Err.Description
    mInitialized = False
End Sub

Private Sub Class_Terminate()
    Debug.Print "[clsFolderWatch] Finalizando wrapper VBA"
    DetenerTodo
    Set fw = Nothing
    Set mWatchedFolders = Nothing
End Sub

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
    
    If Not mInitialized Then
        Err.Raise vbObjectError + 600, "clsFolderWatch", "Componente COM no inicializado"
    End If
    
    ' Normalizar ruta
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' Verificar que la carpeta existe
    If Not RutaExiste(folderPath) Then
        Debug.Print "[clsFolderWatch] ADVERTENCIA: Ruta no existe: " & folderPath
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
    
    Debug.Print "[clsFolderWatch] Monitoreo iniciado: " & folderPath & IIf(foldersOnly, " (solo subcarpetas)", "")
    
    Exit Sub
    
ErrHandler:
    Debug.Print "[clsFolderWatch.IniciarMonitoreo] ERROR: " & Err.Description
    RaiseEvent ErrorOccurred(folderPath, Err.Description)
End Sub

'@Description: Detiene el monitoreo de una carpeta específica
Public Sub DetenerMonitoreo(ByVal folderPath As String)
    On Error Resume Next
    
    If mInitialized Then
        fw.StopWatching folderPath
    End If
    
    If mWatchedFolders.Exists(folderPath) Then
        mWatchedFolders.Remove folderPath
    End If
    
    Debug.Print "[clsFolderWatch] Monitoreo detenido: " & folderPath
End Sub

'@Description: Detiene todos los monitoreos activos
Public Sub DetenerTodo()
    On Error Resume Next
    
    If Not mWatchedFolders Is Nothing Then
        Dim carpeta As Variant
        For Each carpeta In mWatchedFolders.Keys
            If mInitialized Then
                fw.StopWatching CStr(carpeta)
            End If
        Next carpeta
        
        mWatchedFolders.RemoveAll
        Debug.Print "[clsFolderWatch] Todos los monitoreos detenidos"
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
    
    If Not mInitialized Then Exit Sub
    
    fw.SetFilter folderPath, FileFilterType.FileSize, minSize, maxSize
    Debug.Print "[clsFolderWatch] Filtro de tamaño configurado: " & folderPath
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarFiltroTamaño] ERROR: " & Err.Description
End Sub

'@Description: Configura filtro por fecha de archivo
Public Sub ConfigurarFiltroFecha(ByVal folderPath As String, _
                                 ByVal compareDate As Date, _
                                 ByVal dateMode As DateCompareMode)
    On Error GoTo ErrHandler
    
    If Not mInitialized Then Exit Sub
    
    fw.SetFilter folderPath, FileFilterType.FileDate, 0, 0, compareDate, dateMode
    Debug.Print "[clsFolderWatch] Filtro de fecha configurado: " & folderPath
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarFiltroFecha] ERROR: " & Err.Description
End Sub

'@Description: Configura filtro para monitorear solo carpetas
Public Sub ConfigurarFiltroSoloCarpetas(ByVal folderPath As String)
    On Error GoTo ErrHandler
    
    If Not mInitialized Then Exit Sub
    
    fw.SetFilter folderPath, FileFilterType.foldersOnly
    Debug.Print "[clsFolderWatch] Filtro de solo carpetas configurado: " & folderPath
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarFiltroSoloCarpetas] ERROR: " & Err.Description
End Sub

'@Description: Limpia todos los filtros de una carpeta
Public Sub LimpiarFiltros(ByVal folderPath As String)
    On Error Resume Next
    If mInitialized Then fw.ClearFilter folderPath
End Sub

' =====================================================
' CONFIGURACIÓN DE ACCIONES AUTOMÁTICAS
' =====================================================

'@Description: Configura acción de mover archivos automáticamente
Public Sub ConfigurarAccionMover(ByVal folderPath As String, _
                                 ByVal targetFolder As String)
    On Error GoTo ErrHandler
    
    If Not mInitialized Then Exit Sub
    
    fw.SetAutoAction folderPath, AutoActionType.MoveFile, targetFolder
    Debug.Print "[clsFolderWatch] Acción de mover configurada: " & folderPath & " -> " & targetFolder
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarAccionMover] ERROR: " & Err.Description
End Sub

'@Description: Configura acción de copiar archivos automáticamente
Public Sub ConfigurarAccionCopiar(ByVal folderPath As String, _
                                  ByVal targetFolder As String)
    On Error GoTo ErrHandler
    
    If Not mInitialized Then Exit Sub
    
    fw.SetAutoAction folderPath, AutoActionType.CopyFile, targetFolder
    Debug.Print "[clsFolderWatch] Acción de copiar configurada: " & folderPath & " -> " & targetFolder
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarAccionCopiar] ERROR: " & Err.Description
End Sub

'@Description: Configura acción de archivar con timestamp
Public Sub ConfigurarAccionArchivar(ByVal folderPath As String, _
                                    ByVal archiveFolder As String)
    On Error GoTo ErrHandler
    
    If Not mInitialized Then Exit Sub
    
    fw.SetAutoAction folderPath, AutoActionType.Archive, archiveFolder
    Debug.Print "[clsFolderWatch] Acción de archivar configurada: " & folderPath & " -> " & archiveFolder
    
    Exit Sub
ErrHandler:
    Debug.Print "[clsFolderWatch.ConfigurarAccionArchivar] ERROR: " & Err.Description
End Sub

'@Description: Limpia acciones automáticas de una carpeta
Public Sub LimpiarAcciones(ByVal folderPath As String)
    On Error Resume Next
    If mInitialized Then fw.ClearAutoAction folderPath
End Sub

' =====================================================
' PROPIEDADES Y CONSULTAS
' =====================================================

Public Property Get Configuracion() As String
    On Error Resume Next
    If mInitialized Then
        Configuracion = fw.GetConfiguration()
    Else
        Configuracion = "(no inicializado)"
    End If
End Property

Public Function CarpetasActivas() As Variant
    On Error Resume Next
    If mInitialized Then
        CarpetasActivas = fw.ActiveFolders
    Else
        CarpetasActivas = Array()
    End If
End Function

Public Function ObtenerHistorial(Optional ByVal lastMinutes As Long = 0, _
                                 Optional ByVal eventType As String = "") As Variant
    On Error Resume Next
    If mInitialized Then
        ObtenerHistorial = fw.GetEventHistory(lastMinutes, eventType)
    Else
        ObtenerHistorial = Array()
    End If
End Function

Public Function ObtenerEstadisticas(ByVal folderPath As String) As Variant
    On Error Resume Next
    If mInitialized Then
        ObtenerEstadisticas = fw.GetStatistics(folderPath)
    Else
        ObtenerEstadisticas = Array()
    End If
End Function

Public Sub LimpiarHistorial()
    On Error Resume Next
    If mInitialized Then fw.ClearHistory
End Sub

' =====================================================
' MANEJADORES DE EVENTOS COM (propagan a VBA)
' =====================================================

Private Sub fw_FileCreated(ByVal folder As String, ByVal fileName As String)
    RaiseEvent FileCreated(folder, fileName)
End Sub

Private Sub fw_FileDeleted(ByVal folder As String, ByVal fileName As String)
    RaiseEvent FileDeleted(folder, fileName)
End Sub

Private Sub fw_FileChanged(ByVal folder As String, ByVal fileName As String)
    RaiseEvent FileChanged(folder, fileName)
End Sub

Private Sub fw_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    RaiseEvent FileRenamed(folder, oldName, newName)
End Sub

Private Sub fw_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    RaiseEvent Heartbeat(folder, lastUpdate)
End Sub

Private Sub fw_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    RaiseEvent ErrorOccurred(folder, errorMessage)
End Sub

' Nuevos manejadores para eventos de subcarpetas
Private Sub fw_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    RaiseEvent SubfolderCreated(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    RaiseEvent SubfolderDeleted(parentFolder, subfolderName)
End Sub

Private Sub fw_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    RaiseEvent SubfolderRenamed(parentFolder, oldName, newName)
End Sub


