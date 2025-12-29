Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Timers
Imports System.Text.RegularExpressions

' =====================================================
' ENUMERACIONES PÚBLICAS
' =====================================================

<ComVisible(True)>
<Guid("D4F6A8C3-1111-4CDE-BC12-1234567890AB")>
Public Enum FileFilterType
    None = 0
    FileSize = 1
    FileDate = 2
    FileAttributes = 4
    FoldersOnly = 8
    FilesOnly = 16
End Enum

<ComVisible(True)>
<Guid("B1D9F7E1-AAAA-4CDE-BC12-123456789ABC")>
Public Enum AutoActionType
    None = 0
    MoveFile = 1
    CopyFile = 2
    Archive = 3
    Delete = 4
    LogOnly = 5
End Enum

<ComVisible(True)>
<Guid("C3E5F8B2-5678-4CDE-AB12-123456789ABD")>
Public Enum DateCompareMode
    CreatedAfter = 1
    CreatedBefore = 2
    ModifiedAfter = 3
    ModifiedBefore = 4
End Enum

' =====================================================
' INTERFACE COM PARA EVENTOS
' =====================================================

<ComVisible(True)>
<Guid("B1D9F7E1-AAAA-4CDE-BC12-1234567890AC")>
<InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
Public Interface IFolderWatcherEvents
    <DispId(1)> Sub FileCreated(ByVal folder As String, ByVal fileName As String)
    <DispId(2)> Sub FileDeleted(ByVal folder As String, ByVal fileName As String)
    <DispId(3)> Sub FileChanged(ByVal folder As String, ByVal fileName As String)
    <DispId(4)> Sub FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    <DispId(5)> Sub Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    <DispId(6)> Sub ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    <DispId(7)> Sub SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    <DispId(8)> Sub SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    <DispId(9)> Sub SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
End Interface

' =====================================================
' CLASE PRINCIPAL FOLDERWATCHER
' =====================================================

<ComVisible(True)>
<Guid("C3E5F8B2-5678-4CDE-AB12-1234567890AD")>
<ClassInterface(ClassInterfaceType.AutoDual)>
<ComSourceInterfaces(GetType(IFolderWatcherEvents))>
<ProgId("FolderWatcher.Monitor")>
Public Class FolderWatcher
    Implements IDisposable

    ' Diccionarios principales
    Private watchers As New Dictionary(Of String, FileSystemWatcher)
    Private watcherSettings As New Dictionary(Of String, WatcherConfig)
    Private heartbeats As New Dictionary(Of String, Date)
    Private debounceDict As New Dictionary(Of String, Dictionary(Of String, DateTime))

    ' Cola para reintentos de watchers con errores
    Private foldersToRetry As New System.Collections.Concurrent.ConcurrentQueue(Of String)

    ' Filtros y acciones por carpeta
    Private filters As New Dictionary(Of String, FileFilter)
    Private actions As New Dictionary(Of String, AutoAction)

    ' Historial y estadísticas
    Private eventHistory As New Queue(Of FileEventInfo)
    Private Const MAX_HISTORY_SIZE As Integer = 1000
    Private stats As New Dictionary(Of String, FolderStatistics)

    ' Timer y configuración
    Private hbTimer As Timer
    Private disposed As Boolean = False
    Private logWriter As StreamWriter = Nothing

    ' Configuración desde Settings
    Private ReadOnly HEARTBEAT_INTERVAL As Double
    Private ReadOnly DEFAULT_INACTIVITY_MINUTES As Double
    Private ReadOnly DEBOUNCE_SECONDS As Double
    Private ReadOnly ENABLE_AUTO_RESTART As Boolean
    Private ReadOnly MAX_CONCURRENT_WATCHERS As Integer
    Private ReadOnly ENABLE_LOGGING As Boolean
    Private debounceInterval As TimeSpan

    ' =====================================================
    ' CONSTRUCTOR
    ' =====================================================

    Public Sub New()
        Try
            HEARTBEAT_INTERVAL = My.Settings.DefaultHeartbeatInterval
            DEFAULT_INACTIVITY_MINUTES = My.Settings.DefaultInactivityMinutes
            DEBOUNCE_SECONDS = My.Settings.DefaultDebounceSeconds
            ENABLE_AUTO_RESTART = My.Settings.EnableAutoRestart
            MAX_CONCURRENT_WATCHERS = My.Settings.MaxConcurrentWatchers
            ENABLE_LOGGING = My.Settings.EnableDetailedLogging
        Catch ex As Exception
            HEARTBEAT_INTERVAL = 60000
            DEFAULT_INACTIVITY_MINUTES = 10.0
            DEBOUNCE_SECONDS = 1.0
            ENABLE_AUTO_RESTART = True
            MAX_CONCURRENT_WATCHERS = 10
            ENABLE_LOGGING = False
        End Try

        debounceInterval = TimeSpan.FromSeconds(DEBOUNCE_SECONDS)

        If ENABLE_LOGGING AndAlso Not String.IsNullOrEmpty(My.Settings.LogPath) Then
            Try
                Dim logPath As String = My.Settings.LogPath
                Dim logDir As String = Path.GetDirectoryName(logPath)
                If Not String.IsNullOrEmpty(logDir) AndAlso Not Directory.Exists(logDir) Then
                    Directory.CreateDirectory(logDir)
                End If
                logWriter = New StreamWriter(logPath, True)
                logWriter.AutoFlush = True
                WriteLog("FolderWatcher iniciado")
            Catch ex As Exception
            End Try
        End If

        hbTimer = New Timer(HEARTBEAT_INTERVAL)
        AddHandler hbTimer.Elapsed, AddressOf OnHeartbeat
        hbTimer.Start()
    End Sub

    ' =====================================================
    ' EVENTOS PÚBLICOS
    ' =====================================================

    Public Event FileCreated(folder As String, fileName As String)
    Public Event FileDeleted(folder As String, fileName As String)
    Public Event FileChanged(folder As String, fileName As String)
    Public Event FileRenamed(folder As String, oldName As String, newName As String)
    Public Event Heartbeat(folder As String, lastUpdate As Date)
    Public Event ErrorOccurred(folder As String, errorMessage As String)
    Public Event SubfolderCreated(parentFolder As String, subfolderName As String)
    Public Event SubfolderDeleted(parentFolder As String, subfolderName As String)
    Public Event SubfolderRenamed(parentFolder As String, oldName As String, newName As String)

    ' =====================================================
    ' MÉTODOS PÚBLICOS - CONFIGURACIÓN
    ' =====================================================

    Public Function GetConfiguration() As String
        Return String.Format(
            "HeartbeatInterval={0}ms, InactivityMinutes={1}, DebounceSeconds={2}, AutoRestart={3}, MaxWatchers={4}, Logging={5}",
            HEARTBEAT_INTERVAL,
            DEFAULT_INACTIVITY_MINUTES,
            DEBOUNCE_SECONDS,
            ENABLE_AUTO_RESTART,
            MAX_CONCURRENT_WATCHERS,
            ENABLE_LOGGING
        )
    End Function

    Public ReadOnly Property ActiveFolders() As Object
        Get
            Return watchers.Keys.ToArray()
        End Get
    End Property

    Public Function LastHeartbeat(folderPath As String) As Date
        If heartbeats.ContainsKey(folderPath) Then Return heartbeats(folderPath)
        Return New DateTime(2000, 1, 1)
    End Function

    Public ReadOnly Property EventCount() As Integer
        Get
            SyncLock eventHistory
                Return eventHistory.Count
            End SyncLock
        End Get
    End Property

    ' =====================================================
    ' WATCHFOLDER - MÉTODO PRINCIPAL
    ' =====================================================

    Public Sub WatchFolder(folderPath As String,
                           Optional includeSubdirs As Boolean = True,
                           Optional filterPattern As String = "*.*",
                           Optional eventsToWatch As Object = Nothing,
                           Optional inactivityMinutes As Double = -1,
                           Optional foldersOnly As Boolean = False)

        Dim errMsg As String = ""

        If watchers.Count >= MAX_CONCURRENT_WATCHERS Then
            errMsg = String.Format("Límite de watchers alcanzado ({0})", MAX_CONCURRENT_WATCHERS)
            If ENABLE_LOGGING Then
                WriteLog(String.Format("ERROR: {0} - Carpeta: {1}", errMsg, folderPath))
            End If
            RaiseEvent ErrorOccurred(folderPath, errMsg)
            Throw New InvalidOperationException(errMsg)
        End If

        If watchers.ContainsKey(folderPath) Then
            WriteLog(String.Format("ERROR: {0} - Carpeta: {1}", errMsg, folderPath))
            WriteLog(String.Format("ADVERTENCIA: Carpeta ya está siendo monitoreada: {0}", folderPath))
            Exit Sub
        End If

        If Not Directory.Exists(folderPath) Then
            errMsg = String.Format("Carpeta no encontrada: {0}", folderPath)
            WriteLog(String.Format("ERROR: {0}", errMsg))
            RaiseEvent ErrorOccurred(folderPath, errMsg)
            Throw New DirectoryNotFoundException(errMsg)
        End If

        ' Detectar si es ruta de red
        If IsNetworkPath(folderPath) Then
            WriteLog(String.Format("ADVERTENCIA: Ruta de red detectada: {0} - Rendimiento reducido", folderPath))

            ' Verificar accesibilidad real (no solo Exists)
            If Not TestNetworkPathAccess(folderPath) Then
                errMsg = String.Format("Ruta de red no accesible o sin permisos: {0}", folderPath)
                WriteLog(String.Format("ERROR: {0}", errMsg))
                RaiseEvent ErrorOccurred(folderPath, errMsg)
                Throw New UnauthorizedAccessException(errMsg)
            End If
        End If

        WriteLog(String.Format("Iniciando monitoreo: {0} (FoldersOnly={1})", folderPath, foldersOnly))

        If inactivityMinutes < 0 Then
            inactivityMinutes = DEFAULT_INACTIVITY_MINUTES
        End If

        Dim fsw As New FileSystemWatcher(folderPath)
        fsw.IncludeSubdirectories = includeSubdirs

        ' Aumentar buffer para rutas de red (reduce eventos perdidos)
        If IsNetworkPath(folderPath) Then
            fsw.InternalBufferSize = 65536  ' 64KB (máximo permitido)
            WriteLog(String.Format("Buffer aumentado a 64KB para ruta de red: {0}", folderPath))
        Else
            fsw.InternalBufferSize = 8192   ' 8KB (default) para rutas locales
        End If

        ' CLAVE: NotifyFilter según modo (compatible con fw.vbs)
        If foldersOnly Then
            fsw.NotifyFilter = NotifyFilters.DirectoryName
        Else
            fsw.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite Or NotifyFilters.CreationTime
        End If

        fsw.EnableRaisingEvents = True

        AddHandler fsw.Error, Sub(s As Object, e As ErrorEventArgs)
                                        Dim ex As Exception = e.GetException()
                                        Dim errMsgInterno As String = String.Format("Error en FileSystemWatcher: {0}", ex.Message)

                                        ' Detectar errores de red específicos
                                        If TypeOf ex Is IOException AndAlso IsNetworkPath(folderPath) Then
                                            errMsgInterno = String.Format("Error de red detectado: {0}. Programando reintento...", ex.Message)
                                            WriteLog(errMsgInterno)

                                            ' Encolar para reinicio en el siguiente heartbeat
                                            If ENABLE_AUTO_RESTART Then
                                                foldersToRetry.Enqueue(folderPath)
                                            End If
                                        End If

                                        WriteLog(String.Format("ERROR: {0} - Carpeta: {1}", errMsgInterno, folderPath))
                                        UpdateStatistics(folderPath, "Error")
                                        RaiseEvent ErrorOccurred(folderPath, errMsgInterno)
                                    End Sub

        Dim patterns() As String = filterPattern.Split(";"c)

        Dim FileMatches As Func(Of String, Boolean) = Function(name As String)
                                                          For Each p As String In patterns
                                                              If WildcardMatch(name, p.Trim()) Then Return True
                                                          Next
                                                          Return False
                                                      End Function

        Dim evs As New Dictionary(Of String, Boolean)
        If eventsToWatch Is Nothing OrElse Not IsArray(eventsToWatch) Then
            evs("Created") = True
            evs("Deleted") = True
            evs("Changed") = True
            evs("Renamed") = True
        Else
            Dim arr As Array = CType(eventsToWatch, Array)
            For i As Integer = 0 To arr.Length - 1
                Dim eventName As String = arr.GetValue(i).ToString()
                evs(eventName) = True
            Next
        End If

        ' Event handlers
        If evs.ContainsKey("Created") Then
            AddHandler fsw.Created, Sub(s As Object, e As FileSystemEventArgs)
                                        If foldersOnly Then
                                            If Directory.Exists(e.FullPath) Then
                                                HandleSubfolderEvent("Created", folderPath, e.Name)
                                            End If
                                        ElseIf FileMatches(e.Name) AndAlso PassesFilter(folderPath, e.FullPath) Then
                                            RaiseEventIfNotDebounced("Created", folderPath, e.Name, e.FullPath)
                                        End If
                                    End Sub
        End If

        If evs.ContainsKey("Deleted") Then
            AddHandler fsw.Deleted, Sub(s As Object, e As FileSystemEventArgs)
                                        If foldersOnly Then
                                            HandleSubfolderEvent("Deleted", folderPath, e.Name)
                                        ElseIf FileMatches(e.Name) Then
                                            RaiseEventIfNotDebounced("Deleted", folderPath, e.Name, e.FullPath)
                                        End If
                                    End Sub
        End If

        If evs.ContainsKey("Changed") Then
            AddHandler fsw.Changed, Sub(s As Object, e As FileSystemEventArgs)
                                        If Not foldersOnly AndAlso FileMatches(e.Name) AndAlso PassesFilter(folderPath, e.FullPath) Then
                                            RaiseEventIfNotDebounced("Changed", folderPath, e.Name, e.FullPath)
                                        End If
                                    End Sub
        End If

        If evs.ContainsKey("Renamed") Then
            AddHandler fsw.Renamed, Sub(s As Object, e As RenamedEventArgs)
                                        If foldersOnly Then
                                            If Directory.Exists(e.FullPath) Then
                                                HandleSubfolderRenamed(folderPath, e.OldName, e.Name)
                                            End If
                                        ElseIf FileMatches(e.Name) Then
                                            WriteLog(String.Format("Archivo renombrado: {0} -> {1} en {2}", e.OldName, e.Name, folderPath))
                                            AddToHistory("Renamed", folderPath, e.Name, e.OldName)
                                            UpdateStatistics(folderPath, "Renamed")
                                            RaiseEvent FileRenamed(folderPath, e.OldName, e.Name)
                                            heartbeats(folderPath) = DateTime.Now
                                        End If
                                    End Sub
        End If

        watchers(folderPath) = fsw

        Dim config As New WatcherConfig With {
            .InactivityMinutes = inactivityMinutes,
            .FoldersOnly = foldersOnly,
            .IncludeSubdirs = includeSubdirs,
            .FilterPattern = filterPattern
        }
        watcherSettings(folderPath) = config

        heartbeats(folderPath) = DateTime.Now
        debounceDict(folderPath) = New Dictionary(Of String, DateTime)

        If Not stats.ContainsKey(folderPath) Then
            stats(folderPath) = New FolderStatistics With {.FolderPath = folderPath}
        End If

        WriteLog(String.Format("Monitoreo iniciado exitosamente: {0}", folderPath))
    End Sub

    Public Sub StopWatching(folderPath As String)
        If watchers.ContainsKey(folderPath) Then
            WriteLog(String.Format("Deteniendo monitoreo: {0}", folderPath))
            watchers(folderPath).Dispose()
            watchers.Remove(folderPath)
            watcherSettings.Remove(folderPath)
            heartbeats.Remove(folderPath)
            debounceDict.Remove(folderPath)
            filters.Remove(folderPath)
            actions.Remove(folderPath)
        End If
    End Sub

    ' =====================================================
    ' FILTROS AVANZADOS
    ' =====================================================

    Public Sub SetFilter(folderPath As String,
                        filterType As FileFilterType,
                        Optional minSize As Long = 0,
                        Optional maxSize As Long = Long.MaxValue,
                        Optional compareDate As Date = Nothing,
                        Optional dateMode As DateCompareMode = DateCompareMode.ModifiedAfter,
                        Optional attributeMask As FileAttributes = FileAttributes.Normal)

        If Not filters.ContainsKey(folderPath) Then
            filters(folderPath) = New FileFilter()
        End If

        Dim f = filters(folderPath)
        f.FilterType = filterType
        f.MinSize = minSize
        f.MaxSize = maxSize
        f.CompareDate = If(compareDate = Nothing, DateTime.MinValue, compareDate)
        f.DateMode = dateMode
        f.AttributeMask = attributeMask

        WriteLog(String.Format("Filtro configurado para {0}: {1}", folderPath, filterType.ToString()))
    End Sub

    Public Sub ClearFilter(folderPath As String)
        If filters.ContainsKey(folderPath) Then
            filters.Remove(folderPath)
            WriteLog(String.Format("Filtro eliminado: {0}", folderPath))
        End If
    End Sub

    Private Function PassesFilter(folderPath As String, fullPath As String) As Boolean
        If Not filters.ContainsKey(folderPath) Then Return True

        Dim f = filters(folderPath)

        Try
            Dim attrs As FileAttributes = File.GetAttributes(fullPath)
            Dim isFolder As Boolean = (attrs And FileAttributes.Directory) = FileAttributes.Directory
            Dim isFile As Boolean = Not isFolder
            'Dim isFolder As Boolean = Directory.Exists(fullPath)
            'Dim isFile As Boolean = File.Exists(fullPath)

            If (f.FilterType And FileFilterType.FoldersOnly) = FileFilterType.FoldersOnly Then
                If Not isFolder Then Return False
            End If

            If (f.FilterType And FileFilterType.FilesOnly) = FileFilterType.FilesOnly Then
                If Not isFile Then Return False
            End If

            If isFile Then
                Dim fi As New FileInfo(fullPath)

                If (f.FilterType And FileFilterType.FileSize) = FileFilterType.FileSize Then
                    If fi.Length < f.MinSize OrElse fi.Length > f.MaxSize Then
                        Return False
                    End If
                End If

                If (f.FilterType And FileFilterType.FileDate) = FileFilterType.FileDate Then
                    Dim fileDate As DateTime
                    Select Case f.DateMode
                        Case DateCompareMode.CreatedAfter
                            fileDate = fi.CreationTime
                            If fileDate < f.CompareDate Then Return False
                        Case DateCompareMode.CreatedBefore
                            fileDate = fi.CreationTime
                            If fileDate > f.CompareDate Then Return False
                        Case DateCompareMode.ModifiedAfter
                            fileDate = fi.LastWriteTime
                            If fileDate < f.CompareDate Then Return False
                        Case DateCompareMode.ModifiedBefore
                            fileDate = fi.LastWriteTime
                            If fileDate > f.CompareDate Then Return False
                    End Select
                End If

                If (f.FilterType And FileFilterType.FileAttributes) = FileFilterType.FileAttributes Then
                    If (fi.Attributes And f.AttributeMask) = 0 Then
                        Return False
                    End If
                End If
            End If

        Catch ex As Exception
            WriteLog(String.Format("Error al aplicar filtro: {0}", ex.Message))
            Return True
        End Try

        Return True
    End Function

    ' =====================================================
    ' ACCIONES AUTOMÁTICAS
    ' =====================================================

    Public Sub SetAutoAction(folderPath As String,
                            actionType As AutoActionType,
                            Optional targetFolder As String = "")

        If Not actions.ContainsKey(folderPath) Then
            actions(folderPath) = New AutoAction()
        End If

        Dim a = actions(folderPath)
        a.ActionType = actionType
        a.TargetFolder = targetFolder

        WriteLog(String.Format("Acción automática configurada para {0}: {1}", folderPath, actionType.ToString()))
    End Sub

    Public Sub ClearAutoAction(folderPath As String)
        If actions.ContainsKey(folderPath) Then
            actions.Remove(folderPath)
            WriteLog(String.Format("Acción automática eliminada: {0}", folderPath))
        End If
    End Sub

    Private Sub ExecuteAutoAction(folderPath As String, fileName As String, fullPath As String)
        If Not actions.ContainsKey(folderPath) Then Exit Sub

        Dim a = actions(folderPath)

        Try
            Select Case a.ActionType
                Case AutoActionType.MoveFile
                    If Not String.IsNullOrEmpty(a.TargetFolder) AndAlso File.Exists(fullPath) Then
                        If Not Directory.Exists(a.TargetFolder) Then
                            Directory.CreateDirectory(a.TargetFolder)
                        End If
                        Dim targetPath As String = Path.Combine(a.TargetFolder, fileName)
                        File.Move(fullPath, targetPath)
                        WriteLog(String.Format("Archivo movido: {0} -> {1}", fullPath, targetPath))
                    End If

                Case AutoActionType.CopyFile
                    If Not String.IsNullOrEmpty(a.TargetFolder) AndAlso File.Exists(fullPath) Then
                        If Not Directory.Exists(a.TargetFolder) Then
                            Directory.CreateDirectory(a.TargetFolder)
                        End If
                        Dim targetPath As String = Path.Combine(a.TargetFolder, fileName)
                        File.Copy(fullPath, targetPath, True)
                        WriteLog(String.Format("Archivo copiado: {0} -> {1}", fullPath, targetPath))
                    End If

                Case AutoActionType.Archive
                    If Not String.IsNullOrEmpty(a.TargetFolder) AndAlso File.Exists(fullPath) Then
                        If Not Directory.Exists(a.TargetFolder) Then
                            Directory.CreateDirectory(a.TargetFolder)
                        End If
                        Dim dateStamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                        Dim newName As String = String.Format("{0}_{1}{2}",
                            Path.GetFileNameWithoutExtension(fileName),
                            dateStamp,
                            Path.GetExtension(fileName))
                        Dim targetPath As String = Path.Combine(a.TargetFolder, newName)
                        File.Move(fullPath, targetPath)
                        WriteLog(String.Format("Archivo archivado: {0} -> {1}", fullPath, targetPath))
                    End If

                Case AutoActionType.Delete
                    If File.Exists(fullPath) Then
                        File.Delete(fullPath)
                        WriteLog(String.Format("Archivo eliminado: {0}", fullPath))
                    End If

                Case AutoActionType.LogOnly
                    WriteLog(String.Format("Archivo detectado (solo log): {0}", fullPath))
            End Select

        Catch ex As Exception
            Dim errMsg As String = String.Format("Error ejecutando acción: {0}", ex.Message)
            WriteLog(String.Format("ERROR: {0}", errMsg))
            RaiseEvent ErrorOccurred(folderPath, errMsg)
        End Try
    End Sub

    ' =====================================================
    ' MANEJO DE SUBCARPETAS (equivalente a fw.vbs)
    ' =====================================================

    Private Sub HandleSubfolderEvent(eventType As String, parentFolder As String, subfolderName As String)
        WriteLog(String.Format("Subcarpeta {0}: {1} en {2}", eventType, subfolderName, parentFolder))

        AddToHistory(eventType, parentFolder, subfolderName)
        UpdateStatistics(parentFolder, eventType)

        Select Case eventType
            Case "Created"
                RaiseEvent SubfolderCreated(parentFolder, subfolderName)
            Case "Deleted"
                RaiseEvent SubfolderDeleted(parentFolder, subfolderName)
        End Select

        heartbeats(parentFolder) = DateTime.Now
    End Sub

    Private Sub HandleSubfolderRenamed(parentFolder As String, oldName As String, newName As String)
        WriteLog(String.Format("Subcarpeta renombrada: {0} -> {1} en {2}", oldName, newName, parentFolder))

        AddToHistory("Renamed", parentFolder, newName, oldName)
        UpdateStatistics(parentFolder, "Renamed")

        RaiseEvent SubfolderRenamed(parentFolder, oldName, newName)
        heartbeats(parentFolder) = DateTime.Now
    End Sub

    ' =====================================================
    ' HISTORIAL DE EVENTOS
    ' =====================================================

    Private Sub AddToHistory(eventType As String, folder As String, fileName As String, Optional oldName As String = "")
        Dim evt As New FileEventInfo With {
            .EventType = eventType,
            .Folder = folder,
            .FileName = fileName,
            .OldName = oldName,
            .Timestamp = DateTime.Now
        }

        Try
            Dim fullPath As String = Path.Combine(folder, fileName)
            If File.Exists(fullPath) Then
                evt.FileSize = New FileInfo(fullPath).Length
            End If
        Catch
        End Try

        SyncLock eventHistory
            If eventHistory.Count >= MAX_HISTORY_SIZE Then
                eventHistory.Dequeue()
            End If
            eventHistory.Enqueue(evt)
        End SyncLock
    End Sub

    Public Function GetEventHistory(Optional lastMinutes As Integer = 0, Optional eventType As String = "") As Object
        SyncLock eventHistory
            Dim filtered = eventHistory.AsEnumerable()

            If lastMinutes > 0 Then
                Dim cutoff As DateTime = DateTime.Now.AddMinutes(-lastMinutes)
                filtered = filtered.Where(Function(e) e.Timestamp >= cutoff)
            End If

            If Not String.IsNullOrEmpty(eventType) Then
                filtered = filtered.Where(Function(e) e.EventType = eventType)
            End If

            Dim result = filtered.ToArray()

            If result.Length = 0 Then
                Return New Object(0, 4) {}
            End If

            Dim arr(result.Length - 1, 4) As Object

            For i As Integer = 0 To result.Length - 1
                arr(i, 0) = result(i).Timestamp
                arr(i, 1) = result(i).EventType
                arr(i, 2) = result(i).Folder
                arr(i, 3) = result(i).FileName
                arr(i, 4) = result(i).FileSize
            Next

            Return arr
        End SyncLock
    End Function

    Public Sub ClearHistory()
        SyncLock eventHistory
            eventHistory.Clear()
            WriteLog("Historial de eventos limpiado")
        End SyncLock
    End Sub

    ' =====================================================
    ' ESTADÍSTICAS
    ' =====================================================

    Private Sub UpdateStatistics(folder As String, eventType As String)
        If Not stats.ContainsKey(folder) Then
            stats(folder) = New FolderStatistics With {.FolderPath = folder}
        End If

        Dim s = stats(folder)
        s.TotalEvents += 1
        s.LastEventTime = DateTime.Now

        Select Case eventType
            Case "Created"
                s.CreatedCount += 1
            Case "Deleted"
                s.DeletedCount += 1
            Case "Changed"
                s.ChangedCount += 1
            Case "Renamed"
                s.RenamedCount += 1
            Case "Error"
                s.ErrorCount += 1
        End Select
    End Sub

    Public Function GetStatistics(folderPath As String) As Object
        If Not stats.ContainsKey(folderPath) Then
            Return New Object() {"Carpeta no encontrada", 0, 0, 0, 0, 0, DateTime.MinValue, 0.0, 0}
        End If

        Dim s = stats(folderPath)

        If heartbeats.ContainsKey(folderPath) Then
            Dim startTime As DateTime = heartbeats(folderPath)
            Dim hours As Double = (DateTime.Now - startTime).TotalHours
            If hours > 0 Then
                s.AverageEventsPerHour = Math.Round(s.TotalEvents / hours, 2)
            End If
        End If

        Return New Object() {
            s.FolderPath,
            s.TotalEvents,
            s.CreatedCount,
            s.DeletedCount,
            s.ChangedCount,
            s.RenamedCount,
            s.LastEventTime,
            s.AverageEventsPerHour,
            s.ErrorCount
        }
    End Function

    Public Sub ResetStatistics(folderPath As String)
        If stats.ContainsKey(folderPath) Then
            stats(folderPath) = New FolderStatistics With {.FolderPath = folderPath}
            WriteLog(String.Format("Estadísticas reseteadas: {0}", folderPath))
        End If
    End Sub

    Public Function GetAllStatistics() As Object
        If stats.Count = 0 Then
            Return New Object(0, 8) {}
        End If

        Dim arr(stats.Count - 1, 8) As Object
        Dim i As Integer = 0

        For Each kvp In stats
            Dim s = kvp.Value

            If heartbeats.ContainsKey(kvp.Key) Then
                Dim startTime As DateTime = heartbeats(kvp.Key)
                Dim hours As Double = (DateTime.Now - startTime).TotalHours
                If hours > 0 Then
                    s.AverageEventsPerHour = Math.Round(s.TotalEvents / hours, 2)
                End If
            End If

            arr(i, 0) = s.FolderPath
            arr(i, 1) = s.TotalEvents
            arr(i, 2) = s.CreatedCount
            arr(i, 3) = s.DeletedCount
            arr(i, 4) = s.ChangedCount
            arr(i, 5) = s.RenamedCount
            arr(i, 6) = s.LastEventTime
            arr(i, 7) = s.AverageEventsPerHour
            arr(i, 8) = s.ErrorCount
            i += 1
        Next

        Return arr
    End Function

    ' =====================================================
    ' HEARTBEAT Y REINICIO AUTOMÁTICO
    ' =====================================================

    Private Sub OnHeartbeat(sender As Object, e As ElapsedEventArgs)
        Dim foldersToRestart As New List(Of String)

        ' 1. Procesar cola de reintentos por errores
        Dim retryFolder As String = Nothing
        While foldersToRetry.TryDequeue(retryFolder)
            If Not foldersToRestart.Contains(retryFolder) Then
                foldersToRestart.Add(retryFolder)
                WriteLog(String.Format("Reintento por error programado: {0}", retryFolder))
            End If
        End While

        ' 2. Detectar inactividad normal
        For Each folder As String In watchers.Keys.ToList()
            Dim lastUpdate As DateTime = heartbeats(folder)

            ' Timeout diferente para rutas de red

            RaiseEvent Heartbeat(folder, lastUpdate)

            Dim inactivityLimit As Double = DEFAULT_INACTIVITY_MINUTES
            If watcherSettings.ContainsKey(folder) Then
                inactivityLimit = watcherSettings(folder).InactivityMinutes
            End If

            ' Timeout diferente para rutas de red
            Dim timeoutMinutes As Double = inactivityLimit
            If IsNetworkPath(folder) Then
                timeoutMinutes = My.Settings.NetworkPathTimeoutSeconds / 60.0
            End If
            If ENABLE_AUTO_RESTART AndAlso (DateTime.Now - lastUpdate).TotalMinutes > inactivityLimit Then
                If Not foldersToRestart.Contains(folder) Then
                    WriteLog(String.Format("ADVERTENCIA: Inactividad detectada en {0} - Programando reinicio", folder))
                    foldersToRestart.Add(folder)
                End If
            End If
        Next

        ' 3. Reiniciar todos los watchers marcados
        For Each folder As String In foldersToRestart
            Try
                If watcherSettings.ContainsKey(folder) Then
                    Dim config = watcherSettings(folder)
                    WriteLog(String.Format("Reiniciando watcher: {0}", folder))
                    StopWatching(folder)
                    WatchFolder(folder, config.IncludeSubdirs, config.FilterPattern, Nothing, config.InactivityMinutes, config.FoldersOnly)
                End If
            Catch ex As Exception
                WriteLog(String.Format("ERROR al reiniciar {0}: {1}", folder, ex.Message))
                RaiseEvent ErrorOccurred(folder, String.Format("Error al reiniciar: {0}", ex.Message))
            End Try
        Next
    End Sub

    ' =====================================================
    ' MÉTODOS DE UTILIDAD
    ' =====================================================

    Private Sub WriteLog(message As String)
        If logWriter IsNot Nothing Then
            Try
                logWriter.WriteLine(String.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}", DateTime.Now, message))
            Catch
            End Try
        End If
    End Sub

    Private Function WildcardMatch(text As String, pattern As String) As Boolean
        If String.IsNullOrEmpty(pattern) Then Return False
        If pattern = "*.*" OrElse pattern = "*" Then Return True

        Dim regexPattern As String = "^" & Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".") & "$"
        Return Regex.IsMatch(text, regexPattern, RegexOptions.IgnoreCase)
    End Function

    Private Sub RaiseEventIfNotDebounced(eventType As String, folder As String, fileName As String, fullPath As String)
        Dim key As String = eventType & "|" & fileName
        Dim lastTime As DateTime = DateTime.MinValue

        If debounceDict(folder).ContainsKey(key) Then
            lastTime = debounceDict(folder)(key)
            If (DateTime.Now - lastTime) < debounceInterval Then Exit Sub
        End If

        debounceDict(folder)(key) = DateTime.Now

        WriteLog(String.Format("Evento {0}: {1} en {2}", eventType, fileName, folder))

        AddToHistory(eventType, folder, fileName)
        UpdateStatistics(folder, eventType)

        ExecuteAutoAction(folder, fileName, fullPath)

        Select Case eventType
            Case "Created"
                RaiseEvent FileCreated(folder, fileName)
            Case "Deleted"
                RaiseEvent FileDeleted(folder, fileName)
            Case "Changed"
                RaiseEvent FileChanged(folder, fileName)
        End Select

        heartbeats(folder) = DateTime.Now
    End Sub

    ' =====================================================
    ' SOPORTE DE RUTAS DE RED
    ' =====================================================

    Private Function IsNetworkPath(strpath As String) As Boolean
        ' UNC paths: \\servidor\...
        If strpath.StartsWith("\\") Then Return True

        ' Mapped drives: Verificar si es unidad de red
        Try
            Dim drive As String = Path.GetPathRoot(strpath)
            Dim driveInfo As New System.IO.DriveInfo(drive)
            Return driveInfo.DriveType = DriveType.Network
        Catch
            Return False
        End Try
    End Function

    Private Function TestNetworkPathAccess(strpath As String) As Boolean
        Try
            ' Intentar crear/eliminar archivo de test
            Dim testFile As String = Path.Combine(strpath, ".fwtest_" & DateTime.Now.Ticks.ToString())
            File.WriteAllText(testFile, "test")
            File.Delete(testFile)
            Return True
        Catch ex As UnauthorizedAccessException
            WriteLog(String.Format("Sin permisos de escritura en: {0}", strpath))
            Return False
        Catch ex As IOException
            WriteLog(String.Format("Error de E/S en red: {0} - {1}", strpath, ex.Message))
            Return False
        Catch ex As Exception
            WriteLog(String.Format("Error de acceso a red: {0} - {1}", strpath, ex.Message))
            Return False
        End Try
    End Function

    ' =====================================================
    ' IDISPOSABLE
    ' =====================================================

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposed Then
            If disposing Then
                WriteLog("Finalizando FolderWatcher")

                If hbTimer IsNot Nothing Then
                    hbTimer.Stop()
                    hbTimer.Dispose()
                End If

                For Each watcher In watchers.Values
                    watcher.Dispose()
                Next

                watchers.Clear()
                heartbeats.Clear()
                debounceDict.Clear()
                watcherSettings.Clear()
                stats.Clear()
                filters.Clear()
                actions.Clear()

                SyncLock eventHistory
                    eventHistory.Clear()
                End SyncLock

                If logWriter IsNot Nothing Then
                    logWriter.Close()
                    logWriter.Dispose()
                End If
            End If
            disposed = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

End Class
' =====================================================
' CLASES AUXILIARES (NO VISIBLES DESDE COM)
' =====================================================
<ComVisible(False)>
Public Class FileEventInfo
    Public Property EventType As String
    Public Property Folder As String
    Public Property FileName As String
    Public Property OldName As String = ""
    Public Property Timestamp As DateTime
    Public Property FileSize As Long = 0
End Class
<ComVisible(False)>
Public Class FolderStatistics
    Public Property FolderPath As String
    Public Property TotalEvents As Long
    Public Property CreatedCount As Long
    Public Property DeletedCount As Long
    Public Property ChangedCount As Long
    Public Property RenamedCount As Long
    Public Property LastEventTime As DateTime
    Public Property AverageEventsPerHour As Double
    Public Property ErrorCount As Long
End Class
<ComVisible(False)>
Public Class WatcherConfig
    Public Property InactivityMinutes As Double
    Public Property FoldersOnly As Boolean
    Public Property IncludeSubdirs As Boolean
    Public Property FilterPattern As String
End Class
<ComVisible(False)>
Public Class FileFilter
    Public Property FilterType As FileFilterType
    Public Property MinSize As Long
    Public Property MaxSize As Long
    Public Property CompareDate As DateTime
    Public Property DateMode As DateCompareMode
    Public Property AttributeMask As FileAttributes
End Class
<ComVisible(False)>
Public Class AutoAction
    Public Property ActionType As AutoActionType
    Public Property TargetFolder As String
End Class
