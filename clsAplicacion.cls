VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Control de estado"
' gestion de la aplicacion 'gestor de oportunidades ABC'
Option Explicit

Private m_bChartActive As Boolean

' lo siguiente es para controlar **SUCESOS EN LA COLECCION clsOpportunitiesMgr QUE HACEN
' QUE TENGA QUE INTERVENIR clsAplicacion (para coordinar con clsFileManager, o con el Ribbon, ...)
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1
Private WithEvents mChartManager As clsChartEventsManager
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mFolderWatcher As clsFolderWatch
Attribute mFolderWatcher.VB_VarHelpID = -1
Private WithEvents xlApp As Application
Attribute xlApp.VB_VarHelpID = -1

Private WithEvents mRibbonState As clsRibbonState
Attribute mRibbonState.VB_VarHelpID = -1
Private WithEvents evRibbon As clsRibbonEvents
Attribute evRibbon.VB_VarHelpID = -1

Private mConfiguration As clsConfiguration
Private mFileMgr As clsFileManager

Public Property Get bChartActive() As Boolean
    bChartActive = m_bChartActive
End Property

'@Description: Devuelve True si, en el estado actual de la aplicación,
'             tiene sentido habilitar la acción "Invertir Ejes".
'@Scope: público (consulta de estado desde el Ribbon/modRibbonCallbacks)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Boolean | True si es posible invertir ejes ahora mismo.
'@Category: UI / Estado consolidado
Public Property Get bCanInvertAxes() As Boolean
    On Error GoTo ErrHandler
    
    bCanInvertAxes = m_bChartActive
    If bCanInvertAxes Then bCanInvertAxes = mFileMgr.AnalizarArchivoActivo.TipoDetectado = CGASING_CurvasRendimiento
    If Not bCanInvertAxes Then Exit Property
    Dim oExcelFile As clsExcelFile
    Set oExcelFile = mFileMgr.ActiveWb
    If Not oExcelFile Is Nothing Then
        bCanInvertAxes = oExcelFile.EsValidoInvertirEjes
    Else
        bCanInvertAxes = False
    End If
    Exit Property
ErrHandler:
    Debug.Print "[clsAplicacion bCanInvertAxes] error: " & Err.Description
    bCanInvertAxes = False
End Property

Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

Public Property Get FileMgr() As clsFileManager
    Set FileMgr = mFileMgr
End Property

Public Property Get OpportunitiesMgr() As clsOpportunitiesMgr
    Set OpportunitiesMgr = mOpportunities
End Property

Public Property Get ChartManager() As clsChartEventsManager
    Set ChartManager = mChartManager
End Property

Public Property Get Ribbon() As clsRibbonEvents
    Set Ribbon = evRibbon
End Property

Public Property Get RibbonState() As clsRibbonState
    Set RibbonState = mRibbonState
End Property

Public Property Let RibbonHandler(xlRibbon As IRibbonUI)
    If evRibbon Is Nothing Then Set evRibbon = New clsRibbonEvents
    Call evRibbon.Init(xlRibbon)
    Debug.Print vbTab & "[clsAplicacion RibbonHandler] - conectada instancia WithEvents de ribbon para gestion de eventos"
End Property

' -------------------------------------------------------------
' Class_Initialize: SOLO inicialización mínima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    Debug.Print vbTab & "[clsAplicacion Class_Initialize]"
    
    On Error GoTo ErrHandler
    
    If evRibbon Is Nothing Then Set evRibbon = New clsRibbonEvents
    Debug.Print vbTab & "[clsAplicacion Class_Initialize] - CREADA instancia WithEvents de ribbon para gestion de eventos"
    Set mRibbonState = New clsRibbonState
    mRibbonState.Modo = Ribbon_OpportunityOnly
    
    ' 1. Configuración (sin dependencias)
    Set mConfiguration = New clsConfiguration
    
    ' 2. FileManager (solo depende de configuración)
    Set mFileMgr = New clsFileManager
    InitFileManager

    ' 3. OpportunitiesMgr
    Set mOpportunities = New clsOpportunitiesMgr
    ' Inicialización explícita
    mOpportunities.SetBaseFolder mConfiguration.RutaOportunidades
    
    ' 4. ChartManager (gestion de eventos del usuario sobre charts de excel (seleccion; etc)
    Set mChartManager = New clsChartEventsManager
    
    ' 5. Iniciar Application events
    Set xlApp = Application
    Debug.Print "[clsAplicacion Class_Initialize] - xlApp suscrito"
    
    ' 6. Watcher
    'StartFolderWatcher mConfiguration.oDicFoldersToWatch
    Set mFolderWatcher = New clsFolderWatch
    Call IniciarMonitoreo
    
    Debug.Print "[clsAplicacion Class_Initialize] - inicialización completada"
    
    ' como consecuencia de introducir el watcher, habrá que gestionar la actualizacion del ribbon, con los cambios que se graben en el
    ' registro; de ello se encarga la aplicacion (la configuracion NO sabe si hay un control dinámico de las carpetas, o si su contenido
    ' es estático)
    ' LA PREGUNTA: DEL RIBBON SOLO SABE LA APLICACION, o trmaiben sabe el watcher? (encomiendo al modulo del watcher, o a la app, la actualiz del ribbon?)

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsAplicacion.Class_Initialize", _
              "Error inicializando aplicación: " & Err.Description
End Sub

' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    Terminate
End Sub

' Método Terminate para ser llamado desde ThisWorkbook.Workbook_BeforeClose/Uninstall
Public Sub Terminate()
    LogInfo "clsAplicacion", "Terminate: Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza continúa aunque haya errores

    ' 1. Liberar FolderWatcher primero (tiene recursos COM externos)
    If Not mFolderWatcher Is Nothing Then
        LogDebug "clsAplicacion", "Terminate: Liberando FolderWatcher COM"
        mFolderWatcher.Dispose  ' Usar el nuevo metodo Dispose
        Set mFolderWatcher = Nothing
    End If

    ' 2. Liberar Ribbon events
    If Not evRibbon Is Nothing Then
        evRibbon.StopEvents
        Set evRibbon = Nothing
    End If

    ' 3. Liberar managers
    Set mConfiguration = Nothing
    Set mOpportunities = Nothing
    Set mFileMgr = Nothing
    Set mChartManager = Nothing
    Set mRibbonState = Nothing

    ' 4. Liberar referencia a Excel Application
    Set xlApp = Nothing

    LogInfo "clsAplicacion", "Terminate: Limpieza completada"
    On Error GoTo 0
End Sub

' -------------------------------------------------------------
' InitFileManager: GESTIONA LA COLECCIÓN DE WORKBOOKS
' (Es responsabilidad de clsAplicacion, NO de clsFileManager)
' -------------------------------------------------------------
Public Sub InitFileManager()
    Debug.Print "[clsAplicacion InitFileManager]"
    
    On Error GoTo ErrHandler
    
    Dim oWorkBook As Workbook
    Dim TrackedCount As Long
    
    ' Trackear todos los workbooks abiertos
    For Each oWorkBook In Application.Workbooks
        mFileMgr.GetOrTrackWorkbook oWorkBook
        TrackedCount = TrackedCount + 1
    Next oWorkBook
    
    ' Establecer el archivo activo
    If Not Application.ActiveWorkbook Is Nothing Then
        Set mFileMgr.ActiveWb = mFileMgr.GetOrTrackWorkbook(Application.ActiveWorkbook)
    End If
    
    Debug.Print "[clsAplicacion InitFileManager] Workbooks tracked: " & TrackedCount
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsAplicacion.InitFileManager", _
              "Error inicializando tracking de workbooks: " & Err.Description
End Sub

' -------------------------------------------------------------
' IniciarMonitoreo: CONFIGURA EL FOLDER WATCHER
' -------------------------------------------------------------
Private Sub IniciarMonitoreo()
    Debug.Print "[clsAplicacion IniciarMonitoreo]"
    
    On Error GoTo ErrHandler
    
    ' Obtener carpetas a monitorear desde configuración
    Dim oDicFolders As Object
    Set oDicFolders = mConfiguration.oDicFoldersToWatch
    
    If oDicFolders Is Nothing Or oDicFolders.Count = 0 Then
        Debug.Print "[clsAplicacion IniciarMonitoreo] No hay carpetas configuradas para monitorear"
        Exit Sub
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
                ConfigurarMonitoreoOportunidades rutaCarpeta, mFolderWatcher
                Debug.Print "[clsAplicacion] Oportunidades monitoreadas: " & rutaCarpeta
                    
            Case CFG_RUTA_PLANTILLAS
                ' Usar función especializada para plantillas
                ConfigurarMonitoreoPlantillas rutaCarpeta, mFolderWatcher
                Debug.Print "[clsAplicacion] Plantillas monitoreadas: " & rutaCarpeta
                    
            Case CFG_RUTA_GAS_VBNET
                ' Usar función especializada para Gas
                ConfigurarMonitoreoGasVBNet rutaCarpeta, mFolderWatcher
                Debug.Print "[clsAplicacion] Gas monitoreado: " & rutaCarpeta
                    
            Case Else
                ' Monitoreo genérico para otras carpetas
                mFolderWatcher.IniciarMonitoreo _
        folderPath:=rutaCarpeta, _
        includeSubdirs:=False, _
        filterPattern:="*.*", _
        eventsToWatch:=Array("Created", "Changed"), _
        inactivityMinutes:=10, _
        foldersOnly:=False
                    
                Debug.Print "[clsAplicacion] Monitoreo genérico: " & rutaCarpeta
            End Select
        Else
            Debug.Print "[clsAplicacion IniciarMonitoreo] ADVERTENCIA: Ruta no existe: " & rutaCarpeta
        End If
    Next key
    
    Debug.Print "[clsAplicacion IniciarMonitoreo] Monitoreo iniciado en " & oDicFolders.Count & " carpetas"
    Exit Sub
    
ErrHandler:
    Debug.Print "[clsAplicacion.IniciarMonitoreo] Error: " & Err.Description
End Sub

' -------------------------------------------------------------
' Property: Acceso al FolderWatcher (para debugging)
' -------------------------------------------------------------
Public Property Get FolderWatcher() As clsFolderWatch
    Set FolderWatcher = mFolderWatcher
End Property

' ==========================================
' gestion del ribbon
' ==========================================
'@Description: Conmuta cíclicamente el modo del Ribbon (admin ? opp ? user ? admin)
'@Scope: público
'@Category: UI / Ribbon
Public Sub ToggleRibbonMode()
    mRibbonState.ToggleModo
    Call ActivarTabSiProcede
End Sub

Private Sub ActivarTabSiProcede()
    On Error Resume Next

    If evRibbon Is Nothing Then Exit Sub

    If EsFicheroOportunidad() Then
        evRibbon.ActivarTab "tabOfertasEspecial"
    End If
End Sub

Private Sub mRibbonState_StateChanged()
    evRibbon.InvalidarRibbon
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Eventos de Application (con delegación al FileManager, ChartManager etc)
' -------------------------------------------------------------

Private Sub xlApp_WorkbookOpen(ByVal wb As Workbook)
    Debug.Print vbTab & vbTab & vbTab & vbTab & "[clsAplicacion [xlApp_WorkbookOpen] - libro abierto: " & wb.Name
    mFileMgr.GetOrTrackWorkbook wb
End Sub

Private Sub xlApp_WorkbookActivate(ByVal wb As Workbook)
    Debug.Print vbTab & vbTab & vbTab & vbTab & "[clsAplicacion xlApp_WorkbookActivate] - libro activo: " & wb.Name & " (ribbon: " & TypeName(evRibbon) & ")"
    Set mFileMgr.ActiveWb = mFileMgr.GetOrTrackWorkbook(wb) ' actualiza current
    mChartManager.RefreshCurrentSheet
End Sub

Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    Debug.Print vbTab & vbTab & vbTab & vbTab & "[clsAplicacion xlApp_WorkbookBeforeClose] - libro a cerrar: " & wb.Name
    mFileMgr.UntrackWorkbook wb
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object
Private Sub xlApp_SheetActivate(ByVal sh As Object)
    On Error GoTo ErrHandler
    
    Debug.Print "[clsAplicacion xlApp_SheetActivate] [i] Hoja activada: " & TypeName(sh) & " '" & sh.Name & "'"
    mChartManager.WatchSheet sh
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"
    'evRibbon.InvalidarRibbon
    Exit Sub
ErrHandler:
    Debug.Print "[clsAplicacion xlApp_SheetActivate] error: " & Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object
Private Sub xlApp_SheetDeactivate(ByVal sh As Object)
    On Error GoTo ErrHandler
    
    Debug.Print "[clsAplicacion xlApp_SheetDeactivate] [i] Hoja desactivada: '" & sh.Name & "'"
    mChartManager.StopWatching
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"

    Exit Sub
ErrHandler:
    Debug.Print "[clsAplicacion xlApp_SheetDeactivate] error: " & Err.Description
End Sub

' -------------------------------------------------------------
' Evento: cambios detectados en oportunidades (desde OpportunitiesMgr)
' -------------------------------------------------------------
Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    Debug.Print "[clsAplicacion mOpportunities_OpportunityCollectionUpdate] - cambios detectados: " & cambios
    
    ' Actualizar ribbon - usar Invalidate en lugar de InvalidarControl
    If Not evRibbon Is Nothing Then
        'evRibbon.InvalidarRibbon
        evRibbon.InvalidarControl "ddlOportunidades"
        Debug.Print "[clsAplicacion mOpportunities_OpportunityCollectionUpdate] - Ribbon invalidado"
    End If
End Sub
'TODO: Extraer la lógica de clsExcelFile A una función (En principio se mantendrá en Celesa aplicación pero es posible que se lleve a otra clase)
Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    Debug.Print "[clsAplicacion mOpportunities_currOpportunityChanged] - Oportunidad seleccionada: " & Path
    
    Dim f As clsExcelFile
    Set f = mFileMgr.ActiveWb
    If Not f Is Nothing Then
        If f.DebeMoverseAOp(Path, mConfiguration.RutaOportunidades) Then
            f.MoverAOp Path, mConfiguration.RutaOportunidades
        End If
    End If
    ' (Aquí no tiene sentido invalidar el ribbon, esto no tiene nada que ver con ribbon, SOLO LÓGICA DE NEGOCIO)
End Sub

' -------------------------------------------------------------
' Eventos de cambio de seleccion de charts (desde ChartManager)
' -------------------------------------------------------------

Private Sub mChartManager_ChartActivated(cht As chart)
    Debug.Print "[clsAplicacion] Gráfico activado: " & cht.Name
    m_bChartActive = True
    ' Aquí decides:
    ' - ¿Invalidar ribbon?
    ' - ¿Cargar datos del gráfico?
    ' - ¿Actualizar estado de la oportunidad?
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartManager_ChartDeactivated(cht As chart)
    Debug.Print "[clsAplicacion] Gráfico DESactivado: " & cht.Name
    m_bChartActive = False
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

' -------------------------------------------------------------
' Eventos de comandos activados desde el Ribbon
' -------------------------------------------------------------

Private Sub evRibbon_GenerarGraficosDesdeCurvasRto()
    Debug.Print "[callback: clsRibbonEvents GenerarGraficosDesdeCurvasRto]"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call EjecutarGraficoEnLibroActivo
    mChartManager.RefreshCurrentSheet

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Private Sub evRibbon_InvertirEjes()
    Debug.Print "[callback: clsRibbonEvents InvertirEjes]"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call InvertirEjesDelGraficoActivo
ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Private Sub evRibbon_FormatearCGASING()
    Debug.Print "[callback: clsRibbonEvents FormatearCGASING]"
    On Error GoTo ErrorHandler
    'Application.ScreenUpdating = False
    Call FixCGASING
ErrorHandler:
    Application.ScreenUpdating = True
End Sub

'@Description: Muestra el formulario de configuracion de la aplicacion
'@Scope: Private (evento de ribbon)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: frmConfiguracion
'@Note: Usa instanciacion explicita para evitar problemas de memoria con la instancia predeterminada
Private Sub evRibbon_Configurador()
    Dim frm As frmConfiguracion
    Set frm = New frmConfiguracion
    frm.Show vbModal
    Unload frm
    Set frm = Nothing
End Sub

Private Sub evRibbon_NuevaOportunidad()
    OpportunitiesMgr.CreaOportunidad
    'evRibbon.InvalidarControl "ddlOportunidades"
End Sub

Private Sub evRibbon_ReplaceWithNamesInValidations()
    On Error GoTo ErrorHandler
    'Application.ScreenUpdating = False
    Call modAPPBudgetQuotesUtilids.AplicarNombresAValidacionesCeldasActWBConReporte
ErrorHandler:
    'Application.ScreenUpdating = True
End Sub

' -------------------------------------------------------------
' Eventos de cambios en el sistema de archivos (desde FolderWatch)
' -------------------------------------------------------------

' Eventos de subcarpetas (reemplazan los eventos de archivos para oportunidades)
Private Sub mFolderWatcher_SubfolderCreated(ByVal parentFolder As String, ByVal subfolderName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_SubfolderCreated] " & subfolderName & " en " & parentFolder
    
    ' Determinar tipo de carpeta padre
    If InStr(1, parentFolder, mConfiguration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Es una nueva oportunidad - delegar a OpportunitiesMgr
        mOpportunities.ProcesarCambiosEnOportunidades subfolderName
        Debug.Print "[clsAplicacion] Nueva oportunidad detectada: " & subfolderName
    End If
End Sub

Private Sub mFolderWatcher_SubfolderDeleted(ByVal parentFolder As String, ByVal subfolderName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_SubfolderDeleted] " & subfolderName & " en " & parentFolder
    
    If InStr(1, parentFolder, mConfiguration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Oportunidad eliminada - delegar a OpportunitiesMgr
        mOpportunities.ProcesarCambiosEnOportunidades subfolderName
        Debug.Print "[clsAplicacion] Oportunidad eliminada: " & subfolderName
    End If
End Sub

Private Sub mFolderWatcher_SubfolderRenamed(ByVal parentFolder As String, ByVal oldName As String, ByVal newName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_SubfolderRenamed] " & oldName & " -> " & newName & " en " & parentFolder
    
    If InStr(1, parentFolder, mConfiguration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Oportunidad renombrada - delegar a OpportunitiesMgr
        ' Primero eliminar la antigua
        mOpportunities.ProcesarCambiosEnOportunidades oldName
        ' Luego agregar la nueva
        mOpportunities.ProcesarCambiosEnOportunidades newName
        Debug.Print "[clsAplicacion] Oportunidad renombrada: " & oldName & " -> " & newName
    End If
End Sub

' Los eventos de archivos individuales se mantienen para plantillas y Gas
Private Sub mFolderWatcher_FileCreated(ByVal folder As String, ByVal fileName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_FileCreated] " & fileName & " en " & folder
    
    ' Para plantillas o archivos Gas - implementar según necesidad
    If InStr(1, folder, mConfiguration.RutaPlantillas, vbTextCompare) > 0 Then
        Debug.Print "[clsAplicacion] Plantilla nueva detectada: " & fileName
        ' TODO: Implementar lógica de actualización de plantillas
    ElseIf InStr(1, folder, mConfiguration.RutaGasVBNet, vbTextCompare) > 0 Then
        Debug.Print "[clsAplicacion] Archivo Gas nuevo detectado: " & fileName
        ' TODO: Implementar lógica de sincronización Gas
    End If
End Sub

Private Sub mFolderWatcher_FileDeleted(ByVal folder As String, ByVal fileName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_FileDeleted] " & fileName & " en " & folder
    ' Implementar según necesidad
    ' Determinar tipo de carpeta
    If InStr(1, folder, mConfiguration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Es una oportunidad eliminada - delegar a OpportunitiesMgr
        mOpportunities.ProcesarCambiosEnItemsOportunidad fileName
    End If
End Sub

Private Sub mFolderWatcher_FileChanged(ByVal folder As String, ByVal fileName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_FileChanged] " & fileName & " en " & folder
    
    ' Detectar cambios en plantillas o archivos Gas
    If InStr(1, folder, mConfiguration.RutaPlantillas, vbTextCompare) > 0 Then
        Debug.Print "[clsAplicacion] Plantilla modificada: " & fileName
        ' TODO: Notificar cambio en plantilla
    ElseIf InStr(1, folder, mConfiguration.RutaGasVBNet, vbTextCompare) > 0 Then
        Debug.Print "[clsAplicacion] Archivo Gas modificado: " & fileName
        ' TODO: Actualizar referencias
    End If
End Sub

Private Sub mFolderWatcher_FileRenamed(ByVal folder As String, ByVal oldName As String, ByVal newName As String)
    Debug.Print "[clsAplicacion mFolderWatcher_FileRenamed] " & oldName & " -> " & newName & " en " & folder
    ' Implementar según necesidad
    ' Determinar tipo de carpeta
    If InStr(1, folder, mConfiguration.RutaOportunidades, vbTextCompare) > 0 Then
        ' Oportunidad renombrada - delegar a OpportunitiesMgr
        mOpportunities.ProcesarCambiosEnItemsOportunidad newName
    End If
End Sub

' Los eventos Heartbeat y ErrorOccurred se mantienen igual
Private Sub mFolderWatcher_Heartbeat(ByVal folder As String, ByVal lastUpdate As Date)
    ' Heartbeat - solo para debug detallado si es necesario
    ' Debug.Print "[clsAplicacion mFolderWatcher_Heartbeat] " & folder & " - " & lastUpdate
End Sub

Private Sub mFolderWatcher_ErrorOccurred(ByVal folder As String, ByVal errorMessage As String)
    Debug.Print "[clsAplicacion mFolderWatcher_ErrorOccurred] ERROR en " & folder & ": " & errorMessage
    
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
    Debug.Print "[clsAplicacion mFolderWatcher_WatcherReconnected] " & folder & " reconectado tras " & attempts & " intento(s)"
    ' Opcional: notificar al usuario solo si hubo multiples intentos
    If attempts > 1 Then
        LogInfo "clsAplicacion", "Monitoreo restaurado en " & folder & " tras " & attempts & " intentos"
    End If
End Sub

Private Sub mFolderWatcher_WatcherReconnectionFailed(ByVal folder As String, ByVal reason As String)
    Debug.Print "[clsAplicacion mFolderWatcher_WatcherReconnectionFailed] " & folder & ": " & reason

    ' Notificar al usuario que el monitoreo fallo permanentemente
    MsgBox "El monitoreo de la carpeta ha fallado permanentemente:" & vbCrLf & _
           folder & vbCrLf & vbCrLf & _
           "Razon: " & reason & vbCrLf & vbCrLf & _
           "Es posible que la carpeta no este accesible o haya problemas de red.", _
           vbExclamation, "FolderWatcher - Reconexion Fallida"
End Sub



