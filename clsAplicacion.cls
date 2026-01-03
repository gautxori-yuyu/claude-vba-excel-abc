VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' gestion de la aplicacion 'gestor de oportunidades ABC'
'@Folder "3-Aplicac (Coord)"
Option Explicit

Private Const MODULE_NAME As String = "clsAplicacion"

Private m_bChartActive As Boolean

' lo siguiente es para controlar **SUCESOS EN LA COLECCION clsOpportunitiesMgr QUE HACEN
' QUE TENGA QUE INTERVENIR clsAplicacion (para coordinar con clsFileManager, o con el Ribbon, ...)
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1
Private WithEvents mChartManager As clsChartEventsManager
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mFSMonitoringCoord As clsFSMonitoringCoord
Attribute mFSMonitoringCoord.VB_VarHelpID = -1
Private WithEvents ctx As clsExecutionContext
Attribute ctx.VB_VarHelpID = -1

Private WithEvents mRibbonState As clsRibbonState
Attribute mRibbonState.VB_VarHelpID = -1
Private WithEvents evRibbon As clsRibbonEvents
Attribute evRibbon.VB_VarHelpID = -1

Private mConfiguration As clsConfiguration
Private mFileMgr As clsFileManager

Public Property Get bChartActive() As Boolean
    bChartActive = m_bChartActive
End Property

Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = ctx
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
    LogInfo MODULE_NAME, "[RibbonHandler] - conectada instancia WithEvents de ribbon para gestion de eventos"
End Property

' -------------------------------------------------------------
' Class_Initialize: SOLO inicializaci�n m�nima (ligera y segura)
' NO debe hacer trabajo pesado ni depender de App
' -------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrHandler

    LogInfo MODULE_NAME, "[Class_Initialize] - Iniciando aplicacion"
    
    Set evRibbon = New clsRibbonEvents
    Set mRibbonState = New clsRibbonState
    mRibbonState.Modo = Ribbon_OpportunityOnly
    
    ' 1. Configuración (sin dependencias)
    Set mConfiguration = New clsConfiguration

    ' 2. Contexto de ejecución (suscripción a eventos de Application)
    Set ctx = New clsExecutionContext
    ctx.Initialize
    LogInfo MODULE_NAME, "[Class_Initialize] - ctx suscrito a eventos de Application"

    ' 3. FileManager (recibe el contexto compartido)
    Set mFileMgr = New clsFileManager
    mFileMgr.Initialize ctx
    InitFileManager

    ' 4. OpportunitiesMgr
    Set mOpportunities = New clsOpportunitiesMgr
    mOpportunities.SetBaseFolder mConfiguration.RutaOportunidades

    ' 5. ChartManager (gestion de eventos del usuario sobre charts)
    Set mChartManager = New clsChartEventsManager
    
    ' 6. Watcher - recibe el diccionario de carpetas como parametro
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    mFSMonitoringCoord.IniciarMonitoreo mConfiguration.oDicFoldersToWatch
    
    LogInfo MODULE_NAME, "[Class_Initialize] - inicializaci�n completada"
    
    ' como consecuencia de introducir el watcher, habr� que gestionar la actualizacion del ribbon, con los cambios que se graben en el
    ' registro; de ello se encarga la aplicacion (la configuracion NO sabe si hay un control din�mico de las carpetas, o si su contenido
    ' es est�tico)
    ' LA PREGUNTA: DEL RIBBON SOLO SABE LA APLICACION, o trmaiben sabe el watcher? (encomiendo al modulo del watcher, o a la app, la actualiz del ribbon?)

    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "clsAplicacion.Class_Initialize", _
              "Error inicializando aplicaci�n: " & Err.Description
End Sub

' -------------------------------------------------------------
' Class_Terminate: limpieza
' -------------------------------------------------------------
Private Sub Class_Terminate()
    Terminate
End Sub

' Método Terminate para ser llamado desde ThisWorkbook.Workbook_BeforeClose/Uninstall
Public Sub Terminate()
    LogInfo MODULE_NAME, "[Terminate] - Iniciando limpieza de la aplicacion"

    On Error Resume Next  ' Asegurar que la limpieza continúa aunque haya errores

    ' 1. Liberar FolderWatcher primero (tiene recursos COM externos)
    Set mFSMonitoringCoord = Nothing

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

    ' 4. Liberar contexto de ejecución (desuscribe de Application)
    Set ctx = Nothing

    LogInfo MODULE_NAME, "[Terminate] - Limpieza completada"
End Sub

' -------------------------------------------------------------
' InitFileManager: GESTIONA LA COLECCI�N DE WORKBOOKS
' (Es responsabilidad de clsAplicacion, NO de clsFileManager)
' -------------------------------------------------------------
Public Sub InitFileManager()
    LogInfo MODULE_NAME, "[InitFileManager]"
    
    On Error GoTo ErrHandler
    
    Dim oWorkBook As Workbook
    Dim TrackedCount As Long
    
    ' Trackear todos los workbooks abiertos
    For Each oWorkBook In Application.Workbooks
        mFileMgr.GetOrTrackWorkbook oWorkBook
        TrackedCount = TrackedCount + 1
    Next oWorkBook
    
    ' Establecer el archivo activo
    Dim wb As Workbook
    Set wb = ctx.Workbook
    If Not wb Is Nothing Then
        Set mFileMgr.ActiveWb = mFileMgr.GetOrTrackWorkbook(wb)
    End If

    LogInfo MODULE_NAME, "[InitFileManager] Workbooks tracked: " & TrackedCount
    
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[InitFileManager] Error inicializando tracking de workbooks", , Err.Description
    Err.Raise Err.Number, "clsAplicacion.InitFileManager", _
              "Error inicializando tracking de workbooks: " & Err.Description
End Sub

' ==========================================
' gestion del ribbon
' ==========================================
'@Description: Conmuta c�clicamente el modo del Ribbon (admin ? opp ? user ? admin)
'@Scope: p�blico
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
' Eventos de clsExecutionContext (propagados desde Application)
' -------------------------------------------------------------

'@Description: Se dispara cuando se activa un workbook (desde ctx)
Private Sub ctx_WorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[ctx_WorkbookActivated] " & wb.Name
    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
Private Sub ctx_SheetActivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[ctx_SheetActivated] " & TypeName(sh) & " '" & sh.Name & "'"
    mChartManager.WatchSheet sh
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ctx_SheetActivated] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
Private Sub ctx_SheetDeactivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[ctx_SheetDeactivated] '" & sh.Name & "'"
    mChartManager.StopWatching
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[ctx_SheetDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Evento: cambios detectados en oportunidades (desde OpportunitiesMgr)
' -------------------------------------------------------------
Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr OpportunityCollectionUpdate] - cambios detectados: " & cambios
    
    ' Actualizar ribbon - usar Invalidate en lugar de InvalidarControl
    If Not evRibbon Is Nothing Then
        'evRibbon.InvalidarRibbon
        evRibbon.InvalidarControl "ddlOportunidades"
        LogDebug MODULE_NAME, "[mOpportunities_OpportunityCollectionUpdate] - Ribbon invalidado"
    End If
End Sub
'TODO: Extraer la l�gica de clsExcelFile A una funci�n (En principio se mantendr� en Celesa aplicaci�n pero es posible que se lleve a otra clase)
Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr currOpportunityChanged] - Oportunidad seleccionada: " & Path
    
    Dim f As clsExcelFile
    Set f = mFileMgr.ActiveWb
    If Not f Is Nothing Then
        If f.DebeMoverseAOp(Path, mConfiguration.RutaOportunidades) Then
            f.MoverAOp Path, mConfiguration.RutaOportunidades
        End If
    End If
    ' (Aqu� no tiene sentido invalidar el ribbon, esto no tiene nada que ver con ribbon, SOLO L�GICA DE NEGOCIO)
End Sub

' -------------------------------------------------------------
' Eventos de mFSMonitoringCoord (cambios en sistema de archivos)
' clsAplicacion recibe estos eventos y delega a OpportunitiesMgr
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

    ' Mostrar al usuario solo errores criticos
    If InStr(1, errorMessage, "Limite", vbTextCompare) > 0 Or _
       InStr(1, errorMessage, "no encontrada", vbTextCompare) > 0 Then
        MsgBox "Error en monitoreo de carpeta:" & vbCrLf & _
               folder & vbCrLf & vbCrLf & _
               errorMessage, vbExclamation, "FolderWatcher"
    End If
End Sub

Private Sub mFSMonitoringCoord_MonitoringReconnected(ByVal folder As String, ByVal attempts As Long)
    LogInfo MODULE_NAME, "[mFSMonitoringCoord_MonitoringReconnected] " & folder & " tras " & attempts & " intentos"
End Sub

Private Sub mFSMonitoringCoord_MonitoringFailed(ByVal folder As String, ByVal reason As String)
    LogError MODULE_NAME, "[mFSMonitoringCoord_MonitoringFailed] " & folder & ": " & reason

    MsgBox "El monitoreo de la carpeta ha fallado permanentemente:" & vbCrLf & _
           folder & vbCrLf & vbCrLf & _
           "Razon: " & reason & vbCrLf & vbCrLf & _
           "Es posible que la carpeta no este accesible o haya problemas de red.", _
           vbExclamation, "FolderWatcher - Reconexion Fallida"
End Sub

' -------------------------------------------------------------
' Eventos de cambio de seleccion de charts (desde ChartManager)
' -------------------------------------------------------------

Private Sub mChartManager_ChartActivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartEventsManager ChartActivated] - Gr�fico activado: " & cht.Name
    m_bChartActive = True
    ' Aqu� decides:
    ' - �Invalidar ribbon?
    ' - �Cargar datos del gr�fico?
    ' - �Actualizar estado de la oportunidad?
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartManager_ChartDeactivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartEventsManager ChartActivated] - Gr�fico DESactivado: " & cht.Name
    m_bChartActive = False
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

' -------------------------------------------------------------
' Eventos de comandos activados desde el Ribbon
' -------------------------------------------------------------

Private Sub evRibbon_GenerarGraficosDesdeCurvasRto()
    LogDebug MODULE_NAME, "[callback: clsRibbonEvents GenerarGraficosDesdeCurvasRto]"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call EjecutarGraficoEnLibroActivo
    mChartManager.RefreshCurrentSheet

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Private Sub evRibbon_InvertirEjes()
    LogDebug MODULE_NAME, "[callback: clsRibbonEvents InvertirEjes]"
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call InvertirEjesDelGraficoActivo
ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Private Sub evRibbon_FormatearCGASING()
    LogDebug MODULE_NAME, "[callback: clsRibbonEvents FormatearCGASING]"
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

'@Description: Devuelve True si, en el estado actual de la aplicaci�n,
'             tiene sentido habilitar la acci�n "Invertir Ejes".
'@Scope: p�blico (consulta de estado desde el Ribbon/modRibbonCallbacks)
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
    LogError MODULE_NAME, "[bCanInvertAxes] Error", , Err.Description
    bCanInvertAxes = False
End Property


