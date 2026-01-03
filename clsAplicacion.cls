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
Private WithEvents xlApp As Application
Attribute xlApp.VB_VarHelpID = -1

Private WithEvents mRibbonState As clsRibbonState
Attribute mRibbonState.VB_VarHelpID = -1
Private WithEvents evRibbon As clsRibbonEvents
Attribute evRibbon.VB_VarHelpID = -1

Private mConfiguration As clsConfiguration
Private mFileMgr As clsFileManager
Private mExecutionContext As clsExecutionContext

Public Property Get bChartActive() As Boolean
    bChartActive = m_bChartActive
End Property

Public Property Get Configuration() As clsConfiguration
    Set Configuration = mConfiguration
End Property

Public Property Get ExecutionContext() As clsExecutionContext
    Set ExecutionContext = mExecutionContext
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
' Class_Initialize: SOLO inicialización mínima (ligera y segura)
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
    
    ' 1.5 Contexto de ejecución (antes del FileManager y demás)
    Set mExecutionContext = New clsExecutionContext
    mExecutionContext.Initialize

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
    LogInfo MODULE_NAME, "[Class_Initialize] - xlApp suscrito"
    
    ' 6. Watcher
    Set mFSMonitoringCoord = New clsFSMonitoringCoord
    'StartFolderWatcher mConfiguration.oDicFoldersToWatch
    mFSMonitoringCoord.IniciarMonitoreo
    
    LogInfo MODULE_NAME, "[Class_Initialize] - inicialización completada"
    
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

    ' 4. Liberar referencia a Excel Application
    Set xlApp = Nothing

    LogInfo MODULE_NAME, "[Terminate] - Limpieza completada"
End Sub

' -------------------------------------------------------------
' InitFileManager: GESTIONA LA COLECCIÓN DE WORKBOOKS
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
    Set wb = mExecutionContext.Workbook
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
    LogDebug MODULE_NAME, "[callback: Application WorkbookOpen] - libro abierto: " & wb.Name
    mFileMgr.GetOrTrackWorkbook wb
End Sub

Private Sub xlApp_WorkbookActivate(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[callback: Application WorkbookActivate] - libro activo: " & wb.Name & " (ribbon: " & TypeName(evRibbon) & ")"
    Set mFileMgr.ActiveWb = mFileMgr.GetOrTrackWorkbook(wb) ' actualiza current
    mChartManager.RefreshCurrentSheet
End Sub

Private Sub xlApp_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[callback: Application WorkbookBeforeClose] - libro a cerrar: " & wb.Name
    mFileMgr.UntrackWorkbook wb
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object
Private Sub xlApp_SheetActivate(ByVal sh As Object)
    On Error GoTo ErrHandler
    
    LogDebug MODULE_NAME, "[callback: Application SheetActivate] [i] Hoja activada: " & TypeName(sh) & " '" & sh.Name & "'"
    mChartManager.WatchSheet sh
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"
    'evRibbon.InvalidarRibbon
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[callback: Application SheetActivate] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object
Private Sub xlApp_SheetDeactivate(ByVal sh As Object)
    On Error GoTo ErrHandler
    
    LogDebug MODULE_NAME, "[callback: Application SheetDeactivate] [i] Hoja desactivada: '" & sh.Name & "'"
    mChartManager.StopWatching
    evRibbon.InvalidarControl "btnGenerarGraficos"
    evRibbon.InvalidarControl "btnCGASING"

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[callback: Application SheetDeactivate] Error", , Err.Description
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
'TODO: Extraer la lógica de clsExcelFile A una función (En principio se mantendrá en Celesa aplicación pero es posible que se lleve a otra clase)
Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    LogInfo MODULE_NAME, "[callback: clsOpportunitiesMgr currOpportunityChanged] - Oportunidad seleccionada: " & Path
    
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

Private Sub mChartManager_ChartActivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartEventsManager ChartActivated] - Gráfico activado: " & cht.Name
    m_bChartActive = True
    ' Aquí decides:
    ' - ¿Invalidar ribbon?
    ' - ¿Cargar datos del gráfico?
    ' - ¿Actualizar estado de la oportunidad?
    evRibbon.InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartManager_ChartDeactivated(cht As Chart)
    LogDebug MODULE_NAME, "[callback: clsChartEventsManager ChartActivated] - Gráfico DESactivado: " & cht.Name
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
    LogError MODULE_NAME, "[bCanInvertAxes] Error", , Err.Description
    bCanInvertAxes = False
End Property


