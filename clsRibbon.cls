VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================
' CLASE DE EVENTOS DEL RIBBON
' ==========================================
' Esta clase envuelve el objeto IRibbonUI y gestiona
' su ciclo de vida con proteccion y logging.
'
' CAMBIOS IMPORTANTES:
' - ribbonUI ahora es privado con Property Get
' - mWasEverInitialized es informativo, no bloquea
' ==========================================

'@Folder "3-UI.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbon"

' ==========================================
' RIBBON UI Y ESTADO
' ==========================================
Private mribbonUI As IRibbonUI

Private mStateRibbon As clsRibbonState
Attribute mStateRibbon.VB_VarHelpID = -1

Private mIsRecovering As Boolean              ' Flag para evitar recursion durante recuperacion
Private mWasEverInitialized As Boolean        ' Flag para distinguir "nunca inicializado" de "perdido"

' ==========================================
' DEPENDENCIAS INYECTADAS (para escuchar eventos)
' ==========================================
Private WithEvents mOpportunities As clsOpportunitiesMgr
Attribute mOpportunities.VB_VarHelpID = -1

Private WithEvents mFileMgr As clsFileManager
Attribute mFileMgr.VB_VarHelpID = -1

Private WithEvents mChartMgr As clsChartMgr
Attribute mChartMgr.VB_VarHelpID = -1

Private WithEvents mCtxMgr As clsExecutionContextMgr
Attribute mCtxMgr.VB_VarHelpID = -1

' ==========================================
' PROPERTY GET/SET PARA MIEMBROS
' ==========================================

'@Description: Obtiene el puntero al Ribbon (puede ser Nothing)
Public Property Get ribbonUI() As IRibbonUI
    Set ribbonUI = mribbonUI
End Property

Public Property Get State() As clsRibbonState
    Set State = mStateRibbon
End Property

' ==========================================
' INICIALIZACION Y LIMPIEZA
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    
    Set mStateRibbon = New clsRibbonState
    mStateRibbon.Modo = Ribbon_OpportunityOnly
    
    mWasEverInitialized = False
    mIsRecovering = False
End Sub

Private Sub Class_Terminate()
    Set mStateRibbon = Nothing
    
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub

Public Sub Init(ByRef ribbonObj As IRibbonUI)
    LogDebug MODULE_NAME, "[Init] Recibiendo puntero IRibbonUI"
    Set mribbonUI = ribbonObj

    If ribbonObj Is Nothing Then
        LogWarning MODULE_NAME, "[Init] Se paso Nothing como puntero"
    Else
        LogDebug MODULE_NAME, "[Init] Puntero establecido correctamente"
    End If
    mWasEverInitialized = True
End Sub

'@Description: Inyecta dependencias para escuchar eventos de los Managers
Public Sub Initialize(oOpportunities As clsOpportunitiesMgr, _
                      oFileMgr As clsFileManager, _
                      oChartMgr As clsChartMgr, _
                      oCtxMgr As clsExecutionContextMgr)
    Set mOpportunities = oOpportunities
    Set mFileMgr = oFileMgr
    Set mChartMgr = oChartMgr
    Set mCtxMgr = oCtxMgr
End Sub

Public Sub StopEvents()
    LogDebug MODULE_NAME, "[StopEvents] Liberando puntero"
    Set mribbonUI = Nothing
    Set mOpportunities = Nothing
    Set mFileMgr = Nothing
    Set mChartMgr = Nothing
    Set mCtxMgr = Nothing
End Sub


' ==========================================
'
' ==========================================
Public Sub ToggleModo()
    ' Logica de toggle movida desde clsRibbonState (el Mgr decide, no el State)
    Select Case mStateRibbon.Modo
    Case Ribbon_OpportunityOnly
        mStateRibbon.Modo = Ribbon_User
    Case Ribbon_User
        mStateRibbon.Modo = Ribbon_Hidden
    Case Ribbon_Hidden
        mStateRibbon.Modo = Ribbon_Admin
    Case Ribbon_Admin
        mStateRibbon.Modo = Ribbon_OpportunityOnly
    Case Else
        mStateRibbon.Modo = Ribbon_OpportunityOnly
    End Select

    ' Guardar el nuevo estado de vuelta al libro activo
    On Error Resume Next
    Dim f As clsFileXLS
    If Not mFileMgr Is Nothing Then
        If Not mFileMgr.State.CurrentFile Is Nothing Then
            Set f = mFileMgr.State.CurrentFile
            f.RibbonMode = mStateRibbon.Modo
            LogDebug MODULE_NAME, "[ToggleModo] Estado guardado en libro: " & f.Name & " -> " & mStateRibbon.Description
        End If
    End If
    On Error GoTo 0

    ' El Mgr invalida el ribbon despues de cambiar estado
    InvalidarRibbon

    On Error Resume Next
    If EsFicheroOportunidad() Then
        ActivarTab "tabOfertasEspecial"
    End If
End Sub

'@Description: Activa una pestaña del Ribbon por id
'@Scope: Friend (uso exclusivo desde clsApplication)
'@ArgumentDescriptions: tabId | id del tab a activar
'@Category: UI / Ribbon
Friend Sub ActivarTab(tabId As String)
    On Error Resume Next
    If IsRibbonUIAvailable() Then
        mribbonUI.ActivateTab tabId
    End If
End Sub

'@Description: Invalida todo el Ribbon, forzando recarga de callbacks
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarRibbon()
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        ' Esto evita que durante Class_Initialize se intente recuperar un ribbon que
        ' aun no ha sido configurado (el callback RibbonOnLoad aun no ha terminado)
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarRibbon] Ribbon aun no inicializado, omitiendo"
            Exit Sub
        End If

        ' Intentar recuperacion automatica (solo si no estamos ya recuperando)
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarRibbon] Ribbon perdido, intentando recuperacion..."
            If TryAutoRecover() Then
                LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion exitosa"
            Else
                LogError MODULE_NAME, "[InvalidarRibbon] Recuperacion fallida"
                Exit Sub
            End If
        Else
            LogDebug MODULE_NAME, "[InvalidarRibbon] Recuperacion en curso, omitiendo"
            Exit Sub
        End If
    End If

    ' Invalidar el Ribbon
    mribbonUI.Invalidate
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarRibbon] Error", Err.Description
    ' No propagar el error, solo loguear
End Sub

'@Description: Invalida un control especifico del Ribbon
'@Note: Intenta recuperacion automatica si el Ribbon esta perdido
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler

    ' Verificar si el Ribbon esta disponible
    If Not IsRibbonUIAvailable() Then
        ' IMPORTANTE: Solo intentar recuperacion si el ribbon FUE inicializado alguna vez
        If Not mWasEverInitialized Then
            LogDebug MODULE_NAME, "[InvalidarControl] Ribbon aun no inicializado, omitiendo: " & idControl
            Exit Sub
        End If

        ' Intentar recuperacion automatica
        If Not mIsRecovering Then
            LogWarning MODULE_NAME, "[InvalidarControl] Ribbon perdido, intentando recuperacion..."
            If Not TryAutoRecover() Then
                LogError MODULE_NAME, "[InvalidarControl] Recuperacion fallida para control: " & idControl
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' Invalidar el control
    mribbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarControl] Error en control " & idControl, , Err.Description
End Sub

'@Description: Verifica si el objeto ribbonUI esta disponible y funcional
Private Function IsRibbonUIAvailable() As Boolean
    If mribbonUI Is Nothing Then IsRibbonUIAvailable = False: Exit Function
    
    On Error Resume Next
    Dim testType As String
    testType = TypeName(mribbonUI)
    If Err.Number <> 0 Then
        IsRibbonUIAvailable = False
        Err.Clear
    ElseIf testType = "Nothing" Or testType = "Empty" Then
        IsRibbonUIAvailable = False
    Else
        IsRibbonUIAvailable = True
    End If
End Function

'@Description: Intenta recuperar el Ribbon automaticamente
Private Function TryAutoRecover() As Boolean
    On Error GoTo ErrHandler

    mIsRecovering = True
    LogInfo MODULE_NAME, "[TryAutoRecover] Iniciando"

    ' Usar el modulo de recuperacion
    TryAutoRecover = TryRecoverRibbon()

    mIsRecovering = False
    Exit Function

ErrHandler:
    LogCurrentError MODULE_NAME, "[TryAutoRecover]"
    mIsRecovering = False
    TryAutoRecover = False
End Function

'@Description: Indica si el Ribbon esta en proceso de recuperacion
Public Property Get IsRecovering() As Boolean
    IsRecovering = mIsRecovering
End Property


'@Description: Determina si la pestana principal del Ribbon debe ser visible
'@Note: La decision de visibilidad vive en el Mgr (no en State). El Mgr consulta
'       su propio State (Modo) y, si hace falta, el estado de la aplicacion.
'@Returns: True si la pestana debe mostrarse segun el modo actual
Public Function IsTabVisible() As Boolean
    On Error GoTo ErrHandler
    Select Case mStateRibbon.Modo
    Case Ribbon_Admin
        IsTabVisible = True
    Case Ribbon_User
        IsTabVisible = True
    Case Ribbon_OpportunityOnly
        ' En este modo: visible solo si hay archivo de oportunidad activo.
        ' La UI puede consultar estado de la aplicacion (conoce al mundo).
        IsTabVisible = App().AppState.IsOpportunityFileActive()
    Case Else
        IsTabVisible = False
    End Select
    Exit Function
ErrHandler:
    IsTabVisible = False
End Function

'@Description: Diagnostico rapido del estado del Ribbon
Public Function GetQuickDiagnostics() As String
    Dim info As String
    info = "RibbonUI: "

    If Not IsRibbonUIAvailable() Then
        info = info & "Nothing"
    Else
        On Error Resume Next
        info = info & TypeName(mribbonUI)
        If Err.Number <> 0 Then
            info = info & "Corrupted"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    info = info & " | WasInit: " & mWasEverInitialized
    info = info & " | Recovering: " & mIsRecovering

    GetQuickDiagnostics = info
End Function

' ==========================================
' HANDLERS DE EVENTOS (UI escucha a los Managers)
' ==========================================

Private Sub mOpportunities_OpportunityCollectionUpdate(ByVal cambios As String)
    InvalidarControl "ddlOportunidades"
End Sub

Private Sub mOpportunities_currOpportunityChanged(ByVal Index As Long, ByVal Path As String)
    ' En modo OpportunityOnly, mostrar ribbon cuando hay oportunidad
    If mStateRibbon.Modo = Ribbon_OpportunityOnly Then
        mStateRibbon.Visible = True
    End If
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "dropdownOportunidades"
    InvalidarControl "btnCreaOportunidad"
    InvalidarControl "grpOportunidades"
End Sub

Private Sub mFileMgr_ActiveFileChanged(f As clsFileXLS)
    ' Sincronizar estado del ribbon desde el libro activo
    If Not f Is Nothing Then
        mStateRibbon.SyncFromBookState f.RibbonBookState
        LogDebug MODULE_NAME, "[mFileMgr_ActiveFileChanged] Sincronizado estado ribbon desde libro: " & f.Name
    End If

    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnExportarPDF"
    InvalidarControl "btnLimpiarFormato"
    InvalidarControl "btnAjustarFilas"

    ' Invalidar todo el ribbon para reflejar el nuevo estado
    InvalidarRibbon
End Sub

Private Sub mChartMgr_ChartActivated(cht As Chart)
    InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartMgr_ChartDeactivated(cht As Chart)
    InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mCtxMgr_WorkbookActivated(ByVal Wb As Workbook)
    ' Necesario para re-evaluar btnCGASING cuando una hoja se copia a otro libro
    ' (WorkbookActivated se dispara al activar el libro destino)
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnCGASING"
End Sub

Private Sub mCtxMgr_SheetActivated(ByVal Sh As Object)
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnCGASING"
End Sub

Private Sub mCtxMgr_SheetDeactivated(ByVal Sh As Object)
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnCGASING"
End Sub

