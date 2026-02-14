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

'@Folder "4-UI.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbon"

' ==========================================
' RIBBON UI Y ESTADO
' ==========================================
Private mribbonUI As IRibbonUI

Private mStateRibbon As clsRibbonState

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
' INICIALIZACION, CONTROL DEL CICLO DE VIDA
' ==========================================

Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
    
    Set mStateRibbon = New clsRibbonState
    ' Modo inicial: Ribbon_User (seguro, sin dependencias).
    ' Ribbon_OpportunityOnly solo se permite cuando mOpportunities y mFileMgr esten inyectados.
    mStateRibbon.Modo = DEFAULT_RIBBONMODE
    
    mWasEverInitialized = False
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
' CONTROL DE MODO DE VISUALIZACION
' ==========================================
Public Sub ToggleModo()
    ' Logica de toggle movida desde clsRibbonState (el Mgr decide, no el State)
    Select Case mStateRibbon.Modo
    Case Ribbon_OpportunityOnly
        ' Si de algun modo llegamos aqui sin soporte, salir a User
        mStateRibbon.Modo = Ribbon_User
    Case Ribbon_User
        mStateRibbon.Modo = Ribbon_Hidden
    Case Ribbon_Hidden
        mStateRibbon.Modo = Ribbon_Admin
    Case Ribbon_Admin
        ' Solo pasar a OpportunityOnly si las dependencias de dominio estan disponibles
        If App.State.IsOpportunityMgrEnabled() Then
            mStateRibbon.Modo = Ribbon_OpportunityOnly
        Else
            mStateRibbon.Modo = Ribbon_User
        End If
    Case Else
        ' Fallback seguro: nunca aterrizar en OpportunityOnly sin soporte
        If App.State.IsOpportunityMgrEnabled() Then
            mStateRibbon.Modo = Ribbon_OpportunityOnly
        Else
            mStateRibbon.Modo = Ribbon_User
        End If
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
    If Not mribbonUI Is Nothing Then
        mribbonUI.ActivateTab tabId
    End If
End Sub

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
            ' Este modo solo es valido si las dependencias de dominio estan disponibles.
            ' Si no lo estan (p.ej. durante inicializacion progresiva), el modo no debe
            ' haberse alcanzado, pero lo tratamos defensivamente devolviendo False.
            If App().State Is Nothing Then
                IsTabVisible = False
            ElseIf Not App.State.IsOpportunityMgrEnabled() Then
                IsTabVisible = False
            Else
                IsTabVisible = App().State.IsOpportunityFileActive()
            End If
        Case Else
            IsTabVisible = False
    End Select
    Exit Function
ErrHandler:
    IsTabVisible = False
End Function

Public Function IsAdminGroupVisible() As Boolean
    IsAdminGroupVisible = (mStateRibbon.Modo = Ribbon_Admin)
End Function

' ==========================================
' INVALIDACION
' ==========================================
'@Description: Invalida todo el Ribbon, forzando recarga de callbacks
'@Note: Si el ribbon se pierde, usar RecuperarRibbon() en modMACROAppLifecycle
Public Sub InvalidarRibbon()
    On Error GoTo ErrHandler

    If mribbonUI Is Nothing Then
        If mWasEverInitialized Then LogWarning MODULE_NAME, "[InvalidarRibbon] ribbonUI no disponible"
        Exit Sub
    End If

    mribbonUI.Invalidate
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarRibbon] Error - marcando ribbon como perdido"
    Set mribbonUI = Nothing
End Sub

'@Description: Invalida un control especifico del Ribbon
'@Note: Si el ribbon se pierde, usar RecuperarRibbon() en modMACROAppLifecycle
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler

    If mribbonUI Is Nothing Then
        If mWasEverInitialized Then LogDebug MODULE_NAME, "[InvalidarControl] ribbonUI no disponible: " & idControl
        Exit Sub
    End If

    mribbonUI.InvalidateControl idControl
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[InvalidarControl] Error en control " & idControl & " - marcando ribbon como perdido"
    Set mribbonUI = Nothing
End Sub

' ==========================================
' HANDLERS DE EVENTOS (UI escucha a los Managers)
' ==========================================

Private Sub mCtxMgr_WorkbookActivated(ByVal Wb As Workbook)
    ' CRÍTICO: Invalidar TODO el ribbon al cambiar de libro
    ' Esto es necesario para:
    ' 1. Re-evaluar la visibilidad de controles según el libro activo
    ' 2. Sincronizar el estado del ribbon con el libro (cada libro puede tener su propio modo)
    ' 3. Refrescar callbacks getVisible cuando hay múltiples libros abiertos
    LogDebug MODULE_NAME, "[mCtxMgr_WorkbookActivated] Invalidando ribbon para libro: " & Wb.Name
    InvalidarRibbon
End Sub

Private Sub mCtxMgr_SheetActivated(ByVal Sh As Object)
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnCGASING"
End Sub

Private Sub mCtxMgr_SheetDeactivated(ByVal Sh As Object)
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnCGASING"
End Sub

Private Sub mChartMgr_ChartActivated(cht As Chart)
    InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mChartMgr_ChartDeactivated(cht As Chart)
    InvalidarControl "btnInvertirSeries"
End Sub

Private Sub mFileMgr_ActiveFileChanged(f As clsFileXLS)
    ' Sincronizar estado del ribbon desde el libro activo
    If Not f Is Nothing Then
        mStateRibbon.Modo = f.RibbonMode
        LogDebug MODULE_NAME, "[mFileMgr_ActiveFileChanged] Sincronizado estado ribbon desde libro: " & f.Name
    End If

    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "btnExportarPDF"
    InvalidarControl "btnLimpiarFormato"
    InvalidarControl "btnAjustarFilas"

    ' Invalidar todo el ribbon para reflejar el nuevo estado
    InvalidarRibbon
End Sub

Private Sub mOpportunities_OpportunityCollectionUpdated(ByVal reason As String)
    InvalidarControl "ddlOportunidades"
End Sub

Private Sub mOpportunities_CurrentOpportunityChanged(ByVal opportunityName As String, ByVal opportunityPath As String)
    ' En modo OpportunityOnly, mostrar ribbon cuando hay oportunidad
'    If mStateRibbon.Modo = Ribbon_OpportunityOnly Then
'        mStateRibbon.Visible = True
'    End If
    InvalidarControl "btnGenerarGraficos"
    InvalidarControl "dropdownOportunidades"
    InvalidarControl "btnNuevaOp"
    InvalidarControl "grpOportunidades"
End Sub

