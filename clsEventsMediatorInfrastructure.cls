VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMediatorInfrastructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.2-Mediadores"
'@ModuleDescription "Gestor centralizado de eventos (orquestador)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMediatorInfrastructure"

' ==================================================================
' EVENTOS SEMANTICOS DE INFRAESTRUCTURA (para consumo por dominio)
' ==================================================================
' Estos eventos traducen los eventos tecnicos de infraestructura
' a un lenguaje que el dominio puede entender sin conocer Excel/COM.
' El mediador de dominio escucha ESTOS eventos, no los de mCtxMgr directamente.

'@Description: El contexto de ejecucion se ha invalidado (reset VBA, error COM, etc.)
Public Event InfraContextInvalidated()

'@Description: El contexto de ejecucion se ha reinicializado tras un reset
Public Event InfraContextReinitialized()

'@Description: El archivo activo ha cambiado (traduccion de ActiveFileChanged)
'@Param filePath: Ruta del archivo activo (vacio si ninguno)
'@Param isOpportunityFile: True si el archivo es de una oportunidad
Public Event ActiveFileSessionChanged(ByVal filePath As String, ByVal IsOpportunityFile As Boolean)

' ==================================================================
' REFERENCIAS INTERNAS
' ==================================================================
Private WithEvents mCtxMgr As clsExecutionContextMgr
Attribute mCtxMgr.VB_VarHelpID = -1

Private WithEvents mFileMgr As clsFileManager
Attribute mFileMgr.VB_VarHelpID = -1

'--------------------------------------------------------------
' Inicialización
'--------------------------------------------------------------
Private Sub Class_Initialize()
    LogInfo MODULE_NAME, "[Class_Initialize]"
End Sub

Private Sub Class_Terminate()
    Terminate
    LogInfo MODULE_NAME, "[Class_Terminate]"
End Sub
Public Sub Initialize(oFileMgr As clsFileManager, oCtxMgr As clsExecutionContextMgr)
    On Error GoTo ErrHandler
    Set mCtxMgr = oCtxMgr
    Set mFileMgr = oFileMgr
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Initialize]"
    Err.Raise Err.Number, MODULE_NAME & " [Initialize]", _
              "Error inicializando aplicación: " & Err.Description
End Sub

Public Sub Terminate()
    On Error GoTo ErrHandler
    Set mCtxMgr = Nothing
    Set mFileMgr = Nothing
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[Terminate]"
    Err.Raise Err.Number, MODULE_NAME & " [Terminate]", _
              "Error finalizando aplicación: " & Err.Description
End Sub

' ==================================================================
' GESTION DE EVENTOS
' ==================================================================

' -------------------------------------------------------------
' Eventos de ExecutionContext
' -------------------------------------------------------------

'@Description: Se dispara cuando se activa un workbook
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: wb As Workbook

Private Sub mCtxMgr_WorkbookActivated(ByVal Wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookActivated] " & Wb.Name & " -> GetOrTrackWorkbook + SetActiveFile"

    ' AQUI ocurre el tracking real: GetOrTrackWorkbook puede tardar varios segundos
    ' si AnalizarWorkBook necesita leer el contenido del workbook (Excel aun cargando).
    ' Esto dispara ActiveFileChanged -> sincroniza estado del ribbon.
    If Not mFileMgr Is Nothing Then
        Dim f As clsFileXLS
        Set f = mFileMgr.GetOrTrackWorkbook(Wb)
        mFileMgr.SetActiveFile f
    End If
    ' Chart refresh: delegado a clsExecutionContextMgr (composicion interna)
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtxMgr_SheetActivated(ByVal Sh As Object)
    ' WatchSheet delegado a clsExecutionContextMgr (composicion interna con clsChartMgr)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetActivated] " & Sh.Name & "'"

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetActivated] Error", , Err.Description
End Sub

Private Sub mCtxMgr_SheetDeactivated(ByVal Sh As Object)
    ' StopWatching delegado a clsExecutionContextMgr (composicion interna con clsChartMgr)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetDeactivated] '" & Sh.Name & "'"

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de tracking de ficheros (desde clsFileMgr)
' -------------------------------------------------------------

'@Description: Callback cuando se abre un workbook
Private Sub mCtxMgr_WorkbookOpened(ByVal Wb As Workbook)
    ' NOTA: WorkbookOpened solo notifica la apertura. El tracking real se hace en
    ' mCtxMgr_WorkbookActivated (WorkbookActivate), que es cuando Excel tiene el
    ' workbook completamente disponible para operaciones COM.
    LogDebug MODULE_NAME, "[mCtx_WorkbookOpened] " & Wb.Name & " (solo notificacion; tracking en WorkbookActivated)"
End Sub

'@Description: Callback cuando se cierra un workbook
Private Sub mCtxMgr_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtx_WorkbookBeforeClose] " & Wb.Name
    If Not mFileMgr Is Nothing Then
        mFileMgr.UntrackWorkbook Wb
    End If
End Sub

' -------------------------------------------------------------
' Eventos de ciclo de vida del contexto (reset / reenganche)
' -------------------------------------------------------------

'@Description: Se dispara cuando el contexto se invalida (shutdown, error COM, etc.)
'@Note: Libera referencias que puedan estar corruptas para evitar errores en cascada
Private Sub mCtxMgr_ContextInvalidated()
    On Error Resume Next
    LogWarning MODULE_NAME, "[mCtxMgr_ContextInvalidated] Contexto invalidado, liberando recursos"

    ' StopWatching de charts delegado a clsExecutionContextMgr (composicion interna)
    ' Emitir evento semantico para que el dominio pueda reaccionar
    RaiseEvent InfraContextInvalidated

    On Error GoTo 0
End Sub

'@Description: Se dispara cuando el contexto se reinicializa tras un reset
'@Note: Reconecta los managers que necesitan referencias actualizadas
Private Sub mCtxMgr_ContextReinitialized()
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[mCtxMgr_ContextReinitialized] Re-enganchando consumidores"

    ' Re-registrar el workbook activo si existe
    Dim Wb As Workbook
    Set Wb = Excel.Application.ActiveWorkbook
    If Not Wb Is Nothing Then
        mFileMgr.GetOrTrackWorkbook Wb
    End If

    ' RefreshCurrentSheet delegado a clsExecutionContextMgr (composicion interna)
    ' Emitir evento semantico para que el dominio pueda reaccionar
    RaiseEvent InfraContextReinitialized

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtxMgr_ContextReinitialized] Error durante reconexion", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de FileManager (traduccion a eventos semanticos)
' -------------------------------------------------------------

'@Description: Traduce el evento tecnico ActiveFileChanged a evento semantico
' FIXME: este evento se esta capturando DOS VECES!!, aqui, y en clsRibbon!!!
Private Sub mFileMgr_ActiveFileChanged(f As clsFileXLS)
    On Error GoTo ErrHandler

    Dim filePath As String
    Dim isOpp As Boolean

    If f Is Nothing Then
        filePath = ""
        isOpp = False
    Else
        filePath = f.path
        isOpp = f.IsOpportunityFile
    End If

    LogDebug MODULE_NAME, "[mFileMgr_ActiveFileChanged] -> ActiveFileSessionChanged: " & filePath

    ' Emitir evento semantico (sin referencias a clsFileXLS)
    RaiseEvent ActiveFileSessionChanged(filePath, isOpp)

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mFileMgr_ActiveFileChanged] Error traduciendo evento", , Err.Description
End Sub

