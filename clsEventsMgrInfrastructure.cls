VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsMgrInfrastructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura"
'@ModuleDescription "Gestor centralizado de eventos (orquestador)"
Option Explicit

Private Const MODULE_NAME As String = "clsEventsMgrInfrastructure"

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
Public Event ActiveFileSessionChanged(ByVal filePath As String, ByVal isOpportunityFile As Boolean)

' ==================================================================
' REFERENCIAS INTERNAS
' ==================================================================
Private mChartManager As clsChartMgr
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
Public Sub Initialize(oFileMgr As clsFileManager, oChartManager As clsChartMgr, oCtxMgr As clsExecutionContextMgr)
    On Error GoTo ErrHandler
    
    Set mChartManager = oChartManager
    
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
    
    Set mChartManager = Nothing
    
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

Private Sub mCtxMgr_WorkbookActivated(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookActivated] " & wb.Name

    ' Actualizar el archivo activo en FileManager
    ' Esto dispara ActiveFileChanged -> sincroniza estado del ribbon
    Dim f As clsFileXLS
    Set f = mFileMgr.GetOrTrackWorkbook(wb)
    mFileMgr.SetActiveFile f

    mChartManager.RefreshCurrentSheet
End Sub

'@Description: Se dispara al activar una hoja; inicia vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtxMgr_SheetActivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetActivated] " & sh.Name & "'"
    mChartManager.WatchSheet sh
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetActivated] Error", , Err.Description
End Sub

'@Description: Se dispara al desactivar una hoja; detiene vigilancia de graficos
'@Category: Gráficos
'@Scope: Grafico seleccionado
'@ArgumentDescriptions: sh As Object

Private Sub mCtxMgr_SheetDeactivated(ByVal sh As Object)
    On Error GoTo ErrHandler

    LogDebug MODULE_NAME, "[mCtx_SheetDeactivated] '" & sh.Name & "'"
    mChartManager.StopWatching
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mCtx_SheetDeactivated] Error", , Err.Description
End Sub

' -------------------------------------------------------------
' Eventos de tracking de ficheros (desde clsFileMgr)
' -------------------------------------------------------------

'@Description: Callback cuando se abre un workbook
Private Sub mCtxMgr_WorkbookOpened(ByVal wb As Workbook)
    LogDebug MODULE_NAME, "[mCtx_WorkbookOpened] " & wb.Name
    mFileMgr.GetOrTrackWorkbook wb
End Sub

'@Description: Callback cuando se cierra un workbook
Private Sub mCtxMgr_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    LogDebug MODULE_NAME, "[mCtx_WorkbookBeforeClose] " & wb.Name
    mFileMgr.UntrackWorkbook wb
End Sub

' -------------------------------------------------------------
' Eventos de ciclo de vida del contexto (reset / reenganche)
' -------------------------------------------------------------

'@Description: Se dispara cuando el contexto se invalida (shutdown, error COM, etc.)
'@Note: Libera referencias que puedan estar corruptas para evitar errores en cascada
Private Sub mCtxMgr_ContextInvalidated()
    On Error Resume Next
    LogWarning MODULE_NAME, "[mCtxMgr_ContextInvalidated] - Contexto invalidado, liberando recursos"

    ' Detener vigilancia de charts (pueden tener referencias muertas)
    If Not mChartManager Is Nothing Then
        mChartManager.StopWatching
    End If

    ' Emitir evento semantico para que el dominio pueda reaccionar
    RaiseEvent InfraContextInvalidated

    On Error GoTo 0
End Sub

'@Description: Se dispara cuando el contexto se reinicializa tras un reset
'@Note: Reconecta los managers que necesitan referencias actualizadas
Private Sub mCtxMgr_ContextReinitialized()
    On Error GoTo ErrHandler
    LogInfo MODULE_NAME, "[mCtxMgr_ContextReinitialized] - Re-enganchando consumidores"

    ' Re-registrar el workbook activo si existe
    Dim wb As Workbook
    Set wb = Excel.Application.ActiveWorkbook
    If Not wb Is Nothing Then
        mFileMgr.GetOrTrackWorkbook wb
    End If

    ' Refrescar el ChartManager con la hoja actual
    If Not mChartManager Is Nothing Then
        mChartManager.RefreshCurrentSheet
    End If

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
Private Sub mFileMgr_ActiveFileChanged(f As clsFileXLS)
    On Error GoTo ErrHandler

    Dim filePath As String
    Dim isOpp As Boolean

    If f Is Nothing Then
        filePath = ""
        isOpp = False
    Else
        filePath = f.FullName
        isOpp = mFileMgr.State.isOpportunityFile
    End If

    LogDebug MODULE_NAME, "[mFileMgr_ActiveFileChanged] -> ActiveFileSessionChanged: " & filePath

    ' Emitir evento semantico (sin referencias a clsFileXLS)
    RaiseEvent ActiveFileSessionChanged(filePath, isOpp)

    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "[mFileMgr_ActiveFileChanged] Error traduciendo evento", , Err.Description
End Sub
