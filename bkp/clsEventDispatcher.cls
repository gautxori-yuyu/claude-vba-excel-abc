VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("5-Aplicac (Coord)")
Option Explicit

Private Const MODULE_NAME As String = "clsEventDispatcher"


'@Folder("3-Aplicac (Coord)")

'TODO: decidir si las macros se convierten en metodos de clase
Public Sub Dispatch(controlID As String)
    Select Case controlID
        Case "btnCompSheets": MostrarComparador
        Case "btnDirtyRecalc": AplicarDirtyATodasLasHojasConFormulas
        Case "btnEvalUDFs": ReemplazarUDFsEnFormulas
        Case "btnRowsHeight": AjustarAltoFilasSegunColor
        Case "btnEditableBook": modMACROWbkEditableCleaning.LimpiarLibroActual
        Case "btnFitForPrint": modMACROWbkEditableFormatting.AjustarSelWSheetsParaImpresionPDF
        Case "btnExportVBAProject": modMACROImportExportMacros.ExportarComponentesVBA
        Case "btnImportVBAProject": modMACROImportExportMacros.ImportarComponentesVBA
        Case "btnOpenLog": mod_Logger.AbrirLog
        Case "btnBkpVBASrc": CrearBackupCodigoVBA
        Case "btnSyncProcsWithXLAMSheet": modMACROProceduresToWorksheet.WriteProcedimientosSheet_ConBackup
        
        Case "btnGenerarGraficos"
                On Error GoTo ErrorHandler
                Application.ScreenUpdating = False
                Call EjecutarGraficoEnLibroActivo
                App.ChartManager.RefreshCurrentSheet
        Case "btnInvertirSeries"
                On Error GoTo ErrorHandler
                Application.ScreenUpdating = False
                Call InvertirEjesDelGraficoActivo
        Case "btnCGASING"
                On Error GoTo ErrorHandler
                Application.ScreenUpdating = False
                Call FixCGASING
        Case "btnNuevaOp"
                App.OpportunitiesMgr.CreaOportunidad
                App.ribbon.InvalidarControl "ddlOportunidades"
        Case "btnConfig"
                App.OpportunitiesMgr.Conf.ShowUI
        Case "btnValidationsToNames"
                Call modAPPBudgetQuotesUtilids.AplicarNombresAValidacionesCeldasActWBConReporte
        
        Case "btnToggleXLAM"
            ThisWorkbook.IsAddin = Not (ThisWorkbook.IsAddin)
        Case "btnOpRefresh"
            App.OpportunitiesMgr.actualizarColeccionOportunidades
            'FIXME: La invalidacion de ribbon no debería aparecer aqui: ¿tal vez en el coordinador de eventos?
            App.ribbon.InvalidarRibbon
            'App.Ribbon.InvalidarControl "ddlOportunidades"
    End Select

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Public Function GetRibbonItemsNr(controlID As String) As Long
    Select Case controlID
        Case "ddlOportunidades"
            ' MEZCLA LOGICA DE NEGOCIO CON EVENTOS -> deberia resolver el "coordinador de eventos"
            GetRibbonItemsNr = App.OpportunitiesMgr.numOpportunities
    End Select
End Function

Public Function GetRibbonItemLabel(controlID As String, Optional Index As Integer = -1) As String
    Select Case controlID
        Case "ddlOportunidades"
            ' MEZCLA LOGICA DE NEGOCIO CON EVENTOS -> deberia resolver el "coordinador de eventos"
            If Index < 0 Then Err.Raise vbObjectError + 600, , "No se ha proporcionado indice"
            GetRibbonItemLabel = App.OpportunitiesMgr.OportunityLabel(Index)
    End Select
End Function

' resuelve un cambio de seleccion en un item de seleccion multiple del ribbon
Public Sub SetRibbonSelectionIndex(controlID As String, Index As Integer, id As String)
    Select Case controlID
        Case "ddlOportunidades"
            ' Consulta la lógica de negocio real
            App.OpportunitiesMgr.CurrOpportunity = Index
            LogInfo MODULE_NAME, "[SetItemSelectionIndex] - Cambiada oportunidad activa: " & id
    End Select
    ' FIXME: invalidar, refrescar el UI? yo creo que estas funciones son llamadas CUANDO SE HA
    ' SOLICITADO UNA INVALIDACION y que NO debería forzarse una nueva invalidacion desde aqui...
    App.ribbon.InvalidarControl controlID
End Sub

Public Function GetRibbonControlEnabled(controlID As String) As Boolean
    Dim enabled As Boolean
    Select Case controlID
    Case "btnGenerarGraficos"                    ' GetGraficoEnabled
        enabled = EsFicheroOportunidad() And EsValidoGenerarGrafico()
    Case "btnInvertirSeries"                     ' GetInvertirEjesEnabled
        enabled = Not (ActiveChart Is Nothing)
        If enabled Then enabled = App.ChartManager.EsValidoInvertirEjes()
    Case "btnCGASING"                            ' GetCGASINGEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = IsDefaultCGasIngSheet()
    Case "btnNuevaOp"                            ' GetNuevaOportunidadEnabled
        enabled = True
    End Select
    GetRibbonControlEnabled = enabled
    'InvalidarControl controlID ' esto ESTA DE SOBRA!!
    LogDebug MODULE_NAME, "[GetRibbonControlEnabled] - ¿" & controlID & "?: " & CBool(enabled)
End Function

' TODO: PTE PASAR EL RESTO DE LLAMADAS DE CALLBACKS DE RIBBON, supertips, fijar item sel en dropbox, etc
