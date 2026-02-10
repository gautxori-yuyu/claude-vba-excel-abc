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
                App.OpportunitiesMgr.CreateNewOpportunity
                ' Ribbon se invalida automaticamente via evento OpportunityCollectionUpdated
                ' que clsRibbon escucha con WithEvents mOpportunities
        Case "btnConfig"
                App.Configuration.ShowUI
        Case "btnValidationsToNames"
                Call modAPPBudgetQuotesUtilids.AplicarNombresAValidacionesCeldasActWBConReporte

        Case "btnToggleXLAM"
            ThisWorkbook.IsAddin = Not (ThisWorkbook.IsAddin)
        Case "btnOpRefresh"
            App.OpportunitiesMgr.RefreshOpportunities
            ' Ribbon se invalida automaticamente via evento OpportunityCollectionUpdated
            ' que clsRibbon escucha con WithEvents mOpportunities
    End Select

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Public Function GetRibbonItemsNr(controlID As String) As Long
    Select Case controlID
        Case "ddlOportunidades"
            ' MEZCLA LOGICA DE NEGOCIO CON EVENTOS -> deberia resolver el "coordinador de eventos"
            GetRibbonItemsNr = App.OpportunitiesMgr.Count
    End Select
End Function

Public Function GetRibbonItemLabel(controlID As String, Optional Index As Integer = -1) As String
    Select Case controlID
        Case "ddlOportunidades"
            ' MEZCLA LOGICA DE NEGOCIO CON EVENTOS -> deberia resolver el "coordinador de eventos"
            If Index < 0 Then Err.Raise vbObjectError + 600, , "No se ha proporcionado indice"
            GetRibbonItemLabel = App.OpportunitiesMgr.GetOpportunityNameByIndex(Index)
    End Select
End Function

'@Description: Resuelve un cambio de seleccion en un dropdown del Ribbon
'@Note: Este callback se ejecuta DESPUES de que Excel ya proceso el cambio visual.
'       El Ribbon se invalida automaticamente via eventos del Manager correspondiente.
Public Sub SetRibbonSelectionIndex(controlID As String, Index As Integer, id As String)
    Select Case controlID
        Case "ddlOportunidades"
            App.OpportunitiesMgr.SetCurrentOpportunityByIndex Index
            LogInfo MODULE_NAME, "[SetRibbonSelectionIndex] Cambiada oportunidad activa: " & id
            ' clsRibbon escucha CurrentOpportunityChanged via WithEvents mOpportunities
    End Select
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
    LogDebug MODULE_NAME, "[GetRibbonControlEnabled] ¿" & controlID & "?: " & CBool(enabled)
End Function

' TODO: PTE PASAR EL RESTO DE LLAMADAS DE CALLBACKS DE RIBBON, supertips, fijar item sel en dropbox, etc
