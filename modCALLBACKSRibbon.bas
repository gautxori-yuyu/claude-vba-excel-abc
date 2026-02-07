Attribute VB_Name = "modCALLBACKSRibbon"
' Módulo de integración con la Ribbon: gestiona visibilidad y ejecución de macros para gráficos de sensibilidad

'FIXME: DETECCIÓN Y RECUPERACIÓN DE OBJETOS RIBBON; en ocasiones el ribbon se pierde. Es necesario revisar que lo causa
'  Creo que casi siempre tiene que ver con que se desactive el XLAM, o se suspende la ejecución de VBA mediante STOP

'@Folder "2-Infraestructura.Excel.Ribbon"
'@IgnoreModule ProcedureNotUsed
Option Private Module
Option Explicit

Private Const MODULE_NAME As String = "modCALLBACKSRibbon"

' ==========================================
' CALLBACK: Se llama al cargar el Ribbon
' ==========================================
Sub RibbonOnLoad(xlRibbon As IRibbonUI)
Attribute RibbonOnLoad.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: RibbonOnLoad] Inicio"
    On Error GoTo ErrorHandler
    ' inicializamos la referencia al ribbon en la aplicación
    Dim mApp As clsApplication
    Set mApp = App
    mApp.ribbon.Init xlRibbon
    
    LogInfo MODULE_NAME, "[callback: RibbonOnLoad] ribbon cargado en la interfaz de excel"
    mApp.ribbon.InvalidarRibbon
    
    Exit Sub
ErrorHandler:
    LogCurrentError MODULE_NAME, "[callback: RibbonOnLoad]"
End Sub

' ==========================================
' CALLBACKS DE MACROS
' ==========================================

Sub OnCompararHojas(control As IRibbonControl)
Attribute OnCompararHojas.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnCompararHojas]"
End Sub

Sub OnDirtyRecalc(control As IRibbonControl)
Attribute OnDirtyRecalc.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnDirtyRecalc]"
End Sub

Sub OnEvalUDFs(control As IRibbonControl)
Attribute OnEvalUDFs.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnEvalUDFs]"
End Sub

Public Sub OnChangeAlturaFilas(control As IRibbonControl)
Attribute OnChangeAlturaFilas.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnChangeAlturaFilas]"
End Sub

Public Sub OnMakeEditableBook(control As IRibbonControl)
Attribute OnMakeEditableBook.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnMakeEditableBook]"
End Sub

Public Sub OnFitForPrint(control As IRibbonControl)
Attribute OnFitForPrint.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnFitForPrint]"
End Sub

Public Sub OnVBAExport(control As IRibbonControl)
Attribute OnVBAExport.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBAExport]"
End Sub

Public Sub OnVBAImport(control As IRibbonControl)
Attribute OnVBAImport.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBAImport]"
End Sub

Public Sub OnOpenLog(control As IRibbonControl)
Attribute OnOpenLog.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnOpenLog]"
End Sub

Public Sub OnVBABackup(control As IRibbonControl)
Attribute OnVBABackup.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBABackup] Creada copia de seguridad del código en " & ThisWorkbook.Path & "\Backups"
    MsgBox "Creada copia de seguridad del código en " & _
            ThisWorkbook.Path & "\Backups", vbInformation, "Copia de seguridad"
End Sub

Public Sub OnProcMetadataSync(control As IRibbonControl)
Attribute OnProcMetadataSync.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnProcMetadataSync]"
End Sub

Public Sub OnToggleXLAMVisibility(control As IRibbonControl)
Attribute OnToggleXLAMVisibility.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnToggleXLAMVisibility]"
End Sub

' ==========================================
' CALLBACKS DE APLICACION
' ==========================================
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
Attribute OnGenerarGraficosDesdeCurvasRto.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnGenerarGraficosDesdeCurvasRto]"
End Sub

Public Sub OnInvertirEjes(control As IRibbonControl)
Attribute OnInvertirEjes.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnInvertirEjes]"
End Sub

Public Sub OnFormatearCGASING(control As IRibbonControl)
Attribute OnFormatearCGASING.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnFormatearCGASING]"
End Sub

Public Sub OnNuevaOportunidad(control As IRibbonControl)
Attribute OnNuevaOportunidad.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnNuevaOportunidad]"
End Sub

Public Sub OnReplaceWithNamesInValidations(control As IRibbonControl)
Attribute OnReplaceWithNamesInValidations.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnReplaceWithNamesInValidations]"
End Sub

'--------------------------------------------------------------
' CALLBACKS DE CONFIGURACIÓN
'--------------------------------------------------------------

' Callback del botón de configuración
Sub OnConfigurador(control As IRibbonControl)
Attribute OnConfigurador.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnConfigurador]"
End Sub

'--------------------------------------------------------------
' CALLBACKS DEL DROPDOWN DE OPORTUNIDADES
'--------------------------------------------------------------

'FIXME: revisar la secuencia de eventos con el dropdown / box!!:
'  actualmente la sucesión de eventos relacionados con ese drop down no está bien coordinada.
'  revisar los eventos OpportunityChanged y su relación con CurrOpportunity y ProcesarCambiosEnOportunidades,
'  y el resto de eventos relacionados

'--------------------------------------------------------------
' @Description: Callback del botón de refresco de oportunidades.
' Callback for btnOpRefresh CallbackRefrescarOportunidades
' Refresca el listado de subcarpetas y actualiza el desplegable
' del Ribbon.
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: control: control del Ribbon que dispara el evento
'--------------------------------------------------------------
Public Sub CallbackRefrescarOportunidades(control As IRibbonControl)
Attribute CallbackRefrescarOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: CallbackRefrescarOportunidades] control de ribbon activado para actualizar la lista de oportunidades"
End Sub

'--------------------------------------------------------------
' @Description: Devuelve el número de oportunidades disponibles (número de elementos del desplegable).
' Callback for ddlOportunidades getItemCount
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: control: control del Ribbon|getItemCount: valor devuelto
'--------------------------------------------------------------
Sub GetOportunidadesCount(control As IRibbonControl, ByRef returnedVal)
Attribute GetOportunidadesCount.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = App.Dispatcher.GetRibbonItemsNr(control.id)
End Sub

'--------------------------------------------------------------
' @Description: Devuelve la etiqueta de cada oportunidad en el
' desplegable del Ribbon.
' Callback for ddlOportunidades getItemLabel
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: control: control del Ribbon|index: índice (base 0)|label: texto mostrado
'--------------------------------------------------------------
Sub GetOportunidadesLabel(control As IRibbonControl, Index As Integer, ByRef Label)
Attribute GetOportunidadesLabel.VB_ProcData.VB_Invoke_Func = " \n0"
    Label = App.Dispatcher.GetRibbonItemLabel(control.id, Index)
End Sub

'--------------------------------------------------------------
' @Description: Gestiona el evento de selección de oportunidad.
' Dispara el evento OpportunityChanged de la clase clsOpportunitiesMgr.
' Callback for ddlOportunidades onAction
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: control: control del Ribbon|id: identificador del control|index: índice seleccionado
'--------------------------------------------------------------
Sub OnOportunidadesSeleccionada(control As IRibbonControl, id As String, Index As Integer)
Attribute OnOportunidadesSeleccionada.VB_ProcData.VB_Invoke_Func = " \n0"
    ' supongo que el id es el texto de la opcion...
    Call App.Dispatcher.SetRibbonSelectionIndex(control.id, Index, id)
    LogInfo MODULE_NAME, "[callback: OnOportunidadesSeleccionada] modificada la oportunidad seleccionada en el control de ribbon: " & id
End Sub


' TODO: falta pasar al dispatcher el resto de callbacks

'Callback for ddlOportunidades getSelectedItemIndex
' Índice del elemento seleccionado
Sub GetSelectedOportunidadIndex(control As IRibbonControl, ByRef Index)
Attribute GetSelectedOportunidadIndex.VB_ProcData.VB_Invoke_Func = " \n0"
    Index = App.OpportunitiesMgr.CurrOpportunity
End Sub

' ==========================================
' CALLBACKS GetEnabled (habilitar/deshabilitar controles)
' ==========================================
' Habilita el botón de gráfico si el fichero es válido y cumple condiciones internas
Public Sub GetGraficoEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetGraficoEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
End Sub

' Habilita el botón de inversión de ejes si hay gráfico válido en contexto
Public Sub GetInvertirEjesEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetInvertirEjesEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
End Sub

' Habilita el botón de procesado C-GAS-ING si hoja válida en contexto
Public Sub GetCGASINGEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetCGASINGEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
End Sub

' Habilita el botón de creación de nuevas oportunidades
Public Sub GetNuevaOportunidadEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetNuevaOportunidadEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
End Sub

' Habilita el botón de cumplimentación de oferta FULL si hoja válida en contexto
Public Sub GetOfertaFullEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetOfertaFullEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = True                               ' EsValidoRellenarOferta()
End Sub

Public Sub GetOpenLogEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetOpenLogEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = GetLogFilePath <> ""
End Sub

' Habilita el botón del menú contextual del Ribbon si el fichero tiene nombre válido
Public Sub GetMenuEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetMenuEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = EsFicheroOportunidad()
    enabled = True
    'App.Ribbon.InvalidarRibbon
End Sub

' ==========================================
' CALLBACKS DE SUPERTIPS DINÁMICOS
' ==========================================
Sub GetSupertipRutaBaseOportunidades(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.OpportunitiesMgr.Conf.RutaOportunidades)
End Sub

Sub GetSupertipRutaBasePlantillas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBasePlantillas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.OpportunitiesMgr.Conf.RutaPlantillas)
End Sub

Sub GetSupertipRutaBaseOfergas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOfergas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.OpportunitiesMgr.Conf.RutaOfergas)
End Sub

Sub GetSupertipRutaBaseGasVBNet(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseGasVBNet.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.OpportunitiesMgr.Conf.RutaGasVBNet)
End Sub

Sub GetSupertipRutaBaseCalcTmpl(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseCalcTmpl.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.OpportunitiesMgr.Conf.RutaExcelCalcTempl)
End Sub

' Para mostrar la ruta actual en el supertip (dinámico)
Function GetSupertipRutaBase(ByVal strSettingRuta As String)
Attribute GetSupertipRutaBase.VB_ProcData.VB_Invoke_Func = " \n0"
    If strSettingRuta = "" Then strSettingRuta = "No configurada"
    GetSupertipRutaBase = "Ruta actual: " & strSettingRuta & vbCrLf & "Haz clic para cambiar..."
End Function

' ==========================================
' CALLBACKS GetLabel (cambia la etiqueta de controles)
' ==========================================
Public Sub GetLabelToggleXLAM(control As IRibbonControl, ByRef returnedVal)
Attribute GetLabelToggleXLAM.VB_ProcData.VB_Invoke_Func = " \n0"
    If ThisWorkbook.IsAddin Then
        returnedVal = "Mostrar XLAM"
    Else
        returnedVal = "Ocultar XLAM"
    End If
End Sub

Public Sub GetLabelGrpConfiguracion(control As IRibbonControl, ByRef returnedVal)
Attribute GetLabelGrpConfiguracion.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = App.ribbon.State.Description
End Sub

' ==========================================
' CALLBACKS getVisible
' ==========================================
'@Description: Callback getVisible de la pestaña "ABC"
Public Sub GetTabABCVisible(control As IRibbonControl, ByRef Visible)
Attribute GetTabABCVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    Visible = App.ribbon.IsTabVisible()
    Exit Sub

ErrHandler:
    Visible = False
End Sub

Public Sub GetGrpDeveloperAdminVisible(control As IRibbonControl, ByRef Visible)
Attribute GetGrpDeveloperAdminVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    Visible = App.ribbon.State.IsAdminGroupVisible
    Exit Sub

ErrHandler:
    Visible = False
End Sub


