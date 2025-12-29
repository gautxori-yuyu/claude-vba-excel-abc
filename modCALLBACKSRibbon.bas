Attribute VB_Name = "modCALLBACKSRibbon"
' Módulo de integración con la Ribbon: gestiona visibilidad y ejecución de macros para gráficos de sensibilidad

'FIXME: DETECCIÓN Y RECUPERACIÓN DE OBJETOS RIBBON; en ocasiones el ribbon se pierde. Es necesario revisar que lo causa
'  Creo que casi siempre tiene que ver con que se desactive el XLAM, o se suspende la ejecución de VBA mediante STOP

'@Folder "2-Control de estado"
'@IgnoreModule ProcedureNotUsed
Option Explicit
Option Private Module

' ==========================================
' CALLBACK: Se llama al cargar el Ribbon
' ==========================================
Sub RibbonOnLoad(xlRibbon As IRibbonUI)
Attribute RibbonOnLoad.VB_ProcData.VB_Invoke_Func = " \n0"
    Debug.Print "[callback: RibbonOnLoad] - Inicio"
    On Error GoTo ErrorHandler
    ' inicializamos la referencia al ribbon en la aplicación
    App.RibbonHandler = xlRibbon
    
    Debug.Print "[callback: RibbonOnLoad] - ribbon cargado en la interfaz de excel"
    App.Ribbon.InvalidarRibbon
    
    Exit Sub
ErrorHandler:
    Debug.Print "[callback: RibbonOnLoad] - Error: " & vbCrLf & vbTab & Err.Description
End Sub

' ==========================================
' CALLBACKS DE MACROS
' ==========================================

Sub OnCompararHojas(control As IRibbonControl)
Attribute OnCompararHojas.VB_ProcData.VB_Invoke_Func = " \n0"
    MostrarComparador
End Sub

Public Sub OnChangeAlturaFilas(control As IRibbonControl)
Attribute OnChangeAlturaFilas.VB_ProcData.VB_Invoke_Func = " \n0"
    Debug.Print "[callback: CallbackAlturaFilas]"
    On Error GoTo Finalizar
    Call AjustarAltoFilasSegunColor
Finalizar:
End Sub

Public Sub OnMakeEditableBook(control As IRibbonControl)
    Call modMACROWbkEditableCleaning.LimpiarLibroActual
End Sub

Public Sub OnVBAExport(control As IRibbonControl)
    Call modMACROImportExportMacros.ExportarComponentesVBA
End Sub

Public Sub OnVBAImport(control As IRibbonControl)
Attribute OnVBAImport.VB_ProcData.VB_Invoke_Func = " \n0"
    Call modMACROImportExportMacros.ImportarComponentesVBA
End Sub

Public Sub OnToggleXLAMVisibility(control As IRibbonControl)
Attribute OnToggleXLAMVisibility.VB_ProcData.VB_Invoke_Func = " \n0"
    Debug.Print "[callback: OnToggleXLAMVisibility]"
    ThisWorkbook.IsAddin = Not (ThisWorkbook.IsAddin)
End Sub

' ==========================================
' CALLBACKS DE APLICACION
' ==========================================
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
Attribute OnGenerarGraficosDesdeCurvasRto.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnGenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes(control As IRibbonControl)
Attribute OnInvertirEjes.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnInvertirEjes
End Sub

Public Sub OnFormatearCGASING(control As IRibbonControl)
Attribute OnFormatearCGASING.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnFormatearCGASING
End Sub

Public Sub OnNuevaOportunidad(control As IRibbonControl)
Attribute OnNuevaOportunidad.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnNuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations(control As IRibbonControl)
Attribute OnReplaceWithNamesInValidations.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnReplaceWithNamesInValidations
End Sub

'--------------------------------------------------------------
' CALLBACKS DE CONFIGURACIÓN
'--------------------------------------------------------------

' Callback del botón de configuración
Sub OnConfigurador(control As IRibbonControl)
Attribute OnConfigurador.VB_ProcData.VB_Invoke_Func = " \n0"
    App.Ribbon.OnConfigurador
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
    Debug.Print "[callback: CallbackRefrescarOportunidades] - control de ribbon activado para actualizar la lista de oportunidades"
    App.OpportunitiesMgr.actualizarColeccionOportunidades
    App.Ribbon.InvalidarRibbon
    'App.Ribbon.InvalidarControl "ddlOportunidades"
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
    returnedVal = App.OpportunitiesMgr.numOpportunities
End Sub

'--------------------------------------------------------------
' @Description: Devuelve la etiqueta de cada oportunidad en el
' desplegable del Ribbon.
' Callback for ddlOportunidades getItemLabel
'--------------------------------------------------------------
' @Category: Información de archivo
' @ArgumentDescriptions: control: control del Ribbon|index: índice (base 0)|label: texto mostrado
'--------------------------------------------------------------
Sub GetOportunidadesLabel(control As IRibbonControl, Index As Integer, ByRef label)
Attribute GetOportunidadesLabel.VB_ProcData.VB_Invoke_Func = " \n0"
    label = App.OpportunitiesMgr.OportunityLabel(Index)
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
    Debug.Print "[callback: OnOportunidadesSeleccionada] - modificada la oportunidad seleccionada en el control de ribbon"
    App.OpportunitiesMgr.CurrOpportunity = Index
    ' invalidar, refrescar el UI
    App.Ribbon.InvalidarControl "ddlOportunidades"
End Sub

'Callback for ddlOportunidades getSelectedItemIndex
' Índice del elemento seleccionado
Sub GetSelectedOportunidadIndex(control As IRibbonControl, ByRef Index)
Attribute GetSelectedOportunidadIndex.VB_ProcData.VB_Invoke_Func = " \n0"
    Index = App.OpportunitiesMgr.CurrOpportunity
End Sub

' ==========================================
' CALLBACKS DE SUPERTIPS DINÁMICOS
' ==========================================
Sub GetSupertipRutaBaseOportunidades(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOportunidades)
End Sub

Sub GetSupertipRutaBasePlantillas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBasePlantillas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaPlantillas)
End Sub

Sub GetSupertipRutaBaseOfergas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOfergas.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOfergas)
End Sub

Sub GetSupertipRutaBaseGasVBNet(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseGasVBNet.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaGasVBNet)
End Sub

Sub GetSupertipRutaBaseCalcTmpl(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseCalcTmpl.VB_ProcData.VB_Invoke_Func = " \n0"
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaExcelCalcTempl)
End Sub

' Para mostrar la ruta actual en el supertip (dinámico)
Function GetSupertipRutaBase(ByVal strSettingRuta As String)
Attribute GetSupertipRutaBase.VB_ProcData.VB_Invoke_Func = " \n0"
    If strSettingRuta = "" Then strSettingRuta = "No configurada"
    GetSupertipRutaBase = "Ruta actual: " & strSettingRuta & vbCrLf & "Haz clic para cambiar..."
End Function

' ==========================================
' CALLBACKS GetEnabled (habilitar/deshabilitar controles)
' ==========================================
' Habilita el botón de gráfico si el fichero es válido y cumple condiciones internas
Public Sub GetGraficoEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetGraficoEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Ribbon.GetRibbonControlEnabled(control)
End Sub

' Habilita el botón de inversión de ejes si hay gráfico válido en contexto
Public Sub GetInvertirEjesEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetInvertirEjesEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Ribbon.GetRibbonControlEnabled(control)
End Sub

' Habilita el botón de procesado C-GAS-ING si hoja válida en contexto
Public Sub GetCGASINGEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetCGASINGEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Ribbon.GetRibbonControlEnabled(control)
End Sub

' Habilita el botón de creación de nuevas oportunidades
Public Sub GetNuevaOportunidadEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetNuevaOportunidadEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = App.Ribbon.GetRibbonControlEnabled(control)
End Sub

' Habilita el botón de cumplimentación de oferta FULL si hoja válida en contexto
Public Sub GetOfertaFullEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetOfertaFullEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = True                               ' EsValidoRellenarOferta()
End Sub

' Habilita el botón del menú contextual del Ribbon si el fichero tiene nombre válido
Public Sub GetMenuEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetMenuEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    enabled = EsFicheroOportunidad()
    enabled = True
    'App.Ribbon.InvalidarRibbon
End Sub

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
    returnedVal = App.RibbonState.RibbonStateDescription
End Sub

' ==========================================
' CALLBACKS getVisible
' ==========================================
'@Description: Callback getVisible de la pestaña "ABC"
Public Sub GetTabABCVisible(control As IRibbonControl, ByRef Visible)
Attribute GetTabABCVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    If Not App() Is Nothing Then
        Visible = App.RibbonState.IsRibbonTabVisible
    Else
        Visible = False
    End If
End Sub

Public Sub GetGrpDeveloperAdminVisible(control As IRibbonControl, ByRef Visible)
Attribute GetGrpDeveloperAdminVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    Visible = False
    If Not App() Is Nothing Then
        Visible = App.RibbonState.IsAdminGroupVisible
    End If
End Sub
