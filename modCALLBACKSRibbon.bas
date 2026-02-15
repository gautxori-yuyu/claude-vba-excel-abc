Attribute VB_Name = "modCALLBACKSRibbon"
' Módulo de integración con la Ribbon: gestiona visibilidad y ejecución de macros para gráficos de sensibilidad

'FIXME: DETECCIÓN Y RECUPERACIÓN DE OBJETOS RIBBON; en ocasiones el ribbon se pierde. Es necesario revisar que lo causa
'  Creo que casi siempre tiene que ver con que se desactive el XLAM, o se suspende la ejecución de VBA mediante STOP

'@Folder "4-UI.Excel.Ribbon"
'@IgnoreModule ProcedureNotUsed
Option Private Module
Option Explicit

Private Const MODULE_NAME As String = "modCALLBACKSRibbon"


' ==========================================
' RIBBON RECOVERY CON CopyMemory
' ==========================================
' Permite recuperar el objeto IRibbonUI si la variable global se pierde
' (por ejemplo, al depurar, tras un error no manejado, o reset de VBA)
' ==========================================
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private glngRibPtr As LongPtr
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
    Private glngRibPtr As Long
#End If

Private gobjRibbonUI As IRibbonUI
Private bRibbonWasInitialized As Boolean

' ==========================================
' CALLBACK: Se llama al cargar el Ribbon
' ==========================================
Sub RibbonOnLoad(xlRibbon As IRibbonUI)
Attribute RibbonOnLoad.VB_ProcData.VB_Invoke_Func = " \n0"
    LogInfo MODULE_NAME, "[callback: RibbonOnLoad] Inicio"
    On Error GoTo ErrorHandler

    ' Guardar puntero para recuperación con CopyMemory
    Set gobjRibbonUI = xlRibbon
    glngRibPtr = ObjPtr(xlRibbon)
    bRibbonWasInitialized = True
    LogDebug MODULE_NAME, "[callback: RibbonOnLoad] Puntero guardado: " & glngRibPtr

    ' Persistir puntero en nombres Excel4 (sobrevive a resets de VBA)
    StoreRibbonInExcelNames xlRibbon

    ' Inicializamos la referencia al ribbon en la aplicación
    Dim mApp As clsApplication
    Set mApp = App
    
    If Not mApp.ribbon Is Nothing Then
        mApp.ribbon.Init xlRibbon
    
        LogInfo MODULE_NAME, "[callback: RibbonOnLoad] ribbon cargado en la interfaz de excel"
        mApp.ribbon.InvalidarRibbon
    End If
    Exit Sub
ErrorHandler:
    LogCurrentError MODULE_NAME, "[callback: RibbonOnLoad]"
End Sub

'@Description: Obtiene el objeto IRibbonUI, recuperándolo con CopyMemory si se perdió
'@Returns: IRibbonUI o Nothing si no se puede recuperar
Public Function GetRibbonFromMemory() As IRibbonUI
Attribute GetRibbonFromMemory.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Si tenemos la referencia directa, usarla
    If Not gobjRibbonUI Is Nothing Then
        Set GetRibbonFromMemory = gobjRibbonUI
        Exit Function
    End If

    ' Intentar recuperar usando CopyMemory si tenemos el puntero
    If glngRibPtr <> 0 Then
        LogInfo MODULE_NAME, "[GetRibbonFromMemory] Intentando recuperar ribbon desde puntero: " & glngRibPtr

        On Error GoTo RecoveryFailed
        Dim tempObj As Object
        CopyMemory tempObj, glngRibPtr, LenB(glngRibPtr)
        Set gobjRibbonUI = tempObj

        ' Limpiar para evitar errores de referencia circular/memoria (64-bit safe: ptrZero es LongPtr)
        Dim ptrZero As LongPtr
        CopyMemory tempObj, ptrZero, LenB(ptrZero)

        ' Verificar que el objeto recuperado es válido
        Dim testType As String
        testType = TypeName(gobjRibbonUI)
        If testType <> "Nothing" And testType <> "Empty" Then
            LogInfo MODULE_NAME, "[GetRibbonFromMemory] Ribbon recuperado exitosamente"
            Set GetRibbonFromMemory = gobjRibbonUI
            Exit Function
        End If
    End If

    ' Si llegamos aquí, no se pudo recuperar
    If bRibbonWasInitialized Then
        LogWarning MODULE_NAME, "[GetRibbonFromMemory] Ribbon perdido y no recuperable"
    End If
    Set GetRibbonFromMemory = Nothing
    Exit Function

RecoveryFailed:
    LogError MODULE_NAME, "[GetRibbonFromMemory] Error al recuperar ribbon", Err.Number, Err.Description
    Set GetRibbonFromMemory = Nothing
End Function

'@Description: Indica si el ribbon fue inicializado alguna vez en esta sesión
Public Function WasRibbonInitialized() As Boolean
Attribute WasRibbonInitialized.VB_ProcData.VB_Invoke_Func = " \n0"
    WasRibbonInitialized = bRibbonWasInitialized
End Function

'@Description: Obtiene el puntero guardado del Ribbon (para diagnóstico)
#If VBA7 Then
Public Function GetRibbonPointer() As LongPtr
Attribute GetRibbonPointer.VB_ProcData.VB_Invoke_Func = " \n0"
#Else
Public Function GetRibbonPointer() As Long
#End If
    GetRibbonPointer = glngRibPtr
End Function


' ==========================================
' CALLBACKS DE MACROS
' ==========================================

Sub OnCompararHojas(control As IRibbonControl)
Attribute OnCompararHojas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnCompararHojas]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnCompararHojas]"
End Sub

Sub OnDirtyRecalc(control As IRibbonControl)
Attribute OnDirtyRecalc.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnDirtyRecalc]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnDirtyRecalc]"
End Sub

Sub OnEvalUDFs(control As IRibbonControl)
Attribute OnEvalUDFs.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnEvalUDFs]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnEvalUDFs]"
End Sub

Public Sub OnChangeAlturaFilas(control As IRibbonControl)
Attribute OnChangeAlturaFilas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnChangeAlturaFilas]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnChangeAlturaFilas]"
End Sub

Public Sub OnMakeEditableBook(control As IRibbonControl)
Attribute OnMakeEditableBook.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnMakeEditableBook]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnMakeEditableBook]"
End Sub

Public Sub OnFitForPrint(control As IRibbonControl)
Attribute OnFitForPrint.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnFitForPrint]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnFitForPrint]"
End Sub

Public Sub OnVBAExport(control As IRibbonControl)
Attribute OnVBAExport.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBAExport]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnVBAExport]"
End Sub

Public Sub OnVBAImport(control As IRibbonControl)
Attribute OnVBAImport.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBAImport]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnVBAImport]"
End Sub

Public Sub OnOpenLog(control As IRibbonControl)
Attribute OnOpenLog.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnOpenLog]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnOpenLog]"
End Sub

Public Sub OnVBABackup(control As IRibbonControl)
Attribute OnVBABackup.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnVBABackup] Creada copia de seguridad del código en " & ThisWorkbook.Path & "\Backups"
    ' Mensaje al usuario sin revelar la ruta completa (seguridad)
    ShowTaskDialogError "Copia de seguridad", _
                        "Backup completado", _
                        "Se ha creado correctamente la copia de seguridad del código VBA."
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnVBABackup]"
End Sub

Public Sub OnProcMetadataSync(control As IRibbonControl)
Attribute OnProcMetadataSync.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnProcMetadataSync]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnProcMetadataSync]"
End Sub

Public Sub OnToggleXLAMVisibility(control As IRibbonControl)
Attribute OnToggleXLAMVisibility.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnToggleXLAMVisibility]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnToggleXLAMVisibility]"
End Sub

' ==========================================
' CALLBACKS DE APLICACION
' ==========================================
Public Sub OnGenerarGraficosDesdeCurvasRto(control As IRibbonControl)
Attribute OnGenerarGraficosDesdeCurvasRto.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnGenerarGraficosDesdeCurvasRto]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnGenerarGraficosDesdeCurvasRto]"
End Sub

Public Sub OnInvertirEjes(control As IRibbonControl)
Attribute OnInvertirEjes.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnInvertirEjes]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnInvertirEjes]"
End Sub

Public Sub OnFormatearCGASING(control As IRibbonControl)
Attribute OnFormatearCGASING.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnFormatearCGASING]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnFormatearCGASING]"
End Sub

Public Sub OnNuevaOportunidad(control As IRibbonControl)
Attribute OnNuevaOportunidad.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnNuevaOportunidad]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnNuevaOportunidad]"
End Sub

Public Sub OnReplaceWithNamesInValidations(control As IRibbonControl)
Attribute OnReplaceWithNamesInValidations.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnReplaceWithNamesInValidations]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnReplaceWithNamesInValidations]"
End Sub

'--------------------------------------------------------------
' CALLBACKS DE CONFIGURACIÓN
'--------------------------------------------------------------

' Callback del botón de configuración
Sub OnConfigurador(control As IRibbonControl)
Attribute OnConfigurador.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: OnConfigurador]"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnConfigurador]"
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
    On Error GoTo ErrHandler
    App.Dispatcher.Dispatch (control.id)
    LogInfo MODULE_NAME, "[callback: CallbackRefrescarOportunidades] control de ribbon activado para actualizar la lista de oportunidades"
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[CallbackRefrescarOportunidades]"
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
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    returnedVal = App.Dispatcher.GetRibbonItemsNr(control.id)
    If Err.Number <> 0 Then returnedVal = 0 ' Devolver valor por defecto seguro
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
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    Label = App.Dispatcher.GetRibbonItemLabel(control.id, Index)
    If Err.Number <> 0 Then Label = "" ' Devolver valor por defecto seguro
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
    On Error GoTo ErrHandler
    ' supongo que el id es el texto de la opcion...
    Call App.Dispatcher.SetRibbonSelectionIndex(control.id, Index, id)
    LogInfo MODULE_NAME, "[callback: OnOportunidadesSeleccionada] modificada la oportunidad seleccionada en el control de ribbon: " & id
    Exit Sub
ErrHandler:
    LogCurrentError MODULE_NAME, "[OnOportunidadesSeleccionada]"
End Sub


' TODO: falta pasar al dispatcher el resto de callbacks

'Callback for ddlOportunidades getSelectedItemIndex
' Índice del elemento seleccionado
Sub GetSelectedOportunidadIndex(control As IRibbonControl, ByRef Index)
Attribute GetSelectedOportunidadIndex.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    
    Dim tempIndex As Long
    tempIndex = App.OpportunitiesMgr.CurrentIndex

    ' PREVENCIÓN DE ERROR: El Ribbon no acepta -1 como índice.
    ' Si no hay nada seleccionado (CurrentIndex = -1), devolvemos 0 para evitar el crash de UI.
    If tempIndex < 0 Then
        Index = 0
    Else
        Index = tempIndex
    End If
    
    Exit Sub
    
ErrHandler:
    ' En caso de error (App no lista), devolver 0 para evitar crash
    Index = 0
    LogCurrentError MODULE_NAME, "[GetSelectedOportunidadIndex]"
End Sub

' ==========================================
' CALLBACKS GetEnabled (habilitar/deshabilitar controles)
' ==========================================

' Habilita el botón de gráfico si el fichero es válido y cumple condiciones internas
Public Sub GetGraficoEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetGraficoEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
    If Err.Number <> 0 Then enabled = False ' Devolver valor por defecto seguro
End Sub

' Habilita el botón de inversión de ejes si hay gráfico válido en contexto
Public Sub GetInvertirEjesEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetInvertirEjesEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
    If Err.Number <> 0 Then enabled = False ' Devolver valor por defecto seguro
End Sub

' Habilita el botón de procesado C-GAS-ING si hoja válida en contexto
Public Sub GetCGASINGEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetCGASINGEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
    If Err.Number <> 0 Then enabled = False ' Devolver valor por defecto seguro
End Sub

' Habilita el botón de creación de nuevas oportunidades
Public Sub GetNuevaOportunidadEnabled(control As IRibbonControl, ByRef enabled)
Attribute GetNuevaOportunidadEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Dispatcher fallan
    enabled = App.Dispatcher.GetRibbonControlEnabled(control.id)
    If Err.Number <> 0 Then enabled = False ' Devolver valor por defecto seguro
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
    ' enabled = EsFicheroOportunidad()
    enabled = True
    'App.Ribbon.InvalidarRibbon
End Sub

' ==========================================
' CALLBACKS DE SUPERTIPS DINÁMICOS
' ==========================================
Sub GetSupertipRutaBaseOportunidades(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOportunidades.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Configuration fallan
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOportunidades)
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
End Sub

Sub GetSupertipRutaBasePlantillas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBasePlantillas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Configuration fallan
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaPlantillas)
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
End Sub

Sub GetSupertipRutaBaseOfergas(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseOfergas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Configuration fallan
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaOfergas)
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
End Sub

Sub GetSupertipRutaBaseGasVBNet(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseGasVBNet.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Configuration fallan
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaGasVBNet)
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
End Sub

Sub GetSupertipRutaBaseCalcTmpl(control As IRibbonControl, ByRef returnedVal)
Attribute GetSupertipRutaBaseCalcTmpl.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error Resume Next ' Proteger UI si App() o Configuration fallan
    returnedVal = GetSupertipRutaBase(App.Configuration.RutaExcelCalcTempl)
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
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
    On Error Resume Next ' Proteger UI si App() o Ribbon fallan
    returnedVal = App.ribbon.State.Description
    If Err.Number <> 0 Then returnedVal = "" ' Devolver valor por defecto seguro
End Sub

' ==========================================
' CALLBACKS getVisible
' ==========================================
'@Description: Callback getVisible de la pestaña "ABC"
Public Sub GetTabABCVisible(control As IRibbonControl, ByRef Visible)
Attribute GetTabABCVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    If App.ribbon Is Nothing Then
        LogDebug MODULE_NAME, "[GetTabABCVisible] Ribbon no disponible"
        Exit Sub
    End If
    On Error GoTo ErrHandler
    LogDebug MODULE_NAME, "[GetTabABCVisible] comprobando visibilidad Ribbon"
    Visible = App.ribbon.IsTabVisible()
    Exit Sub

ErrHandler:
    Visible = False
    LogCurrentError MODULE_NAME, "[GetTabABCVisible]"
    Err.Raise Err.Number, MODULE_NAME & "[GetTabABCVisible]", _
              "Error determinando la visibilidad del ribbon: " & Err.Description
End Sub
Public Sub GetOpGrpEnabled(control As IRibbonControl, ByRef Visible)
Attribute GetOpGrpEnabled.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    Visible = App.State.IsOpportunityMgrEnabled
    LogDebug MODULE_NAME, "[GetOpGrpEnabled] IsOpportunityMgrEnabled=" & CStr(Visible)
    Exit Sub

ErrHandler:
    Visible = False
    LogCurrentError MODULE_NAME, "[GetOpGrpEnabled]"
    Err.Raise Err.Number, MODULE_NAME & "[GetOpGrpEnabled]", _
              "Error determinando la visibilidad del grupo de gestion de oportunidad actual: " & Err.Description
End Sub

Public Sub GetGrpDeveloperAdminVisible(control As IRibbonControl, ByRef Visible)
Attribute GetGrpDeveloperAdminVisible.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    Visible = App.ribbon.IsAdminGroupVisible
    Exit Sub

ErrHandler:
    Visible = False
    LogCurrentError MODULE_NAME, "[GetGrpDeveloperAdminVisible]"
    Err.Raise Err.Number, MODULE_NAME & "[GetOpGrpEnabled]", _
              "Error determinando la visibilidad del grupo de Administración: " & Err.Description
End Sub
