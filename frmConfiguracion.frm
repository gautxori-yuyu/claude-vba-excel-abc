VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfiguracion 
   Caption         =   "Configuración"
   ClientHeight    =   8715.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9180.001
   OleObjectBlob   =   "frmConfiguracion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================================
' FORMULARIO: frmConfiguracion
' DESCRIPCIÓN: Formulario para configurar los parámetros de la aplicación ABC_ofertas. Permite al usuario
'              establecer rutas de carpetas, configurar números SAM, seleccionar carpetas de imágenes y
'              planos de compresores, etc. Todos los cambios se persisten en la configuración de la aplicación.
' ==============================================================================================================

'@Folder "2-Servicios.Configuracion"
Option Explicit

' -------------------------------------------------------------------------------------------------------------
' INICIALIZACIÓN DEL FORMULARIO
' -------------------------------------------------------------------------------------------------------------

'@Description: Inicializa el formulario cargando la configuración actual desde la aplicación
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: clsAplicacion (variable App)
'@Note: Se ejecuta automáticamente al abrir el formulario. Carga todos los valores actuales de
'        configuración en los controles del formulario
Private Sub UserForm_Initialize()
    CargarConfiguracion
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

'@Description: Carga los valores de configuración actuales desde la aplicación a los controles del formulario
'@Scope: Private
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: App.Configuration (objeto de configuración)
'@Note: Sincroniza todos los TextBox y ListBox con los valores almacenados en la configuración
Private Sub CargarConfiguracion()
    
    On Error GoTo ErrHandler
    ' Cargar cada ruta desde el registro y mostrarla
    
    ' Cargar cada ruta desde el registro y mostrarla
    TextBoxRutaOportunidades.text = App.Configuration.RutaOportunidades
    TextBoxRutaPlantillas.text = App.Configuration.RutaPlantillas
    TextBoxRutaOfergas.text = App.Configuration.RutaOfergas
    TextBoxRutaGasVBNet.text = App.Configuration.RutaGasVBNet
    TextBoxRutaExcelCalcTempl.text = App.Configuration.RutaExcelCalcTempl
    ListBoxComprImgs.List = App.Configuration.ListComprImgs
    ListBoxComprDrawPIDs.List = App.Configuration.ListComprDrawPIDs
    ListBoxComprImgs.ControlTipText = ListToText(ListBoxComprImgs)
    ListBoxComprDrawPIDs.ControlTipText = ListToText(ListBoxComprDrawPIDs)
    TextBoxSAM.text = App.Configuration.SAM
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "frmConfiguracion.CargarConfiguracion", _
              "Error inicializando formulario: " & Err.Description
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE CAMBIO DE CONTROLES - SINCRONIZACIÓN CON CONFIGURACIÓN
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el evento de cambio en el TextBox del número SAM, sincronizando el valor con
'              la configuración de la aplicación
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: App.Configuration.SAM
'@Note: Previene bucles infinitos verificando que el valor haya cambiado realmente antes de asignarlo.
'        Registra en Debug cuando se asigna el valor
Private Sub TextBoxSAM_AfterUpdate()
    ' Evitar bucle infinito: solo asignar si el valor ha cambiado
    If TextBoxSAM.text = "" Then Exit Sub
    If Not IsNumeric(TextBoxSAM.text) Then
        MsgBox "Valor no válido. Introduce un número entre 0 y 255", vbExclamation
    ElseIf CInt(TextBoxSAM.text) <> App.Configuration.SAM Then
        Debug.Print "[frmConfiguracion TextBoxSAM_Change] Asignación de SAM desde UserForm"
        App.Configuration.SAM = CInt(TextBoxSAM.text)
    End If
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE BOTONES - SELECCIÓN DE CARPETAS INDIVIDUALES
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el clic en el botón de selección de carpeta para Oportunidades
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaATextBox, App.Configuration.RutaOportunidades
'@Note: Abre un diálogo de selección de carpeta y actualiza el TextBox y la configuración
Private Sub CommandButtonSelFldOportunidades_Click()
    SeleccionarCarpetaATextBox TextBoxRutaOportunidades, "Seleccionar carpeta para Oportunidades"
    App.Configuration.RutaOportunidades = TextBoxRutaOportunidades.text
End Sub

'@Description: Maneja el clic en el botón de selección de carpeta para Plantillas
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaATextBox, App.Configuration.RutaPlantillas
'@Note: Abre un diálogo de selección de carpeta y actualiza el TextBox y la configuración
Private Sub CommandButtonSelFldPlantillas_Click()
    SeleccionarCarpetaATextBox TextBoxRutaPlantillas, "Seleccionar carpeta para Plantillas"
    App.Configuration.RutaPlantillas = TextBoxRutaPlantillas.text
End Sub

'@Description: Maneja el clic en el botón de selección de carpeta para Ofergas
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaATextBox, App.Configuration.RutaOfergas
'@Note: Abre un diálogo de selección de carpeta y actualiza el TextBox y la configuración
Private Sub CommandButtonSelFldOfergas_Click()
    SeleccionarCarpetaATextBox TextBoxRutaOfergas, "Seleccionar carpeta para Ofergas"
    App.Configuration.RutaOfergas = TextBoxRutaOfergas.text
End Sub

'@Description: Maneja el clic en el botón de selección de carpeta para GasVBNet
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaATextBox, App.Configuration.RutaGasVBNet
'@Note: Abre un diálogo de selección de carpeta y actualiza el TextBox y la configuración
Private Sub CommandButtonSelFldGasVBNet_Click()
    SeleccionarCarpetaATextBox TextBoxRutaGasVBNet, "Seleccionar carpeta para GasVBNet"
    App.Configuration.RutaGasVBNet = TextBoxRutaGasVBNet.text
End Sub

'@Description: Maneja el clic en el botón de selección de carpeta para plantillas de cálculos Excel
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaATextBox, App.Configuration.RutaExcelCalcTempl
'@Note: Abre un diálogo de selección de carpeta y actualiza el TextBox y la configuración
Private Sub CommandButtonSelFldExcelCalcTempl_Click()
    SeleccionarCarpetaATextBox TextBoxRutaExcelCalcTempl, "Seleccionar carpeta para plantillas de calculos"
    App.Configuration.RutaExcelCalcTempl = TextBoxRutaExcelCalcTempl.text
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE BOTONES - GESTIÓN DE LISTAS DE CARPETAS (IMÁGENES Y PLANOS)
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el clic en el botón para agregar carpetas de imágenes de compresores a la lista
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaAListBox, App.Configuration.ListComprImgs
'@Note: Permite múltiples carpetas de imágenes, se agregan al ListBox y se guardan en la configuración
Private Sub CommandButtonComprImgs_Click()
    SeleccionarCarpetaAListBox ListBoxComprImgs, "Seleccionar carpeta de imágenes de compresores"
    App.Configuration.ListComprImgs = ListBoxComprImgs.List
    ListBoxComprImgs.ControlTipText = ListToText(ListBoxComprImgs)
End Sub

'@Description: Maneja el clic en el botón para agregar carpetas de planos PID de compresores a la lista
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpetaAListBox, App.Configuration.ListComprDrawPIDs
'@Note: Permite múltiples carpetas de planos, se agregan al ListBox y se guardan en la configuración
Private Sub CommandButtonComprDrawPIDs_Click()
    SeleccionarCarpetaAListBox ListBoxComprDrawPIDs, "Seleccionar carpeta de planos de compresores"
    App.Configuration.ListComprDrawPIDs = ListBoxComprDrawPIDs.List
    ListBoxComprDrawPIDs.ControlTipText = ListToText(ListBoxComprDrawPIDs)
End Sub

'@Description: Maneja el clic en el botón para eliminar carpetas seleccionadas de la lista de imágenes
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: BorraItemsListBox, App.Configuration.ListComprImgs
'@Note: Elimina los items seleccionados del ListBox y actualiza la configuración
Private Sub CommandButtonDelComprImgs_Click()
    Call BorraItemsListBox(ListBoxComprImgs)
    App.Configuration.ListComprImgs = ListBoxComprImgs.List
End Sub

'@Description: Maneja el clic en el botón para eliminar carpetas seleccionadas de la lista de planos
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: BorraItemsListBox, App.Configuration.ListComprDrawPIDs
'@Note: Elimina los items seleccionados del ListBox y actualiza la configuración
Private Sub CommandButtonDelComprDrawPIDs_Click()
    Call BorraItemsListBox(ListBoxComprDrawPIDs)
    App.Configuration.ListComprDrawPIDs = ListBoxComprDrawPIDs.List
End Sub

' -------------------------------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES - SELECCIÓN Y GESTIÓN DE CARPETAS
' -------------------------------------------------------------------------------------------------------------

'@Description: Muestra un diálogo de selección de carpeta y devuelve la ruta seleccionada
'@Scope: Private
'@ArgumentDescriptions: rutaActual (String): Ruta actual para inicializar el diálogo
'   | titulo (String): Título del diálogo
'@Returns: String - Ruta seleccionada, o cadena vacía si se canceló
'@Dependencies: Application.FileDialog, RutaExiste
'@Note: Valida que la ruta seleccionada exista antes de devolverla
Private Function SeleccionarCarpeta(rutaActual As String, titulo As String)
    Dim dlg As FileDialog
    
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = titulo
        .AllowMultiSelect = False
        
        ' Inicializar en la ruta actual si existe
        If rutaActual <> "" And RutaExiste(rutaActual) Then
            .InitialFileName = rutaActual
        End If
        
        If .Show = -1 Then
            If RutaExiste(.SelectedItems(1)) Then SeleccionarCarpeta = .SelectedItems(1)
        End If
    End With
    
    Set dlg = Nothing
End Function

'@Description: Selecciona una carpeta y actualiza un TextBox con la ruta seleccionada
'@Scope: Private
'@ArgumentDescriptions: txtDestino (MSForms.TextBox): TextBox donde se mostrará la ruta seleccionada
'   | titulo (String): Título del diálogo de selección
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpeta
'@Note: Solo actualiza el TextBox si se seleccionó una carpeta válida
Private Sub SeleccionarCarpetaATextBox(txtDestino As MSForms.TextBox, titulo As String)
    Dim nuevaRuta As String
    
    nuevaRuta = SeleccionarCarpeta(txtDestino.text, titulo)
    If nuevaRuta <> "" Then txtDestino.text = nuevaRuta
End Sub

'@Description: Selecciona una carpeta y la agrega a un ListBox si no existe ya en la lista
'@Scope: Private
'@ArgumentDescriptions: LBDestino (MSForms.ListBox): ListBox donde se agregará la ruta
'   | titulo (String): Título del diálogo de selección
'@Returns: (ninguno)
'@Dependencies: SeleccionarCarpeta
'@Note: Verifica que la carpeta no esté duplicada antes de agregarla. Agrega en posición 0 (inicio)
Private Sub SeleccionarCarpetaAListBox(LBDestino As MSForms.ListBox, titulo As String)
    Dim nuevaRuta As String, i As Integer
    
    nuevaRuta = SeleccionarCarpeta(LBDestino.Value, titulo)
    If nuevaRuta <> "" Then
        For i = 0 To ListBoxComprImgs.ListCount - 1
            If StrComp(ListBoxComprImgs.List(i), nuevaRuta & "\", vbTextCompare) = 0 Then Exit Sub
        Next i
        Call LBDestino.AddItem(nuevaRuta, 0)
    End If
End Sub

'@Description: Elimina los items seleccionados de un ListBox
'@Scope: Private
'@ArgumentDescriptions: LBDestino (MSForms.ListBox): ListBox del cual eliminar items
'@Returns: (ninguno)
'@Dependencies: Ninguna
'@Note: Muestra mensaje si no hay items seleccionados. Actualiza el tooltip después de eliminar
Private Sub BorraItemsListBox(LBDestino As MSForms.ListBox)
    Dim idx As Long
    If LBDestino.ListIndex = 0 Then
        MsgBox "Selecciona una carpeta de la lista para borrar.", vbInformation, "Información"
        Exit Sub
    End If
    ' Read through each item in the listbox
    For idx = 0 To LBDestino.ListCount - 1
        ' Check if item at position i is selected
        If LBDestino.Selected(idx) Then
            LBDestino.RemoveItem idx
        End If
    Next idx
    LBDestino.ControlTipText = ListToText(LBDestino)
End Sub

' -------------------------------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES - CONVERSIÓN Y VISUALIZACIÓN
' -------------------------------------------------------------------------------------------------------------

'@Description: Convierte los items de un ListBox a texto separado por saltos de línea
'@Scope: Private
'@ArgumentDescriptions: LB (MSForms.ListBox): ListBox a convertir
'@Returns: String - Texto con todos los items separados por vbCrLf
'@Dependencies: Ninguna
'@Note: Útil para mostrar el contenido completo en tooltips
Function ListToText(LB As MSForms.ListBox)
    Dim idx As Integer
    For idx = 0 To LB.ListCount - 1
        If ListToText <> "" Then ListToText = ListToText & vbCrLf
        ListToText = ListToText & LB.List(idx)
    Next
End Function

' -------------------------------------------------------------------------------------------------------------
' FUNCIONES DE TOOLTIPS (DESHABILITADAS)
' -------------------------------------------------------------------------------------------------------------

'@Description: Muestra un tooltip personalizado con texto descriptivo
'@Scope: Private
'@ArgumentDescriptions: Ctrl (Control): Control sobre el cual mostrar el tooltip
'   | texto (String): Texto a mostrar
'@Returns: (ninguno)
'@Dependencies: lblTooltip (label del formulario)
'@Note: Posiciona el tooltip debajo del control. Actualmente no utilizado
Private Sub MostrarTooltip(Ctrl As control, ByVal texto As String)
    With Me.lblTooltip
        .Caption = texto
        .Width = 200                             ' ajusta según necesites
        .Top = Ctrl.Top + Ctrl.Height + 2
        .Left = Ctrl.Left
        .Visible = True
        .ZOrder 0
    End With
End Sub

'@Description: Oculta el tooltip personalizado
'@Scope: Private
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: lblTooltip (label del formulario)
'@Note: Actualmente no utilizado
Private Sub OcultarTooltip()
    Me.lblTooltip.Visible = False
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE MOUSE (DESHABILITADOS)
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el evento MouseMove sobre el ListBox de planos de compresores
'@Scope: Private (evento)
'@ArgumentDescriptions: Parámetros estándar del evento MouseMove
'@Returns: (ninguno)
'@Dependencies: Ninguna
'@Note: Código comentado - podría usarse para mostrar tooltip con contenido completo del ListBox
Private Sub ListBoxComprDrawPIDs_MouseMove( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)
    
    'MostrarTooltip Me.ListBoxComprDrawPIDs, _
    ListToText(ListBoxComprDrawPIDs)
End Sub

'@Description: Maneja el evento MouseMove sobre el UserForm
'@Scope: Private (evento)
'@ArgumentDescriptions: Parámetros estándar del evento MouseMove
'@Returns: (ninguno)
'@Dependencies: Ninguna
'@Note: Código comentado - podría usarse para ocultar tooltips al mover el mouse fuera de los controles
Private Sub UserForm_MouseMove( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)
    
    'OcultarTooltip
End Sub


