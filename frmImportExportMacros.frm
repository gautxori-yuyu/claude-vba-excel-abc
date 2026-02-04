VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportExportMacros 
   Caption         =   "Seleccionar Proyecto de macros"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmImportExportMacros.frx":0000
   StartUpPosition =   1  'Centrar en propietario
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmImportExportMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================================
' FORMULARIO: frmImportExportMacros
' DESCRIPCIÃN: Formulario modal para seleccionar un libro (Workbook) de entre los abiertos actualmente o
'              los complementos (Add-ins) instalados. Usado para operaciones de importaciÃ³n/exportaciÃ³n de
'              componentes VBA.
' ==============================================================================================================

'@Folder "0-Developer"
Option Explicit

Private Const MODULE_NAME As String = "frmImportExportMacros"

' -------------------------------------------------------------------------------------------------------------
' VARIABLES PRIVADAS
' -------------------------------------------------------------------------------------------------------------

Private libroSeleccionado As Workbook

' -------------------------------------------------------------------------------------------------------------
' PROPIEDADES PÃBLICAS
' -------------------------------------------------------------------------------------------------------------

'@Description: Propiedad de solo lectura que devuelve el libro seleccionado por el usuario
'@Scope: Public
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Workbook - El libro seleccionado, o Nothing si no se seleccionÃ³ ninguno
'@Dependencies: libroSeleccionado (variable privada)
'@Note: El formulario debe cerrarse antes de acceder a esta propiedad
Public Property Get WorkbookSeleccionado() As Workbook
    Set WorkbookSeleccionado = libroSeleccionado
End Property

' -------------------------------------------------------------------------------------------------------------
' INICIALIZACIÃN DEL FORMULARIO
' -------------------------------------------------------------------------------------------------------------

'@Description: Inicializa el formulario cargando la lista de libros abiertos y complementos disponibles
'              en el ComboBox
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: Application.Workbooks, Application.AddIns
'@Note: Se ejecuta automÃ¡ticamente al crear el formulario. Incluye tanto libros normales como Add-ins
Private Sub UserForm_Initialize()
    Dim Wb As Workbook, wbaddin As AddIn
    For Each Wb In Application.Workbooks
        Me.cmbLibros.AddItem Wb.Name
    Next Wb
    For Each wbaddin In Application.AddIns
        Me.cmbLibros.AddItem wbaddin.Name
    Next wbaddin
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE BOTONES
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el clic en el botÃ³n Aceptar, validando y guardando la selecciÃ³n del usuario
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: libroSeleccionado, cmbLibros
'@Note: Valida que se haya seleccionado un libro y que exista en la colecciÃ³n Workbooks. Oculta el
'        formulario si la selecciÃ³n es vÃ¡lida
Private Sub btnAceptar_Click()
    Dim nombre As String
    nombre = Me.cmbLibros.Value
    If nombre <> "" Then
        On Error Resume Next
        Set libroSeleccionado = Workbooks(nombre)
        On Error GoTo 0
        If Not libroSeleccionado Is Nothing Then
            Me.hide
        Else
            MsgBox "No se pudo encontrar el libro.", vbExclamation
        End If
    Else
        MsgBox "Selecciona un libro.", vbInformation
    End If
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE CIERRE
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el evento de cierre del formulario, interceptando el cierre con la X
'@Scope: Private (evento)
'@ArgumentDescriptions: Cancel (Integer): Permite cancelar el cierre
'   | CloseMode (Integer): Indica el modo de cierre (X, cÃ³digo, etc)
'@Returns: (ninguno)
'@Dependencies: Ninguna
'@Note: Si el usuario cierra con la X, cancela el cierre real y solo oculta el formulario, permitiendo
'        que el cÃ³digo principal detecte que no se seleccionÃ³ ningÃºn libro
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then        ' CerrÃ³ con la X
        Cancel = True                            ' Evitar cerrar directamente
        Me.hide
    End If
End Sub

