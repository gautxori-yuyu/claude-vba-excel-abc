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
' DESCRIPCIÓN: Formulario modal para seleccionar un libro (Workbook) de entre los abiertos actualmente o
'              los complementos (Add-ins) instalados. Usado para operaciones de importación/exportación de
'              componentes VBA.
' ==============================================================================================================

'@Folder "0-Developer"
Option Explicit

' -------------------------------------------------------------------------------------------------------------
' VARIABLES PRIVADAS
' -------------------------------------------------------------------------------------------------------------

Private libroSeleccionado As Workbook

' -------------------------------------------------------------------------------------------------------------
' PROPIEDADES PÚBLICAS
' -------------------------------------------------------------------------------------------------------------

'@Description: Propiedad de solo lectura que devuelve el libro seleccionado por el usuario
'@Scope: Public
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Workbook - El libro seleccionado, o Nothing si no se seleccionó ninguno
'@Dependencies: libroSeleccionado (variable privada)
'@Note: El formulario debe cerrarse antes de acceder a esta propiedad
Public Property Get WorkbookSeleccionado() As Workbook
    Set WorkbookSeleccionado = libroSeleccionado
End Property

' -------------------------------------------------------------------------------------------------------------
' INICIALIZACIÓN DEL FORMULARIO
' -------------------------------------------------------------------------------------------------------------

'@Description: Inicializa el formulario cargando la lista de libros abiertos y complementos disponibles
'              en el ComboBox
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: Application.Workbooks, Application.AddIns
'@Note: Se ejecuta automáticamente al crear el formulario. Incluye tanto libros normales como Add-ins
Private Sub UserForm_Initialize()
    Dim wb As Workbook, wbaddin As AddIn
    For Each wb In Application.Workbooks
        Me.cmbLibros.AddItem wb.Name
    Next wb
    For Each wbaddin In Application.AddIns
        Me.cmbLibros.AddItem wbaddin.Name
    Next wbaddin
End Sub

' -------------------------------------------------------------------------------------------------------------
' EVENTOS DE BOTONES
' -------------------------------------------------------------------------------------------------------------

'@Description: Maneja el clic en el botón Aceptar, validando y guardando la selección del usuario
'@Scope: Private (evento)
'@ArgumentDescriptions: (sin argumentos)
'@Returns: (ninguno)
'@Dependencies: libroSeleccionado, cmbLibros
'@Note: Valida que se haya seleccionado un libro y que exista en la colección Workbooks. Oculta el
'        formulario si la selección es válida
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
'   | CloseMode (Integer): Indica el modo de cierre (X, código, etc)
'@Returns: (ninguno)
'@Dependencies: Ninguna
'@Note: Si el usuario cierra con la X, cancela el cierre real y solo oculta el formulario, permitiendo
'        que el código principal detecte que no se seleccionó ningún libro
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then        ' Cerró con la X
        Cancel = True                            ' Evitar cerrar directamente
        Me.hide
    End If
End Sub

