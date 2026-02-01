VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComparadorHojas 
   Caption         =   "UserForm1"
   ClientHeight    =   1785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   OleObjectBlob   =   "frmComparadorHojas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmComparadorHojas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder "MACROS"
Option Explicit

Private Const MODULE_NAME As String = "frmComparadorHojas"

Private colorSeleccionado As Long

' Evento al inicializar el formulario
Private Sub UserForm_Initialize()
    Dim Wb As Workbook
    
    On Error GoTo ErrHandler
    
    ' Llenar ComboBox de Libros
    cboLibro1.Clear
    cboLibro2.Clear
    For Each Wb In Application.Workbooks
        cboLibro1.AddItem Wb.Name
        cboLibro2.AddItem Wb.Name
    Next Wb
    
    ' Valores por defecto - libro activo
    If cboLibro1.ListCount > 0 Then
        cboLibro1.Value = ActiveWorkbook.Name
        cboLibro2.Value = ActiveWorkbook.Name
    End If
    
    ' Llenar hojas del libro activo
    Call ActualizarHojas1
    Call ActualizarHojas2
    
    ' Checkboxes por defecto
    chkColorear1.Value = False
    chkColorear2.Value = True
    chkSoloBlanco.Value = False
    
    ' Color por defecto
    colorSeleccionado = RGB(255, 100, 200)
    btnColor.BackColor = colorSeleccionado
    
    ' Deshabilitar botón deshacer
    btnDeshacer.enabled = False
    
    btnSelRango1.ControlTipText = "Seleccionar rango de celdas a comparar."
    btnSelRango2.ControlTipText = "Seleccionar rango de celdas a comparar."
    
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "frmComparadorHojas.UserForm_Initialize", _
              "Error inicializando formulario: " & Err.Description
End Sub

' Actualizar hojas cuando cambia el libro 1
Private Sub cboLibro1_Change()
    Call ActualizarHojas1
End Sub

' Actualizar hojas cuando cambia el libro 2
Private Sub cboLibro2_Change()
    Call ActualizarHojas2
End Sub

' Actualizar lista de hojas para Libro 1
Private Sub ActualizarHojas1()
    Dim ws As Worksheet
    Dim Wb As Workbook
    
    cboHoja1.Clear
    
    If cboLibro1.Value <> "" Then
        On Error Resume Next
        Set Wb = Workbooks(cboLibro1.Value)
        On Error GoTo 0
        
        If Not Wb Is Nothing Then
            For Each ws In Wb.Worksheets
                If ws.Visible Then
                    cboHoja1.AddItem ws.Name
                End If
            Next ws
            If cboHoja1.ListCount > 0 Then cboHoja1.ListIndex = 0
        End If
    End If
End Sub

' Actualizar lista de hojas para Libro 2
Private Sub ActualizarHojas2()
    Dim ws As Worksheet
    Dim Wb As Workbook
    
    cboHoja2.Clear
    
    If cboLibro2.Value <> "" Then
        On Error Resume Next
        Set Wb = Workbooks(cboLibro2.Value)
        On Error GoTo 0
        
        If Not Wb Is Nothing Then
            For Each ws In Wb.Worksheets
                If ws.Visible Then
                    cboHoja2.AddItem ws.Name
                End If
            Next ws
            If cboHoja2.ListCount > 0 Then cboHoja2.ListIndex = 0
        End If
    End If
End Sub

' Botón para seleccionar Rango 1
Private Sub btnSelRango1_Click()
    Dim rng As range
    
    On Error Resume Next
    Set rng = Application.InputBox("Seleccione el rango en " & cboLibro1.Value & " - " & cboHoja1.Value, _
                                   "Rango Hoja 1", Type:=8)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        txtRango1.Value = rng.Address
    End If
End Sub

' Botón para seleccionar Rango 2
Private Sub btnSelRango2_Click()
    Dim rng As range
    
    On Error Resume Next
    Set rng = Application.InputBox("Seleccione el rango en " & cboLibro2.Value & " - " & cboHoja2.Value, _
                                   "Rango Hoja 2", Type:=8)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        txtRango2.Value = rng.Address
    End If
End Sub

Private Sub cboLibro1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    cboLibro1.Value = ActiveWorkbook.Name
    Call ActualizarHojas1
    cboHoja1.Value = ActiveSheet.Name
    On Error GoTo 0
End Sub

Private Sub cboHoja1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    cboLibro1.Value = ActiveWorkbook.Name
    Call ActualizarHojas1
    cboHoja1.Value = ActiveSheet.Name
    On Error GoTo 0
End Sub

Private Sub cboLibro2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    cboLibro2.Value = ActiveWorkbook.Name
    Call ActualizarHojas2
    cboHoja2.Value = ActiveSheet.Name
    On Error GoTo 0
End Sub

Private Sub cboHoja2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    cboLibro2.Value = ActiveWorkbook.Name
    Call ActualizarHojas2
    cboHoja2.Value = ActiveSheet.Name
    On Error GoTo 0
End Sub

' Botón Comparar
Private Sub btnComparar_Click()
    Dim wb1 As Workbook, wb2 As Workbook
    Dim Hoja1 As Worksheet, Hoja2 As Worksheet
    Dim rango1 As range, rango2 As range
    
    ' Validaciones
    If cboLibro1.Value = "" Or cboLibro2.Value = "" Then
        MsgBox "Debe seleccionar ambos libros.", vbExclamation
        Exit Sub
    End If
    
    If cboHoja1.ListIndex = -1 Or cboHoja2.ListIndex = -1 Then
        MsgBox "Debe seleccionar ambas hojas.", vbExclamation
        Exit Sub
    End If
    
    If Not chkColorear1.Value And Not chkColorear2.Value Then
        MsgBox "Debe seleccionar al menos una hoja para colorear.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener libros, hojas y rangos
    On Error GoTo ErrorHandler
    Set wb1 = Workbooks(cboLibro1.Value)
    Set wb2 = Workbooks(cboLibro2.Value)
    Set Hoja1 = wb1.Worksheets(cboHoja1.Value)
    Set Hoja2 = wb2.Worksheets(cboHoja2.Value)
    
    If Trim(txtRango1.Value) = "" Or Trim(txtRango2.Value) = "" Then
        If vbNo = MsgBox("No se han especificado rangos." & vbCrLf & _
                         "¿Desea comparar las hojas completas?", _
                         vbQuestion + vbYesNo, "Comparar hojas completas") Then
            Exit Sub
        End If
        Set rango1 = Hoja1.UsedRange
        Set rango2 = Hoja2.UsedRange
    Else
        Set rango1 = Hoja1.range(txtRango1.Value)
        Set rango2 = Hoja2.range(txtRango2.Value)
    End If
    On Error GoTo 0
    
    If wb1 Is Nothing Or wb2 Is Nothing Or Hoja1 Is Nothing Or Hoja2 Is Nothing Or _
       rango1 Is Nothing Or rango2 Is Nothing Then
        MsgBox "Error al obtener libros, hojas o rangos. Verifique los datos.", vbCritical
        Exit Sub
    End If
    
    Call CompararRangos(rango1, chkColorear1.Value, _
                        rango2, chkColorear2.Value, _
                        colorSeleccionado, chkSoloBlanco.Value)
    
    ' Habilitar botón deshacer
    btnDeshacer.enabled = True
    
    ' Mostrar hojas en paralelo automáticamente
    Call VerHojasEnParalelo(Hoja1, Hoja2, rango1, rango2)
    
    'MsgBox "Comparación completada. Las diferencias han sido resaltadas.", vbInformation
    
    Exit Sub
ErrorHandler:
    
End Sub

' Botón Deshacer
Private Sub btnDeshacer_Click()
    Dim wb1 As Workbook, wb2 As Workbook
    Dim Hoja1 As Worksheet, Hoja2 As Worksheet
    
    If Not HayComparacionActiva() Then
        MsgBox "No hay comparación activa para deshacer.", vbInformation
        Exit Sub
    End If
    
    ' Obtener libros y hojas (no necesitamos rangos para deshacer)
    On Error Resume Next
    Set wb1 = Workbooks(cboLibro1.Value)
    Set wb2 = Workbooks(cboLibro2.Value)
    Set Hoja1 = wb1.Worksheets(cboHoja1.Value)
    Set Hoja2 = wb2.Worksheets(cboHoja2.Value)
    On Error GoTo 0
    
    If Hoja1 Is Nothing Or Hoja2 Is Nothing Then
        MsgBox "Error al obtener hojas.", vbCritical
        Exit Sub
    End If
    
    ' Deshacer
    Call DeshacerComparacion
    
    ' Deshabilitar botón deshacer
    btnDeshacer.enabled = False
    
    MsgBox "Comparación deshecha. Los colores han sido restaurados.", vbInformation
End Sub

' Botón seleccionar color - CON COLOR PICKER NATIVO
Private Sub btnColor_Click()
    Dim RGBRed As Long, RGBGreen As Long, RGBBlue As Long
    Dim FullColorCode As Long
    
    ' Extraer componentes RGB del color actual
    RGBRed = colorSeleccionado Mod 256
    RGBGreen = (colorSeleccionado \ 256) Mod 256
    RGBBlue = (colorSeleccionado \ 65536) Mod 256
    
    ' Abrir el ColorPicker, aplicando el color actual como predeterminado
    If Application.Dialogs(xlDialogEditColor).Show(1, RGBRed, RGBGreen, RGBBlue) = True Then
        ' Obtener el color seleccionado
        FullColorCode = ActiveWorkbook.Colors(1)
        
        ' Actualizar el color seleccionado
        colorSeleccionado = FullColorCode
        btnColor.BackColor = colorSeleccionado
        
        ' Extraer nuevos componentes para el caption
        RGBRed = colorSeleccionado Mod 256
        RGBGreen = (colorSeleccionado \ 256) Mod 256
        RGBBlue = (colorSeleccionado \ 65536) Mod 256
        
        btnColor.Caption = "RGB(" & RGBRed & "," & RGBGreen & "," & RGBBlue & ")"
    End If
End Sub

' Evento al cerrar el formulario
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If HayComparacionActiva() Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("Los cambios de la comparación se convertirán en definitivos." & vbCrLf & _
                           "¿Continuar?", vbQuestion + vbYesNo, "Confirmar cierre")
        
        If respuesta = vbNo Then
            Cancel = True
        End If
    End If
End Sub


