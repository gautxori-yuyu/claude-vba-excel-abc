Attribute VB_Name = "modMACROUtilsExcelCheckbox"
'@Folder "Funciones auxiliares"
Option Explicit

Private Const MODULE_NAME As String = "modMACROUtilsExcelCheckbox"

' Inserta un checkbox vinculado a una celda de datos con validaciones completas
Sub InsertarCheckbox(Optional ByVal HojaDestino As String = "C.DATA", _
                     Optional ByVal ColumnaVinculo As String = "B", _
                     Optional ByVal MostrarCaption As Boolean = False, _
                     Optional ByVal BuscarTextoIzquierda As Boolean = True, _
                     Optional ByVal ValorInicial As Boolean = False, _
                     Optional ByVal TextoPersonalizado As String = "")
Attribute InsertarCheckbox.VB_ProcData.VB_Invoke_Func = " \n0"
    
    '----------------------------------------------------------------------
    ' PROCEDIMIENTO: InsertarCheckbox
    ' DESCRIPCIÓN:   Inserta un checkbox vinculado a una celda de datos
    '                con validaciones completas y manejo robusto de errores
    '
    ' PARÁMETROS OPCIONALES:
    '   - HojaDestino: Nombre de la hoja donde guardar el estado (por defecto "C.DATA")
    '   - ColumnaVinculo: Columna donde guardar TRUE/FALSE (por defecto "B")
    '   - MostrarCaption: Si muestra el texto del checkbox (por defecto False)
    '   - BuscarTextoIzquierda: Si busca texto en celdas a la izquierda (por defecto True)
    '   - ValorInicial: Estado inicial del checkbox (por defecto desmarcado)
    '   - TextoPersonalizado: Texto específico para el checkbox (anula búsqueda automática)
    '
    ' USO: Llamar desde la celda donde se quiere insertar el checkbox
    '----------------------------------------------------------------------
    
    On Error GoTo ManejoError
    
    '--- VALIDACIÓN 1: VERIFICAR QUE EXISTA UNA APLICACIÓN ACTIVA ---
    If Application Is Nothing Then
        MsgBox "No hay una instancia de Excel activa.", vbCritical, "Error de aplicación"
        Exit Sub
    End If
    
    '--- VALIDACIÓN 2: VERIFICAR QUE HAY UNA HOJA ACTIVA ---
    If ActiveSheet Is Nothing Then
        MsgBox "No hay ninguna hoja de cálculo activa.", vbExclamation, "Seleccione una hoja"
        Exit Sub
    End If
    
    Dim checkboxSheet As Worksheet
    Set checkboxSheet = ActiveSheet
    
    '--- VALIDACIÓN 3: VERIFICAR QUE EL ELEMENTO ACTIVO ES UNA CELDA ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, seleccione una celda antes de insertar el checkbox.", _
               vbExclamation, "Selección requerida"
        Exit Sub
    End If
    
    '--- VALIDACIÓN 4: VERIFICAR QUE EXISTE LA HOJA DESTINO ---
    Dim HojaExiste As Boolean
    HojaExiste = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = HojaDestino Then
            HojaExiste = True
            Exit For
        End If
    Next ws
    
    If ws Is checkboxSheet Then
        MsgBox "El checkbox no se puede insertar en la misma hoja en que se guarda el estado.", vbExclamation, "Operación cancelada"
        Exit Sub
    ElseIf Not HojaExiste Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("La hoja '" & HojaDestino & "' no existe." & vbCrLf & _
                           "¿Desea crearla?", vbYesNo + vbQuestion, "Hoja no encontrada")
        
        If respuesta = vbYes Then
            With Worksheets.Add(After:=Worksheets(Worksheets.Count))
                .Name = HojaDestino
            End With
            checkboxSheet.Activate
            ' Crear encabezado en la primera fila
            Worksheets(HojaDestino).Range(ColumnaVinculo & "1").Value = "Checkbox_States"
        Else
            MsgBox "No se puede continuar sin la hoja de destino.", vbExclamation, "Operación cancelada"
            Exit Sub
        End If
    End If
    
    '--- VALIDACIÓN 5: VERIFICAR COLUMNA VÁLIDA ---
    If Len(ColumnaVinculo) = 0 Or Not EsColumnaValida(ColumnaVinculo) Then
        MsgBox "La columna '" & ColumnaVinculo & "' no es válida.", vbExclamation, "Columna inválida"
        Exit Sub
    End If
    
    '--- ENCONTRAR PRÓXIMA CELDA DISPONIBLE ---
    Dim FilaSiguiente As Long
    With Worksheets(HojaDestino)
        Dim RangoBusqueda As Range
        Set RangoBusqueda = .Range(ColumnaVinculo & "2:" & ColumnaVinculo & .Rows.Count)
        
        ' Manejar caso donde no hay celdas vacías
        On Error Resume Next
        Dim CeldaVacia As Range
        Set CeldaVacia = RangoBusqueda.Cells.SpecialCells(xlCellTypeBlanks).Cells(1)
        On Error GoTo ManejoError
        
        If CeldaVacia Is Nothing Then
            ' Si no hay celdas vacías, usar la última fila + 1
            FilaSiguiente = .Cells(.Rows.Count, ColumnaVinculo).End(xlUp).Row + 1
        Else
            FilaSiguiente = CeldaVacia.Row
        End If
        
        ' Verificar que la fila no exceda el límite de Excel
        If FilaSiguiente > .Rows.Count Then
            MsgBox "No hay espacio disponible en la hoja '" & HojaDestino & "'.", vbExclamation, "Límite alcanzado"
            Exit Sub
        End If
    End With
    
    '--- OBTENER TEXTO PARA EL CHECKBOX ---
    Dim TextoCheckbox As String
    TextoCheckbox = ""
    
    If Len(TextoPersonalizado) > 0 Then
        ' Usar texto personalizado si se proporciona
        TextoCheckbox = TextoPersonalizado
    Else
        ' Buscar texto automáticamente
        Dim CeldaTexto As Range
        Set CeldaTexto = ActiveCell
        
        If BuscarTextoIzquierda Then
            ' Buscar texto hacia la izquierda hasta encontrar celda no vacía
            Dim ColumnaOriginal As Long
            ColumnaOriginal = CeldaTexto.Column
            
            Do While CeldaTexto.Value = "" And CeldaTexto.Column > 1
                Set CeldaTexto = CeldaTexto.Offset(0, -1)
            Loop
            
            ' Si no se encontró texto después de buscar, usar texto genérico
            If CeldaTexto.Value = "" Then
                TextoCheckbox = "Checkbox_" & FilaSiguiente
            Else
                TextoCheckbox = CStr(CeldaTexto.Value)
            End If
        Else
            ' Usar el texto de la celda actual
            If CeldaTexto.Value <> "" Then
                TextoCheckbox = CStr(CeldaTexto.Value)
            Else
                TextoCheckbox = "Checkbox_" & FilaSiguiente
            End If
        End If
    End If
    
    '--- INSERTAR Y CONFIGURAR CHECKBOX ---
    Dim CheckboxActual As CheckBox
    
    ' Verificar que la celda activa es válida para insertar
    If ActiveCell.Width = 0 Or ActiveCell.Height = 0 Then
        MsgBox "La celda seleccionada no tiene dimensiones válidas.", vbExclamation, "Celda inválida"
        Exit Sub
    End If
    
    Set CheckboxActual = checkboxSheet.CheckBoxes.Add( _
                         Left:=ActiveCell.Left, _
                         Top:=ActiveCell.Top, _
                         Width:=ActiveCell.Width, _
                         Height:=ActiveCell.Height)
    
    With CheckboxActual
        If MostrarCaption Then
            .Caption = TextoCheckbox
        Else
            .Caption = ""
        End If
        .LinkedCell = HojaDestino & "!" & ColumnaVinculo & FilaSiguiente
        .Value = ValorInicial
        .Display3DShading = False
        .Name = "CheckBox_" & HojaDestino & "_" & FilaSiguiente ' Nombre único
        .Placement = xlMoveAndSize               ' Se mueve y redimensiona con las celdas
    End With
    
    '--- INICIALIZAR VALOR EN HOJA DE DATOS ---
    Worksheets(HojaDestino).Range(ColumnaVinculo & FilaSiguiente).Value = (ValorInicial = True)
    Worksheets(HojaDestino).Range(ColumnaVinculo & FilaSiguiente).Offset(0, -1).Value = TextoCheckbox
    
    '--- CONFIRMACIÓN DE ÉXITO ---
    Dim MensajeExito As String
    MensajeExito = "Checkbox insertado correctamente:" & vbCrLf & _
                   " Vinculado a: " & HojaDestino & "!" & ColumnaVinculo & FilaSiguiente & vbCrLf & _
                   " Estado inicial: " & IIf(ValorInicial = True, "Marcado", "Desmarcado")
    
    If MostrarCaption And Len(TextoCheckbox) > 0 Then
        MensajeExito = MensajeExito & vbCrLf & " Texto: " & TextoCheckbox
    End If
    
    '--- SELECCIONAR CELDA ORIGINAL ---
    ActiveCell.Select
    
    ' Mostrar mensaje de éxito (opcional)
    ' MsgBox MensajeExito, vbInformation, "Checkbox insertado"
    
    Exit Sub
    
ManejoError:
    Select Case Err.Number
    Case 1004                                    ' Error general de Excel
        MsgBox "Error al acceder a la hoja de cálculo: " & Err.Description, _
               vbCritical, "Error de acceso"
    Case 9                                       ' Subíndice fuera de intervalo
        MsgBox "Error: Referencia a hoja o rango no válida.", vbCritical, "Error de referencia"
    Case 13                                      ' Tipo no coincide
        MsgBox "Error de tipo de dato en los parámetros.", vbCritical, "Error de tipo"
    Case Else
        MsgBox "Error inesperado (" & Err.Number & "): " & Err.Description, _
               vbCritical, "Error"
    End Select
    
    ' Limpiar recursos
    Set CheckboxActual = Nothing
    Set CeldaTexto = Nothing
    Set RangoBusqueda = Nothing
End Sub

'--- FUNCIÓN AUXILIAR PARA VALIDAR COLUMNAS ---
Private Function EsColumnaValida(ByVal Columna As String) As Boolean
    ' Verificar que la columna es válida (A-XFD)
    On Error GoTo ErrorHandler
    
    If Len(Columna) = 0 Then
        EsColumnaValida = False
        Exit Function
    End If
    
    ' Intentar convertir a número de columna
    Dim NumeroColumna As Long
    NumeroColumna = Range(Columna & "1").Column
    
    ' Si llegó aquí, la columna es válida
    EsColumnaValida = True
    Exit Function
    
ErrorHandler:
    EsColumnaValida = False
End Function

'--- PROCEDIMIENTOS DE EJEMPLO PARA USO RÁPIDO ---
Sub InsertarCheckboxConTexto()
Attribute InsertarCheckboxConTexto.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox con texto visible
    Call InsertarCheckbox(MostrarCaption:=True)
End Sub

Sub InsertarCheckboxMarcado()
Attribute InsertarCheckboxMarcado.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox marcado por defecto
    Call InsertarCheckbox(ValorInicial:=True, MostrarCaption:=False)
End Sub

Sub InsertarCheckboxPersonalizado()
Attribute InsertarCheckboxPersonalizado.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox con texto personalizado
    Call InsertarCheckbox(TextoPersonalizado:="Opción Personalizada", _
                          MostrarCaption:=True, _
                          HojaDestino:="CONFIG")
End Sub


