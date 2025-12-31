Attribute VB_Name = "modUTILS"
' ==========================================
' Módulo de utilidades generales
' ==========================================
Option Explicit

'@Description: Muestra el libro que contiene este XLAM, haciéndolo visible en la interfaz de Excel.
'@Scope: Manipula el libro host del complemento XLAM cargado.
'@ArgumentDescriptions: (no tiene argumentos)
'@Returns: Boolean | True si el libro se muestra correctamente; False en caso contrario.
'@Category: ComplementosExcel
Public Function MostrarLibroXLAM() As Boolean
Attribute MostrarLibroXLAM.VB_Description = "[modUTILS] Muestra el libro que contiene este XLAM, haciéndolo visible en la interfaz de Excel. Aplica a: ThisWorkbook\r\nM.D.:Manipula el libro host del complemento XLAM cargado."
Attribute MostrarLibroXLAM.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error GoTo ErrHandler
    
    ' El objeto ThisWorkbook, en un XLAM, apunta al libro host del complemento.
    ThisWorkbook.IsAddin = False                 ' Hace que el libro se muestre
    
    MostrarLibroXLAM = True
    Exit Function

ErrHandler:
    MostrarLibroXLAM = False
End Function

'@Description: Oculta el libro que contiene este XLAM, dejando el complemento operativo pero sin mostrar su ventana.
'@Scope: Manipula el libro host del complemento XLAM cargado.
'@ArgumentDescriptions: (no tiene argumentos)
'@Returns: Boolean | True si el libro se oculta correctamente; False en caso contrario.
'@Category: ComplementosExcel

Public Function OcultarLibroXLAM() As Boolean
Attribute OcultarLibroXLAM.VB_Description = "[modUTILS] Oculta el libro que contiene este XLAM, dejando el complemento operativo pero sin mostrar su ventana. Aplica a: ThisWorkbook\r\nM.D.:Manipula el libro host del complemento XLAM cargado."
Attribute OcultarLibroXLAM.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error GoTo ErrHandler
    
    ' Vuelve a marcar el libro como AddIn para que desaparezca de la vista.
    ThisWorkbook.IsAddin = True
    
    OcultarLibroXLAM = True
    Exit Function

ErrHandler:
    OcultarLibroXLAM = False
End Function

Function LongToRGB(colorValue As Long) As String
Attribute LongToRGB.VB_Description = "[modUTILS] Long To RGB (función personalizada)"
Attribute LongToRGB.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim r As Long, g As Long, b As Long
    
    r = colorValue And &HFF
    g = (colorValue And &HFF00&) \ &H100
    b = (colorValue And &HFF0000) \ &H10000
    
    LongToRGB = "RGB(" & r & ", " & g & ", " & b & ")"
End Function

' Función auxiliar para detectar filas vacías
Function IsEmptyRow(r As Range) As Boolean
Attribute IsEmptyRow.VB_Description = "[modUTILS] Función auxiliar para detectar filas vacías. Aplica a: Cells Range"
Attribute IsEmptyRow.VB_ProcData.VB_Invoke_Func = " \n23"
    IsEmptyRow = (WorksheetFunction.CountA(r) = 0)
End Function

'@Description: Verifica si una hoja existe en un workbook
'@Scope: Privado
'@ArgumentDescriptions: wb: Workbook donde buscar | sheetName: Nombre de la hoja
'@Returns: Boolean | True si la hoja existe
Function SheetExists(wb As Workbook, ByVal sheetName As String) As Boolean
Attribute SheetExists.VB_Description = "[modUTILS] Verifica si una hoja existe en un workbook"
Attribute SheetExists.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function HojaEstaSeleccionada(nombreHoja As String) As Boolean
Attribute HojaEstaSeleccionada.VB_Description = "[modUTILS] Hoja Esta Seleccionada (función personalizada)"
Attribute HojaEstaSeleccionada.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error Resume Next
    Dim sh As Object
    Set sh = ActiveWindow.SelectedSheets(nombreHoja)
    HojaEstaSeleccionada = (Not sh Is Nothing)
    On Error GoTo 0
End Function

' Reemplaza texto en todas las celdas de un rango
' NOTA: Esta es una función auxiliar (no UDF) - modifica celdas, no retorna valor
Function ReplaceInAllCells(rng As Range, strFrom As String, strTo As String, ByRef bSave As Boolean) As Boolean
Attribute ReplaceInAllCells.VB_Description = "[modUTILS] Reemplaza texto en todas las celdas de un rango. NOTA: Esta es una función auxiliar (no UDF) - modifica celdas, no retorna valor. Aplica a: Cells Range"
Attribute ReplaceInAllCells.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim oCell As Range
    Dim firstAddress As String, bNext As Boolean
    
    On Error GoTo ErrorHandler
    
    With rng
        Set oCell = .Find(What:=strFrom, After:=ActiveCell, LookIn:=xlValues, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, MatchCase:=True)
        
        If Not oCell Is Nothing Then
            firstAddress = oCell.Address
            Do
                oCell.Value = Replace(oCell.Value, strFrom, strTo)
                bSave = True
                Set oCell = .FindNext(oCell)
                bNext = Not oCell Is Nothing
                If bNext Then bNext = oCell.Address <> firstAddress
            Loop While bNext
        End If
    End With
    
    ReplaceInAllCells = bSave
    Exit Function
    
ErrorHandler:
    ReplaceInAllCells = False
End Function

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
                   "• Vinculado a: " & HojaDestino & "!" & ColumnaVinculo & FilaSiguiente & vbCrLf & _
                   "• Estado inicial: " & IIf(ValorInicial = True, "Marcado", "Desmarcado")
    
    If MostrarCaption And Len(TextoCheckbox) > 0 Then
        MensajeExito = MensajeExito & vbCrLf & "• Texto: " & TextoCheckbox
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     RANGE VALIDATION FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function bContentsErrorFree(ByRef refOutput As String, ByRef refValues As String, ByRef refInput As String) As Boolean
Attribute bContentsErrorFree.VB_Description = "[modUTILS] RANGE VALIDATION FUNCTIONS. . Aplica a: Cells Range"
Attribute bContentsErrorFree.VB_ProcData.VB_Invoke_Func = " \n23"
    'This function checks whether the values(s) of a range throw an error in an output cell.
    'If True then no error
    'If False then error

    Dim origInputContents As Variant
    Dim arrValues() As Variant
    Dim n As Variant

    'Assume by default that the contents are error free
    bContentsErrorFree = True

    'Store the formula of the input cell, just in case
    origInputContents = Range(refInput).formula

    If Range(refValues).Count > 1 Then

        arrValues = Range(refValues).Value

        For Each n In arrValues
            'Set the input cell equal to that value
            Range(refInput).Value = n

            'If the value causes an error in the output cell, return false
            If IsError(Range(refOutput).Value) Then

                'Return False
                bContentsErrorFree = False

                Exit For
            End If
        Next n
    
    Else
    
        'Set the input cell equal to the single value
        Range(refInput).Value = Range(refValues).Value

        'If the value causes an error in the output cell, return false
        If IsError(Range(refOutput).Value) Then bContentsErrorFree = False
    
    End If
    
    'Restore origional contents
    Range(refInput).formula = origInputContents

End Function

Function bAllNumbers(ByVal ref As String) As Boolean
Attribute bAllNumbers.VB_Description = "[modUTILS] b All Numbers (función personalizada). Aplica a: Cells Range"
Attribute bAllNumbers.VB_ProcData.VB_Invoke_Func = " \n23"
    ' This function checks whether the value(s) of a range are numeric.
    ' If True all are numeric
    ' IF False then at least one value is non-numeric

    Dim arr As Variant, n As Variant
    arr = Range(ref).Value
    
    'Assume by default all values are numeric
    bAllNumbers = True
    
    If Range(ref).Count > 1 Then
        ' Make sure all values are numeric
        For Each n In arr
            ' If not numeric, return False
            If Not IsNumeric(n) Then
                bAllNumbers = False
                Exit Function
            End If
        Next
    Else
        'Test if single value is numeric
        If Not IsNumeric(Range(ref).Value) Then
            bAllNumbers = False
        End If
    End If

End Function

Function bIsAddress(ByVal Str As String) As Boolean
Attribute bIsAddress.VB_Description = "[modUTILS] b Is Address (función personalizada). Aplica a: Cells Range"
Attribute bIsAddress.VB_ProcData.VB_Invoke_Func = " \n23"
    'This function checks whether a string is a reference to a range.
    On Error Resume Next
    
    Dim Var As Long
    Var = Range(Str).Count                       'Fails if the str is not an address
    
    If Err.Number <> 0 Then
        bIsAddress = False
    Else
        bIsAddress = True
    End If
    
    On Error GoTo 0
End Function


