Attribute VB_Name = "modMACROFixCGAS"
' ===== Módulo: modFixCGAS =====

'@Folder "6-DOMINIO-Oportunidades y compresores.b-Calculos técnicos"
Option Explicit

Private Const MODULE_NAME As String = "modMACROFixCGAS"

Public Sub FixCGASING()
Attribute FixCGASING.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    ' Validar libro activo
    If ActiveWorkbook Is Nothing Then
        Debug.Print "FixCGASING: No hay libro activo."
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = Nothing
    Dim c As Double, d As Double, vTmp As Variant
    Dim bSave As Boolean
    Dim regEx As Object
    On Error Resume Next
    Set ws = ActiveSheet
    '    Set ws = ActiveWorkbook.Worksheets("C-GAS-ING")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Debug.Print "FixCGASING: No existe la hoja en formato 'C-GAS-ING'."
        Exit Sub
    End If
    
    ' Guardar estado previo y desactivar actualizaciones
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Application.ActiveWindow.Zoom = 100
    
    ' Reemplazos en la hoja
    Call ReplaceInAllCells(ws.Cells, "Vapor de agua", "Water vapor", bSave)
    Call ReplaceInAllCells(ws.Cells, "Agua", "Water", bSave)
    Call ReplaceInAllCells(ws.Cells, "Límite RPM", "RPM Limit", bSave)
    Call ReplaceInAllCells(ws.Cells, " / 0 ( 0 )", "", bSave)
    Call ReplaceInAllCells(ws.Cells, "Seco-LT", "Dry-LT", bSave)
    Call ReplaceInAllCells(ws.Cells, "o Dry-LT", "or Dry-LT", bSave)
    Call ReplaceInAllCells(ws.Cells, "Atmosférico (Normal)", "Atmospheric (Standard)", bSave)
    Call ReplaceInAllCells(ws.Cells, "Metros", "Meters", bSave)
    Call ReplaceInAllCells(ws.Cells, "Composición del gas en Volumen :", "Gas composition by volume :", bSave)
    Call ReplaceInAllCells(ws.Cells, "Aire seco", "Dry air", bSave)
    Call ReplaceInAllCells(ws.Cells, "Aire", "Air", bSave)
    Call ReplaceInAllCells(ws.Cells, "Monóxido de Carbono", "Carbon monoxide", bSave)
    Call ReplaceInAllCells(ws.Cells, "Anhídrido Carbónico, Dióxido de Carbono", "Carbon dioxide", bSave)
    Call ReplaceInAllCells(ws.Cells, "Acido Sulfhídrico, Sulfuro de Hidrógeno", "Hydrogen sulfide", bSave)
    Call ReplaceInAllCells(ws.Cells, "Nitrógeno", "Nitrogen", bSave)
    Call ReplaceInAllCells(ws.Cells, "Hidrógeno", "Hydrogen", bSave)
    Call ReplaceInAllCells(ws.Cells, "Oxígeno", "Oxygen", bSave)
    Call ReplaceInAllCells(ws.Cells, "Metano", "Methane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Etano", "Ethane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Propano", "Propane", bSave)
    Call ReplaceInAllCells(ws.Cells, "propano", "propane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Etileno, Eteno", "Ethylene, Ethene", bSave)
    Call ReplaceInAllCells(ws.Cells, "Propileno, Propeno", "Propylene, Propene", bSave)
    Call ReplaceInAllCells(ws.Cells, "Butano", "Buthane", bSave)
    Call ReplaceInAllCells(ws.Cells, "butano", "buthane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Metil", "Methyl", bSave)
    Call ReplaceInAllCells(ws.Cells, "metil", "methyl", bSave)
    Call ReplaceInAllCells(ws.Cells, "Argón", "Argon", bSave)
    Call ReplaceInAllCells(ws.Cells, "Pentano", "Penthane", bSave)
    Call ReplaceInAllCells(ws.Cells, "pentano", "penthane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Hexano", "Hexane", bSave)
    Call ReplaceInAllCells(ws.Cells, "Autor :", "Author :", bSave)
    Call ReplaceInAllCells(ws.Cells, "Fecha :", "Date :", bSave)
    Call ReplaceInAllCells(ws.Cells, "- Pressure ", "- Exhaust pressure ", bSave)
    Call ReplaceInAllCells(ws.Cells, "CV/KW", "HP/kW", bSave)
    Call ReplaceInAllCells(ws.Cells, " CV", " HP", bSave)
    Call ReplaceInAllCells(ws.Cells, Environ("username"), Application.UserName, bSave)
    
    Dim cell As range
    
    Set cell = ws.Cells.Find("CH4O       : ", After:=ActiveCell, LookIn:=xlValues, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    If Not cell Is Nothing Then
        If cell.Offset(0, 2).Value <> "Methanol" Then
            cell.Offset(0, 2).Value = "Methanol"
            bSave = True
        End If
    End If
    Debug.Print "FixCGASING: Corregidos errores de idioma y texto en C-GAS-ING."
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    
    Set cell = ws.Cells.Find("Total mechanical losses : ", After:=ActiveCell, LookIn:=xlValues, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    If Not cell Is Nothing Then
        Dim result
        regEx.Pattern = "([\d,]+) HP"
        If regEx.Test(cell.Offset(0, 1).Value) Then
            Set result = regEx.Execute(cell.Offset(0, 1).Value)
            If result.Count > 0 Then
                c = CDbl(result(0).SubMatches(0)) ' manejo de coma decimal
                cell.Offset(0, 1).Value = c & " / " & Format(c * 0.7457, "0.00") & " HP/kW"
                bSave = True
            End If
        End If
    End If
    regEx.Pattern = "\s*:\s*"
    For Each cell In ws.range("F19:F29")
        cell.Value = regEx.Replace(cell.Value, "")
    Next
    
    ' mostrar celdas ocultas, para eliminarlas
    If ws.range("A60:A60").Value <> "" Then
        'stop
        ws.Rows("1:100").Select
        Application.Selection.EntireRow.Hidden = False
        If ws.Cells.Find("Motor at max.  : ") Is Nothing Or ws.Cells.Find("Isothermal efficiency : ") Is Nothing Then
            'Stop
        Else
            'xlShiftUp = -4162' CÓMO SE DESPLAZAN LAS CELDAS PARA SUSTITUIR A LAS ELIMINADAS
            ws.Rows("52:53").Delete
            ws.Rows("53:55").Delete
            ws.Rows("63:64").Delete
            ws.Rows("64:87").Delete
            ws.Rows("39:39").Delete
            Debug.Print "Eliminadas filas ocultas en C-GAS-ING"
        End If
        bSave = True
    End If
    
    If ws.range("E45:E45").Value <> "" Then
        ' EL FLOW DRY / WET
        ' xlDown, -4121 (inserta desplazando filas hacia abajo); xlFormatFromLeftOrAbove = 0 (el formato de las celdas insertadas es el de las de encima)
        ws.Rows("46:46").Insert -4121, 1
        ws.range("E45:F45").Cut ws.range("B46")
        ws.range("A45").Value = "Actual flow :"
        If InStr(ws.range("B46").Value, "kg") > 0 Then
            ws.range("A46").Value = "Mass flow (dry / wet):"
        Else
            ws.range("A46").Value = "Normal flow (dry / wet):"
        End If
        ws.range("C45:D45").Copy
        ws.range("C46").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        bSave = True
    Else
        'Stop ' será que la hoja está en un fichero ya modificado... pero asegurarse
    End If
    
    
    If ws.range("G30").Value <> "Suction pressure :" And ws.range("F30").Value <> "Specific weight in normal conditions:" Then
        ' dimensiona la lista de gases, PARA QUE TODAS LAS CELDAS TENGAN EL FORMATO CORRECTO
        regEx.Pattern = "([\d,]+)\%"             '13,99%
        ' For Each Cell In ws.Range("G19:G27")
        ' If regex.Test(Cell.value) Then Cell.value = regex.Execute(Cell.value).Item(0).SubMatches(0) * 1 & "%"
        ' Next
        ' ws.Range("E19").FormulaR1C1 = "1"
        ' ws.Range("E19").Copy
        ' ws.Range("G19:G27").PasteSpecial -4163, 4, False, False ' CONVERSION A NUMEROS, xlPasteValues = -4163, xlMultiply = 4
        ' Application.CutCopyMode = False
        ' ws.Range("G19:G28").Style = "Percent"
        ' ws.Range("G19:G28").NumberFormat = "0.00%"
        ' 'ws.Range("E19").Clear
        
        
        c = 19
        Do
            Set cell = ws.range("G" & c & ":G" & c)
            
            If regEx.Test(cell.Value) Then cell.Value = regEx.Execute(cell.Value).Item(0).SubMatches(0) * 1 & "%"
            cell.NumberFormat = "General"
            cell.Value = Replace(Trim(Replace(cell.Value, "'", "")), "%", "") / 100
            cell.NumberFormat = "0.00%"
            
            d = 19 - c
            Do
                If cell.Value > cell.Offset(d, 0).Value Then
                    'Stop
                    vTmp = cell.Offset(d, 0).Value: cell.Offset(d, 0).Value = cell.Value: cell.Value = vTmp
                    vTmp = cell.Offset(d, -1).Value: cell.Offset(d, -1).Value = cell.Offset(0, -1).Value: cell.Offset(0, -1).Value = vTmp
                    vTmp = cell.Offset(d, 1).Value: cell.Offset(d, 1).Value = cell.Offset(0, 1).Value: cell.Offset(0, 1).Value = vTmp
                End If
                d = d + 1
            Loop While d <= 0
            c = c + 1
        Loop While ws.range("F" & c & ":F" & c).Value <> ""
        
        
        ws.range("F28").Value = "Other     : "
        ws.range("G28").FormulaR1C1 = "=1-SUM(R[-9]C:R[-1]C)"
        ws.range("H28").ClearContents
        
        '  corregir las celdas de gases:
        c = 29
        Do While ws.range("F" & c).Value <> ""
            c = c + 1
        Loop
        ws.range("F29:H" & c - 1).Clear
        ws.range("G30").Value = "Suction pressure :"
        ws.range("G31").Value = "Atmospheric pressure :"
        ws.range("G32").Value = "Suction temperature :"
        ws.range("G33").Value = "Ambient temperature :"
        ws.range("G34").Value = "Relative humidity :"
        ws.range("G35").Value = "Water temperature :"
        bSave = True
        Debug.Print "Redimensionada la lista de gases en C-GAS-ING"
    End If
    
    ' recoloca primera y segunda columnas de INPUT DATA
    If ws.range("F29").Value = "" And ws.range("A24").Value = "Specific weight in normal conditions:" _
       And ws.range("A30").Value = "Compressor series: " And ws.range("G30").Value = "Suction pressure :" Then
        ' SI NO SE CUMPLE ws.range ("F29").value = ""... LAS CELDAS A MOVER SE HABRIAN REEMPLAZADO POR NOMBRES DE GASES!!!
        '       me aseguro además de que el resto de la hoja no se haya modificado, que sea "la original"; por si acaso
        ' PRESENTACION ALTERNATIVA: RECOLOCA LAS FILAS ORDENANDO MEJOR LOS CONCEPTOS DE ENTRADA.. OJO!!, ESTO AFECTA A LAS OFERTAS GENERADAS
        ' ASEGURARSE DE CAMBIAR LAS PLANTILLAS DE OFERTAS, LAS QUE HACEN REF A C-GAS-ING, AL HACER ESTE CAMBIO!!
        ws.range("A24:C26").Cut ws.range("F37")
        
        ws.range("A30:D30").Cut ws.range("A17")
        ws.range("A34:D34").Cut ws.range("A18")
        ws.range("A33:D33").Cut ws.range("A19")
        ws.range("A31:D31").Cut ws.range("A20")
        ws.range("A35:D36").Cut ws.range("A21")
        
        If ws.range("F34").Value = "" Then
            ws.range("G34").Cut ws.range("A23")
        Else
            ws.range("A23").Value = "Relative humidity : "
            ws.range("A17").Copy
            ws.range("A23").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I34:J34").Cut ws.range("B23")
        ws.range("C22:D22").Copy
        ws.range("C23").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        If ws.range("F30").Value = "" Then
            ws.range("G30").Cut ws.range("A24")
        Else
            ws.range("A24").Value = "Suction pressure :"
            ws.range("A17").Copy
            ws.range("A24").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I30:J30").Cut ws.range("B24")
        ws.range("C23:D23").Copy
        ws.range("C24").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        If ws.range("F32").Value = "" Then
            ws.range("G32").Cut ws.range("A25")
        Else
            ws.range("A25").Value = "Suction temperature : "
            ws.range("A17").Copy
            ws.range("A25").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I32:J32").Cut ws.range("B25")
        ws.range("C23:D23").Copy
        ws.range("C25").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        If ws.range("F33").Value = "" Then
            ws.range("G33").Cut ws.range("A26")
        Else
            ws.range("A26").Value = "Ambient temperature : "
            ws.range("A17").Copy
            ws.range("A26").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I33:J33").Cut ws.range("B26")
        ws.range("C24:D24").Copy
        ws.range("C26").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        If ws.range("F31").Value = "" Then
            ws.range("G31").Cut ws.range("A27")
        Else
            ws.range("A27").Value = "Atmospheric pressure :"
            ws.range("A17").Copy
            ws.range("A27").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I31:J31").Cut ws.range("B27")
        ws.range("C25:D25").Copy
        ws.range("C27").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        ws.range("A37:D37").Cut ws.range("A28")
        
        ws.range("A32:D32").Cut ws.range("A29")
        
        If ws.range("F35").Value = "" Then
            ws.range("G35").Cut ws.range("A30")
        Else
            ws.range("A30").Value = "Water temperature : "
            ws.range("A17").Copy
            ws.range("A30").PasteSpecial -4122, -4142, False, False ' PEGA EL FORMATO
            Application.CutCopyMode = False
        End If
        ws.range("I35:J35").Cut ws.range("B30")
        ws.range("C28:D28").Copy
        ws.range("C30").PasteSpecial -4122, -4142, False, False
        Application.CutCopyMode = False
        
        ws.range("A38:D40").Cut ws.range("A31")
        
        If ws.range("F29").Value = "" Then
            ws.range("F37:F39").Cut ws.range("F30")
        Else
            ' si no son blancos... primero reajustar la lista de gases, luego se actualizaría estas celdas
        End If
        
        ws.range("G37:H39").Cut ws.range("I30")
        ws.Rows("35:42").Delete -4162
        bSave = True
    Else
        ' LAS CELDAS A MOVER EN LA SEGUNDA COLUMNA PODRIAN HABER REEMPLAZADO POR NOMBRES DE GASES!! --> HAY QUE REVISAR EL FORMATO y el script...
    End If
    
    If ws.Columns("J:J").ColumnWidth > 10 Then
        ' ajusta anchos de columnas, para hacer la tabla mas presentable
        ws.Columns("A:A").ColumnWidth = 31.6
        ws.Columns("H:H").ColumnWidth = 12
        ws.Columns("I:I").ColumnWidth = 9
        ws.Columns("J:J").ColumnWidth = 6.5
        ws.Columns("B:G").ColumnWidth = 9.8
        bSave = True
    End If
    
    ' OPERACIONES A REALIZAR AL FINAL, UNA VEZ MOVIDAS TODAS LAS CELDAS
    Set cell = ws.Cells.Find("Compressor model : ", After:=ActiveCell, LookIn:=xlValues, _
                             LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    regEx.Pattern = MODEL_PATTERN
    If Not cell Is Nothing Then
        cell.Offset(0, 1).Value = regEx.Execute(strModelName()).Item(0).Value
    End If
    
    ' Añado unas conversiones de unidades...
    If InStr(ws.range("B25:B25").Value, "ºF") > 0 Then
        ws.range("A53:A53").Value = Replace(ws.range("A53:A53").Value, "ºC", "ºF")
        ws.range("A54:A54").Value = Replace(ws.range("A54:A54").Value, "ºC", "ºF")
        For c = Asc("B") To Asc("G")
            If ws.range(Chr(c) & "53:" & Chr(c) & "53").Value <> "" Then
                ws.range(Chr(c) & "53:" & Chr(c) & "53").Value = ws.range(Chr(c) & "53:" & Chr(c) & "53").Value * 9 / 5 + 32
            End If
            If ws.range(Chr(c) & "54:" & Chr(c) & "54").Value <> "" Then
                ws.range(Chr(c) & "54:" & Chr(c) & "54").Value = ws.range(Chr(c) & "54:" & Chr(c) & "54").Value * 9 / 5 + 32
            End If
        Next
    End If
    ' Eliminar RPM en los datos de entrada, si se ha puesto caudal > 0
    regEx.Pattern = "^\d+\s*(\(\s*RPM Limit = \d+\s*\))?"
    If ws.range("B21") <> "-" Then If regEx.Test(ws.range("B31")) Then ws.range("B31").Value = regEx.Replace(ws.range("B31").Value, "--$1")
    
    ' (Opcional) Formatear encabezados principales en negrita
    With ws
        .range("B2").Font.Bold = True            ' Título CALCULATION - GAS
        .range("A15,A35,A47,F17").Font.Bold = True ' INPUT DATA, OUTPUT DATA, STAGES, Coolers
        '.Range("F18,G18").Font.Bold = True   ' Encabezados Gas/Percentage
    End With
    
    'Stop
    ws.range("A1").Select
    
    ' Recalcular si se requiere
    Debug.Print "FixCGASING: Recalculando hoja 'C-GAS-ING'."
    ws.Calculate
    
CleanUp:
    ' Restaurar propiedades de Excel
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    
    If bSave Then If MsgBox("¿Guardar los cambios?", vbYesNo Or vbDefaultButton2 Or vbQuestion) = 6 Then ActiveWorkbook.Save
    Exit Sub
    
ErrHandler:
    Debug.Print "FixCGASING: Error " & Err.Number & " - " & Err.Description
    Resume CleanUp
End Sub
