Attribute VB_Name = "UDFs_UtilsExcel"
' ==========================================
' Módulo UDFS que requieren CONTEXTO EXCEL
' ==========================================

'@Folder "UDFS.Excel"
Option Explicit

Private Const MODULE_NAME As String = "UDFs_UtilsExcel"

' Función auxiliar para detectar filas vacías
Function IsEmptyRow(r As Range) As Boolean
Attribute IsEmptyRow.VB_Description = "[UDFs_UtilsExcel] Función auxiliar para detectar filas vacías. Aplica a: Cells Range"
Attribute IsEmptyRow.VB_ProcData.VB_Invoke_Func = " \n23"
    IsEmptyRow = (WorksheetFunction.CountA(r) = 0)
End Function

'@Description: Verifica si una hoja existe en un workbook
'@Scope: Privado
'@ArgumentDescriptions: wb: Workbook donde buscar | sheetName: Nombre de la hoja
'@Returns: Boolean | True si la hoja existe
Function SheetExists(Wb As Workbook, ByVal sheetName As String) As Boolean
Attribute SheetExists.VB_Description = "[UDFs_UtilsExcel] Verifica si una hoja existe en un workbook"
Attribute SheetExists.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function HojaEstaSeleccionada(nombreHoja As String) As Boolean
Attribute HojaEstaSeleccionada.VB_Description = "[UDFs_UtilsExcel] Hoja Esta Seleccionada (función personalizada)"
Attribute HojaEstaSeleccionada.VB_ProcData.VB_Invoke_Func = " \n23"
    On Error Resume Next
    Dim Sh As Object
    Set Sh = ActiveWindow.SelectedSheets(nombreHoja)
    HojaEstaSeleccionada = (Not Sh Is Nothing)
    On Error GoTo 0
End Function

' Reemplaza texto en todas las celdas de un rango
' NOTA: Esta es una función auxiliar (no UDF) - modifica celdas, no retorna valor
Function ReplaceInAllCells(rng As Range, strFrom As String, strTo As String, ByRef bSave As Boolean) As Boolean
Attribute ReplaceInAllCells.VB_Description = "[UDFs_UtilsExcel] Reemplaza texto en todas las celdas de un rango. NOTA: Esta es una función auxiliar (no UDF) - modifica celdas, no retorna valor. Aplica a: Cells Range"
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

'@UDF
'@Description: Obtiene el nombre de la primera tabla de una hoja especificada
'@Category: Tablas
'@ArgumentDescriptions: Nombre de la hoja donde buscar la tabla
Public Function GetFirstTableName(wsName As String) As String
Attribute GetFirstTableName.VB_Description = "[UDFs_UtilsExcel] Obtiene el nombre de la primera tabla de una hoja especificada. Aplica a: ActiveWorkbook"
Attribute GetFirstTableName.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ws As Worksheet
    Application.Volatile
    On Error GoTo ErrorHandler
    
    Set ws = ActiveWorkbook.Worksheets(wsName)
    
    If ws.ListObjects.Count > 0 Then
        GetFirstTableName = ws.ListObjects(1).Name
    Else
        GetFirstTableName = ""
    End If
    
    Exit Function
    
ErrorHandler:
    GetFirstTableName = "#ERROR"
End Function

'@UDF
'@Description: Busca un patrón de expresión regular en un rango de celdas
'@Category: Búsqueda
'@ArgumentDescriptions: Rango donde buscar|Patrón de expresión regular|Si TRUE devuelve la coincidencia, si FALSE devuelve la dirección
Public Function BuscarRegex(rango As Range, patron As String, Optional devolverCoincidencia As Boolean = False) As Variant
Attribute BuscarRegex.VB_Description = "[UDFs_UtilsExcel] Busca un patrón de expresión regular en un rango de celdas. Aplica a: Cells Range"
Attribute BuscarRegex.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim regEx As Object
    Dim celda As Range
    Dim coincidencias As Object
    
    On Error GoTo ErrorHandler
    
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        .Pattern = patron
        .Global = True
        .IgnoreCase = True
    End With
    
    For Each celda In rango
        If regEx.Test(celda.Value) Then
            If devolverCoincidencia Then
                Set coincidencias = regEx.Execute(celda.Value)
                BuscarRegex = coincidencias(0).Value
            Else
                BuscarRegex = celda.Address
            End If
            Exit Function
        End If
    Next celda
    
    BuscarRegex = CVErr(xlErrNA)
    Exit Function
    
ErrorHandler:
    BuscarRegex = CVErr(xlErrValue)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                     RANGE VALIDATION FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function bContentsErrorFree(ByRef refOutput As String, ByRef refValues As String, ByRef refInput As String) As Boolean
Attribute bContentsErrorFree.VB_Description = "[UDFs_UtilsExcel] RANGE VALIDATION FUNCTIONS. Aplica a: Cells Range"
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
Attribute bAllNumbers.VB_Description = "[UDFs_UtilsExcel] b All Numbers (función personalizada). Aplica a: Cells Range"
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
Attribute bIsAddress.VB_Description = "[UDFs_UtilsExcel] b Is Address (función personalizada). Aplica a: Cells Range"
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

