Attribute VB_Name = "UDFs_Utilids"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

'@UDF
'@Description: Busca un patrón de expresión regular en un rango de celdas
'@Category: Búsqueda
'@ArgumentDescriptions: Rango donde buscar|Patrón de expresión regular|Si TRUE devuelve la coincidencia, si FALSE devuelve la dirección
Public Function BuscarRegex(rango As Range, patron As String, Optional devolverCoincidencia As Boolean = False) As Variant
Attribute BuscarRegex.VB_Description = "[UDFs_Utilids] Busca un patrón de expresión regular en un rango de celdas. Aplica a: Cells Range"
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

'@UDF
'@Description: Extrae la parte numérica inicial de un texto (soporta decimales con punto o coma). Sirve por ejemplo para separar el valor numerico, de las unidades, en celdas de gas_vbnet etc.
'@Category: Texto
'@ArgumentDescriptions: Texto del que extraer el número
Public Function ExtraerNumeroInicial(texto As String) As Double
Attribute ExtraerNumeroInicial.VB_Description = "[UDFs_Utilids] Extrae la parte numérica inicial de un texto (soporta decimales con punto o coma). Sirve por ejemplo para separar el valor numerico, de las unidades, en celdas de gas_vbnet etc."
Attribute ExtraerNumeroInicial.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim i As Integer
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    For i = 1 To Len(texto)
        If IsNumeric(Mid(texto, i, 1)) Or _
           Mid(texto, i, 1) = "." Or _
           Mid(texto, i, 1) = "," Then
            resultado = resultado & Mid(texto, i, 1)
        Else
            If resultado <> "" Then Exit For
        End If
    Next i
    
    If resultado <> "" Then
        ExtraerNumeroInicial = CDbl(Replace(resultado, ",", "."))
    Else
        ExtraerNumeroInicial = 0
    End If
    
    Exit Function
    
ErrorHandler:
    ExtraerNumeroInicial = 0
End Function

'@UDF
'@Description: Obtiene el nombre de la primera tabla de una hoja especificada
'@Category: Tablas
'@ArgumentDescriptions: Nombre de la hoja donde buscar la tabla
Public Function GetFirstTableName(wsName As String) As String
Attribute GetFirstTableName.VB_Description = "[UDFs_Utilids] Obtiene el nombre de la primera tabla de una hoja especificada. Aplica a: ActiveWorkbook"
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


