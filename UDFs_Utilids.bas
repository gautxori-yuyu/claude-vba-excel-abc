Attribute VB_Name = "UDFs_Utilids"
'@IgnoreModule MissingAnnotationArgument
'@Folder "UDFS"
Option Explicit

Private Const MODULE_NAME As String = "UDFs_Utilids"

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

Function LongToRGB(colorValue As Long) As String
Attribute LongToRGB.VB_Description = "[UDFs_Utilids] Long To RGB (función personalizada)"
Attribute LongToRGB.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim r As Long, g As Long, b As Long
    
    r = colorValue And &HFF
    g = (colorValue And &HFF00&) \ &H100
    b = (colorValue And &HFF0000) \ &H10000
    
    LongToRGB = "RGB(" & r & ", " & g & ", " & b & ")"
End Function
