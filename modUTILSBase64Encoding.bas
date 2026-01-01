Attribute VB_Name = "modUTILSBase64Encoding"
' ==========================================
' Módulo de codificación/decodificación Base64
' Usado internamente por el sistema de instalación
' ==========================================

'@IgnoreModule MissingAnnotationArgument
Option Explicit

' Decodifica una cadena Base64 a texto plano
Function Base64Decode(texto As String) As String
Attribute Base64Decode.VB_Description = "[modUTILSBase64Encoding] Decodifica una cadena Base64 a texto plano"
Attribute Base64Decode.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim xml As Object
    Dim node As Object
    
    On Error GoTo ErrorHandler
    
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    
    node.DataType = "bin.base64"
    node.text = texto
    Base64Decode = StrConv(node.nodeTypedValue, vbUnicode)
    
    Exit Function
    
ErrorHandler:
    Base64Decode = ""
End Function

' Codifica un array de bytes en Base64
Function Base64EncodeFromBytes(bytes() As Byte) As String
Attribute Base64EncodeFromBytes.VB_Description = "[modUTILSBase64Encoding] Codifica un array de bytes en Base64"
Attribute Base64EncodeFromBytes.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim xml As Object
    Dim nodo As Object
    
    On Error GoTo ErrorHandler
    
    ' Codificar en Base64 usando MSXML
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set nodo = xml.createElement("b64")
    nodo.DataType = "bin.base64"
    nodo.nodeTypedValue = bytes
    Base64EncodeFromBytes = nodo.text
    
    Exit Function
    
ErrorHandler:
    Base64EncodeFromBytes = ""
End Function

' Codifica un archivo en Base64 leyéndolo como binario
Function Base64EncodeFromFile(rutaArchivo As String) As String
Attribute Base64EncodeFromFile.VB_Description = "[modUTILSBase64Encoding] Codifica un archivo en Base64 leyéndolo como binario"
Attribute Base64EncodeFromFile.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim stream As Object
    Dim xml As Object
    Dim nodo As Object
    Dim bytes() As Byte
    
    On Error GoTo ErrorHandler
    
    ' Leer archivo como binario
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1                                ' Binario
        .Open
        .LoadFromFile rutaArchivo
        bytes = .Read
        .Close
    End With
    
    ' Codificar en Base64
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set nodo = xml.createElement("b64")
    nodo.DataType = "bin.base64"
    nodo.nodeTypedValue = bytes
    Base64EncodeFromFile = nodo.text
    
    Exit Function
    
ErrorHandler:
    Base64EncodeFromFile = ""
End Function

'@Description: Convierte un script VBScript al texto de una funcion que se puede pegar en el VBA
'@ArgumentDescriptions: rutaInput:ruta del fichero vbscript a convertir (normalmente ext. .vbs)
' |rutaOutput: ruta del fichero b64 convertido (normalmente ext. .Base64)
' |FUNC_NAME: nombre que tendrá la funcion VBA (Function FUNC_NAME() As String)
Sub ScriptToFunctionBase64RC4(rutaInput As String, rutaOutput As String, FUNC_NAME As String)
Attribute ScriptToFunctionBase64RC4.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim salida As String
    Dim f As Integer, l, strlin
    
    salida = Base64EncodeFromFile((rutaInput))
    salida = """" & Replace(salida, vbLf, """ & _" & vbCrLf & """") & """"
    l = 0
    For Each strlin In Split(salida, vbCrLf)
        If l = 0 Then salida = "Function " & FUNC_NAME & "() As String" & vbCrLf & vbTab & "" & FUNC_NAME & " = _" & vbCrLf
        salida = salida & vbTab & vbTab & strlin & vbCrLf
        l = l + 1
        If l Mod 20 = 0 And Right(salida, 6) = " & _" & vbCrLf Then
            salida = Left(salida, Len(salida) - 6) & vbCrLf & vbTab & "" & FUNC_NAME & " = " & FUNC_NAME & " & _" & vbCrLf
        End If
    Next
    salida = salida & "End Function"
    
    f = FreeFile
    Open rutaOutput For Output As #f
    Print #f, salida
    Close #f
End Sub

