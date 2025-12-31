Attribute VB_Name = "modUTILSProcedureParsing"
' ==========================================
' FUNCIONES DE PARSING
' ==========================================
'@Folder "1-Inicio e Instalacion"
'@IgnoreModule MissingAnnotationArgument, ProcedureNotUsed
Option Explicit

' Parsea todos los procedimientos del proyecto VBA (CON Y SIN metadatos)
Public Function ParsearProcsDelProyecto() As Object
Attribute ParsearProcsDelProyecto.VB_Description = "[modUTILSProcedureParsing] Parsea todos los procedimientos del proyecto VBA (CON Y SIN metadatos). Aplica a: ThisWorkbook"
Attribute ParsearProcsDelProyecto.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim vbProj As Object, vbComp As VBIDE.VBComponent
    
    Dim procName As String
    Dim PKind As ProcKind
    
    Dim CodeBlock As T_CodeBlock
    '    Dim procStartLine As Long, procNumLines As Long
    '    Dim procSignatureLine As Long, strCode As String
    
    Dim oVBAProcedure As clsVBAProcedure
    Dim funciones As Object
    Set funciones = CreateObject("Scripting.Dictionary")
    
    On Error GoTo ErrorHandler
    ' Intentar acceder al VBA Project
    Set vbProj = ThisWorkbook.VBProject
    
    If vbProj Is Nothing Then
        Debug.Print "[ParsearProcsDelProyecto] - No hay acceso al VBA Project."
        Debug.Print "  -> Habilita 'Confiar en el acceso al modelo de objetos de proyectos de VBA'"
        Debug.Print "  -> En: Archivo > Opciones > Centro de confianza > Configuración"
        Set ParsearProcsDelProyecto = Nothing
        Exit Function
    End If
    
    ' Recorrer todos los módulos estándar
    For Each vbComp In vbProj.VBComponents
        With vbComp
            ' Usar Members para enumerar todos los procedimientos
            Dim lineNum As Long
            lineNum = IIf(.CodeModule.CountOfDeclarationLines = 0, 1, .CodeModule.CountOfDeclarationLines)
        
            Do While lineNum < .CodeModule.CountOfLines
                ' Obtener siguiente procedimiento
                procName = .CodeModule.ProcOfLine(lineNum, PKind)
            
                If procName <> "" Then
                    CodeBlock = getProcCode(.CodeModule, procName, PKind)
                    ' Intentar parsear metadatos
                    Set oVBAProcedure = New clsVBAProcedure
                    Call oVBAProcedure.Init(.Name, EsModuloPrivado(.CodeModule), .Type, _
                                            PKind, procName, CodeBlock)
                
                    If oVBAProcedure.Name <> "" Then
                        funciones.Add funciones.Count, oVBAProcedure
                    End If
                
                    ' Saltar al final del procedimiento
                    lineNum = .CodeModule.procStartLine(procName, PKind) + .CodeModule.ProcCountLines(procName, PKind) + 1
                Else
                    lineNum = lineNum + 1
                End If
            Loop
        End With
    Next vbComp
    
    Set ParsearProcsDelProyecto = funciones
    
    If funciones.Count > 0 Then
        Debug.Print "[ParsearProcsDelProyecto] - " & funciones.Count & " procedimientos encontrados."
    End If
    
    Exit Function
ErrorHandler:
    Debug.Print "[ParsearProcsDelProyecto] - Error al parsear procedimientos: " & Err.Description
End Function

'@Description: Corrige los desplazamientos erroneos en los modulos de codigo detectados por las funciones
' del 'modelo de objetos de proyectos de VBA', y obtiene el bloque de codigo corregido
Private Function getProcCode(CodeModule As Object, procName As String, PKind As ProcKind) As T_CodeBlock
    Dim CodeBlock As T_CodeBlock
    Dim i As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    
    With CodeBlock
        .procStartLine = CodeModule.procStartLine(procName, PKind)
        .procNumLines = CodeModule.ProcCountLines(procName, PKind)
        .procSignatureLine = CodeModule.ProcBodyLine(procName, PKind)
        
        ' se reajusta el comienzo del bloque de código, VBE no lo pone bien
        re.Pattern = "^\s*'.+"
        On Error GoTo ErrorHandler
        Do While .procStartLine > 1
            If Not (re.Test(CodeModule.Lines(.procStartLine - 1, 1)) Or _
                    CodeModule.Lines(.procStartLine - 1, 1) = "") Then Exit Do
            .procStartLine = .procStartLine - 1
        Loop
        
        ' ... y hay que corregir el final, tampoco termina bien los bloques de función
        re.Pattern = "\bFunction|Sub|Property\b"
        re.Pattern = "^\s*End\s+" & re.Execute(CodeModule.Lines(.procSignatureLine, 1)).Item(0).Value
        i = .procStartLine
        .strCode = CodeModule.Lines(i, 1)
        Do
            i = i + 1
            .strCode = .strCode & vbCrLf & CodeModule.Lines(i, 1)
            If i - .procStartLine > 500 Then Stop
        Loop Until (i = CodeModule.CountOfLines) Or re.Test(CodeModule.Lines(i, 1))
        
        '.procWrongEndLines = .procNumLines - (i - .procStartLine + 1) ' ESTAS LINEAS DEBEN ASOCIARSE AL PROCEDIMIENTO SIGUIENTE
        .procNumLines = i - .procStartLine + 1
    End With
    
    getProcCode = CodeBlock
    Exit Function
ErrorHandler:
    Debug.Print "[getProcCode] - Error: " & Err.Description
End Function

' Verifica si un módulo tiene Option Private Module
Private Function EsModuloPrivado(CodeModule As Object) As Boolean
    EsModuloPrivado = False
    Dim i As Long, lineText As String
        
    On Error GoTo ErrorHandler
    
    For i = 1 To CodeModule.CountOfDeclarationLines
        lineText = Trim$(CodeModule.Lines(i, 1))
        If InStr(1, lineText, "Option Private Module", vbTextCompare) > 0 Then
            EsModuloPrivado = True: Exit For
        ElseIf lineText <> "" And _
               Left$(lineText, 1) <> "'" And _
               InStr(1, lineText, "Option", vbTextCompare) = 0 And _
               InStr(1, lineText, "Attribute", vbTextCompare) = 0 Then
            ' Si encontramos código (no opciones/comentarios), dejar de buscar
            Exit For
        End If
    Next i
    Exit Function
ErrorHandler:
    Debug.Print "[EsModuloPrivado] - Error: " & Err.Description
End Function


