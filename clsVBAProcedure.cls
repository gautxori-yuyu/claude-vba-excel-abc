VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Exposed
'@Folder "1-Inicio e Instalacion.Gestion de modulos y procs"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "clsVBAProcedure"

Public Name As String
Public Module As String
Public ContainerType As ProcContainerType
Public bPrivateModule As Boolean

Public strCode As String
Public procNumLines As Long
Public procSignatureLine As Long
Public procStartLine As Long

Public PKind As ProcKind
Public NormalizedSignature As String
Public ProcedureType As ProcType                 ' Macro (Sub sin params), UDF (Function no Private), Function, Sub con params

Public Scope As String                           ' ambito al que se aplica: selecciÃ³n, hoja activa, libro activo, argumentos de la funcion, ...
Public ArgumentDescriptions As String            ' Separado por "|"
Public Returns As String                         ' QuÃ© devuelve

Public HasMetadata As Boolean                    ' TRUE si tiene comentarios @UDF
Public Description As String
Public Category As String
Public Example As String                         ' Ejemplo de uso
Public Raises As String                          ' Errores que puede lanzar
Public Dependencies As String                    ' Dependencias de otros procedimientos

' Constructor para fÃ¡cil inicializaciÃ³n
'TODO: revisar el codigo para evitar las restricciones en el orden de las llamadas
Public Sub Init(ByVal modl As String, ByVal bModuloPrivado_ As Boolean, ByVal ContainerType_ As ProcContainerType, _
                ByVal PKind_ As ProcKind, ByVal procName As String, CodeBlock As T_CodeBlock)
        
    On Error GoTo ErrorHandler
    
    Name = procName
    Module = modl
    ContainerType = ContainerType_
    bPrivateModule = bModuloPrivado_
    
    strCode = CodeBlock.strCode
    procNumLines = CodeBlock.procNumLines
    procSignatureLine = CodeBlock.procSignatureLine
    procStartLine = CodeBlock.procStartLine

    PKind = PKind_
    ' Es IMPORTANTE mantener el orden de estas llamadas! (al menos de momento, PTE de mejorar)
    
    NormalizedSignature = NormalizarLineasFirma(CodeBlock.strCode, CodeBlock.procSignatureLine - CodeBlock.procStartLine + 1)
    
    ProcedureType = TipoProcedimiento(NormalizedSignature)
    
    Scope = ParsearCodeParaDeducirContexto(CodeBlock.strCode)
    
    ArgumentDescriptions = ComponerArgumentDescriptions(NormalizedSignature)
    
    Dim arrMetadata()
    arrMetadata = ParsearMetadataCompleta(CodeBlock.strCode)
    
    ' VALIDACION DE METADATOS: deben corresponder a los deducidos del procedimiento... o extenderlos.
    Dim procTypeTag As String
    Select Case ProcedureType
        Case eventHandler
            procTypeTag = "(VBA, Manejador de eventos) "
        Case internalPrivate, internalSubPublicWithParams
            procTypeTag = "(VBA, Interna, no expuesta) "
    End Select
    Description = arrMetadata(0)
    If Left(Description, Len(procTypeTag)) <> procTypeTag Then Description = procTypeTag & Description
    ' Si no hay descripciÃ³n del procedimiento, generar una bÃ¡sica
    If Description = "" Then
        Description = GenerarDescripcionAutomatica(Name)
    End If
    
    ' validar metadatos, con los deducidos de la firma y contenido del codigo
    If arrMetadata(1) <> "" And InStr(arrMetadata(1), Category) = 0 Then
        Category = Category & vbCrLf & "M.D.:" & arrMetadata(1)
    End If
    If arrMetadata(2) <> "" And InStr(arrMetadata(2), ArgumentDescriptions) = 0 Then
        ArgumentDescriptions = ArgumentDescriptions & vbCrLf & "M.D.:" & arrMetadata(2)
    End If
    If arrMetadata(3) <> "" And InStr(arrMetadata(3), Scope) = 0 Then
        Scope = Scope & vbCrLf & "M.D.:" & arrMetadata(3)
    End If
    If arrMetadata(4) <> "" And InStr(arrMetadata(4), Returns) = 0 Then
        Returns = Returns & vbCrLf & "M.D.:" & arrMetadata(4)
    End If
    
    'Debug.Print "[clsVBAProcedure Init] - " & dumpProcedimiento
    
    Exit Sub
ErrorHandler:
    LogCurrentError MODULE_NAME, "[Init]"
    Err.Raise Err.Number, MODULE_NAME & " [Init]", Err.Description
End Sub

'@Description: Determina si una funciÃ³n es visible para excel como macro o udf, o si es "interna",
' a partir de NormalizedSignature,PKind,ContainerType,bModuloPrivado,...
Private Function TipoProcedimiento(ByVal firmaNormalizada As String) As ProcType
    Dim tipo As String, acceso As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True

    ' Acceso
    Set re = Nothing: Set re = CreateObject("VBScript.RegExp")
    On Error GoTo ErrorHandler
    re.Pattern = "^\s*((?:(?:Public|Private|Friend|Static)\s+){1,2})?\s*(Function|Sub|Property\s+(?:Get|Set|Let))\s*(\S+)\s*\(\s*(.*)\)"
    If re.Test(firmaNormalizada) Then
        Dim m As Object: Set m = re.Execute(firmaNormalizada)(0)
        acceso = Replace(LCase$(Trim$(m.SubMatches(0))), " static", "")
        tipo = LCase$(Trim$(m.SubMatches(1)))
    Else
        'Stop
    End If

    If acceso = "" Then acceso = "public"
    
    If InStr(m.SubMatches(2), "_") > 0 Then
        TipoProcedimiento = ProcType.eventHandler
        ' Normalmente deberÃ­an ser privados, dentro de una clase que instancie un objeto de otra clase que implemente eventos
        'If acceso <> "private" Then Stop
        ' y no se si puede llegar a haber manejadores que sean funciones... por si acaso:
        'If tipo <> "sub" Then Stop
    ElseIf tipo = "sub" And acceso = "public" And Not bPrivateModule And ContainerType = StdModule _
           And m.SubMatches(3) <> "" Then
        TipoProcedimiento = ProcType.internalSubPublicWithParams
    ElseIf acceso <> "public" Or bPrivateModule Or ContainerType <> StdModule Or InStr(tipo, "property") > 0 Then
        TipoProcedimiento = ProcType.internalPrivate
    Else
        Select Case LCase$(tipo)
        Case "sub"
            TipoProcedimiento = ProcType.Macro
        Case "function"
            TipoProcedimiento = ProcType.udf
        Case Else
            Debug.Print "[TipoProcedimiento] - Error al procesar Firma Normalizada con expresiones regulares"
            Stop                                 ' depurar la causa
        End Select
    End If
    
    ' FIXME: establece los valores de PKind y Returns, en funcion del parsing realizado en este procedimiento
    If tipo = "function" Then
        If PKind = ProcKind.proc Then PKind = ProcKind.ProcFunction Else Stop ' debug
    ElseIf InStr(tipo, "property") > 0 And (PKind < ProcKind.PropLet Or PKind < ProcKind.PropGet) Then
        If InStr(tipo, "let") > 0 Then
            PKind = ProcKind.PropLet
        ElseIf InStr(tipo, "set") > 0 Then
            PKind = ProcKind.PropSet
        ElseIf InStr(tipo, "get") > 0 Then
            PKind = ProcKind.PropGet
        Else
            Stop ' debug
        End If
    ElseIf tipo = "sub" Then
        Returns = DEFAULT_NORETURNS
        If PKind = ProcKind.proc Then
            PKind = ProcKind.ProcSub
        Else
            Stop ' debug
        End If
    End If
    If PKind = ProcSub Then Returns = DEFAULT_NORETURNS
    
    Exit Function
ErrorHandler:
    LogCurrentError MODULE_NAME, "[TipoProcedimiento] - Error al procesar Code con expresiones regulares"
    Err.Raise Err.Number, MODULE_NAME & " [TipoProcedimiento]", Err.Description
End Function

'@Description: Determina si una funciÃ³n es visible para excel como macro o udf, o si es "interna",
' a partir de NormalizedSignature,PKind,ContainerType,bModuloPrivado,...
Private Function ParsearCodeParaDeducirContexto(CodeText As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\bThisWorkBook\b"
    ' lo siguiente serÃ­a una 'condicion de error' muy probable: no deberÃ­a, en general, manipular el XLAM
    If re.Test(CodeText) Then
        If ParsearCodeParaDeducirContexto <> "" Then _
                ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "|"
        ParsearCodeParaDeducirContexto = "ThisWorkbook"
    End If
    re.Pattern = "\bSelection\b"
    If re.Test(CodeText) Then
        If ParsearCodeParaDeducirContexto <> "" Then _
                ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "|"
        ParsearCodeParaDeducirContexto = "Selection"
    End If
    re.Pattern = "\bActiveWorkbook\b"
    If re.Test(CodeText) Then
        If ParsearCodeParaDeducirContexto <> "" Then _
                ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "|"
        ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "ActiveWorkbook"
    End If
    re.Pattern = "\bActiveSheet\b"
    If re.Test(CodeText) Then
        If ParsearCodeParaDeducirContexto <> "" Then _
                ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "|"
        ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "ActiveSheet"
    End If
    re.Pattern = "\b(?:Range|Cells)\b"
    If re.Test(CodeText) Then
        If ParsearCodeParaDeducirContexto <> "" Then _
                ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "|"
        ParsearCodeParaDeducirContexto = ParsearCodeParaDeducirContexto & "Cells Range"
    End If
End Function

'@Description: Normaliza firma del procedimiento: elimina continuaciones "_" y deja la firma en una sola lÃ­nea
Private Function NormalizarLineasFirma(ByVal raw As String, ByVal SignatureLine As Long) As String
    ' raw: texto del codio, con cabecera de procedimiento, e incluso varias lÃ­neas que conforman la firma (puede contener vbCrLf y "_" al final de lÃ­neas)
    Dim s As String
    s = raw
    ' eliminar vbCrLf que tienen _ al final, y unir
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    On Error GoTo ErrHandler                     '  & "," & SignatureLine
    re.Pattern = "^(?:.*(?:\r\n|\n|\r)){" & SignatureLine - 1 & "}" _
               & "(.+(?:_\s*(?:\r\n|\n|\r).+)*\))"
    s = re.Execute(s).Item(0).SubMatches(0)
    ' Quitar secuencias de continuation: " _" al final de lÃ­nea + CRLF -> vacÃ­o
    re.Pattern = "_\s*(\r\n|\n|\r)"
    s = re.Replace(s, " ")
    ' ahora sustituir saltos de lÃ­nea sobrantes por espacio y normalizar espacios
    re.Pattern = "(\r\n|\n|\r)"
    s = re.Replace(s, " ")
    ' colapsar mÃºltiples espacios
    re.Pattern = "\s+"
    s = Trim$(re.Replace(s, " "))
    NormalizarLineasFirma = s
    Exit Function
ErrHandler:
    LogCurrentError MODULE_NAME, "[NormalizarLineasFirma]"
    Err.Raise Err.Number, MODULE_NAME & " [NormalizarLineasFirma]", Err.Description
End Function

'@Description: Extrae contenido entre parÃ©ntesis de una firma normalizada y reemplaza comas por "|"
Private Function ComponerArgumentDescriptions(ByVal firmaNormalizada As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[^'\(]+\((.*?)\)\s*(?:'.+)?$"
    If re.Test(firmaNormalizada) Then
        Dim args As String
        args = Trim$(re.Execute(firmaNormalizada)(0).SubMatches(0))
        If args = "" Then
            ComponerArgumentDescriptions = DEFAULT_NOARGS
        Else
            ' remplazar "," por "|", respetando que pueda haber comas dentro de literales (caso raro en VBA)
            ' asumimos que no hay comas embebidas; si las hubiera, se necesitarÃ­a parsing mÃ¡s sofisticado.
            re.Pattern = "\s*,\s*"
            ComponerArgumentDescriptions = re.Replace(args, "|")
        End If
    Else
        ComponerArgumentDescriptions = DEFAULT_NOARGS
    End If
End Function

' Parsea metadatos de comentarios estructurados (@Description, @Category, etc.)
' Soporta formatos:
'   '@Tag: Valor
'   '@Tag("Valor")
'   '@Tag "Valor"
Private Function ParsearMetadataCompleta(CodeText As String)
    Dim lineText As Variant
    Dim regEx As Object
    Dim matches As Object

    Dim sDescription As String, sCategory As String, sArgumentDescriptions As String
    Dim sScope As String, sReturns As String, sExample As String
    Dim sRaises As String, sDependencies As String
    ' Inicializar con valores por defecto
    HasMetadata = False
    sDescription = ""
    sCategory = DEFAULT_CATEGORY
    sArgumentDescriptions = ""
    sScope = ""
    sReturns = ""
    sExample = ""
    sRaises = ""
    sDependencies = ""

    ' Configurar expresion regular
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True ' Ahora ignora mayusculas/minusculas

    Dim tag As String, value As String
    ' Procesar hacia delante
    For Each lineText In Split(CodeText, vbCrLf)
        If lineText <> "" Then
            ' Patron extendido que soporta:
            '   '@Tag: Valor
            '   '@Tag("Valor")
            '   '@Tag "Valor"
            ' Tags soportados: UDF, Macro, Description, Note, Nota, Scope, Category, Dependencies,
            '                  ArgumentDescriptions, Param, Returns, Return, Example, Raises, Throws
            regEx.Pattern = "^\s*'\s*(?:['\s=\-_]+|(?:@\s*(UDF|Macro|Description|Note|Nota|Scope|Category|Dependencies|ArgumentDescriptions|Param|Returns|Return|Example|Raises|Throws)(?:\s*[:\s])?))?\s*(.*?)\s*$"
            Set matches = regEx.Execute(lineText)
            If matches.Count = 0 Then Exit For
            If matches(0).SubMatches(0) <> "" Then
                tag = matches(0).SubMatches(0) ' La palabra clave (UDF, Description, etc.)
            End If
            value = Trim(matches(0).SubMatches(1)) ' El valor

            ' Extraer valor de formatos ("valor") o "valor"
            regEx.Pattern = "^\s*\(?\s*""(.*?)""\s*\)?\s*$"
            If regEx.Test(value) Then value = regEx.Execute(value).Item(0).SubMatches(0)

            ' Tambien soportar formato con comillas simples ('valor')
            If value = "" Then
                regEx.Pattern = "^\s*\(?\s*'(.*?)'\s*\)?\s*$"
                If regEx.Test(Trim(matches(0).SubMatches(1))) Then
                    value = regEx.Execute(Trim(matches(0).SubMatches(1))).Item(0).SubMatches(0)
                End If
            End If

            Select Case UCase$(tag)
            Case "UDF", "MACRO"
                HasMetadata = True
            Case "DESCRIPTION", "NOTE", "NOTA"
                If sDescription <> "" Then sDescription = sDescription & " "
                sDescription = sDescription & value
            Case "SCOPE"
                If sScope <> "" Then sScope = sScope & " "
                sScope = sScope & value
            Case "RETURNS", "RETURN"
                If sReturns <> "" Then sReturns = sReturns & " "
                sReturns = sReturns & value
            Case "CATEGORY"
                If sCategory <> DEFAULT_CATEGORY And sCategory <> "" Then sCategory = sCategory & " "
                If sCategory = DEFAULT_CATEGORY Then sCategory = ""
                sCategory = sCategory & value
                Select Case LCase$(sCategory)
                Case "hidden", "oculta", "ocultar", "-1"
                    sCategory = "-1"
                End Select
            Case "ARGUMENTDESCRIPTIONS", "PARAM"
                If sArgumentDescriptions <> "" Then sArgumentDescriptions = sArgumentDescriptions & "|"
                sArgumentDescriptions = sArgumentDescriptions & value
            Case "DEPENDENCIES"
                If sDependencies <> "" Then sDependencies = sDependencies & "|"
                sDependencies = sDependencies & value
            Case "EXAMPLE"
                If sExample <> "" Then sExample = sExample & vbCrLf
                sExample = sExample & value
            Case "RAISES", "THROWS"
                If sRaises <> "" Then sRaises = sRaises & "|"
                sRaises = sRaises & value
            Case Else ' comentarios SIN TAG: se atribuyen a Description
                If value <> "" Then
                    If sDescription <> "" Then sDescription = sDescription & ". "
                    sDescription = sDescription & value
                End If
            End Select
        End If
    Next

    ' Asignar propiedades de la clase
    Example = sExample
    Raises = sRaises
    Dependencies = sDependencies

    ParsearMetadataCompleta = Array(sDescription, sCategory, sArgumentDescriptions, sScope, sReturns)
End Function

'@Description: Genera descripciÃ³n automÃ¡tica basada en el nombre de la funciÃ³n
Private Function GenerarDescripcionAutomatica(ByVal nombreFuncion As String) As String
    If InStr(nombreFuncion, "_") > 0 Then GenerarDescripcionAutomatica = "(manejador de evento " & nombreFuncion & ")": Exit Function
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "([A-Z]+[a-z]*)"
    GenerarDescripcionAutomatica = Trim$(re.Replace(nombreFuncion, " $1")) & " (funciÃ³n personalizada)"
    Select Case PKind
        Case PropGet: GenerarDescripcionAutomatica = "Get " & GenerarDescripcionAutomatica
        Case PropLet: GenerarDescripcionAutomatica = "Let " & GenerarDescripcionAutomatica
        Case PropSet: GenerarDescripcionAutomatica = "Set " & GenerarDescripcionAutomatica
    End Select
End Function

Private Function dumpProcedimiento()
    Dim strdbg As String, strGen As String
    
    dumpProcedimiento = "Procedimiento: " & Name & ", en " & Module & vbCrLf & vbTab
    Select Case PKind
        Case PropGet: strdbg = strdbg & "property get"
        Case PropLet: strdbg = strdbg & "property let"
        Case PropSet: strdbg = strdbg & "property set"
        Case proc: strdbg = strdbg & "procedimiento"
        Case ProcSub: strdbg = strdbg & "Sub"
        Case ProcFunction: strdbg = strdbg & "Function"
    End Select
    strdbg = "en "
    Select Case ContainerType
        Case StdModule: strdbg = strdbg & "mÃ³dulo estandar": strGen = "o"
        Case ClassModule: strdbg = strdbg & "mÃ³dulo de clase": strGen = "o"
        Case Form: strdbg = strdbg & "formulario": strGen = "o"
        Case Sheet: strdbg = strdbg & "hoja de excel": strGen = "a"
    End Select
    strdbg = strdbg & " " & IIf(bPrivateModule, "privad", "pÃºblic") & strGen
 
    strdbg = IIf(ProcedureType = Macro, "Macro", IIf(ProcedureType = udf, "UDF", "interno a VBA")) & " (" & strdbg & ")"
    dumpProcedimiento = dumpProcedimiento & "Tipo:" & strdbg & vbCrLf & vbTab
    dumpProcedimiento = dumpProcedimiento & "Scope: " & IIf(Scope = "", "----------", Scope) & _
            vbTab & "Categoria: " & IIf(Category = "", "----------", Category)
    dumpProcedimiento = dumpProcedimiento & vbCrLf & vbTab
    dumpProcedimiento = dumpProcedimiento & "Argumentos: " & IIf(ArgumentDescriptions = "", "----------", ArgumentDescriptions) & _
            vbCrLf & vbTab & "Returns: " & IIf(Returns = "", "----------", Returns)
    ' Nuevos atributos
    If Example <> "" Then
        dumpProcedimiento = dumpProcedimiento & vbCrLf & vbTab & "Ejemplo: " & Example
    End If
    If Raises <> "" Then
        dumpProcedimiento = dumpProcedimiento & vbCrLf & vbTab & "Raises: " & Raises
    End If
    If Dependencies <> "" Then
        dumpProcedimiento = dumpProcedimiento & vbCrLf & vbTab & "Dependencies: " & Dependencies
    End If
End Function


