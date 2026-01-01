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
'@Folder "1-Inicio e Instalacion"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

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

Public Scope As String                           ' ambito al que se aplica: selección, hoja activa, libro activo, argumentos de la funcion, ...
Public ArgumentDescriptions As String            ' Separado por "|"
Public Returns As String                         ' Qué devuelve

Public HasMetadata As Boolean                    ' TRUE si tiene comentarios @UDF
Public Description As String
Public Category As String

' Constructor para fácil inicialización
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
    ' Si no hay descripción del procedimiento, generar una básica
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
    Debug.Print "[clsVBAProcedure Init] - Error: " & Err.Description
End Sub

'@Description: Determina si una función es visible para excel como macro o udf, o si es "interna",
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
        ' Normalmente deberían ser privados, dentro de una clase que instancie un objeto de otra clase que implemente eventos
        'If acceso <> "private" Then Stop
        ' y no se si puede llegar a haber manejadores que sean funciones... por si acaso:
        'If tipo <> "sub" Then Stop
    ElseIf tipo = "sub" And acceso = "public" And Not bPrivateModule And ContainerType = StdModule _
           And m.SubMatches(3) <> "" Then
        TipoProcedimiento = ProcType.internalSubPublicWithParams
    ElseIf acceso <> "public" Or bPrivateModule Or ContainerType <> StdModule Then
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
    Debug.Print "[TipoProcedimiento] - Error al procesar Code con expresiones regulares: " & Err.Description
End Function

'@Description: Determina si una función es visible para excel como macro o udf, o si es "interna",
' a partir de NormalizedSignature,PKind,ContainerType,bModuloPrivado,...
Private Function ParsearCodeParaDeducirContexto(CodeText As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\bThisWorkBook\b"
    ' lo siguiente sería una 'condicion de error' muy probable: no debería, en general, manipular el XLAM
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

'@Description: Normaliza firma del procedimiento: elimina continuaciones "_" y deja la firma en una sola línea
Private Function NormalizarLineasFirma(ByVal raw As String, ByVal SignatureLine As Long) As String
    ' raw: texto del codio, con cabecera de procedimiento, e incluso varias líneas que conforman la firma (puede contener vbCrLf y "_" al final de líneas)
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
    ' Quitar secuencias de continuation: " _" al final de línea + CRLF -> vacío
    re.Pattern = "_\s*(\r\n|\n|\r)"
    s = re.Replace(s, " ")
    ' ahora sustituir saltos de línea sobrantes por espacio y normalizar espacios
    re.Pattern = "(\r\n|\n|\r)"
    s = re.Replace(s, " ")
    ' colapsar múltiples espacios
    re.Pattern = "\s+"
    s = Trim$(re.Replace(s, " "))
    NormalizarLineasFirma = s
    Exit Function
ErrHandler:
    Debug.Print "[NormalizarLineasFirma] - Error al procesar Code con expresiones regulares: " & Err.Description
End Function

'@Description: Extrae contenido entre paréntesis de una firma normalizada y reemplaza comas por "|"
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
            ' asumimos que no hay comas embebidas; si las hubiera, se necesitaría parsing más sofisticado.
            re.Pattern = "\s*,\s*"
            ComponerArgumentDescriptions = re.Replace(args, "|")
        End If
    Else
        ComponerArgumentDescriptions = DEFAULT_NOARGS
    End If
End Function

'TODO: añadir el metadato dependencies (vendrá bien para representar la jerarquia de funciones en el excel)
Private Function ParsearMetadataCompleta(CodeText As String)
    Dim lineText As Variant
    Dim regEx As Object
    Dim matches As Object
    
    Dim Description As String, Category As String, ArgumentDescriptions As String
    Dim Scope As String, Returns As String
    ' Inicializar con valores por defecto
    HasMetadata = False
    Description = ""
    Category = DEFAULT_CATEGORY
    ArgumentDescriptions = ""
    Scope = ""
    Returns = ""
    
    ' Configurar expresión regular
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = False
    
    Dim tag As String, Value As String
    ' Procesar hacia delante
    For Each lineText In Split(CodeText, vbCrLf)
        If lineText <> "" Then
            'regex.Pattern = "\s*'(?:['\s=\-_]+|(?:@\s*(UDF|Description|Not[a-e]|Scope|Category|ArgumentDescriptions|Param|Returns)(?:\s*:)?))?(?:\s*\(\s*""?)?\s*(.*)\s*(?:""?\s*\)?\s*)"
            regEx.Pattern = "\s*'\s*(?:['\s=\-_]+|(?:@\s*(UDF|Description|Not[a-e]|Scope|Category|Dependencies|ArgumentDescriptions|Param|Returns)(?:\s*:)?))?\s*(.*?)\s*$"
            Set matches = regEx.Execute(lineText)
            If matches.Count = 0 Then Exit For
            If tag <> matches(0).SubMatches(0) And matches(0).SubMatches(0) <> "" Then
                tag = matches(0).SubMatches(0) ' La palabra clave (UDF, Description, etc.)
            End If
            Value = Trim(matches(0).SubMatches(1)) ' El valor
            regEx.Pattern = "^\s*\(""\s*(.*?)\s*""\)\s*$"
            If regEx.Test(Value) Then Value = regEx.Execute(Value).Item(0).SubMatches(0)
            
            Select Case UCase$(tag)
            Case "UDF", "MACRO"
                HasMetadata = True
            Case "DESCRIPTION", "NOTE", "NOTA"
                If Description <> "" Then Description = Description & " "
                Description = Description & Value
            Case "SCOPE"
                If Scope <> "" Then Scope = Scope & " "
                Scope = Scope & Value
            Case "RETURNS", "RETURN"
                If Returns <> "" Then Returns = Returns & " "
                Returns = Returns & Value
            Case "CATEGORY"
                If Category <> "" Then Category = Category & " "
                Category = Category & Value
                Select Case LCase$(Category)
                Case "hidden", "oculta", "ocultar", "-1"
                    Category = -1
                End Select
            Case "ARGUMENTDESCRIPTIONS", "PARAM"
                If ArgumentDescriptions <> "" Then ArgumentDescriptions = ArgumentDescriptions & " "
                ArgumentDescriptions = ArgumentDescriptions & Value
            Case "DEPENDENCIES"
                ' de momento no se procesa, PTE añadir
            Case Else ' comentarios SIN TAG: se atribuyen a Description
                If Description <> "" Then Description = Description & ". "
                Description = Description & Value
            End Select
        End If
    Next
    If InStr(Description, "@Description") > 0 Then Stop
    ParsearMetadataCompleta = Array(Description, Category, ArgumentDescriptions, Scope, Returns)
End Function

'@Description: Genera descripción automática basada en el nombre de la función
Private Function GenerarDescripcionAutomatica(ByVal nombreFuncion As String) As String
    If InStr(nombreFuncion, "_") > 0 Then GenerarDescripcionAutomatica = "(manejador de evento " & nombreFuncion & ")": Exit Function
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "([A-Z]+[a-z]*)"
    GenerarDescripcionAutomatica = Trim$(re.Replace(nombreFuncion, " $1")) & " (función personalizada)"
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
        Case StdModule: strdbg = strdbg & "módulo estandar": strGen = "o"
        Case ClassModule: strdbg = strdbg & "módulo de clase": strGen = "o"
        Case Form: strdbg = strdbg & "formulario": strGen = "o"
        Case Sheet: strdbg = strdbg & "hoja de excel": strGen = "a"
    End Select
    strdbg = strdbg & " " & IIf(bPrivateModule, "privad", "públic") & strGen
 
    strdbg = IIf(ProcedureType = Macro, "Macro", IIf(ProcedureType = udf, "UDF", "interno a VBA")) & " (" & strdbg & ")"
    dumpProcedimiento = dumpProcedimiento & "Tipo:" & strdbg & vbCrLf & vbTab
    dumpProcedimiento = dumpProcedimiento & "Scope: " & IIf(Scope = "", "----------", Scope) & _
            vbTab & "Categoria: " & IIf(Category = "", "----------", Category)
    dumpProcedimiento = dumpProcedimiento & vbCrLf & vbTab
    dumpProcedimiento = dumpProcedimiento & "Argumentos: " & IIf(ArgumentDescriptions = "", "----------", ArgumentDescriptions) & _
            vbCrLf & vbTab & "Returns: " & IIf(Returns = "", "----------", Returns)
End Function


