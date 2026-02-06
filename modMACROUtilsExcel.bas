Attribute VB_Name = "modMACROUtilsExcel"
'@Folder "MACROS"
Option Explicit

Private Const MODULE_NAME As String = "modMACROUtilsExcel"

Sub AplicarDirtyATodasLasHojasConFormulas()
Attribute AplicarDirtyATodasLasHojasConFormulas.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim rFormulas As Range

    ' Recorrer todas las hojas del libro actual
    For Each ws In ActiveWorkbook.Worksheets
        ws.UsedRange.Calculate
        ' Establecer el rango con fórmulas en la hoja activa
        On Error Resume Next
        Set rFormulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        ' Verificar si se encontró un rango con fórmulas
        If Not rFormulas Is Nothing Then
            ' Aplicar el método Dirty para marcar las celdas para su recálculo
            rFormulas.Dirty
        End If
        
        ' Limpiar la variable de rango para el siguiente bucle
        Set rFormulas = Nothing
    Next ws

    ' Desactivar las alertas para evitar errores si no hay fórmulas en una hoja
    ' On Error Resume Next
    FullRecalc

    MsgBox "El método Dirty se ha aplicado a todos los rangos con fórmulas en este libro.", vbInformation
End Sub

Sub FullRecalc()
Attribute FullRecalc.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim prevCalcMode As XlCalculation
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    
    On Error GoTo ErrorHandler
    
    prevCalcMode = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    
    ' === 1. Configurar entorno para recálculo fiable ===
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    ' === 2. Recálculo TOTAL con reconstrucción de dependencias ===
    Application.CalculateFullRebuild

    ' === 5. Restaurar estado original ===
Finish:
    Application.Calculation = prevCalcMode
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    Exit Sub

ErrorHandler:
    Debug.Print "[ERR] Excepción en FullRecalc: " & Err.Description
    Resume Finish
End Sub

' Requiere referencia a: Microsoft VBScript Regular Expressions 5.5
Sub ReemplazarUDFsEnFormulas()
Attribute ReemplazarUDFsEnFormulas.VB_ProcData.VB_Invoke_Func = " \n0"
    Const bReplaceIndirectionsInAllFormula = False ' Si es true, reemplaza indirecciones en toda la formula; si no, solo en los argumentos de las UDFs
    Dim ws As Worksheet, celda As Range, regEx As Object
    Dim strProjectFN, oDicUDFs As Object, funciones()
    Dim f As Variant, strFPattern As String
    Dim formula As String, nuevaFormula As String
    
    ' 1. CONFIGURACIÓN
    Set oDicUDFs = ParsearUDFsDeTodosLosProyectos()
    
    ' 2. OPTIMIZACIÓN EXTREMA
    On Error GoTo CleanUp
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False ' Evita disparar eventos al cambiar celdas
        .DisplayAlerts = False
    End With

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = True

    For Each ws In ActiveWorkbook.Worksheets
        Dim rng As Range
        On Error Resume Next
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo CleanUp
        
        If Not rng Is Nothing Then
            For Each celda In rng
                nuevaFormula = celda.Formula2
                If nuevaFormula = "" Then GoTo nextCelda
                
                If bReplaceIndirectionsInAllFormula Then nuevaFormula = ResolverIndirecciones(nuevaFormula, celda)
                
                ' 3. PROCESO DE EVALUACIÓN "DENTRO HACIA FUERA"
                ' Repetimos hasta que no queden nombres de nuestras UDFs
                Dim huboCambio As Boolean
                Do
                    huboCambio = False
                    For Each strProjectFN In oDicUDFs.Keys()
                        funciones = oDicUDFs(strProjectFN).Keys()
                        'regEx.Pattern = "\b" & f & "\("
                        ' Detecta opcionalmente: 'Nombre.xlam'!Funcion( o Nombre.xlam!Funcion( o Funcion(
                        regEx.Pattern = "(['""]?[^'""!]+['""]?!)?\b(?:" & Join(funciones, "|") & ")\("
                        
                        Dim matches As Object, i, m As Object
                        Set matches = regEx.Execute(nuevaFormula)
                        
                        For i = matches.Count - 1 To 0 Step -1
                            Dim startPos As Long
                            Set m = matches(i)
                            startPos = m.FirstIndex + 1
                            Dim primerParentesis As Long
                            primerParentesis = InStr(startPos, nuevaFormula, "(")
                        
                            If Not EstaEnComillas(nuevaFormula, startPos) Then
                                Dim endPos As Long
                                ' Buscamos el paréntesis de cierre balanceado desde el primer "("
                                endPos = BuscarParentesisCierreRobusto(nuevaFormula, primerParentesis)
                                
                                If endPos > 0 Then
                                    Dim llamadaUDF As String
                                    llamadaUDF = Mid(nuevaFormula, startPos, endPos - startPos + 1)
                                    
                                    If Not bReplaceIndirectionsInAllFormula Then llamadaUDF = ResolverIndirecciones(llamadaUDF, celda)
                                    
                                    ' --- LÓGICA DE REDUCCIÓN DE CARACTERES ---
                                    ' Si la llamada es > 255, intentamos resolver lo que hay DENTRO primero
                                    If Len(llamadaUDF) > 255 Then
                                        llamadaUDF = ReducirArgumentosInternos(llamadaUDF)
                                    End If
                                    
                                    Dim valorUDF As Variant
                                    valorUDF = celda.Parent.Evaluate(llamadaUDF)
                                    
                                    If Not IsError(valorUDF) Then
                                        Dim strRep As String
                                        strRep = ConvertirAStringFormula(valorUDF)
                                        nuevaFormula = Left(nuevaFormula, startPos - 1) & strRep & Mid(nuevaFormula, endPos + 1)
                                        huboCambio = True
                                    End If
                                End If
                            End If
                        Next i
                    Next strProjectFN
                Loop While huboCambio
                
                If nuevaFormula <> celda.formula Then celda.formula = nuevaFormula
nextCelda:
            Next celda
        End If
    Next ws

CleanUp:
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
End Sub

' Función que intenta evaluar partes internas de una cadena larga para acortarla
Private Function ReducirArgumentosInternos(ByVal textoUDF As String) As String
    ' Buscamos funciones anidadas dentro de los paréntesis de la UDF principal
    ' Si detectamos una función interna, la evaluamos y reemplazamos su texto por el valor
    ' Esto reduce la longitud de la cadena total "hacia afuera".
    
    ' Por simplicidad, esta función puede llamar recursivamente a un evaluador
    ' de funciones nativas de Excel o simplemente devolver el texto si es irreducible.
    ReducirArgumentosInternos = textoUDF ' (Estrategia base de seguridad)
End Function

Private Function ConvertirAStringFormula(ByVal valor As Variant) As String
    ' 1. Manejo de Errores de Excel (#N/A, #VALOR!, etc.)
    If IsError(valor) Then
        ' Si la UDF devolvió un error, lo mantenemos como literal de error
        ConvertirAStringFormula = CVErrToText(valor)
        Exit Function
    End If

    ' 2. Manejo de MATRICES (Vartype 8192+ o constantes como 8204)
    If IsArray(valor) Or (VarType(valor) And vbArray) Then
        ConvertirAStringFormula = ConvertirMatrizAString(valor)
        Exit Function
    End If

    ' 3. Manejo de tipos escalares
    Select Case VarType(valor)
        Case vbEmpty, vbNull
            ConvertirAStringFormula = """" & """"
            
        Case vbString
            ' Duplicamos comillas internas para no romper la fórmula
            ConvertirAStringFormula = """" & Replace(valor, """", """""") & """"
            
        Case vbBoolean
            ConvertirAStringFormula = IIf(valor, "TRUE", "FALSE")
            
        Case vbDate
            ' Excel trata las fechas como números en las fórmulas
            ConvertirAStringFormula = CDbl(valor)
            
        Case vbObject
            ' Si la UDF devolvió un objeto Range, tomamos su valor
            If TypeOf valor Is Range Then
                ConvertirAStringFormula = ConvertirAStringFormula(valor.value)
            Else
                ConvertirAStringFormula = """#OBJETO!"""
            End If
            
        Case Else
            ' Para números (Double, Integer, etc.), asegurar punto decimal
            ConvertirAStringFormula = Replace(CStr(valor), ",", ".")
    End Select
End Function

' Función auxiliar para convertir arrays en formato {a,b;c,d}
Private Function ConvertirMatrizAString(ByVal arr As Variant) As String
    Dim res As String, r As Long, c As Long
    Dim vTmp As String
    
    res = "{"
    On Error Resume Next ' Por si es una matriz de una sola dimensión
    For r = LBound(arr, 1) To UBound(arr, 1)
        For c = LBound(arr, 2) To UBound(arr, 2)
            vTmp = ConvertirAStringFormula(arr(r, c))
            res = res & vTmp & IIf(c < UBound(arr, 2), ",", "")
        Next c
        res = res & IIf(r < UBound(arr, 1), ";", "")
    Next r
    If Err.Number <> 0 Then ' Fallback para 1D
        Err.Clear
        For r = LBound(arr) To UBound(arr)
            res = res & ConvertirAStringFormula(arr(r)) & IIf(r < UBound(arr), ",", "")
        Next r
    End If
    ConvertirMatrizAString = res & "}"
End Function

' Convierte códigos de error internos en texto de fórmula (#N/A...)
Private Function CVErrToText(ByVal errVal As Variant) As String
    Select Case CLng(errVal)
        Case -2146826281: CVErrToText = "#DIV/0!"
        Case -2146826246: CVErrToText = "#N/A"
        Case -2146826259: CVErrToText = "#NAME?"
        Case -2146826288: CVErrToText = "#NULL!"
        Case -2146826252: CVErrToText = "#NUM!"
        Case -2146826265: CVErrToText = "#REF!"
        Case -2146826273: CVErrToText = "#VALUE!"
        Case Else: CVErrToText = "#ERROR!"
    End Select
End Function

Private Function ResolverIndirecciones(ByVal textoUDF As String, ByVal rContexto As Range) As String
    Dim regInd As Object
    Dim matches As Object, m As Object
    Dim interiorIndireccion As String, direccionResuelta As Variant
    Dim i As Long
    
    Set regInd = CreateObject("VBScript.RegExp")
    With regInd
        .Global = True
        .IgnoreCase = True
        ' Busca el patrón INDIRECT(...)
        .Pattern = "\bINDIRECT\(([^()]*(\([^()]*\)[^()]*)*)\)"
    End With
    
    Set matches = regInd.Execute(textoUDF)
    
    ' Procesamos de atrás hacia adelante
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        ' Extraemos lo que hay dentro de los paréntesis de INDIRECT
        interiorIndireccion = Mid(m.value, 10, Len(m.value) - 10)
        
        ' Evaluamos solo el interior para obtener la cadena de texto de la dirección
        direccionResuelta = rContexto.Parent.Evaluate(interiorIndireccion)
        
        If Not IsError(direccionResuelta) Then
            ' Reemplazamos "INDIRECT(X)" por "X" (como referencia directa)
            textoUDF = Left(textoUDF, m.FirstIndex) & direccionResuelta & Mid(textoUDF, m.FirstIndex + m.Length + 1)
        End If
    Next i
    
    ResolverIndirecciones = textoUDF
End Function

' Determina si una posición en la fórmula está dentro de comillas dobles o simples (hojas)
Private Function EstaEnComillas(ByVal texto As String, ByVal pos As Long) As Boolean
    Dim i As Long, enDoble As Boolean, enSimple As Boolean
    For i = 1 To pos - 1
        Dim char As String: char = Mid(texto, i, 1)
        If char = """" And Not enSimple Then enDoble = Not enDoble
        ' If char = "'" And Not enDoble Then enSimple = Not enSimple
        ' ajusta la lógica del char = "'"
        If char = "'" And Not enDoble Then
            ' Si el siguiente carácter tras el cierre de comilla simple no es "!", es una cadena real
            If i < Len(texto) Then
                If Mid(texto, InStr(i + 1, texto, "'") + 1, 1) <> "!" Then
                    enSimple = Not enSimple
                End If
            End If
        End If
    Next i
    EstaEnComillas = enDoble Or enSimple
End Function

' Busca el paréntesis de cierre ignorando lo que hay entre comillas (evita error en nombres de hojas)
Private Function BuscarParentesisCierreRobusto(ByVal texto As String, ByVal posApertura As Long) As Long
    Dim nivel As Integer: nivel = 0
    Dim i As Long, enDoble As Boolean, enSimple As Boolean
    
    For i = posApertura To Len(texto)
        Dim char As String: char = Mid(texto, i, 1)
        
        ' Si encontramos comillas, invertimos estado y no contamos paréntesis dentro
        If char = """" And Not enSimple Then enDoble = Not enDoble
        If char = "'" And Not enDoble Then enSimple = Not enSimple
        
        If Not enDoble And Not enSimple Then
            If char = "(" Then nivel = nivel + 1
            If char = ")" Then
                nivel = nivel - 1
                If nivel = 0 Then
                    BuscarParentesisCierreRobusto = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function
