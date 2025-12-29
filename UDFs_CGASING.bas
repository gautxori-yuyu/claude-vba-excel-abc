Attribute VB_Name = "UDFs_CGASING"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

' Variable global para evitar múltiples mensajes
Private bMessageGases As Boolean

'@UDF
'@Description: Comprueba si la hoja activa tiene el formato C-GAS-ING por defecto (sin modificar)
'@Category: Comprobación de formato de ficheros
'@ArgumentDescriptions: (sin argumentos)
Public Function IsDefaultCGasIngSheet() As Boolean
Attribute IsDefaultCGasIngSheet.VB_Description = "[UDFs_CGASING] Comprueba si la hoja activa tiene el formato C-GAS-ING por defecto (sin modificar). Aplica a: ActiveSheet|Cells Range"
Attribute IsDefaultCGasIngSheet.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo NoSheet
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then GoTo NoSheet
    
    ' Verificar etiquetas clave
    If UCase(Trim(ws.Cells(2, 2).Value)) <> "CALCULATION - GAS" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(15, 1).Value)) <> "INPUT DATA" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(44, 1).Value)) <> "OUTPUT DATA" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(60, 1).Value)) <> "STAGES" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(85, 1).Value)) <> "COOLERS" Then GoTo NoSheet
    
    ' Verificar celdas específicas del formato por defecto
    If Trim(ws.Cells(24, 1).Value) <> "Specific weight in normal conditions:" Or _
       Trim(ws.Cells(35, 7).Value) <> "Water temperature :" Or _
       Trim(ws.Cells(46, 8).Value) = "Factor =" Or _
       Trim(ws.Cells(47, 1).Value) <> "RPM :" Then GoTo NoSheet
    
    IsDefaultCGasIngSheet = True
    Exit Function
    
NoSheet:
    IsDefaultCGasIngSheet = False
End Function

'@UDF
'@Description: Verifica si una hoja tiene el formato C-GAS-ING estándar (modificado)
'@Category: Comprobación de formato de ficheros
'@ArgumentDescriptions: Hoja de excel a verificar
Public Function IsCGASING(ws As Worksheet) As Boolean
Attribute IsCGASING.VB_Description = "[UDFs_CGASING] Verifica si una hoja tiene el formato C-GAS-ING estándar (modificado). Aplica a: Cells Range"
Attribute IsCGASING.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo NoSheet
    
    If UCase(Trim(ws.Cells(2, 2).Value)) <> "CALCULATION - GAS" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(15, 1).Value)) <> "INPUT DATA" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(35, 1).Value)) <> "OUTPUT DATA" Then GoTo NoSheet
    If UCase(Trim(ws.Cells(47, 1).Value)) <> "STAGES" Then GoTo NoSheet
    
    IsCGASING = True
    Exit Function
    
NoSheet:
    IsCGASING = False
End Function

'@UDF
'@Description: Extrae y concatena nombres de gases con sus porcentajes, identificando tipos comunes (SYNGAS, BIOGAS, NATURAL GAS)
'@Category: Análisis de Gases
'@ArgumentDescriptions: Rango de celdas con nombres de gases|Separador entre gases (por defecto ", ")|Hoja C-GAS-ING (opcional)
Public Function Gases(r As Range, Optional d As String = ", ", Optional CGASINGSheet As Worksheet) As String
Attribute Gases.VB_Description = "[UDFs_CGASING] Extrae y concatena nombres de gases con sus porcentajes, identificando tipos comunes (SYNGAS, BIOGAS, NATURAL GAS). Aplica a: Cells Range"
Attribute Gases.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim s As String, c As Range
    Dim cCellPc As Double
    Dim oGasesPC As Object
    Dim oH2OPCCell As Range
    Dim total As Double, pcMultiplier As Double
    
    Set oGasesPC = CreateObject("Scripting.Dictionary")
    
    If CGASINGSheet Is Nothing Then Set CGASINGSheet = Worksheets("C-GAS-ING")
    
    ' Identificar si es gas saturado en agua
    On Error Resume Next
    Set oH2OPCCell = CGASINGSheet.Cells.SpecialCells(xlCellTypeFormulas).Find( _
                     What:="Relative humidity :", _
                     LookIn:=xlValues, _
                     LookAt:=xlPart, _
                     SearchOrder:=xlByRows, _
                     SearchDirection:=xlNext, _
                     MatchCase:=False, _
                     SearchFormat:=False)
    
    If Not oH2OPCCell Is Nothing Then
        Set oH2OPCCell = oH2OPCCell.Offset(0, 2)
    End If
    On Error GoTo 0
    
    ' Determinar multiplicador de porcentaje
    pcMultiplier = 1
    On Error GoTo ErrorHandler
    total = Application.WorksheetFunction.Sum(r.Offset(0, 1))
    On Error GoTo 0
    If total < 1.5 And total > 0 Then pcMultiplier = 100
    
    ' Procesar cada gas
    For Each c In r
        If Len(c.text) > 0 Then
            ' Validación de coherencia (solo mostrar una vez)
            If Not bMessageGases Then
                If InStr(LCase(c.text), "air") > 0 And _
                   InStr(fileName, "LT") > 0 And _
                   InStr(fileName, "GT") > 0 Then
                    MsgBox "La referencia del compresor debe terminar en 'LT', por ser compresor de aire"
                End If
                
                If Trim(LCase(Replace(c.text, ":", ""))) = "h2" And _
                   InStr(fileName(), "LGT") = 0 Then
                    MsgBox "La referencia del compresor debe terminar en 'LGT', por ser compresor de HIDRÓGENO - requiere distanciador largo"
                End If
                
                bMessageGases = True
            End If
            
            cCellPc = CDbl(Replace(c.Offset(0, 1).Value, "%", "")) * pcMultiplier
            
            If cCellPc > 1 Then
                ' Agregar gas a la cadena
                s = s & IIf(Len(s), d, vbNullString) & Trim(Replace(c.text, ":", ""))
                
                ' Marcar si es gas saturado
                If Trim(Replace(c.text, ":", "")) = "H2O" And _
                   Not oH2OPCCell Is Nothing Then
                    If oH2OPCCell.Value = "100 %" Then
                        If Worksheets("BUDGET_ENTRY").Cells.SpecialCells(xlCellTypeFormulas).Find( _
                           What:="LANGUAGE", LookIn:=xlValues).Offset(0, 3).Value = "EN" Then
                            s = s & " (saturated)"
                        Else
                            s = s & " (saturado)"
                        End If
                    End If
                End If
                
                ' Guardar porcentaje en diccionario
                oGasesPC.Add Trim(Replace(c.text, ":", "")), cCellPc
            End If
        End If
    Next c
    
    ' Identificar tipos de gas especiales
    On Error Resume Next
    If oGasesPC("H2") > 10 And oGasesPC("CO") > 10 Then s = "SYNGAS"
    If oGasesPC("CH4") > 50 And oGasesPC("CH4") < 75 And oGasesPC("CO2") > 25 And oGasesPC("CO2") < 45 Then s = "BIOGAS"
    If oGasesPC("CH4") > 75 Then s = "NATURAL GAS"
    On Error GoTo 0
    
    Gases = s
    
    Exit Function
ErrorHandler:
    Gases = CVErr(xlErrNA)
End Function

'@UDF
'@Description: Genera el nombre del modelo de compresor basado en configuración de cilindros y presiones
'@Category: Información del compresor
'@ArgumentDescriptions: Hoja C-GAS-ING con los datos del compresor (opcional)
Public Function strModelName(Optional CGASINGSheet As Worksheet = Nothing) As String
Attribute strModelName.VB_Description = "[UDFs_CGASING] Genera el nombre del modelo de compresor basado en configuración de cilindros y presiones. Aplica a: ActiveSheet|Cells Range"
Attribute strModelName.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim strCils As String, nEtapa As Long
    Dim c As Long, d As String, i As Long
    Dim bEtapaReqForjado As Boolean, bEtapaConvieneCamisa As Boolean
    Dim bATEX_Inflamable As Boolean, bAire As Boolean, bNACE_Corrosivo As Boolean, bSafeZone As Boolean
    Dim oCell_Compressor_Serie As Range
    Dim oCell_Stage_CilsDiam As Range, oCell_Stage_Pout As Range
    Dim ncils As Long
    
    On Error GoTo ErrorHandler
    
    If CGASINGSheet Is Nothing Then Set CGASINGSheet = ActiveSheet
    If Not IsCGASING(CGASINGSheet) Then strModelName = CVErr(xlErrNA): Exit Function
    
    ' Obtener serie del compresor
    c = 17
    Set oCell_Compressor_Serie = CGASINGSheet.Range("B" & c)
    
    ' Detectar tipo de gas
    ' bATEX_Inflamable es "TODO ATEX"; y bSafeZone es "ZONIFICACION", poner armarios en zona segura, y compresor NO ATEX
    bSafeZone = Not IsError(BuscarRegex(CGASINGSheet.Range("F19:F29"), "\bH2\b|C\d*H\d*"))
    bAire = Not IsError(BuscarRegex(CGASINGSheet.Range("F19:F29"), "\bAire?\b"))
    bATEX_Inflamable = CGASINGSheet.Range("I32").Value < 12
    
    ' Procesar cada etapa
    For i = Asc("B") To Asc("G")
        d = Chr(i)
        c = 50
        Set oCell_Stage_CilsDiam = CGASINGSheet.Range(d & c)
        
        If Trim(oCell_Stage_CilsDiam.Value) = "" Then Exit For
        
        c = c + 2
        Set oCell_Stage_Pout = CGASINGSheet.Range(d & c)
        
        nEtapa = nEtapa + 1
        If InStr(UCase(oCell_Stage_CilsDiam), "T") > 0 Then
            ncils = ncils + 0.5 * Split(UCase(oCell_Stage_CilsDiam), " X ")(0) * 1
        Else
            ncils = ncils + Split(UCase(oCell_Stage_CilsDiam), " X ")(0) * 1
        End If
        
        strCils = strCils & "-" & Replace(UCase(oCell_Stage_CilsDiam.Value), " X ", "x")
        
        ' Determinar si requiere forjado
        bEtapaReqForjado = (oCell_Stage_Pout.Value > 80)
        
        ' Diámetros pequeños requieren forjado
        bEtapaReqForjado = bEtapaReqForjado Or _
                           (CDbl(Replace(Split(UCase(oCell_Stage_CilsDiam.Value), " X ")(1), "T", "")) <= 75)
        
        bEtapaConvieneCamisa = bEtapaReqForjado
        
        ' ATEX requiere forjado desde presiones más bajas
        bEtapaReqForjado = bEtapaReqForjado Or _
                           (bATEX_Inflamable And oCell_Stage_Pout.Value > 80 * 0.85)
        ' compresores HP o HX, DE DIAMETROS GRANDES, Y A ALTAS PRESIONES (ya incluso inferiores a 80 bar)... convendría que fuesen encamisados
        ' (se hace en fundicion nodular o acero fundido el cuerpo, y el liner añade proteccion)
        bEtapaReqForjado = bEtapaReqForjado Or _
                           (CDbl(Replace(Split(UCase(oCell_Stage_CilsDiam.Value), " X ")(1), "T", "")) >= 450 _
                            And oCell_Stage_Pout.Value > 65)
        
        ' Añadir sufijos de construcción
        If bEtapaReqForjado Then
            strCils = strCils & "FC"
            'ElseIf bNACE_Corrosivo Then ' EN ESTE CONTEXTO NO ME COMPENSA DETERMINAR SI se aplica NACE
            ' en gases corrosivos, la camisa protege el cuerpo del cilindro. API 618 practicamente LO EXIGE en caso de NACE
            '    strCils = strCils & "C"
        ElseIf bEtapaConvieneCamisa Then
            strCils = strCils & "(C)"
        End If
    Next i
    
    ' Construir nombre del modelo
    strModelName = nEtapa
    If InStr(UCase(strCils), "T") > 0 Then strModelName = strModelName & "T"
    strModelName = strModelName & "E" & oCell_Compressor_Serie.Value & "-" & ncils & "-"
    
    If bATEX_Inflamable Then strModelName = strModelName & "L"
    
    If bAire Then
        strModelName = strModelName & "LT" & strCils
    Else
        strModelName = strModelName & "GT" & strCils
    End If
    
    If bNACE_Corrosivo Then strModelName = strModelName & " NACE"
    If bATEX_Inflamable Then strModelName = strModelName & " ATEX"
    If bSafeZone Then strModelName = strModelName & " (ATEX)"
    
    Exit Function
    
ErrorHandler:
    strModelName = "#ERROR: " & Err.Description
End Function

'@UDF
'@Description: Devuelve un array con los nombres de todas las hojas C-GAS-ING del libro activo
'@Category: Comprobación de formato de ficheros
'@ArgumentDescriptions: Libro de excel al que se aplica (opcional)
Public Function HojasCGASING(Optional wb As Workbook = Nothing) As Variant
Attribute HojasCGASING.VB_Description = "[UDFs_CGASING] Devuelve un array con los nombres de todas las hojas C-GAS-ING del libro activo. Aplica a: ActiveWorkbook"
Attribute HojasCGASING.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim ws As Worksheet
    Dim tmp() As String
    Dim n As Long
    
    On Error GoTo ErrorHandler
    
    n = 0
    If wb Is Nothing Then Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        If IsCGASING(ws) Then
            ReDim Preserve tmp(0 To n)
            tmp(n) = ws.Name
            n = n + 1
        End If
    Next ws
    
    If n = 0 Then
        HojasCGASING = Array("")
    Else
        HojasCGASING = Application.Transpose(tmp)
    End If
    
    Exit Function
    
ErrorHandler:
    HojasCGASING = CVErr(xlErrNA)
End Function

'@UDF
'@Description: Devuelve el máximo valor de potencia encontrado en todas las hojas C-GAS-ING del libro
'@Category: Análisis de Gases
'@ArgumentDescriptions: Celda donde buscar el valor de potencia
Public Function MaximaPotencia(ByVal CeldaBuscada As Range) As Variant
Attribute MaximaPotencia.VB_Description = "[UDFs_CGASING] Devuelve el máximo valor de potencia encontrado en todas las hojas C-GAS-ING del libro. Aplica a: Cells Range"
Attribute MaximaPotencia.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim strHoja As Variant
    Dim maxVal As Double, val As Variant
    Dim valorTexto As String
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    
    maxVal = 0
    Set wb = CeldaBuscada.Worksheet.Parent
    For Each strHoja In HojasCGASING(wb)
        If CStr(strHoja) <> "" Then
            val = wb.Worksheets(CStr(strHoja)).Range(CeldaBuscada.Address).Value
            
            ' Extraer valor numérico del formato "xxx / yyy HP/kW"
            valorTexto = CStr(val)
            If InStr(valorTexto, "/") > 0 Then
                valorTexto = Mid(valorTexto, InStr(valorTexto, "/") + 1)
                If InStr(valorTexto, "HP/") > 0 Then
                    valorTexto = Left(valorTexto, InStr(valorTexto, "HP/") - 1)
                End If
            End If
            
            If IsNumeric(Trim(valorTexto)) Then
                If CDbl(Trim(valorTexto)) > maxVal Then
                    maxVal = CDbl(Trim(valorTexto))
                End If
            End If
        End If
    Next strHoja
    
    If maxVal > 0 Then
        MaximaPotencia = maxVal
    Else
        MaximaPotencia = CVErr(xlErrNA)
    End If
    
    Exit Function
    
ErrorHandler:
    MaximaPotencia = CVErr(xlErrNA)
End Function


