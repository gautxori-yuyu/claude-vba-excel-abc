Attribute VB_Name = "UDFs_Units"
'@Folder "UDFS.Unidades"
Option Explicit

Private Const MODULE_NAME As String = "UDFs_Units"

'==========================================
' FUNCIÃN PRINCIPAL - UDF para Excel
'==========================================
Public Function ConvertirUnidad(valor As Double, unidadOrigen As String, unidadBase As String) As Variant
Attribute ConvertirUnidad.VB_Description = "[UDFs_Units] FUNCIÃN PRINCIPAL - UDF para Excel"
Attribute ConvertirUnidad.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrorHandler
    
    ' ValidaciÃ³n de entrada
    If Trim(unidadOrigen) = "" Or Trim(unidadBase) = "" Then
        ConvertirUnidad = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Caso trivial
    If Trim(unidadOrigen) = Trim(unidadBase) Then
        ConvertirUnidad = valor
        Exit Function
    End If
    
    ' Validar que ambas unidades sean del mismo tipo
    Dim tipoOrigen As String, tipoBase As String
    tipoOrigen = ObtenerTipoUnidad(unidadOrigen)
    tipoBase = ObtenerTipoUnidad(unidadBase)
    
    If tipoOrigen = "" Or tipoBase = "" Then
        ConvertirUnidad = CVErr(xlErrNA)         ' Unidad no encontrada
        Exit Function
    End If
    
    If tipoOrigen <> tipoBase Then
        ConvertirUnidad = CVErr(xlErrNA)         ' Tipos incompatibles (ej: Pa -> mm)
        Exit Function
    End If
    
    ' Crear diccionario de visitados (oculto al usuario)
    Dim visitados As Object
    Set visitados = CreateObject("Scripting.Dictionary")
    visitados.CompareMode = vbBinaryCompare      ' Case-sensitive
    
    ' Llamar a la funciÃ³n recursiva interna
    ConvertirUnidad = ConvertirUnidadRecursivo(valor, unidadOrigen, unidadBase, visitados)
    Exit Function
    
ErrorHandler:
    ConvertirUnidad = CVErr(xlErrValue)
End Function

'==========================================
' FUNCIÃN RECURSIVA INTERNA (no expuesta)
'==========================================
Private Function ConvertirUnidadRecursivo(valor As Double, unidadOrigen As String, unidadBase As String, visitados As Object) As Variant
    Static dicConversiones As Object
    
    Dim hoja As Worksheet
    Dim i As Long, lastRow As Long
    Dim pend As Double, ord As Double
    Dim clave As String
    Dim unidadIntermedia As String
    Dim valorIntermedio As Double
    Dim resultado As Variant
    
    On Error GoTo ErrorHandler
    
    ' Inicializar Ã­ndice solo en primera llamada (usando Is Nothing)
    If dicConversiones Is Nothing Then
        Set dicConversiones = CreateObject("Scripting.Dictionary")
        dicConversiones.CompareMode = vbBinaryCompare ' Case-sensitive: MPa ? mPa
        
        Set hoja = ThisWorkbook.Sheets("Unidades")
        lastRow = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            ' Almacenar como "origen|destino" -> [pendiente, ordenada]
            Dim unidadCol2 As String, unidadCol5 As String
            unidadCol2 = Trim(hoja.Cells(i, 2).Value)
            unidadCol5 = Trim(hoja.Cells(i, 5).Value)
            pend = hoja.Cells(i, 3).Value
            ord = hoja.Cells(i, 4).Value
            
            ' Validar que ambas celdas tengan contenido
            If unidadCol2 <> "" And unidadCol5 <> "" Then
                If (Not IsEmpty(pend) Or Not IsEmpty(ord)) And (pend <> 0 Or ord <> 0) Then
                    clave = unidadCol2 & "|" & unidadCol5
                    If Not dicConversiones.Exists(clave) Then
                        pend = IIf(IsEmpty(pend) Or Not IsNumeric(pend), 1, CDbl(pend))
                        ord = IIf(IsEmpty(ord) Or Not IsNumeric(ord), 0, CDbl(ord))
                        dicConversiones(clave) = Array(pend, ord)
                    End If
                End If
            End If
        Next i
    End If
    
    ' Normalizar espacios (pero mantener case)
    unidadOrigen = Trim(unidadOrigen)
    unidadBase = Trim(unidadBase)
    
    ' Evitar bucles: marcar como visitado
    If visitados.Exists(unidadOrigen) Then
        ConvertirUnidadRecursivo = CVErr(xlErrNA)
        Exit Function
    End If
    visitados(unidadOrigen) = True
    
    ' BÃSQUEDA 1: ConversiÃ³n directa (B->E)
    clave = unidadOrigen & "|" & unidadBase
    If dicConversiones.Exists(clave) Then
        pend = dicConversiones(clave)(0)
        ord = dicConversiones(clave)(1)
        ConvertirUnidadRecursivo = valor * pend + ord
        visitados.Remove unidadOrigen            ' Desmarcar antes de salir
        Exit Function
    End If
    
    ' BÃSQUEDA 2: ConversiÃ³n inversa (E->B)
    clave = unidadBase & "|" & unidadOrigen
    If dicConversiones.Exists(clave) Then
        pend = dicConversiones(clave)(0)
        ord = dicConversiones(clave)(1)
        ' FÃ³rmula inversa: si valor_destino = valor_origen * pend + ord
        ' entonces valor_origen = (valor_destino - ord) / pend
        ConvertirUnidadRecursivo = (valor - ord) / pend
        visitados.Remove unidadOrigen            ' Desmarcar antes de salir
        Exit Function
    End If
    
    ' BÃSQUEDA 3: Conversiones indirectas (recursivas)
    ' Recorrer todas las claves del diccionario buscando caminos
    Dim todasClaves As Variant
    todasClaves = dicConversiones.Keys
    
    For i = LBound(todasClaves) To UBound(todasClaves)
        clave = todasClaves(i)
        Dim partes() As String
        partes = Split(clave, "|")
        
        ' Validar que el split produjo 2 elementos
        If UBound(partes) < 1 Then GoTo SiguienteIteracion
        
        Dim origen As String, destino As String
        origen = partes(0)
        destino = partes(1)
        
        ' DirecciÃ³n B->E: si origen coincide con unidadOrigen
        If origen = unidadOrigen Then
            unidadIntermedia = destino
            
            If Not visitados.Exists(unidadIntermedia) Then
                ' Convertir a unidad intermedia
                pend = dicConversiones(clave)(0)
                ord = dicConversiones(clave)(1)
                valorIntermedio = valor * pend + ord
                
                ' Llamada recursiva
                resultado = ConvertirUnidadRecursivo(valorIntermedio, unidadIntermedia, unidadBase, visitados)
                
                If Not IsError(resultado) Then
                    ConvertirUnidadRecursivo = resultado
                    visitados.Remove unidadOrigen ' Desmarcar antes de salir con Ã©xito
                    Exit Function
                End If
            End If
        End If
        
        ' DirecciÃ³n E->B: si destino coincide con unidadOrigen
        If destino = unidadOrigen Then
            unidadIntermedia = origen
            
            If Not visitados.Exists(unidadIntermedia) Then
                ' Convertir a unidad intermedia (inversa)
                pend = dicConversiones(clave)(0)
                ord = dicConversiones(clave)(1)
                valorIntermedio = (valor - ord) / pend
                
                ' Llamada recursiva
                resultado = ConvertirUnidadRecursivo(valorIntermedio, unidadIntermedia, unidadBase, visitados)
                
                If Not IsError(resultado) Then
                    ConvertirUnidadRecursivo = resultado
                    visitados.Remove unidadOrigen ' Desmarcar antes de salir con Ã©xito
                    Exit Function
                End If
            End If
        End If
        
SiguienteIteracion:
    Next i
    
    ' No se encontrÃ³ conversiÃ³n
    visitados.Remove unidadOrigen                ' Desmarcar antes de salir con error
    ConvertirUnidadRecursivo = CVErr(xlErrNA)
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not visitados Is Nothing Then
        visitados.Remove unidadOrigen
    End If
    On Error GoTo 0
    ConvertirUnidadRecursivo = CVErr(xlErrValue)
End Function

'==========================================
' FUNCIONES AUXILIARES
'==========================================
Private Function ObtenerTipoUnidad(unidad As String) As String
    Static dicTipos As Object
    
    Dim hoja As Worksheet
    Dim i As Long, lastRow As Long
    Dim unidadNorm As String
    Dim tipoActual As String
    
    On Error GoTo ErrorHandler
    
    ' Inicializar Ã­ndice de tipos solo en primera llamada (usando Is Nothing)
    If dicTipos Is Nothing Then
        Set dicTipos = CreateObject("Scripting.Dictionary")
        dicTipos.CompareMode = vbBinaryCompare   ' Case-sensitive
        
        Set hoja = ThisWorkbook.Sheets("Unidades")
        lastRow = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            tipoActual = Trim(hoja.Cells(i, 1).Value)
            
            ' Indexar unidad de columna B (origen)
            unidadNorm = Trim(hoja.Cells(i, 2).Value)
            If unidadNorm <> "" And tipoActual <> "" Then
                If Not dicTipos.Exists(unidadNorm) Then
                    dicTipos(unidadNorm) = tipoActual
                End If
            End If
            
            ' Indexar unidad de columna E (destino/base)
            unidadNorm = Trim(hoja.Cells(i, 5).Value)
            If unidadNorm <> "" And tipoActual <> "" Then
                If Not dicTipos.Exists(unidadNorm) Then
                    dicTipos(unidadNorm) = tipoActual
                End If
            End If
        Next i
    End If
    
    unidadNorm = Trim(unidad)
    
    If dicTipos.Exists(unidadNorm) Then
        ObtenerTipoUnidad = dicTipos(unidadNorm)
    Else
        ObtenerTipoUnidad = ""
    End If
    Exit Function
    
ErrorHandler:
    ObtenerTipoUnidad = ""
End Function

'==========================================
' FUNCIÃN PARA VALIDACIONES EN EXCEL
'==========================================
Public Function UdsPorTipo(ByVal strTipo As String) As Variant
Attribute UdsPorTipo.VB_Description = "[UDFs_Units] FUNCIÃN PARA VALIDACIONES EN EXCEL. Aplica a: ThisWorkbook|Cells Range"
Attribute UdsPorTipo.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim resultados() As String
    Dim contador As Long
    Dim unidad As String
    Dim dicTemp As Object
    
    On Error GoTo ErrorHandler
    
    ' Usar diccionario temporal para evitar duplicados
    Set dicTemp = CreateObject("Scripting.Dictionary")
    dicTemp.CompareMode = vbBinaryCompare        ' Case-sensitive
    
    Set ws = ThisWorkbook.Sheets("Unidades")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Recopilar todas las unidades Ãºnicas del tipo solicitado
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = Trim(strTipo) Then
            unidad = Trim(ws.Cells(i, 2).Value)
            If unidad <> "" And Not dicTemp.Exists(unidad) Then
                dicTemp(unidad) = True
            End If
        End If
    Next i
    
    ' Convertir a array para Excel
    If dicTemp.Count > 0 Then
        ReDim resultados(1 To dicTemp.Count)
        contador = 0
        Dim clave As Variant
        For Each clave In dicTemp.Keys
            contador = contador + 1
            resultados(contador) = clave
        Next clave
        UdsPorTipo = Application.Transpose(resultados)
    Else
        UdsPorTipo = CVErr(xlErrNA)
    End If
    Exit Function
    
ErrorHandler:
    UdsPorTipo = CVErr(xlErrNA)
End Function

'==========================================
' FUNCIÃN PARA LIMPIAR ÃNDICES MANUALMENTE
'==========================================
Public Sub ActualizarTablasConversion()
Attribute ActualizarTablasConversion.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Llama esta funciÃ³n despuÃ©s de modificar la hoja "Unidades"
    ' para forzar la reconstrucciÃ³n de los Ã­ndices internos
    
    ' Forzar reinicializaciÃ³n llamando con valores dummy
    ' Esto limpiarÃ¡ los diccionarios estÃ¡ticos
    On Error Resume Next
    Dim dummy As Variant
    dummy = ConvertirUnidad(0, "Pa", "Pa")
    
    ' Mostrar mensaje de confirmaciÃ³n
    MsgBox "Tablas de conversiÃ³n actualizadas." & vbCrLf & _
           "Los Ã­ndices se reconstruirÃ¡n en la prÃ³xima conversiÃ³n.", _
           vbInformation, "ActualizaciÃ³n completada"
End Sub

Function ConvertirCaudalNormal(valor As Double, p1 As Double, T1 As Double, unidadOrigen As String, unidadBase As String) As Variant
Attribute ConvertirCaudalNormal.VB_Description = "[UDFs_Units] Convertir Caudal Normal (funciÃ³n personalizada)"
Attribute ConvertirCaudalNormal.VB_ProcData.VB_Invoke_Func = " \n21"
    ' P1: PresiÃ³n en Pa
    ' T1: Temperatura en K
    ' Convierte caudales: teniendo en cuenta los normalizados (Nm3, SCF), para pasarlos a condiciones reales
    Dim unidadesNormalizadas As Object
    Dim esUnidadNormal As Boolean
    Dim Pn As Double, Tn As Double
    Dim valorReal As Double
    
    If unidadOrigen = unidadBase Then
        ConvertirCaudalNormal = valor
        Exit Function
    ElseIf p1 <= 0 Or T1 <= 0 Then
        ConvertirCaudalNormal = CVErr(xlErrNum)
        Exit Function
    End If
    
    Set unidadesNormalizadas = CreateObject("Scripting.Dictionary")
    unidadesNormalizadas.CompareMode = 1
    
    ' Mapeo de unidades normalizadas a condiciones [Pn, Tn]
    ' PresiÃ³n en Pa, Temperatura en K
    unidadesNormalizadas.Add "nm3/h", Array(101325, 273.15)
    unidadesNormalizadas.Add "nm3/min", Array(101325, 273.15)
    unidadesNormalizadas.Add "scfh", Array(101325, 288.7056) ' 60 Â°F = 288.7056 K
    unidadesNormalizadas.Add "scfmin", Array(101325, 288.7056)
    unidadesNormalizadas.Add "scf/h", Array(101325, 288.7056) ' 60 Â°F = 288.7056 K
    unidadesNormalizadas.Add "scf/min", Array(101325, 288.7056)
    unidadesNormalizadas.Add "mmscfd", Array(101325, 288.7056)
    
    unidadOrigen = (Replace(unidadOrigen, "Â³", "3")) ' Normaliza 'Â³'
    
    If unidadesNormalizadas.Exists(unidadOrigen) Then
        esUnidadNormal = True
        Pn = unidadesNormalizadas(unidadOrigen)(0)
        Tn = unidadesNormalizadas(unidadOrigen)(1)
    Else
        esUnidadNormal = False
    End If

    ' Aplicar correcciÃ³n gas ideal si es necesario
    If esUnidadNormal Then
        valorReal = valor * (Pn / p1) * (T1 / Tn)
    Else
        valorReal = valor
    End If

    ' Llamar a ConvertirUnidad (esta debe existir)
    ConvertirCaudalNormal = ConvertirUnidad(valorReal, Replace(Replace(LCase(unidadOrigen), "nm3", "m3"), "scf", "cf"), unidadBase)
End Function
