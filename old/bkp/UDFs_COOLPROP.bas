Attribute VB_Name = "UDFs_COOLPROP"
'@IgnoreModule MissingAnnotationArgument
'@Folder "UDFS"
Option Explicit

Private Const MODULE_NAME As String = "UDFs_COOLPROP"

'@UDF
'@Description: Construye una cadena HEOS::... para CoolProp procesando uno o varios rangos
'              disjuntos. Cada área del rango es tratada como una tabla independiente con
'              encabezados ("nombre/gas" y "%/percentage") en la primera fila.
'@Scope:
'@ArgumentDescriptions: rango: Rango contiguo o disjunto, donde cada área contiene columnas válidas
'                       de nombres y porcentajes. (columnas: Nombre/Gas y %/Percentage)
'@Returns: String|Cadena HEOS para CoolProp o mensaje de error si la suma ? 100%
'@Category: Análisis de Gases
Public Function ConstruirCadenaCoolPropDesdeTabla(rango As range) As String
Attribute ConstruirCadenaCoolPropDesdeTabla.VB_Description = "[UDFs_COOLPROP] Construye una cadena HEOS::.. para CoolProp procesando uno o varios rangos disjuntos. Cada área del rango es tratada como una tabla independiente con encabezados (""nombre/gas"" y ""%/percentage"") en la primera fila. Aplica a: Cells Range"
Attribute ConstruirCadenaCoolPropDesdeTabla.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim aliasDict As Object
    Set aliasDict = CreateObject("Scripting.Dictionary")
    aliasDict.CompareMode = 1
    aliasDict.Add "C2H6", "n-C2H6"
    aliasDict.Add "CH4O", "METHANOL"
    aliasDict.Add "Ar", "ARGON"
    
    Dim mezcla As String: mezcla = "HEOS::"
    Dim suma As Double: suma = 0
    Dim primera As Boolean: primera = True
    
    ' Rangos colectivos que contendrán TODAS las columnas de Nombres/Porcentajes
    Dim rngNombresGlobal As range
    Dim rngPorcentajesGlobal As range
    
    Dim area As range
    Dim cabecera As range
    
    ' --- Fase 1: Identificar todas las columnas relevantes en todas las áreas ---
    For Each area In rango.Areas
        For Each cabecera In area.Rows(1).Cells
            Select Case LCase(Trim(cabecera.Value))
            Case "nombre", "gas"
                ' Añadir la columna completa al rango global de nombres
                If rngNombresGlobal Is Nothing Then
                    Set rngNombresGlobal = cabecera.EntireColumn
                Else
                    ' Usar Union para combinar columnas, aunque esten disjuntas
                    Set rngNombresGlobal = Application.Union(rngNombresGlobal, cabecera.EntireColumn)
                End If
            Case "%", "percentage"
                ' Añadir la columna completa al rango global de porcentajes
                If rngPorcentajesGlobal Is Nothing Then
                    Set rngPorcentajesGlobal = cabecera.EntireColumn
                Else
                    Set rngPorcentajesGlobal = Application.Union(rngPorcentajesGlobal, cabecera.EntireColumn)
                End If
            End Select
        Next cabecera
    Next area
    
    ' --- Fase 2: Recolectar datos y emparejarlos ---
    
    If rngNombresGlobal Is Nothing Or rngPorcentajesGlobal Is Nothing Then
        ConstruirCadenaCoolPropDesdeTabla = "Error: Faltan columnas 'Nombre' o 'Porcentaje'."
        Exit Function
    End If
    
    ' Intersectamos las columnas globales con el rango original del usuario.
    ' Esto crea un rango multi-area que solo incluye las celdas seleccionadas que son relevantes.
    Dim celdasNombresValidas As range
    Dim celdasPorcentajesValidas As range
    
    Set celdasNombresValidas = Application.Intersect(rango, rngNombresGlobal)
    Set celdasPorcentajesValidas = Application.Intersect(rango, rngPorcentajesGlobal)
    
    If celdasNombresValidas Is Nothing Or celdasPorcentajesValidas Is Nothing Then
        ConstruirCadenaCoolPropDesdeTabla = "Error: No hay datos válidos en las áreas seleccionadas."
        Exit Function
    End If
    
    ' Asumimos que la lista de celdas de nombres y porcentajes son del mismo tamaño
    ' y están en el mismo orden (e.g., A8, A9, A10; F8, F9, F10).
    ' Este bucle itera sobre las celdas válidas (excepto la fila 1, que son cabeceras)
    
    Dim c As Long
    ' Usamos arrays para acceder por índice, ya que collections no lo permiten fácilmente con Rangos disjuntos
    Dim arrNombres() As Variant, nombre As String, nombreRaw As String, valP As String, porcentaje As Double
    Dim arrPorcentajes() As Variant
    
    ' Convertir rangos disjuntos a arrays simples
    arrNombres = CellsToArray(celdasNombresValidas)
    arrPorcentajes = CellsToArray(celdasPorcentajesValidas)
    
    If UBound(arrNombres) <> UBound(arrPorcentajes) Then
        ConstruirCadenaCoolPropDesdeTabla = "Error: Desajuste en el número de valores."
        Exit Function
    End If

    ' Iterar sobre los arrays emparejados (empezando por el índice 1, asumiendo que la fila 1 es cabecera y ya la hemos ignorado)
    ' Nota: La función CellsToArray maneja la omisión de la primera fila.
    For c = 1 To UBound(arrNombres)
        nombreRaw = Trim(CStr(arrNombres(c)))
        valP = arrPorcentajes(c)
        
        If nombreRaw <> "" And IsNumeric(valP) And valP <> "" Then
            porcentaje = CDbl(valP)
            If porcentaje > 1 Then porcentaje = porcentaje / 100
            suma = suma + porcentaje
            nombre = nombreRaw
            If aliasDict.Exists(nombre) Then nombre = aliasDict(nombre)
            If Not primera Then mezcla = mezcla & "&"
            mezcla = mezcla & nombre & "[" & Replace(Format(porcentaje, "0.0000"), ",", ".") & "]"
            primera = False
        End If
    Next c
    
    ' --- Manejo de errores y retorno final ---
    If Abs(suma - 1) > 0.001 Then
        ConstruirCadenaCoolPropDesdeTabla = "Error: suma <> 100% (" & Format(suma * 100, "0.00") & "%)"
    Else
        ConstruirCadenaCoolPropDesdeTabla = mezcla
    End If
End Function

' --- Función de ayuda para convertir un rango disjunto en un array plano ---
Private Function CellsToArray(inputRange As range) As Variant()
    Dim coll As New Collection
    Dim area As range
    Dim cell As range
    
    For Each area In inputRange.Areas
        ' Omitimos la primera fila de cada área, ya que son cabeceras
        If area.Rows.Count > 1 Then
            For Each cell In area.Offset(1).Resize(area.Rows.Count - 1).Cells
                coll.Add cell.Value
            Next cell
        End If
    Next area
    
    Dim tempArr() As Variant
    ReDim tempArr(1 To coll.Count)               ' Reindexamos a base 1 para facilitar el bucle posterior
    Dim i As Long
    For i = 1 To coll.Count
        tempArr(i) = coll(i)
    Next i
    
    CellsToArray = tempArr
End Function
