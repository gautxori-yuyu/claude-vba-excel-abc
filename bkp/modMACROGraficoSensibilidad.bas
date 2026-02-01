Attribute VB_Name = "modMACROGraficoSensibilidad"

'@Folder "6-DOMINIO-Oportunidades y compresores.b-Calculos técnicos"
Option Explicit

Private Const MODULE_NAME As String = "modMACROGraficoSensibilidad"

Dim iSeriesNr As Integer

Public Function EsFicheroOportunidad() As Boolean
Attribute EsFicheroOportunidad.VB_Description = "[modMACROGraficoSensibilidad] Es Fichero Oportunidad (función personalizada). Aplica a: ActiveWorkbook"
Attribute EsFicheroOportunidad.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Z]{3}\d{5}_\d{2}"          ' patrón esperado en el nombre del fichero
    re.IgnoreCase = True
    Select Case True
    Case ActiveWindow Is Nothing, ActiveWindow.Visible = False, ActiveWorkbook Is Nothing
        Exit Function
    End Select
    EsFicheroOportunidad = re.Test(ActiveWorkbook.Name)
End Function

Public Function EsValidoGenerarGrafico() As Boolean
Attribute EsValidoGenerarGrafico.VB_Description = "[modMACROGraficoSensibilidad] Es Valido Generar Grafico (función personalizada). Aplica a: ActiveWorkbook|Cells Range"
Attribute EsValidoGenerarGrafico.VB_ProcData.VB_Invoke_Func = " \n23"
    'On Error GoTo ErrorHandler
    Dim hoja As Worksheet
    Dim r As range
    Dim encabezados As range, datos As range
    Dim formula As String
    
    Select Case True
    Case ActiveWindow Is Nothing, ActiveWindow.Visible = False, ActiveWorkbook Is Nothing, Application.Workbooks.Count = 0
        GoTo ErrorHandler
    End Select
    
    For Each hoja In ActiveWorkbook.Worksheets
        If Not IsNumeric(hoja.Name) Then GoTo Siguiente
        
        ' Validar que solo hay una tabla en la hoja
        If hoja.ListObjects.Count <> 1 Then GoTo ErrorHandler
        
        ' Validar que la tabla comience en A1
        If hoja.ListObjects(1).range.Cells(1, 1).Address <> "$A$1" Then GoTo ErrorHandler
        
        Set r = hoja.range("A1").CurrentRegion
        If r.Rows.Count < 2 Or r.Columns.Count < 2 Then GoTo Siguiente
        
        ' Validar encabezados (fila 1, desde la segunda columna)
        Set encabezados = r.Rows(1).Offset(0, 1).Resize(1, r.Columns.Count - 1)
        formula = "SUMPRODUCT(--ISNUMBER(SEARCH(""("", " & encabezados.Address(External:=True) & ")))"
        If Evaluate(formula) <> encabezados.Columns.Count Then GoTo ErrorHandler
        
        ' Validar datos numéricos (todo menos la primera fila y primera columna)
        Set datos = r.Offset(1, 1).Resize(r.Rows.Count - 1, r.Columns.Count - 1)
        formula = "SUMPRODUCT(--ISNUMBER(" & datos.Address(External:=True) & "))"
        If Evaluate(formula) <> datos.Cells.Count Then GoTo ErrorHandler
        
        EsValidoGenerarGrafico = True
        Exit Function
Siguiente:
    Next hoja
    
ErrorHandler:
    EsValidoGenerarGrafico = False
End Function

' Comprueba si el gráfico activo es válido para invertir ejes
Public Function EsValidoInvertirEjes() As Boolean
Attribute EsValidoInvertirEjes.VB_Description = "[modMACROGraficoSensibilidad] Comprueba si el gráfico activo es válido para invertir ejes. Aplica a: Selection|ActiveWorkbook|ActiveSheet"
Attribute EsValidoInvertirEjes.VB_ProcData.VB_Invoke_Func = " \n23"
    'On Error Resume Next
    Dim Ch As Chart
    Select Case True
    Case ActiveWindow Is Nothing, ActiveWindow.Visible = False, ActiveWorkbook Is Nothing, Application.Workbooks.Count = 0
        Exit Function
    End Select
    Set Ch = ActiveChart
    If Ch Is Nothing Then Exit Function
    ' 3. Caso de hoja de gráfico activa (ChartSheet)
    If Not Application.ActiveChart Is Nothing Then
        ' Validar que la hoja activa es de tipo Chart
        If TypeName(Application.ActiveSheet) = "Chart" Then
            EsValidoInvertirEjes = True
        End If
    End If
    ' 4. Caso de gráfico incrustado en hoja de cálculo (ChartObject)
    Dim sel
    Set sel = Application.Selection
    If Not sel Is Nothing Then
        Select Case TypeName(sel)
        Case "ChartObject", "ChartArea"
            ' No tengo nada claro que en este tipo de objetos se puedan invertir los ejes; en todo caso, NO son los que yo creo
            EsValidoInvertirEjes = True
        Case "DrawingObjects", "Picture", "Shape", "GroupObject", "OLEObject", "TextBox"
            ' Explícitamente NO es un gráfico
            EsValidoInvertirEjes = False
        Case Else
            ' Otras selecciones no válidas
            EsValidoInvertirEjes = False
        End Select
    End If
    
    If Not EsValidoInvertirEjes Then Exit Function
    
    ' Debe haber al menos una serie en cada eje para que tenga sentido
    Dim tienePrimario As Boolean, tieneSecundario As Boolean
    Dim s As Series
    For Each s In Ch.SeriesCollection
        If s.AxisGroup = xlPrimary Then
            tienePrimario = True
        ElseIf s.AxisGroup = xlSecondary Then
            tieneSecundario = True
        End If
    Next s
    
    EsValidoInvertirEjes = tienePrimario And tieneSecundario
End Function

' Ejecuta la macro para cada hoja válida del libro activo
Public Sub EjecutarGraficoEnLibroActivo()
Attribute EjecutarGraficoEnLibroActivo.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler
    Const SHEET_NAME As String = "Graficos"
    Const A4_WIDTH_POINTS As Double = 595        ' Aproximado a A4 horizontal en puntos
    Const GRAPH_ASPECT_RATIO As Double = 3 / 2   ' Proporción ancho/alto (3:2)

    ' Validaciones iniciales (ANTES de modificar estado)
    If Not EsValidoGenerarGrafico Then
        MsgBox "El libro no cumple los requisitos para generar el gráfico.", vbExclamation
        Exit Sub
    End If

    Dim Wb As Workbook
    Set Wb = ActiveWorkbook

    If Wb.FileFormat = xlOpenXMLAddIn Or Wb.FileFormat = xlAddIn Then
        MsgBox "No se puede ejecutar este comando sobre un archivo de tipo complemento (.xlam o .xla).", vbCritical
        Exit Sub
    End If

    ' Determinar hojas de datos válidas
    Dim hojasProcesar As Collection
    Set hojasProcesar = New Collection

    Dim i As Long, c As Long: c = 1
    For i = 1 To Wb.Sheets.Count
        If Wb.Sheets(i).Type = xlWorksheet Then
            If IsNumeric(Wb.Sheets(i).Name) Then
                If CLng(Wb.Sheets(i).Name) = c Then
                    hojasProcesar.Add Wb.Sheets(i)
                    c = c + 1
                End If
            End If
        End If
    Next i
    If hojasProcesar.Count = 0 Then
        MsgBox "No hay hojas de datos que procesar, no es un fichero de curvas de rendimiento", vbInformation
        Exit Sub
    End If

    ' En el caso de que solo se genere UNA serie de variaciones (presion, temp, etc),
    ' SOLO SE CREA UNA hoja de datos; pero para DOS O MAS SERIES, se crea una hoja adicional, con TODOS LOS DATOS.
    ' en este segundo caso SE DESCARTA LA ULTIMA HOJA
    If hojasProcesar.Count > 1 Then hojasProcesar.Remove (hojasProcesar.Count)

    ' Verificar si existe la hoja "Graficos"
    Dim chartSheet As Worksheet
    Dim SheetExists As Boolean: SheetExists = False
    On Error Resume Next
    Set chartSheet = Wb.Sheets(SHEET_NAME)
    On Error GoTo ErrHandler
    If Not chartSheet Is Nothing Then SheetExists = True

    If SheetExists Then
        Dim totalShapes As Long: totalShapes = chartSheet.ChartObjects.Count
        Dim msg As String

        If totalShapes > hojasProcesar.Count Then
            msg = "La hoja '" & SHEET_NAME & "' ya contiene " & totalShapes & " gráficos." & vbCrLf & _
                  "Esto incluye gráficos personalizados añadidos manualmente." & vbCrLf & _
                  "¿Deseas eliminar todos los gráficos y generar nuevos?"
        Else
            msg = "La hoja '" & SHEET_NAME & "' ya existe." & vbCrLf & _
                  "¿Deseas eliminar sus gráficos y generar nuevos?"
        End If

        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox(msg, vbQuestion + vbYesNoCancel, "Reemplazar gráficos")

        If respuesta = vbCancel Then Exit Sub
        If respuesta = vbYes Then
            chartSheet.ChartObjects.Delete
        Else
            Exit Sub
        End If
    ElseIf hojasProcesar.Count > 0 Then
        Set chartSheet = Wb.Sheets.Add(After:=Wb.Sheets(Wb.Sheets.Count))
        chartSheet.Name = SHEET_NAME
    End If

    ' === GUARDAR ESTADO ORIGINAL ===
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating

    ' === MODIFICAR ESTADO PARA MEJOR RENDIMIENTO ===
    DoEvents
    Application.ScreenUpdating = False

    ' Configurar impresión de la hoja de gráficos
    With chartSheet.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ' Crear gráficos en la hoja "Graficos"
    Dim topOffset As Double: topOffset = 20
    Dim graficoAltura As Double: graficoAltura = (A4_WIDTH_POINTS - 40) / GRAPH_ASPECT_RATIO
    Dim espacio As Double: espacio = 30
    Dim chartObj As ChartObject
    Dim ws As Worksheet

    For Each ws In hojasProcesar
        Call TraducirEncabezados(ws)
        Set chartObj = chartSheet.ChartObjects.Add(Left:=20, Width:=A4_WIDTH_POINTS - 40, Top:=topOffset, Height:=graficoAltura)
        Call GenerarGraficoSensibilidad(ws, chartObj)
        topOffset = topOffset + graficoAltura + espacio
    Next ws
    DoEvents

    ' Exportar la hoja "Graficos" a PDF
    Dim rutaArchivo As String
    rutaArchivo = Wb.Path & "\" & Left(Wb.Name, InStrRev(Wb.Name, ".") - 1) & "_Graficos.pdf"

    On Error Resume Next
    chartSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=rutaArchivo, Quality:=xlQualityStandard
    On Error GoTo ErrHandler

    ' Mostrar mensaje de éxito (solo si no hubo error)
    MsgBox "Se han generado " & hojasProcesar.Count & " gráficos en la hoja '" & SHEET_NAME & "'." & vbCrLf & _
           "Además, se ha guardado un PDF en la carpeta del fichero Excel, con el nombre:" & vbCrLf & Mid(rutaArchivo, InStrRev(rutaArchivo, "\") + 1) & vbCrLf & vbCrLf & _
           "(si se hacen cambios en la hoja de gráficos, para preservar su contenido basta con cambiar su nombre)", vbInformation

CleanUp:
    ' === RESTAURAR ESTADO ORIGINAL ===
    Application.ScreenUpdating = prevScreenUpdating
    DoEvents
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "[EjecutarGraficoEnLibroActivo]", Err.Number, Err.Description
    MsgBox "Error en GenerarGraficoSensibilidad: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Sub GenerarGraficoSensibilidad(ws As Worksheet, chartObj As ChartObject)
    Dim col As Long, xCol As Long, lastCol As Long
    Dim variableCols() As Long
    Dim group1() As Long, group2() As Long
    Dim i As Long
    On Error GoTo ManejoErrores
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ReDim variableCols(1 To lastCol - 1)
    Dim countVar As Long: countVar = 0
    
    ' Identificar columnas con valores que varían (excepto encabezados)
    For col = 1 To lastCol
        If InStr(ws.Cells(1, col).Value, "Agua") = 0 Then
            Dim firstVal As Variant
            ' Buscar el primer valor numérico en la columna para usar como referencia
            firstVal = Empty
            Dim filaTemp As Long
            For filaTemp = 2 To ws.Cells(ws.Rows.Count, col).End(xlUp).Row
                If IsNumeric(ws.Cells(filaTemp, col).Value) Then
                    firstVal = CDbl(ws.Cells(filaTemp, col).Value)
                    Exit For
                End If
            Next filaTemp
            
            ' Si no hay valores numéricos, pasamos a la siguiente columna
            If IsEmpty(firstVal) Then GoTo SiguienteColumna
            
            For i = filaTemp + 1 To ws.Cells(ws.Rows.Count, col).End(xlUp).Row
                If IsNumeric(ws.Cells(i, col).Value) Then
                    Dim currentVal As Double
                    currentVal = CDbl(ws.Cells(i, col).Value)
                    If Abs(currentVal - firstVal) > 0.00000001 Then
                        countVar = countVar + 1
                        variableCols(countVar) = col
                        Exit For
                    End If
                End If
            Next i
        End If
SiguienteColumna:
    Next col
    If countVar = 0 Then Exit Sub
    ReDim Preserve variableCols(1 To countVar)
    
    ' La primera columna variable es el eje X
    xCol = variableCols(1)
    
    ' Excluir eje X de columnas a representar
    Dim dataCols() As Long
    If countVar > 1 Then
        ReDim dataCols(1 To countVar - 1)
        For i = 2 To countVar
            dataCols(i - 1) = variableCols(i)
        Next i
    Else
        Exit Sub                                 ' No hay columnas para representar
    End If
    
    ' Agrupar columnas por variación y etiquetas de eje
    Call AgruparColumnasPorVariacion(ws, dataCols, group1, group2)
    If IsEmpty(group1) Then group1 = group2: group2 = Empty
    
    ' Formatear columnas como número con dos decimales
    If Not IsEmpty(group1) Then Call FormatColumnsAsDecimal(ws, group1)
    If Not IsEmpty(group2) Then Call FormatColumnsAsDecimal(ws, group2)
    
    ' modificar gráfico
    With chartObj.Chart
        .ChartType = xlLineMarkers
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' Colores
    Dim palette() As Long
    palette = GetColorPalette()
    
    ' Añadir series
    iSeriesNr = 0
    If Not IsEmpty(group1) Then Call AddGroupSeriesToChart(chartObj.Chart, ws, xCol, group1, False, palette)
    If Not IsEmpty(group2) Then Call AddGroupSeriesToChart(chartObj.Chart, ws, xCol, group2, True, palette)
    
    ' Ajustar ejes
    If Not IsEmpty(group1) Then
        Call AjustarEjeDesdeDatos(chartObj.Chart.Axes(xlValue), ws, group1)
    End If
    If Not IsEmpty(group2) Then
        Call AjustarEjeDesdeDatos(chartObj.Chart.Axes(xlValue, xlSecondary), ws, group2)
    End If
    
    ' Títulos de ejes
    With chartObj.Chart
        .HasTitle = True
        .ChartTitle.Text = "Correlation to " & Trim(Split(ws.Cells(1, xCol).Value, "(")(0))
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = ws.Cells(1, xCol).Value
        If Not IsEmpty(group1) Then
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = GenerarTituloEjeVertical(ws, group1)
            '.Axes(xlValue).AxisTitle.Text = ConcatenarTextosEntreParentesis(ws, group1)
            '.Axes(xlValue).AxisTitle.Text = ExtraerTextoEnParentesis(ws.Cells(1, group1(0)).value)
        End If
        If Not IsEmpty(group2) Then
            .Axes(xlValue, xlSecondary).HasTitle = True
            .Axes(xlValue, xlSecondary).AxisTitle.Text = GenerarTituloEjeVertical(ws, group2)
            '.Axes(xlValue).AxisTitle.Text = ConcatenarTextosEntreParentesis(ws, group2)
            '.Axes(xlValue).AxisTitle.Text = ExtraerTextoEnParentesis(ws.Cells(1, group2(0)).value)
        End If
    End With
    Exit Sub
ManejoErrores:
    Application.ScreenUpdating = True
    MsgBox "Error en GenerarGraficoSensibilidad para la hoja '" & ws.Name & "': " & Err.Description, vbCritical
End Sub

' Intercambia ejes primario/secundario en el gráfico activo
' Preserva títulos, rangos, propiedades, y extiende los ejes temporalmente para evitar desaparición de series
Public Sub InvertirEjesDelGraficoActivo()
Attribute InvertirEjesDelGraficoActivo.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo SinGrafico
    
    Dim chrt As Chart
    Set chrt = ActiveChart
    If chrt Is Nothing Then
        If Not (ActiveWindow.Selection Is Nothing) Then
            MsgBox ("tipo del objeto seleccionado:" & TypeName(ActiveWindow.Selection))
        Else
            MsgBox ("no hay nada seleccionado")
        End If
        GoTo SinGrafico
    End If
    
    Dim seriesPrimarias As New Collection
    Dim seriesSecundarias As New Collection
    
    ' Clasificar series
    Dim s As Series
    For Each s In chrt.SeriesCollection
        If s.AxisGroup = xlPrimary Then
            seriesPrimarias.Add s
        ElseIf s.AxisGroup = xlSecondary Then
            seriesSecundarias.Add s
        End If
    Next s
    
    If seriesPrimarias.Count = 0 And seriesSecundarias.Count = 0 Then
        MsgBox "El gráfico no contiene series con ejes diferenciados.", vbExclamation
        Exit Sub
    End If
    
    ' Capturar propiedades de ejes (valores, títulos, etc.)
    Dim propsPrimario As Object, propsSecundario As Object
    Set propsPrimario = CapturarPropiedadesDeEje(chrt.Axes(xlValue, xlPrimary))
    If chrt.HasAxis(xlValue, xlSecondary) Then
        Set propsSecundario = CapturarPropiedadesDeEje(chrt.Axes(xlValue, xlSecondary))
    Else
        Set propsSecundario = CreateObject("Scripting.Dictionary") ' vacío
    End If
    
    ' Expandir temporalmente el rango de ambos ejes para contener todos los valores
    Dim minTotal As Double, maxTotal As Double
    minTotal = WorksheetFunction.Min(chrt.Axes(xlValue, xlPrimary).MinimumScale, _
                                     chrt.Axes(xlValue, xlSecondary).MinimumScale)
    maxTotal = WorksheetFunction.Max(chrt.Axes(xlValue, xlPrimary).MaximumScale, _
                                     chrt.Axes(xlValue, xlSecondary).MaximumScale)
    
    On Error Resume Next
    chrt.HasAxis(xlValue, xlPrimary) = True
    chrt.HasAxis(xlValue, xlSecondary) = True
    chrt.Axes(xlValue, xlPrimary).MinimumScale = minTotal
    chrt.Axes(xlValue, xlPrimary).MaximumScale = maxTotal
    chrt.Axes(xlValue, xlSecondary).MinimumScale = minTotal
    chrt.Axes(xlValue, xlSecondary).MaximumScale = maxTotal
    On Error GoTo 0
    
    ' === Cambio ordenado de series ===
    
    ' 1. Mover series del secundario al primario primero
    Dim serie As Series
    For Each serie In seriesSecundarias
        serie.AxisGroup = xlPrimary
    Next serie
    
    ' 2. Luego mover las series del primario al secundario
    For Each serie In seriesPrimarias
        serie.AxisGroup = xlSecondary
    Next serie
    
    ' === Restaurar propiedades de los ejes ===
    AplicarPropiedadesAEje chrt.Axes(xlValue, xlPrimary), propsSecundario
    AplicarPropiedadesAEje chrt.Axes(xlValue, xlSecondary), propsPrimario
    
    MsgBox "Ejes y propiedades intercambiados correctamente.", vbInformation
    Exit Sub
    
SinGrafico:
    Application.ScreenUpdating = True
    MsgBox "Selecciona un gráfico válido antes de ejecutar este comando.", vbCritical
End Sub

Private Sub TraducirEncabezados(ws As Worksheet)
    Dim reemplazos As Variant
    reemplazos = Array( _
                 Array("Presión Aspiración", "Suction Pressure"), _
                 Array("Metros", "Meters"), _
                 Array("Presión Escape", "Exhaust Pressure"), _
                 Array("Temperatura Aspiración", "Suction Temperature"), _
                 Array("Temperatura Agua", "Water Temperature"), _
                 Array("Temperatura Ambiente", "Ambient Temperature"), _
                 Array("Caudal", "Flow Rate"), _
                 Array("- seco", "- Dry"), _
                 Array("Temperatura Escape", "Exhaust Temperature"), _
                 Array("Potencia Consumida", "Power Consumption"), _
                 Array("Potencia Instalar", "Installed Power") _
                 )
    
    Dim col As Long, i As Long
    Dim celda As range
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set celda = ws.Cells(1, col)
        For i = LBound(reemplazos) To UBound(reemplazos)
            If InStr(1, celda.Value, reemplazos(i)(0), vbTextCompare) > 0 Then
                celda.Value = Replace(celda.Value, reemplazos(i)(0), reemplazos(i)(1))
            End If
        Next i
    Next col
End Sub

' Formato numérico a dos decimales
Private Sub FormatColumnsAsDecimal(ws As Worksheet, cols() As Long)
    Dim colIndex As Variant, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For Each colIndex In cols
        ws.range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)).NumberFormat = "0.00"
    Next colIndex
End Sub

' Agrupa columnas por variación y etiquetas sin solapamiento
Private Sub AgruparColumnasPorVariacion(ws As Worksheet, columnas() As Long, ByRef grupo1() As Long, ByRef grupo2() As Long)
    On Error GoTo ManejoErrores
    Dim i As Long, r As Long, lastRow As Long
    Dim minVal As Double, maxVal As Double
    Dim variaciones() As Double
    Dim etiquetas() As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ReDim variaciones(LBound(columnas) To UBound(columnas))
    ReDim etiquetas(LBound(columnas) To UBound(columnas))
    
    ' Calcular variaciones y etiquetas
    For i = LBound(columnas) To UBound(columnas)
        minVal = CDbl(ws.Cells(2, columnas(i)).Value)
        maxVal = minVal
        For r = 3 To lastRow
            If IsNumeric(ws.Cells(r, columnas(i)).Value) Then
                Dim valor As Double
                valor = Round(CDbl(ws.Cells(r, columnas(i)).Value), 8)
                If valor < Round(minVal, 8) Then minVal = valor
                If valor > Round(maxVal, 8) Then maxVal = valor
            End If
        Next r
        variaciones(i) = Round(maxVal - minVal, 8)
        etiquetas(i) = ExtraerTextoEnParentesis(ws.Cells(1, columnas(i)).Value)
    Next i
    
    ' Agrupar por etiquetas idénticas
    Dim temp1() As Long, temp2() As Long
    Dim e1 As String
    Dim g1 As Long: g1 = -1
    Dim g2 As Long: g2 = -1
    
    e1 = etiquetas(LBound(columnas))
    For i = LBound(columnas) To UBound(columnas)
        If etiquetas(i) = e1 Then
            g1 = g1 + 1
            ReDim Preserve temp1(0 To g1)
            temp1(g1) = columnas(i)
        Else
            g2 = g2 + 1
            ReDim Preserve temp2(0 To g2)
            temp2(g2) = columnas(i)
        End If
    Next i
    
    If g1 >= 0 Then
        ReDim grupo1(0 To g1)
        For i = 0 To g1: grupo1(i) = temp1(i): Next i
    End If
    
    If g2 >= 0 Then
        ReDim grupo2(0 To g2)
        For i = 0 To g2: grupo2(i) = temp2(i): Next i
    End If
    Exit Sub
ManejoErrores:
    MsgBox "Error en AgruparColumnasPorVariacion: " & Err.Description, vbCritical
End Sub

' Añade series al gráfico desde un grupo
Private Sub AddGroupSeriesToChart(chartObj As Chart, ws As Worksheet, xCol As Long, groupCols() As Long, useSecondaryAxis As Boolean, palette() As Long)
    On Error GoTo ManejoErrores
    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, xCol).End(xlUp).Row
    
    For i = LBound(groupCols) To UBound(groupCols)
        With chartObj.SeriesCollection.NewSeries
            .XValues = ws.range(ws.Cells(2, xCol), ws.Cells(lastRow, xCol))
            .values = ws.range(ws.Cells(2, groupCols(i)), ws.Cells(lastRow, groupCols(i)))
            .Name = ws.Cells(1, groupCols(i)).Value
            .ChartType = xlLineMarkers
            .AxisGroup = IIf(useSecondaryAxis, xlSecondary, xlPrimary)
            .Format.Line.ForeColor.RGB = palette((iSeriesNr Mod UBound(palette)) + 1)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerForegroundColor = palette((iSeriesNr Mod UBound(palette)) + 1)
            .MarkerBackgroundColor = palette((iSeriesNr Mod UBound(palette)) + 1)
            .MarkerSize = 6
            .HasDataLabels = True
            .DataLabels.NumberFormat = "#.##"
            .DataLabels.Position = IIf(useSecondaryAxis, xlLabelPositionAbove, xlLabelPositionBelow)
        End With
        iSeriesNr = iSeriesNr + 1
    Next i
    Exit Sub
ManejoErrores:
    MsgBox "Error en AddGroupSeriesToChart: " & Err.Description, vbCritical
End Sub

' Ajusta el eje vertical basándose en los valores reales de las series
Private Sub AjustarEjeDesdeDatos(axis As axis, ws As Worksheet, cols() As Long)
    On Error GoTo ManejoErrores
    Dim minVal As Double: minVal = WorksheetFunction.Max(ws.Cells.Rows.Count, 1)
    Dim maxVal As Double: maxVal = WorksheetFunction.Min(ws.Cells.Rows.Count, 1)
    Dim i As Long, r As Long, Value As Variant
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cols(0)).End(xlUp).Row
    
    For i = LBound(cols) To UBound(cols)
        For r = 2 To lastRow
            Value = ws.Cells(r, cols(i)).Value
            If IsNumeric(Value) Then
                If Value < minVal Then minVal = Value
                If Value > maxVal Then maxVal = Value
            End If
        Next r
    Next i
    
    ' Ajustes según múltiplos de 5 y 10
    Dim minScale As Double, maxScale As Double
    minScale = RoundDown(minVal, 5)
    maxScale = RoundUp(maxVal, 5)
    
    If minScale > minVal Then minScale = minScale - 5
    If maxScale < maxVal Then maxScale = maxScale + 5
    
    If minScale = 5 Then minScale = 0
    
    ' Alineación con múltiplos de 10 si uno lo es
    If (minScale Mod 10 = 0) Xor (maxScale Mod 10 = 0) Then
        If minScale Mod 10 = 0 Then maxScale = RoundUp(maxScale, 10)
        If maxScale Mod 10 = 0 Then minScale = RoundDown(minScale, 10)
    End If
    
    With axis
        .MinimumScale = minScale
        .MaximumScale = maxScale
        .MajorUnitIsAuto = True
        .MinorUnitIsAuto = True
    End With
    Exit Sub
ManejoErrores:
    MsgBox "Error en AjustarEjeDesdeDatos: " & Err.Description, vbCritical
End Sub

' ================================
' Genera título del eje vertical concatenando todos los textos entre paréntesis (sin repetir)
Private Function GenerarTituloEjeVertical(ws As Worksheet, cols() As Long) As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim texto As String, inicio As Long, fin As Long
    Dim parte As String
    
    For i = LBound(cols) To UBound(cols)
        texto = ws.Cells(1, cols(i)).Value
        inicio = InStr(texto, "(")
        fin = InStr(texto, ")")
        If inicio > 0 And fin > inicio Then
            parte = Mid(texto, inicio + 1, fin - inicio - 1)
            If Not dict.Exists(parte) Then dict.Add parte, 1
        End If
    Next i
    
    If dict.Count > 0 Then
        GenerarTituloEjeVertical = "(" & Join(dict.Keys, ", ") & ")"
    Else
        GenerarTituloEjeVertical = ""
    End If
End Function

' Colores contrastados
Private Function GetColorPalette() As Long()
    Dim palette(1 To 10) As Long
    palette(1) = RGB(0, 102, 204)
    palette(2) = RGB(255, 102, 0)
    palette(3) = RGB(0, 153, 0)
    palette(4) = RGB(204, 0, 102)
    palette(5) = RGB(102, 0, 204)
    palette(6) = RGB(255, 153, 0)
    palette(7) = RGB(0, 204, 204)
    palette(8) = RGB(153, 51, 102)
    palette(9) = RGB(102, 204, 0)
    palette(10) = RGB(204, 51, 0)
    GetColorPalette = palette
End Function

Private Function CapturarPropiedadesDeEje(ax As axis) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    dict("MinimumScale") = ax.MinimumScale
    dict("MaximumScale") = ax.MaximumScale
    dict("MajorUnit") = ax.MajorUnit
    dict("MinorUnit") = ax.MinorUnit
    dict("MajorUnitIsAuto") = ax.MajorUnitIsAuto
    dict("MinorUnitIsAuto") = ax.MinorUnitIsAuto
    dict("MinimumScaleIsAuto") = ax.MinimumScaleIsAuto
    dict("MaximumScaleIsAuto") = ax.MaximumScaleIsAuto
    dict("ReversePlotOrder") = ax.ReversePlotOrder
    dict("Crosses") = ax.Crosses
    dict("DisplayUnit") = ax.DisplayUnit
    dict("HasTitle") = ax.HasTitle
    If ax.HasTitle Then
        dict("TitleText") = ax.AxisTitle.Text
    End If
    On Error GoTo 0
    
    Set CapturarPropiedadesDeEje = dict
End Function

Private Sub AplicarPropiedadesAEje(ax As axis, dict As Object)
    If dict.Exists("MinimumScaleIsAuto") Then ax.MinimumScaleIsAuto = dict("MinimumScaleIsAuto")
    If dict.Exists("MaximumScaleIsAuto") Then ax.MaximumScaleIsAuto = dict("MaximumScaleIsAuto")
    If dict.Exists("MajorUnitIsAuto") Then ax.MajorUnitIsAuto = dict("MajorUnitIsAuto")
    If dict.Exists("MinorUnitIsAuto") Then ax.MinorUnitIsAuto = dict("MinorUnitIsAuto")
    
    If dict.Exists("MinimumScale") Then
        On Error Resume Next: ax.MinimumScale = dict("MinimumScale"): On Error GoTo 0
    End If
    If dict.Exists("MaximumScale") Then
        On Error Resume Next: ax.MaximumScale = dict("MaximumScale"): On Error GoTo 0
    End If
    If dict.Exists("MajorUnit") Then
        On Error Resume Next: ax.MajorUnit = dict("MajorUnit"): On Error GoTo 0
    End If
    If dict.Exists("MinorUnit") Then
        On Error Resume Next: ax.MinorUnit = dict("MinorUnit"): On Error GoTo 0
    End If
    
    If dict.Exists("ReversePlotOrder") Then ax.ReversePlotOrder = dict("ReversePlotOrder")
    If dict.Exists("Crosses") Then ax.Crosses = dict("Crosses")
    If dict.Exists("DisplayUnit") Then ax.DisplayUnit = dict("DisplayUnit")
    
    If dict.Exists("HasTitle") Then
        ax.HasTitle = dict("HasTitle")
        If dict("HasTitle") And dict.Exists("TitleText") Then
            ax.AxisTitle.Text = dict("TitleText")
        End If
    End If
End Sub

' Texto dentro de paréntesis del encabezado
Private Function ExtraerTextoEnParentesis(texto As String) As String
    Dim re As Object, matches As Object, Match As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\((.*?)\)"
    re.Global = True
    Set matches = re.Execute(texto)
    
    Dim resultado As String
    For Each Match In matches
        If InStr(resultado, Match.SubMatches(0)) = 0 Then
            resultado = resultado & Match.SubMatches(0) & ", "
        End If
    Next Match
    
    If Len(resultado) > 2 Then
        resultado = Left(resultado, Len(resultado) - 2)
    End If
    
    ExtraerTextoEnParentesis = resultado
End Function

' Concatena textos entre paréntesis en primera fila para un grupo de columnas
Private Function ConcatenarTextosEntreParentesis(ws As Worksheet, cols() As Long) As String
    Dim i As Long
    Dim texto As String
    Dim dictUniq As Object
    Set dictUniq = CreateObject("Scripting.Dictionary")
    
    For i = LBound(cols) To UBound(cols)
        texto = Trim(ws.Cells(1, cols(i)).Value)
        If Len(texto) > 0 Then
            texto = ExtraerTextoEnParentesis(texto)
            If Len(texto) > 0 Then
                Dim partes() As String, j As Long
                partes = Split(texto, ",")
                For j = LBound(partes) To UBound(partes)
                    Dim parte As String
                    parte = Trim(partes(j))
                    If Len(parte) > 0 And Not dictUniq.Exists(parte) Then
                        dictUniq.Add parte, True
                    End If
                Next j
            End If
        End If
    Next i
    
    ConcatenarTextosEntreParentesis = Join(dictUniq.Keys, ", ")
End Function

' Redondeos a múltiplos de ciónco
Private Function RoundDown(Value As Double, base As Long) As Double
    RoundDown = base * Int(Value / base)
End Function

Private Function RoundUp(Value As Double, base As Long) As Double
    RoundUp = base * -Int(-Value / base)
End Function
