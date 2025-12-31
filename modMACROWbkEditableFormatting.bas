Attribute VB_Name = "modMACROWbkEditableFormatting"
Option Explicit

' Declaraciones para Win32 API (64 bits)
Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal Wt As Long, ByVal IsIt As Long, ByVal Und As Long, ByVal Str As Long, ByVal CharSet As Long, ByVal OutPr As Long, ByVal ClipPr As Long, ByVal Qual As Long, ByVal PitchAndFam As Long, ByVal Facename As String) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hgdiobj As LongPtr) As LongPtr
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
Private Type SIZE: cx As Long: cy As Long: End Type

Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" _
    (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, _
     lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long

' Constantes para DrawText
Private Const DT_CALCRECT As Long = &H400
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_EDITCONTROL As Long = &H2000
Private Const OUT_TT_PRECIS As Long = &H4
Private Const OUT_DEVICE_PRECIS As Long = &H5
Private Const OUT_RASTER_PRECIS As Long = &H6
Private Const OUT_TT_ONLY_PRECIS As Long = &H7

' === TIPO PARA GUARDAR ESTADO DE COMPACTACIÓN ===
Public Type CompactacionEstado
    PrintAreaOriginal As String
    FilasEstado() As Boolean    ' True = estaba visible; False = estaba oculta
    ColumnasEstado() As Boolean
    MinFila As Long
    MaxFila As Long
    minCol As Long
    maxCol As Long
    EsValido As Boolean         ' Flag de seguridad: evitar restaurar basura
End Type

Sub AjustarAltoFilasSegunColor()
    Dim rng As Range
    Dim celda As Range, fila As Range
    Dim colorFondo As Long
    Dim colorBlanco As Long
    Dim alturaOriginal As Double
    Dim nuevaAltura As Double
    
    ' Definir el color blanco (RGB 255, 255, 255)
    colorBlanco = RGB(255, 255, 255)
    
    ' Verificar que hay celdas seleccionadas
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, seleccione algunas celdas o filas primero.", vbExclamation
        Exit Sub
    End If
    
    ' Establecer el rango de trabajo como las celdas seleccionadas
    Set rng = Selection
    
    ' Deshabilitar actualización de pantalla para mejor rendimiento
    'Application.ScreenUpdating = False
    
    ' Recorrer cada celda en el rango seleccionado
    For Each fila In rng.Rows
        Set celda = fila.Cells(1, 1)
        ' Obtener el color de fondo de la celda
        colorFondo = celda.Interior.Color
        
        ' Obtener la altura actual de la fila
        alturaOriginal = celda.RowHeight
        
        ' Calcular nueva altura según el color de fondo
        If colorFondo = colorBlanco Then
            nuevaAltura = alturaOriginal + 6     ' * 1.05
        Else
            nuevaAltura = alturaOriginal + 12    '* 1.1
        End If
        
        ' Aplicar la nueva altura a toda la fila
        fila.RowHeight = nuevaAltura
        fila.VerticalAlignment = xlCenter
        If colorFondo = colorBlanco And InStr(ActiveWorkbook.Name, "API") > 0 Then
            fila.Cells(1, 5).HorizontalAlignment = xlJustify
            fila.Cells(1, 8).HorizontalAlignment = xlJustify
        End If
    Next fila
    
    ' Habilitar actualización de pantalla
    Application.ScreenUpdating = True
    
    'MsgBox "Ajuste de altura de filas completado.", vbInformation
End Sub

Private Function HideQTRow(r As Range) As Double
    If r.Hidden Then HideQTRow = True: Exit Function
    Select Case r.Parent.Name
        Case "GEN_FEATS_STDS"
        Case "1._SCOPE_OF_SUPPLY"
        Case "2._SALES_TERMS"
        Case "3._TECHNICAL_DESCRIPTION"
            HideQTRow = Left(r.Cells(1, 1), 3) = "TC." And Right(r.Cells(1, 2), 6) = " STAGE" And r.Cells(1, 4) = "-"
    End Select
    If HideQTRow Then r.Hidden = True
End Function

' Función auxiliar para detectar Marcadores
Private Function EsMarcador(c As Range) As Boolean
    ' Título Calibri 14 o Encabezado Verde
    Dim esTitulo As Boolean
    esTitulo = (c.Font.Name = "Calibri" And c.Font.SIZE = 14)
    Dim esVerde As Boolean
    Select Case c.Parent.Name
        Case "GEN_FEATS_STDS", "1._SCOPE_OF_SUPPLY", "3._TECHNICAL_DESCRIPTION"
                esVerde = (c.Interior.Color = RGB(141, 184, 34))
        Case "2._SALES_TERMS"
            If c.Row > 1 Then
                esVerde = (c.Offset(-1, 0).Interior.Color = RGB(141, 184, 34))
                EsMarcador = InStr(c.Offset(-1, 0).Value, "Quotation valid for ") > 0
            End If
        Case Else
    End Select
    EsMarcador = EsMarcador Or esTitulo Or esVerde
End Function

Private Function ObtenerIncremento(c As Range) As Double
    Select Case c.Parent.Name
        Case "3._TECHNICAL_DESCRIPTION", "1._SCOPE_OF_SUPPLY"
            If Trim(c.EntireRow.Cells(1, 1).Value) <> "" And _
                    c.Font.Name = "Calibri" And c.Font.SIZE = 10 Then ObtenerIncremento = 12
    End Select
End Function

Private Function ObtenerFactorCorreccion(c As Range) As Double
    Dim factor As Double
    factor = 1#  ' Factor base por defecto (5% de margen)

    ' Corrección por Tipo de Fuente
    ' Calibri suele necesitar un poco más de espacio que Arial en PDF
    Select Case True
        Case c.Font.Name = "Calibri" And c.Font.SIZE = 10:      factor = factor * 1
        Case c.Font.Name = "Calibri" And c.Font.SIZE = 11:      factor = factor * 1.02
        Case c.Font.Name = "Calibri" And c.Font.SIZE = 12:      factor = factor * 1.03
        Case c.Font.Name = "Calibri" And c.Font.SIZE = 14:      factor = factor * 1
        Case c.Font.Name = "Calibri" And c.Font.SIZE = 16:      factor = factor * 1.04
        Case c.Font.Name = "Arial" And c.Font.SIZE = 11:        factor = factor * 1.01
        Case c.Font.Name = "Times New Roman" And c.Font.SIZE = 11: factor = factor * 1.03
        Case Else
            Debug.Print "Fuente: " & c.Font.Name & " - " & c.Font.SIZE
            Stop
    End Select

    ' Corrección por Color de Fondo (Interior.Color)
    ' A veces, celdas con fondos oscuros o rellenos sólidos requieren
    ' un margen extra para que el texto no parezca "pegado" al borde
    If c.Interior.ColorIndex <> xlNone Then
        Select Case c.Interior.Color
            Case RGB(0, 0, 0):   factor = factor * 1.1 ' Negro
            Case RGB(89, 89, 89): factor = factor * 1.02 ' gris oscuro quotation
            Case RGB(141, 184, 34): factor = factor * 1.02 ' verde abc
            Case RGB(255, 0, 0): factor = factor * 1.05 ' Rojo
            Case RGB(255, 255, 255): factor = factor * 1 ' blanco
            Case RGB(217, 217, 217), RGB(191, 191, 191): factor = factor * 1 ' gris claro quotation
            Case Else
                Debug.Print "Color de fondo: " & LongToRGB(c.Interior.Color)
                Stop 'factor = factor * 1.02 ' Otros colores
        End Select
    End If

    ' Corrección por Estilo
    If c.Font.Bold Then factor = factor * 1.01

    ObtenerFactorCorreccion = factor
End Function
Sub AjustarSelWSheetsParaImpresionPDF()
    On Error GoTo Cleanup
    
    Dim selSheets As Object
    Set selSheets = ActiveWindow.SelectedSheets
    
    If selSheets.Count = 0 Then
        MsgBox "[WARN] No hay hojas seleccionadas en el libro '" & ActiveWindow.Parent.Name & "'.", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "=== AjustarSelWSheetsParaImpresionPDF: " & selSheets.Count & " hoja(s) en '" & ActiveWindow.Parent.Name & "' ==="
    
    Dim sh As Object
    Dim hojasProcesadas As Long, hojasFallidas As Long
    Dim estados() As CompactacionEstado
    ReDim estados(1 To selSheets.Count)
    
    Dim i As Long: i = 0
    
    For Each sh In selSheets
        i = i + 1
        If TypeName(sh) = "Worksheet" Then
            On Error Resume Next
            estados(i) = AjustarWSParaPDF_ImpresionMaestra(sh)
            If Err.Number = 0 Then
                hojasProcesadas = hojasProcesadas + 1
            Else
                Debug.Print "[ERR] Falló en '" & sh.Name & "': " & Err.Description
                hojasFallidas = hojasFallidas + 1
            End If
            On Error GoTo 0
        Else
            Debug.Print "[INFO] Saltada hoja no Worksheet: '" & sh.Name & "' (tipo: " & TypeName(sh) & ")"
        End If
    Next sh
    
    Debug.Print "[OK] Proceso finalizado: " & hojasProcesadas & " hojas ajustadas, " & hojasFallidas & " con error."
    
    ' ? Problema: `AjustarWS...` compacta, pero no devuelve el estado.
    ' Solución real: modificar `AjustarWS...` para que acepte ByRef estado (mejor diseño).
    
    ' --------------------------
    ' ? SOLUCIÓN CORRECTA (recomendada):
    ' Modificamos ligeramente `AjustarWS...` para que reciba ByRef estado.
    ' Así evitamos doble compactación.
    ' --------------------------
    
    If 6 = MsgBox("Proceso completado." & vbCrLf & _
           hojasProcesadas & " hoja(s) ajustada(s)." & vbCrLf & _
           "¿Deseas exportar el documento a PDF y restaurar Las áreas de impresión Del mismo (SI), " & _
           "o quieres conservarlo reemplazando las áreas no imprimibles por filas y columnas ocultas (NO)?", _
           vbQuestion + vbYesNo, "Ajuste listo") Then
        ' grabar el documento como PDF
        ' Exportar la hoja "Graficos" a PDF
        Dim wb As Workbook
        Set wb = ActiveWindow.Parent
        If wb Is Nothing Then
            MsgBox "[ERR] No hay libro activo.", vbCritical
            Exit Sub
        End If
        On Error Resume Next
        wb.ExportAsFixedFormat Type:=xlTypePDF, _
                fileName:=wb.Path & "\" & Left(wb.Name, InStrRev(wb.Name, ".") - 1) & ".pdf", _
                Quality:=xlQualityStandard, _
                OpenAfterPublish:=True
        On Error GoTo 0
        
        ' a continuacion se restauran las areas no imprimibles
    Else
        Exit Sub
    End If
    

Cleanup:
    ' RESTAURACION DE LAS AREAS NO IMPRIMIBLES, tras generar documentos
    Debug.Print "Restauración de las áreas de impresión de cada hoja del libro:"
    hojasProcesadas = 0
    hojasFallidas = 0
    i = 0
    For Each sh In selSheets
        i = i + 1
        If TypeName(sh) = "Worksheet" Then
            On Error Resume Next
            RestaurarAreaDeImpresion sh, estados(i)
            If Err.Number = 0 Then
                hojasProcesadas = hojasProcesadas + 1
            Else
                Debug.Print "[ERR] Falló en '" & sh.Name & "': " & Err.Description
                hojasFallidas = hojasFallidas + 1
            End If
            On Error GoTo 0
        Else
            Debug.Print "[INFO] Saltada hoja no Worksheet: '" & sh.Name & "' (tipo: " & TypeName(sh) & ")"
        End If
    Next sh
    
    Debug.Print "[OK] Proceso finalizado: " & hojasProcesadas & " hojas ajustadas, " & hojasFallidas & " con error."
    
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, "AjustarSelWSheetsParaImpresionPDF", Err.Description
    End If
End Sub
Function AjustarWSParaPDF_ImpresionMaestra(Optional ws As Worksheet = Nothing) As CompactacionEstado
    If ws Is Nothing Then Set ws = ActiveSheet
    
    Dim rangoImpresion As Range
    Dim ultimaFila As Long, ultimaColumna As Long
    
    Dim prevScreenUpdating As Boolean, prevCalcMode As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalcMode = Application.Calculation
    
    'Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual ' Acelerar cálculos

    ' --- Tarea 1: Determinar el estado del area de impresión: Rango Máximo Ocupado ---
    On Error Resume Next
    Set rangoImpresion = GetAbsoluteUsedRange(ws)
    On Error GoTo 0
    ultimaFila = rangoImpresion.Rows(rangoImpresion.Rows.Count).Row
    ultimaColumna = rangoImpresion.Columns(rangoImpresion.Columns.Count).Column
'    Set rangoImpresion = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaColumna))
'    ws.PageSetup.PrintArea = rangoImpresion.Address

    ' --- Tarea 2: Compactar area de impresión, convirtiendo en ocultas las lineas ---
    Dim EstadoCompactacionPrintAreaAOcultas As CompactacionEstado
    Call CompactarAreaDeImpresion(ws, EstadoCompactacionPrintAreaAOcultas)
    AjustarWSParaPDF_ImpresionMaestra = EstadoCompactacionPrintAreaAOcultas

    ' --- 3. Ajustes adicionales (tolerantes a errores) ---
    On Error Resume Next
    If SaltoVerticalCortaRango(ws, rangoImpresion) Then
        ' eliminar salto vertical
        If ws.VPageBreaks.Count > 0 Then ws.VPageBreaks(1).Delete
    End If
    
    ' ajusta los rangos
    ReajustarRangoAreasImpresion ws
    On Error GoTo 0
    
    ' --- Tarea 2: Ajustar Alturas de Fila y Configuración de Página ---
    On Error Resume Next
    AjustarAlturaFilasSegunTexto ws, ultimaFila, ultimaColumna
    On Error GoTo 0
    
    ' b) Configurar para que ocupe 1 página de ancho y el largo necesario
    With ws.PageSetup
        .FitToPagesWide = 1
        .Zoom = False
        .FitToPagesTall = False ' Automático
        .CenterHorizontally = True
    End With
    
    ' --- Tarea 3: Ajustes Manuales de Saltos de Página ---
    ConfigurarSaltosDePagina ws, ultimaFila
    
    ' Restaurar configuración y rendimiento
    'ws.PageSetup.PrintQuality = 600 ' Restaurar calidad alta para la impresión final
    
Cleanup:
    Application.Calculation = prevCalcMode
    Application.ScreenUpdating = prevScreenUpdating
    
    If Err.Number <> 0 Then
        Debug.Print "[ERR] AjustarWSParaPDF_ImpresionMaestra('" & ws.Name & "'): " & Err.Description & " (Err " & Err.Number & ")"
        Err.Raise Err.Number, "AjustarWSParaPDF_ImpresionMaestra", Err.Description
    End If
End Function

Private Sub AjustarAlturaFilasSegunTexto(ws As Worksheet, ultimaFila As Long, ultimaColumna As Long)
    Dim i As Long
    Dim altoMaximoFila As Double
    ' Forzar DPI de diseño para la medición
    'ws.PageSetup.PrintQuality = 72
    
    ' a) Recorrer y ajustar alturas (usa las funciones previas GetPrintHeightMultiLine y ObtenerFactorCorreccion)
    For i = 1 To ultimaFila
        If IsEmptyRow(ws.Rows(i)) Then GoTo NextRow
        altoMaximoFila = 0
        Dim celda As Range
        For Each celda In ws.Range(ws.Cells(i, 1), ws.Cells(i, ultimaColumna))
            If Trim(celda.Value) <> "" Then
                Dim altoActual As Double
                altoActual = GetPrintHeightMultiLine(celda) * ObtenerFactorCorreccion(celda) + ObtenerIncremento(celda)
                If altoActual > altoMaximoFila Then altoMaximoFila = altoActual
            End If
        Next celda
        
        If altoMaximoFila > 0 And Not HideQTRow(ws.Rows(i)) Then
            ' Aplicar alto con margen de seguridad
            ws.Rows(i).RowHeight = Round(altoMaximoFila, 1)
        End If
NextRow:
    Next i
End Sub

Private Function GetCleanPrinterName() As String
    Dim fullPrinter As String
    fullPrinter = Application.ActivePrinter
    ' Busca la posición de " en " (formato español) o " on " (inglés)
    Dim pos As Integer
    pos = InStr(1, fullPrinter, " en ", vbTextCompare) ' Español
    If pos = 0 Then pos = InStr(1, fullPrinter, " on ", vbTextCompare) ' Inglés
    
    If pos > 0 Then
        GetCleanPrinterName = Left(fullPrinter, pos - 1)
    Else
        GetCleanPrinterName = fullPrinter
    End If
End Function

Sub ReajustarRangoAreasImpresion(ws As Worksheet)
    Dim areasOriginales As Variant
    Dim areaActual As Range, nuevoPrintRange As Range
    Dim i As Integer
    Dim ultFila As Long, ultCol As Long
    
    ' 1. Verificar si hay un área de impresión definida
    If ws.PageSetup.PrintArea = "" Then Exit Sub
    
    ' 2. Dividir el área de impresión por si contiene múltiples rangos (separados por comas)
    areasOriginales = Split(ws.PageSetup.PrintArea, ";")
    
    For i = LBound(areasOriginales) To UBound(areasOriginales)
        Set areaActual = ws.Range(areasOriginales(i))
        
        ' 3. Encontrar la última celda real CON VALORES dentro de este rango específico
        On Error Resume Next
        ultFila = areaActual.Find(What:="*", _
                                  After:=areaActual.Cells(1, 1), _
                                  LookIn:=xlValues, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious).Row
                                  
        ultCol = areaActual.Find(What:="*", _
                                 After:=areaActual.Cells(1, 1), _
                                 LookIn:=xlValues, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlPrevious).Column
        With areaActual.Cells(areaActual.Rows.Count, areaActual.Columns.Count).CurrentRegion
            ultFila = .Row + Rows.Count - 1
            ultCol = .Column + Columns.Count - 1
        End With
        On Error GoTo 0
        
        ' 4. Redefinir el área si se encontró contenido
        If ultFila > 0 And ultCol > 0 Then
            ' El nuevo rango empieza donde el original, pero acaba en ultFila/ultCol
            ' asegurándonos de no exceder los límites del rango original
            Dim filaFin As Long: filaFin = IIf(ultFila < areaActual.Row, areaActual.Row, ultFila)
            Dim colFin As Long: colFin = IIf(ultCol < areaActual.Column, areaActual.Column, ultCol)
            
            Dim rangoRecortado As Range
            Set rangoRecortado = ws.Range(ws.Cells(areaActual.Row, areaActual.Column), _
                                      ws.Cells(filaFin, colFin))
            
            ' 5. Concatenar para la nueva PrintArea
            If nuevoPrintRange Is Nothing Then
                Set nuevoPrintRange = rangoRecortado
            Else
                Set nuevoPrintRange = Union(nuevoPrintRange, rangoRecortado)
            End If
        End If
        
        ' Reset para la siguiente área
        ultFila = 0: ultCol = 0
    Next i
    
    ' 6. Asignar el área de impresión final limpia
    ws.PageSetup.PrintArea = nuevoPrintRange.Address
End Sub

Function GetPrintHeightMultiLine(ByVal c As Range) As Double
Attribute GetPrintHeightMultiLine.VB_Description = "[modMACROWbkEditableFormatting] Get Print Height Multi Line (función personalizada). Aplica a: Cells Range"
Attribute GetPrintHeightMultiLine.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim hdc As LongPtr, hFont As LongPtr, hOldFont As LongPtr
    Dim r As RECT
    Dim sz As SIZE
    Dim dpiX As Long, dpiY As Long
    Dim colWidthPts As Double, colWidthPx As Long
    Dim fontSizePx As Long
    Dim printerName As String
    Dim angulo As Long, bVert As Boolean
    
    If IsError(c.Value) Then Stop: Exit Function
    
    ' Obtener nombre limpio para la API
    printerName = GetCleanPrinterName()
    
    ' Crear el contexto de dispositivo (HDC)
    ' "WINSPOOL" es el driver estándar para impresoras en Win32
    hdc = CreateDC("WINSPOOL", printerName, vbNullString, 0)
    If hdc = 0 Then hdc = CreateDC("DISPLAY", vbNullString, vbNullString, 0)
    
    dpiX = GetDeviceCaps(hdc, 88) ' LOGPIXELSX
    dpiY = GetDeviceCaps(hdc, 90) ' LOGPIXELSY
    
    angulo = c.Orientation * 10
    bVert = (c.Orientation = xlVertical Or Abs(c.Orientation) = 90)
    ' 1. Detectar si es celda combinada y obtener el área total
    If c.MergeCells Then
        Set c = c.MergeArea
    Else
        ' queda como está
    End If
    
    ' 2. Calcular ancho de columna en píxeles de impresión
    ' Excel c.Width está en puntos. Convertimos a pulgadas y luego a píxeles del dispositivo.
    ' Restamos un pequeño margen para los bordes internos de la celda (aprox 5 pts)
    colWidthPts = c.Width - 4
    colWidthPx = (colWidthPts * dpiX) / 72
    
    ' 3. Configurar Fuente
    fontSizePx = -(c.Cells(1, 1).Font.SIZE * dpiY) / 72
    ' CreateFont permite especificar el ángulo en el 3er y 4º parámetro (Escapement y Orientation)
    hFont = CreateFont(fontSizePx, 0, angulo, angulo, IIf(c.Cells(1, 1).Font.Bold, 700, 400), _
                       IIf(c.Cells(1, 1).Font.Italic, 1, 0), 0, 0, 0, OUT_TT_ONLY_PRECIS, 0, 0, 0, c.Cells(1, 1).Font.Name)
    hOldFont = SelectObject(hdc, hFont)
    
    If c.WrapText = False And Abs(angulo) <> 900 Then
        GetTextExtentPoint32 hdc, c.Cells(1, 1).Value, Len(c.Cells(1, 1).Value), sz
        GetPrintHeightMultiLine = (sz.cy * 72) / dpiY ' Altura de una sola línea
    ElseIf bVert Then
        ' Si es puramente vertical, medimos la extensión de una sola línea pero
        ' el valor de 'ancho' de esa línea será el 'alto' de nuestra fila.
        GetTextExtentPoint32 hdc, c.Cells(1, 1).Value, Len(c.Cells(1, 1).Value), sz
        ' En vertical, el ancho del texto (sz.cx) se convierte en la altura de la fila
        GetPrintHeightMultiLine = (sz.cx * 72) / dpiY
    Else
        ' 4.2. Definir rectángulo de cálculo (ancho fijo, alto variable)
        r.Left = 0
        r.Top = 0
        r.Right = colWidthPx
        r.Bottom = 0
        
        ' 5. Ejecutar DrawText con DT_CALCRECT para que calcule el alto (r.Bottom)
        ' DT_WORDBREAK permite el ajuste de línea como en Excel
        DrawText hdc, c.Cells(1, 1).Value, Len(c.Cells(1, 1).Value), r, DT_CALCRECT Or DT_WORDBREAK Or DT_EDITCONTROL
        
        ' 6. Convertir el alto calculado de píxeles a Puntos de Excel
        GetPrintHeightMultiLine = (r.Bottom * 72) / dpiY
    End If
    
    ' Limpieza
    SelectObject hdc, hOldFont
    DeleteDC hdc
End Function

Function GetAbsoluteUsedRange(ws As Worksheet) As Range
Attribute GetAbsoluteUsedRange.VB_Description = "[modMACROWbkEditableFormatting] Get Absolute Used Range (función personalizada). Aplica a: Cells Range"
Attribute GetAbsoluteUsedRange.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ultFila As Long
    Dim ultCol As Long
    
    ' Asegurar que haya área de impresión definida
    If ws.PageSetup.PrintArea = "" Then
        Debug.Print "[WARN] No hay área de impresión definida en la hoja '" & ws.Name & "'."
    End If

    ' 1. Resetear temporalmente el área de impresión para que .Find vea todo
    Dim areaActual As Range, area
    For Each area In Split(ws.PageSetup.PrintArea, ";")
        If areaActual Is Nothing Then
            Set areaActual = ws.Range(Trim(area))
        Else
            Set areaActual = Union(areaActual, ws.Range(Trim(area)))
        End If
    Next

    ws.PageSetup.PrintArea = ""

    ' 2. Buscar la última columna con datos reales en TODA la hoja
    ' Al usar LookIn:=xlFormulas incluimos celdas con valores y con fórmulas (aunque den "")
    On Error Resume Next
    ultCol = ws.Cells.Find(What:="*", _
                           After:=ws.Cells(1, 1), _
                           LookIn:=xlFormulas, _
                           SearchOrder:=xlByColumns, _
                           SearchDirection:=xlPrevious).Column
    
    ' 3. Buscar la última fila con datos reales
    ultFila = ws.Cells.Find(What:="*", _
                            After:=ws.Cells(1, 1), _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious).Row
    On Error GoTo 0

    ' 4. Restaurar el área previa si fuera necesario o dejar que el macro la actualice
    ws.PageSetup.PrintArea = areaActual.Address

    ' 5. Devolver el rango desde A1 hasta el límite absoluto detectado
    If ultFila = 0 Then ultFila = 1
    If ultCol = 0 Then ultCol = 1
    
    Set GetAbsoluteUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(ultFila, ultCol))
End Function

Function GetRealUsedRange(ws As Worksheet) As Range
Attribute GetRealUsedRange.VB_Description = "[modMACROWbkEditableFormatting] Get Real Used Range (función personalizada). Aplica a: Cells Range"
Attribute GetRealUsedRange.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim ultimaFila As Long
    Dim ultimaCol As Long
    
    ' Buscar la última fila con datos (independientemente de la columna)
    On Error Resume Next
    ultimaFila = ws.Cells.Find(What:="*", _
                               After:=ws.Cells(1, 1), _
                               LookAt:=xlPart, _
                               LookIn:=xlValues, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlPrevious).Row
                               
    ' Buscar la última columna con datos (independientemente de la fila)
    ultimaCol = ws.Cells.Find(What:="*", _
                              After:=ws.Cells(1, 1), _
                              LookAt:=xlPart, _
                              LookIn:=xlValues, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    
    ' Si la hoja está totalmente vacía, devolvemos A1 para evitar errores
    If ultimaFila = 0 Then ultimaFila = 1
    If ultimaCol = 0 Then ultimaCol = 1
    
    ' Retornar el rango desde A1 hasta el límite real encontrado
    Set GetRealUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaCol))
End Function

Function GetPhysicalPaperHeight(ws As Worksheet) As Double
Attribute GetPhysicalPaperHeight.VB_Description = "[modMACROWbkEditableFormatting] Get Physical Paper Height (función personalizada)"
Attribute GetPhysicalPaperHeight.VB_ProcData.VB_Invoke_Func = " \n23"
    ' Altura física total del papel en puntos (1 pulgada = 72 pts)
    Select Case ws.PageSetup.PaperSize
        Case xlPaperA4:      GetPhysicalPaperHeight = 841.89 ' (210mm x 297mm)
        Case xlPaperLetter:  GetPhysicalPaperHeight = 792     ' (8.5" x 11")
        Case xlPaperLegal:   GetPhysicalPaperHeight = 1008    ' (8.5" x 14")
        Case Else:           GetPhysicalPaperHeight = 842     ' A4 por defecto
    End Select
End Function

Function GetUsablePrintHeight(ws As Worksheet) As Double
Attribute GetUsablePrintHeight.VB_Description = "[modMACROWbkEditableFormatting] Get Usable Print Height (función personalizada)"
Attribute GetUsablePrintHeight.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim totalHeight As Double
    totalHeight = GetPhysicalPaperHeight(ws)
    
    With ws.PageSetup
        ' Restamos márgenes físicos y áreas de encabezado/pie
        GetUsablePrintHeight = totalHeight - (.TopMargin + .BottomMargin)
    End With
End Function

Sub ConfigurarSaltosDePagina(ws As Worksheet, ultimaFila As Long)
    Dim i As Long, j As Long
    Dim npb As Long
    Dim prevScreenUpdating As Boolean
    
    ' --- Estado original ---
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Cleanup
    
    ' --- 1. Asegurar que los saltos automáticos estén actualizados ---
    ws.ResetAllPageBreaks
    Application.Calculate  ' ? clave: fuerza que Excel genere HPageBreaks automáticos
    
    npb = ws.HPageBreaks.Count
    If npb = 0 Then
        Debug.Print "[WARN] ConfigurarSaltosDePagina: no hay saltos automáticos en '" & ws.Name & "'. ¿Rango imprimible vacío?"
        GoTo Cleanup
    End If
    
    Debug.Print "=== AJUSTE DE SALTOS (basado en " & npb & " saltos automáticos) ==="
    
    ' --- 2. Procesar cada página (de arriba a abajo) ---
    Dim pagina As Long
    For pagina = 1 To npb + 1  ' +1 para incluir la última página (después del último salto)
        Dim filaInicio As Long, filaFin As Long
        
        ' Determinar rango de la página actual
        If pagina = 1 Then
            filaInicio = 1
        Else
            filaInicio = ws.HPageBreaks(pagina - 1).Location.Row
        End If
        
        If pagina <= npb Then
            filaFin = ws.HPageBreaks(pagina).Location.Row - 1
        Else
            filaFin = ultimaFila
        End If
        
        ' Acotar al rango útil
        If filaInicio > ultimaFila Then Exit For
        If filaFin > ultimaFila Then filaFin = ultimaFila
        
        ' --- 3. Buscar el mejor marcador DENTRO de esta página ---
        Dim mejorDistancia As Double, mejorFilaMarcador As Long, mejorFactorF As Double
        mejorDistancia = 999999
        mejorFilaMarcador = 0
        mejorFactorF = 1#
        
        ' Explorar factores de compresión/estiramiento de filas vacías
        Dim factorF As Double
        For factorF = 0.8 To 1.2 Step 0.05
            Dim offsetAcumulado As Double
            offsetAcumulado = 0
            
            ' Recorrer filas *dentro de la página*, de abajo hacia arriba (más eficiente)
            For i = filaFin To filaInicio Step -1
                If ws.Rows(i).Hidden Then
                    ' ignorar
                ElseIf IsEmptyRow(ws.Rows(i)) Then
                    offsetAcumulado = offsetAcumulado + (ws.Rows(i).RowHeight * (1 - factorF))
                    ' (1 - f) > 0 ? compresión; < 0 ? estiramiento
                End If
                
                If EsMarcador(ws.Cells(i, 1)) Then
                    ' Distancia vertical desde el marcador al final de la página
                    ' (sin recalcular alturas: usamos el offset acumulado por filas vacías ajustadas)
                    Dim distancia As Double
                    distancia = Abs(offsetAcumulado)  ' cuanto más cerca de 0, mejor (marcador al final)
                    
                    If distancia < mejorDistancia Then
                        mejorDistancia = distancia
                        mejorFilaMarcador = i
                        mejorFactorF = factorF
                    End If
                End If
            Next i
        Next factorF
        
        ' --- 4. Aplicar ajuste si se encontró un marcador ---
        If mejorFilaMarcador > 0 And mejorDistancia < 15 Then  ' umbral: <15 pts ˜ 0.2 cm
            ' Ajustar filas vacías *solo en esta página*
            For j = filaInicio To mejorFilaMarcador - 1
                If IsEmptyRow(ws.Rows(j)) Then
                    ws.Rows(j).RowHeight = ws.Rows(j).RowHeight * mejorFactorF
                End If
            Next j
            
            ' Insertar salto *justo antes* del marcador (reemplazando el automático)
            If pagina <= npb Then
                ws.HPageBreaks(pagina).DragOff Direction:=xlToRight, RegionIndex:=1  ' eliminar salto automático
            End If
            ws.HPageBreaks.Add Before:=ws.Rows(mejorFilaMarcador)
            
            Debug.Print "[OK] Página " & pagina & ": salto ajustado a fila " & mejorFilaMarcador & _
                        " (f=" & Round(mejorFactorF, 2) & ", d=" & Round(mejorDistancia, 1) & " pts)"
        Else
            Debug.Print "[INFO] Página " & pagina & ": sin marcador óptimo (rango " & filaInicio & "-" & filaFin & ")"
        End If
    Next pagina
    
Cleanup:
    Application.ScreenUpdating = prevScreenUpdating
    If Err.Number <> 0 Then
        Debug.Print "[ERR] ConfigurarSaltosDePagina: " & Err.Description & " (Err " & Err.Number & ")"
    Else
        Debug.Print "[OK] ConfigurarSaltosDePagina: finalizado. Nuevos saltos: " & ws.HPageBreaks.Count
    End If
End Sub

' PENDIENTE DE PROBAR, MEJOR ESTA QUE LA OLD si la otra no funciona
Sub ConfigurarSaltosDePagina_v2(ws As Worksheet, ultimaFila As Long)
    Dim altoPaginaReal As Double
    Dim i As Long, j As Long, filaInicioPagina As Long
    Dim altoAcumulado As Double
    Dim factorF As Double
    Dim mejorF As Double, minimaDistancia As Double
    Dim filaMarcador As Long
    Dim npb As Long
    
    Dim prevScreenUpdating As Boolean
    
    ' --- Guardar estado ---
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    On Error GoTo Cleanup
    
    ' --- 1. Reiniciar y preparar ---
    ws.ResetAllPageBreaks
    npb = 0
    altoPaginaReal = GetUsablePrintHeight(ws)
    filaInicioPagina = 1
    
    Do While filaInicioPagina <= ultimaFila
        minimaDistancia = altoPaginaReal * 0.15  ' Límite de búsqueda (15% del alto)
        mejorF = 1#
        filaMarcador = 0
        
        ' --- 2. SIMULACIÓN: buscar marcador óptimo con ajuste de filas vacías ---
        ' Probamos variando f de 0.8 a 1.2 en pasos de 0.05
        For factorF = 0.8 To 1.2 Step 0.05
            altoAcumulado = 0
            For i = filaInicioPagina To ultimaFila
                ' Si es fila en blanco, aplicamos el factor de elasticidad
                If ws.Rows(i).Hidden Then
                    ' saltar las filas ocultas
                ElseIf IsEmptyRow(ws.Rows(i)) Then
                    altoAcumulado = altoAcumulado + (ws.Rows(i).RowHeight * factorF)
                Else
                    altoAcumulado = altoAcumulado + ws.Rows(i).RowHeight
                End If
                
                ' Si encontramos un marcador (Título/Encabezado) dentro del rango de interés
                If EsMarcador(ws.Cells(i, 1)) Then
                    Dim distancia As Double
                    distancia = Abs(altoPaginaReal - altoAcumulado)
                    
                    ' Si este marcador con este 'f' es el más cercano al borde inferior
                    If distancia < minimaDistancia Then
                        minimaDistancia = distancia
                        mejorF = factorF
                        filaMarcador = i
                    End If
                End If
                
                ' Si nos pasamos mucho del alto de página, dejamos de simular este 'f'
                If altoAcumulado > altoPaginaReal * 1.1 Then Exit For
            Next i
        Next factorF
        
        ' --- 2. APLICACIÓN DE LOS RESULTADOS ---
        If filaMarcador > 0 Then
            ' Retroceder hasta la primera fila no vacía/marcador
            Do While j > filaInicioPagina And (EsMarcador(ws.Cells(filaMarcador - 1, 1)) Or IsEmptyRow(ws.Rows(filaMarcador - 1)))
                filaMarcador = filaMarcador - 1
            Loop
            ' Aplicamos el factor 'mejorF' a las filas en blanco de esta página
            altoAcumulado = 0
            For j = filaInicioPagina To filaMarcador - 1
                If IsEmptyRow(ws.Rows(j)) Then
                    ws.Rows(j).RowHeight = ws.Rows(j).RowHeight * mejorF
                    Debug.Print "[INFO] Fila " & j & ": RowHeight ajustado a " & Round(ws.Rows(j).RowHeight, 1) & " (f=" & Round(mejorF, 2) & ")"
                End If
            Next j
            
            ' Insertar salto manual justo antes del marcador óptimo
            ws.HPageBreaks.Add Before:=ws.Rows(filaMarcador)
            npb = npb + 1
            filaInicioPagina = filaMarcador
            Debug.Print "[OK] Salto insertado antes de fila " & filaMarcador & " (posición de marcador)"
        Else
            ' Forzar cálculo de saltos automáticos (por si estaban desactivados)
            Application.Calculate  ' asegura que los saltos automáticos estén actualizados
            
            ' Contar cuántos saltos automáticos hay *desde filaInicioPagina*
            Dim autoBreakCount As Long, k As Long
            autoBreakCount = 0
            For k = 1 To ws.HPageBreaks.Count
                If ws.HPageBreaks(k).Location.Row > filaInicioPagina And ws.HPageBreaks(k).Location.Row <= ultimaFila Then
                    autoBreakCount = autoBreakCount + 1
                End If
            Next k
            
            If autoBreakCount > 0 Then
                ' Tomar el primer salto automático válido
                For k = 1 To ws.HPageBreaks.Count
                    If ws.HPageBreaks(k).Location.Row > filaInicioPagina Then
                        j = ws.HPageBreaks(k).Location.Row
                        ws.HPageBreaks.Add Before:=ws.Rows(j)
                        npb = npb + 1
                        filaInicioPagina = j
                        Debug.Print "[OK] Salto automático adoptado en fila " & j
                        Exit For
                    End If
                Next k
            Else
                ' Último recurso: avanzar una página "ciega" (poco probable)
                filaInicioPagina = filaInicioPagina + 40  ' ˜ una página estándar
                Debug.Print "[WARN] Sin saltos automáticos ni marcadores: avance forzado a fila " & filaInicioPagina
            End If
        End If
        
        ' Evitar bucle infinito si algo falla
        If filaInicioPagina > ultimaFila Or i > ultimaFila Then Exit Do
        ' Protección contra bucle infinito (solo por si acaso)
        If npb > 100 Then Exit Do
    Loop
    
    Debug.Print "[OK] ConfigurarSaltosDePagina_v2: " & npb & " saltos en '" & ws.Name & "'"
    
Cleanup:
    Application.ScreenUpdating = prevScreenUpdating
    If Err.Number <> 0 Then
        Debug.Print "[ERR] ConfigurarSaltosDePagina_v2: " & Err.Description & " (Err " & Err.Number & ")"
    End If
End Sub

Sub ConfigurarSaltosDePagina_old(ws As Worksheet, ultimaFila As Long)
    Dim altoPaginaReal As Double
    Dim i As Long, j As Long, filaInicioPagina As Long
    Dim altoAcumulado As Double
    Dim factorF As Double
    Dim mejorF As Double, minimaDistancia As Double
    Dim marcadorOptimo As Long, filaMarcador As Long, npb As Long
    
    'Application.ScreenUpdating = False
    ws.ResetAllPageBreaks
    altoPaginaReal = GetUsablePrintHeight(ws) ' Función definida anteriormente
    filaInicioPagina = 1
    
    Do While filaInicioPagina <= ultimaFila
        minimaDistancia = altoPaginaReal * 0.15  ' Límite de búsqueda (15% del alto)
        mejorF = 1#
        filaMarcador = 0
        
        ' --- 1. SIMULACIÓN DE OPTIMIZACIÓN ---
        ' Probamos variando f de 0.8 a 1.2 en pasos de 0.05
        For factorF = 0.8 To 1.2 Step 0.05
            altoAcumulado = 0
            For i = filaInicioPagina To ultimaFila
                Dim altoBase As Double
                altoBase = ws.Rows(i).RowHeight
                
                ' Si es fila en blanco, aplicamos el factor de elasticidad
                If ws.Rows(i).Hidden Then
                    ' saltar las filas ocultas
                ElseIf IsEmptyRow(ws.Rows(i)) Then
                    altoAcumulado = altoAcumulado + (altoBase * factorF)
                Else
                    altoAcumulado = altoAcumulado + altoBase
                End If
                
                ' Si encontramos un marcador (Título/Encabezado) dentro del rango de interés
                If EsMarcador(ws.Cells(i, 1)) Then
                    Dim distancia As Double
                    distancia = Abs(altoPaginaReal - altoAcumulado)
                    
                    ' Si este marcador con este 'f' es el más cercano al borde inferior
                    If distancia < minimaDistancia Then
                        minimaDistancia = distancia
                        mejorF = factorF
                        filaMarcador = i
                    End If
                End If
                
                ' Si nos pasamos mucho del alto de página, dejamos de simular este 'f'
                If altoAcumulado > altoPaginaReal * 1.1 Then Exit For
            Next i
        Next factorF
        
        ' --- 2. APLICACIÓN DE LOS RESULTADOS ---
        If filaMarcador > 0 Then
            Do While EsMarcador(ws.Cells(filaMarcador - 1, 1)) Or IsEmptyRow(ws.Rows(filaMarcador - 1))
                filaMarcador = filaMarcador - 1
            Loop
            ' Aplicamos el factor 'mejorF' a las filas en blanco de esta página
            altoAcumulado = 0
            For j = filaInicioPagina To filaMarcador - 1
                If IsEmptyRow(ws.Rows(j)) Then
                    ws.Rows(j).RowHeight = ws.Rows(j).RowHeight * mejorF
                End If
            Next j
            
            ' Insertar salto manual justo antes del marcador óptimo
            ws.HPageBreaks.Add Before:=ws.Rows(filaMarcador)
            npb = npb + 1
            filaInicioPagina = filaMarcador
        ElseIf i < ultimaFila Then
            ' Si no hubo marcador en el rango del 10%, buscamos el salto natural
            'Stop
            ws.HPageBreaks.Add Before:=ws.HPageBreaks(npb + 1).Location
            filaInicioPagina = ws.HPageBreaks(npb + 1).Location.Row
            npb = npb + 1
'            altoAcumulado = 0
'            For j = filaInicioPagina To ultimaFila
'                altoAcumulado = altoAcumulado + ws.Rows(j).RowHeight
'                If altoAcumulado > altoPaginaReal Then
'                    ws.HPageBreaks.Add Before:=ws.Rows(j + 1)
'                    filaInicioPagina = j + 1
'                    Exit For
'                End If
'                If j = ultimaFila Then filaInicioPagina = ultimaFila + 1
'            Next j
        End If
        
        ' Evitar bucle infinito si algo falla
        If filaInicioPagina > ultimaFila Or i > ultimaFila Then Exit Do
    Loop
    
    Application.ScreenUpdating = True
    'MsgBox "Paginación optimizada aplicada.", vbInformation
End Sub
Function SaltoVerticalCortaRango(ws As Worksheet, targetRange As Range) As Boolean
Attribute SaltoVerticalCortaRango.VB_Description = "[modMACROWbkEditableFormatting] Salto Vertical Corta Rango (función personalizada). Aplica a: Cells Range"
Attribute SaltoVerticalCortaRango.VB_ProcData.VB_Invoke_Func = " \n23"
    ' Devuelve True si el rango excede el primer salto de página vertical
    
    Dim primerSaltoColumna As Long
    Dim rangoFinColumna As Long
    
    SaltoVerticalCortaRango = False
    
    ' 1. Determinar dónde está el primer salto de página vertical
    If ws.VPageBreaks.Count > 0 Then
        ' Obtiene el número de columna donde se produce el salto
        primerSaltoColumna = ws.VPageBreaks(1).Location.Column
    Else
        ' Si no hay saltos manuales, no hay nada que cortar
        Exit Function
    End If
    
    ' 2. Determinar hasta dónde llega el rango de datos
    ' Obtenemos el número de la última columna del rango
    rangoFinColumna = targetRange.Column + targetRange.Columns.Count - 1
    
    ' 3. Comparar: ¿El final del rango está después del salto de página?
    If rangoFinColumna >= primerSaltoColumna Then
        ' El rango de datos es más ancho que el área de impresión definida por el salto.
        ' Se imprimirá en varias páginas.
        SaltoVerticalCortaRango = True
    End If
End Function

Sub CompactarAreaDeImpresion(ws As Worksheet, ByRef estado As CompactacionEstado)
    Dim printAreas As Variant
    Dim i As Long, j As Long
    Dim totalRange As Range
    Dim addr As String
    Dim minRow As Long, maxRow As Long, minCol As Long, maxCol As Long
    Dim hayMultiples As Boolean
    
    On Error GoTo ErrorHandler
    
    ' --- Inicializar estado como inválido (por seguridad) ---
    estado.EsValido = False
    
    ' --- 1. Verificar área de impresión ---
    If ws.PageSetup.PrintArea = "" Then
        Debug.Print "[WARN] CompactarAreaDeImpresion: no hay área de impresión en '" & ws.Name & "'."
        Exit Sub
    End If
    
    ' --- 2. Guardar estado original ---
    estado.PrintAreaOriginal = ws.PageSetup.PrintArea
    
    printAreas = Split(ws.PageSetup.PrintArea, ";")
    hayMultiples = (UBound(printAreas) - LBound(printAreas) >= 1)
    
    If Not hayMultiples Then
        Debug.Print "[OK] CompactarAreaDeImpresion: solo un rango. Nada que hacer."
        estado.EsValido = True  ' aún así, marcamos como válido (restaurable)
        Exit Sub
    End If
    
    Debug.Print "=== CompactarAreaDeImpresion: " & (UBound(printAreas) + 1) & " rangos en '" & ws.Name & "' ==="
    
    ' --- 3. Construir rango total y límites ---
    Set totalRange = Nothing
    minRow = 1048576: maxRow = 1
    minCol = 16384:   maxCol = 1
    
    For i = LBound(printAreas) To UBound(printAreas)
        addr = Trim(printAreas(i))
        If addr <> "" Then
            Dim rng As Range
            Set rng = ws.Range(addr)
            
            If totalRange Is Nothing Then
                Set totalRange = rng
            Else
                Set totalRange = Union(totalRange, rng)
            End If
            
            If rng.Row < minRow Then minRow = rng.Row
            If rng.Row + rng.Rows.Count - 1 > maxRow Then maxRow = rng.Row + rng.Rows.Count - 1
            If rng.Column < minCol Then minCol = rng.Column
            If rng.Column + rng.Columns.Count - 1 > maxCol Then maxCol = rng.Column + rng.Columns.Count - 1
        End If
    Next i
    
    ' --- 4. Guardar estado de visibilidad SOLO en el rango relevante ---
    estado.MinFila = minRow
    estado.MaxFila = maxRow
    estado.minCol = minCol
    estado.maxCol = maxCol
    
    ReDim estado.FilasEstado(minRow To maxRow)
    ReDim estado.ColumnasEstado(minCol To maxCol)
    
    For i = minRow To maxRow
        estado.FilasEstado(i) = Not ws.Rows(i).Hidden  ' True = visible originalmente
    Next i
    
    For j = minCol To maxCol
        estado.ColumnasEstado(j) = Not ws.Columns(j).Hidden
    Next j
    
    ' --- 5. Ocultar huecos internos (solo dentro del rango global) ---
    Dim filasOcultas As Long, colsOcultas As Long
    
    For i = minRow To maxRow
        Dim debeOcultarFila As Boolean
        debeOcultarFila = True
        For Each rng In totalRange.Areas
            If i >= rng.Row And i <= (rng.Row + rng.Rows.Count - 1) Then
                debeOcultarFila = False
                Exit For
            End If
        Next rng
        
        If debeOcultarFila And Not ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = True
            filasOcultas = filasOcultas + 1
        End If
    Next i
    
    For j = minCol To maxCol
        Dim debeOcultarCol As Boolean
        debeOcultarCol = True
        For Each rng In totalRange.Areas
            If j >= rng.Column And j <= (rng.Column + rng.Columns.Count - 1) Then
                debeOcultarCol = False
                Exit For
            End If
        Next rng
        
        If debeOcultarCol And Not ws.Columns(j).Hidden Then
            ws.Columns(j).Hidden = True
            colsOcultas = colsOcultas + 1
        End If
    Next j
    
    ' --- 6. Unificar área de impresión ---
    Dim unifiedRange As Range
    Set unifiedRange = ws.Range(ws.Cells(minRow, minCol), ws.Cells(maxRow, maxCol))
    ws.PageSetup.PrintArea = unifiedRange.Address
    
    ' --- 7. Marcar estado como válido ---
    estado.EsValido = True
    
    Debug.Print "[OK] Filas ocultas (huecos): " & filasOcultas
    Debug.Print "[OK] Columnas ocultas (huecos): " & colsOcultas
    Debug.Print "[OK] Área unificada: " & unifiedRange.Address
    Debug.Print "[OK] CompactarAreaDeImpresion: finalizado."
    
    Exit Sub

ErrorHandler:
    Debug.Print "[ERR] CompactarAreaDeImpresion: " & Err.Description & " (Err " & Err.Number & ")"
    estado.EsValido = False
End Sub

Sub RestaurarAreaDeImpresion(ws As Worksheet, estado As CompactacionEstado)
    On Error GoTo ErrorHandler
    
    If Not estado.EsValido Then
        Debug.Print "[WARN] RestaurarAreaDeImpresion: estado no válido. Nada restaurado."
        Exit Sub
    End If
    
    ' --- 1. Restaurar área de impresión original ---
    If estado.PrintAreaOriginal = "" Then
        ws.PageSetup.PrintArea = ""
    Else
        ws.PageSetup.PrintArea = estado.PrintAreaOriginal
    End If
    
    ' --- 2. Restaurar visibilidad de filas (solo en rango guardado) ---
    Dim i As Long, j As Long
    Dim filasRestauradas As Long, colsRestauradas As Long
    
    For i = estado.MinFila To estado.MaxFila
        If i >= LBound(estado.FilasEstado) And i <= UBound(estado.FilasEstado) Then
            ws.Rows(i).Hidden = Not estado.FilasEstado(i)  ' False si estaba visible
            If Not ws.Rows(i).Hidden Then filasRestauradas = filasRestauradas + 1
        End If
    Next i
    
    For j = estado.minCol To estado.maxCol
        If j >= LBound(estado.ColumnasEstado) And j <= UBound(estado.ColumnasEstado) Then
            ws.Columns(j).Hidden = Not estado.ColumnasEstado(j)
            If Not ws.Columns(j).Hidden Then colsRestauradas = colsRestauradas + 1
        End If
    Next j
    
    Debug.Print "[OK] RestaurarAreaDeImpresion: área original = '" & estado.PrintAreaOriginal & "'"
    Debug.Print "[OK] Filas restauradas a visible: " & filasRestauradas
    Debug.Print "[OK] Columnas restauradas a visible: " & colsRestauradas
    
    Exit Sub

ErrorHandler:
    Debug.Print "[ERR] RestaurarAreaDeImpresion: " & Err.Description & " (Err " & Err.Number & ")"
End Sub
