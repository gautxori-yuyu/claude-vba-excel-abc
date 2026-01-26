Attribute VB_Name = "modAPPBudgetQuotesUtilids"
' ==========================================
' Módulo de utilidades para gestión de presupuestos
' ==========================================
'FIXME: Las funciones de este modulo son GENERICAS, no solo para Budgets (aunque alguna ESTA PENSADA SOLO PARA CIERTAS PLANTILLAS)
'   mejor ponerlas en un modulo "generico".

'@Folder "6-DOMINIO-Oportunidades y compresores.d-Ofertas.Plantillas"
Option Explicit

Private Const MODULE_NAME As String = "modAPPBudgetQuotesUtilids"

Private Const bSheetReport As Boolean = True

' ------------------------------------------
' Gestion de nombres de rangos y validaciones de datos
' ------------------------------------------

' Recalcula y crea nombres de rangos basados en la columna A de hojas específicas
Sub recalcNames()
Attribute recalcNames.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim begin As Variant, rangename As String
    Dim rRange As Range, rCell As Range
    Dim namesCol As String, rangeCol As String
    
    ' Solo aplica a hojas específicas
    ' 1. Determinar namesCol según la hoja
    Select Case ActiveSheet.Name
    Case "BUDGET_SELECTOR" ' BUDGET_QUOTE_TEMPLATE
        namesCol = "A"
    Case "B._OPTIONS_SELECTOR" ' QUOTATION_TEMPLATE-OILGAS-ES-EN-FR-DE-CN-PO-RU
        namesCol = "B"
    Case Else
        Exit Sub
    End Select
    
    ' Obtener la letra de la siguiente columna de forma robusta
    ' Convertimos la letra a número, sumamos 1, y volvemos a obtener la letra
    rangeCol = Split(Cells(1, Columns(namesCol).Column + 1).Address, "$")(1)
    
    Application.Calculation = xlManual
    Application.EnableEvents = False             ' Disable events to prevent triggering the change event recursively
    
    Set rRange = ActiveSheet.UsedRange.Columns(namesCol).Cells
    If rRange.Areas.Count > 1 Then Exit Sub
    If rRange.Columns.Count > 1 Then Exit Sub
    
    Set rCell = rRange.Cells(1, 1)               ' Take the first cell in the range
    
    Do
        If rCell.Value <> "" Then
            If Not IsEmpty(begin) Then
                Debug.Print rangename & "==" & "=$" & rangeCol & "$" & begin & ":$" & rangeCol & "$" & rCell.Row - 1
                ActiveWorkbook.Names.Add Name:=rangename, _
                    RefersTo:="=$" & rangeCol & "$" & begin & ":$" & rangeCol & "$" & rCell.Row - 1
            End If
            begin = rCell.Row
            rangename = Replace(Replace(Replace(rCell.Value, "-", ""), " / ", " "), " ", "_")
        End If
        Set rCell = rCell.Offset(1, 0)           ' Jump 1 row down to the next cell
    Loop Until (rCell.Row > (rRange.Row + rRange.Rows.Count - 1))
    
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True              ' Re-enable events
End Sub

' Aplica nombres de rangos a todas las hojas del libro
Sub AplicarNombresARangosUsadosActWB()
Attribute AplicarNombresARangosUsadosActWB.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim rangoUsado As Range
    
    Application.ScreenUpdating = False
    
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next                     ' Por si la hoja está vacía
        Set rangoUsado = ws.UsedRange
        On Error GoTo 0
        
        If Not rangoUsado Is Nothing Then
            ' Aplicar nombres al rango usado de cada hoja
            On Error Resume Next
            rangoUsado.ApplyNames
            
            If Err Then
                Debug.Print "Error al aplicar nombres a " & ws.Name
            Else
                Debug.Print "Nombres aplicados a " & ws.Name
            End If
            
            On Error GoTo 0
            Set rangoUsado = Nothing
        End If
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Nombres aplicados a todas las hojas del libro"
End Sub

'@UDF
'@Description: Reemplaza referencias de rango por nombres definidos dentro de una fórmula
'@Scope: Libro Activo
'@ArgumentDescriptions: fórmula a procesar, con referencias del tipo "A2", "$G$7", ...
'@Returns: String - fórmula resultante
'@Category: Validaciones de datos
Public Function AplicarNombresAFormula(ByVal formula As String) As String
Attribute AplicarNombresAFormula.VB_Description = "[modAPPBudgetQuotesUtilids] Reemplaza referencias de rango por nombres definidos dentro de una fórmula. Aplica a: ActiveWorkbook|Cells Range\r\nM.D.:Libro Activo"
Attribute AplicarNombresAFormula.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim nm As Name
    Dim rngNombre As Range
    Dim formulaResultado As String
    
    formulaResultado = formula
    
    For Each nm In ActiveWorkbook.Names
        On Error Resume Next
        Set rngNombre = nm.RefersToRange
        On Error GoTo 0
        
        If Not rngNombre Is Nothing Then
            ' Reemplazar diferentes formatos
            formulaResultado = Replace(formulaResultado, rngNombre.Address(True, True, xlA1, True), nm.Name)
            formulaResultado = Replace(formulaResultado, rngNombre.Address(False, False, xlA1, True), nm.Name)
            formulaResultado = Replace(formulaResultado, rngNombre.Address(True, False, xlA1, True), nm.Name)
            formulaResultado = Replace(formulaResultado, rngNombre.Address(False, True, xlA1, True), nm.Name)
            
            ' Referencias externas
            If InStr(rngNombre.Worksheet.Name, " ") > 0 Then
                formulaResultado = Replace(formulaResultado, _
                                           "'" & rngNombre.Worksheet.Name & "'!" & rngNombre.Address, nm.Name)
            Else
                formulaResultado = Replace(formulaResultado, _
                                           rngNombre.Worksheet.Name & "!" & rngNombre.Address, nm.Name)
            End If
        End If
    Next nm
    
    AplicarNombresAFormula = formulaResultado
End Function

'@Description: Actualiza validaciones del libro y genera un reporte detallado
'@Scope: libro completo
'@ArgumentDescriptions: (sin argumentos)
'@Returns: Nothing
'@Category: Validaciones de datos
Public Sub AplicarNombresAValidacionesCeldasActWBConReporte()
Attribute AplicarNombresAValidacionesCeldasActWBConReporte.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim rangoValidacion As Range
    Dim celda As Range
    Dim reporte As String
    Dim contadorTotal As Long
    Dim contadorHoja As Long
    
    Application.ScreenUpdating = False
    contadorTotal = 0
    reporte = "REPORTE DE ACTUALIZACIÓN DE VALIDACIONES" & vbCrLf & vbCrLf
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            contadorHoja = 0
            reporte = reporte & "Hoja: " & ws.Name & vbCrLf
            
            On Error Resume Next
            Set rangoValidacion = ws.Cells.SpecialCells(xlCellTypeAllValidation)
            On Error GoTo 0
            
            If Not rangoValidacion Is Nothing Then
                For Each celda In rangoValidacion
                    If ActualizarValidacionCelda(celda) Then
                        contadorHoja = contadorHoja + 1
                        reporte = reporte & "  • " & celda.Address & vbCrLf
                    End If
                Next celda
                
                If contadorHoja > 0 Then
                    reporte = reporte & "  Total actualizadas: " & contadorHoja & vbCrLf
                Else
                    reporte = reporte & "  Sin cambios necesarios" & vbCrLf
                End If
                
                contadorTotal = contadorTotal + contadorHoja
                Set rangoValidacion = Nothing
            Else
                reporte = reporte & "  Sin validaciones de datos" & vbCrLf
            End If
            
            reporte = reporte & vbCrLf
        End If
    Next ws
    
    Application.ScreenUpdating = True
    
    ' Mostrar reporte
    reporte = reporte & "TOTAL GENERAL: " & contadorTotal & " validaciones actualizadas"
    
    ' Mostrar en mensaje (para pocos datos) o en hoja nueva (para muchos datos)
    If Len(reporte) < 1000 Or Not bSheetReport Then
        MsgBox reporte, vbInformation, "Reporte de Actualización"
    ElseIf bSheetReport Then
        MostrarReporteEnHoja reporte
    End If
End Sub

'@Description: Actualiza una validación de datos de una celda reemplazando rangos por nombres
'@Scope: celda individual
'@ArgumentDescriptions: celda: celda con validación
'@Returns: Boolean - indica si se ha actualizado
'@Category: Validaciones de datos
Public Function ActualizarValidacionCelda(ByVal celda As Range) As Boolean
Attribute ActualizarValidacionCelda.VB_Description = "[modAPPBudgetQuotesUtilids] Actualiza una validación de datos de una celda reemplazando rangos por nombres. Aplica a: Cells Range\r\nM.D.:celda individual"
Attribute ActualizarValidacionCelda.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim actualizado As Boolean
    Dim formula1Original As String, formula1Nueva As String
    Dim formula2Original As String, formula2Nueva As String
    
    actualizado = False
    
    With celda.Validation
        If .Type <> 0 Then
            ' Procesar Formula1
            formula1Original = .Formula1
            formula1Nueva = AplicarNombresAFormula(formula1Original)
            
            ' Procesar Formula2 (si existe)
            formula2Original = .Formula2
            If formula2Original <> "" Then
                formula2Nueva = AplicarNombresAFormula(formula2Original)
            Else
                formula2Nueva = ""
            End If
            
            ' Aplicar cambios si hay diferencias
            If formula1Nueva <> formula1Original Or formula2Nueva <> formula2Original Then
                On Error Resume Next
                .Modify Type:=.Type, _
                        AlertStyle:=.AlertStyle, _
                        Operator:=.Operator, _
                        Formula1:=formula1Nueva, _
                        Formula2:=formula2Nueva
                On Error GoTo 0
                actualizado = True
            End If
        End If
    End With
    
    ActualizarValidacionCelda = actualizado
End Function

'@Description: Muestra un reporte de texto en una hoja del libro
'@Scope: libro activo
'@ArgumentDescriptions: reporte: texto a mostrar
'@Returns: Nothing
'@Category: Reporting
Sub MostrarReporteEnHoja(reporte As String)
Attribute MostrarReporteEnHoja.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim wsReporte As Worksheet
    Dim lineas() As String
    Dim i As Long
    
    ' Crear o limpiar hoja de reporte
    On Error Resume Next
    Set wsReporte = ActiveWorkbook.Worksheets("Reporte_Validaciones")
    On Error GoTo 0
    
    If wsReporte Is Nothing Then
        Set wsReporte = ActiveWorkbook.Worksheets.Add
        wsReporte.Name = "Reporte_Validaciones"
    Else
        wsReporte.Cells.Clear
    End If
    
    ' Dividir reporte en líneas y escribir en hoja
    lineas = Split(reporte, vbCrLf)
    For i = 0 To UBound(lineas)
        wsReporte.Cells(i + 1, 1).Value = lineas(i)
    Next i
    
    wsReporte.Columns("A").AutoFit
    wsReporte.Activate
    MsgBox "Reporte generado en hoja '" & wsReporte.Name & "'", vbInformation
End Sub


