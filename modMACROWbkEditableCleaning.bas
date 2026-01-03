Attribute VB_Name = "modMACROWbkEditableCleaning"
' ------------------------------------------
' LIMPIEZA Y PREPARACIÓN DE LIBROS / HOJAS
' Convertir un libro de Excel de oferta en
' "editable para enviar a agente comercial"
' ------------------------------------------

'@Folder "4-Oportunidades y compresores.d-Ofertas.Plantillas"
Option Explicit

Public Sub LimpiarLibroActual()
    If SheetExists(ActiveWorkbook, "BUDGET_QUOTE") And SheetExists(ActiveWorkbook, "BUDGET_QUOTE") Then
        MsgBox ("DE MOMENTO ESTE PROCEDIMIENTO NO ES APLICABLE A BUDGET QUOTES, PTE REVISAR ERRORES EN FORMULAS")
        Exit Sub
    End If
    Call LimpiarLibroYHojas(ActiveWorkbook)
End Sub

'@Description: Aplica rutinas de limpieza a un conjunto de hojas y al libro que las contiene
'@Scope: libro y hojas indicadas
'@ArgumentDescriptions: wb: libro objetivo | hojas: array de Worksheet
'@Returns: Nothing
'@Category: Limpieza de datos
Public Sub LimpiarLibroYHojas(Optional ByVal wb As Workbook = Nothing, Optional ByVal hojas As Variant = Empty)
    Dim ws As Worksheet
    Dim hojaInicial As Worksheet
    Dim nErroresWs As Long
    
    If wb Is Nothing And IsEmpty(hojas) Then Exit Sub
    If IsEmpty(hojas) Then
        Set hojas = wb.Worksheets 'Application.Transpose(Application.Transpose(wb.Worksheets))
    End If
    If wb Is Nothing Then
        For Each ws In hojas
            If wb Is Nothing Then
                Set wb = ws.Parent
            ElseIf Not wb Is ws.Parent Then
                MsgBox "Todas las hojas deben pertenecer al mismo libro de Excel", vbExclamation
                Exit Sub
            End If
        Next
     End If
   
    Set hojaInicial = ActiveSheet
    
    ' —— Forzar recálculo completo de *todo el libro*
    FullRecalc
    
    'Application.ScreenUpdating = False
    'Application.EnableEvents = False             ' Disable events to prevent triggering the change event recursively
    
    EjecutarInspectorDeDocumentoVBA wb
    
    For Each ws In hojas
        If Not ws Is Nothing Then
            nErroresWs = ContarYListarErroresEnHoja(ws)
            If nErroresWs > 0 Then
                If 6 = MsgBox("La hoja """ & ws.Name & """ contiene " & nErroresWs & " celda(s) con error(es) de cálculo." & vbCrLf & _
                       "Ver detalles en la Ventana Inmediato (Ctrl+G)." & vbCrLf & "¿DESEAS ELIMINAR TODAS LAS FORMULAS DE LA HOJA?", _
                       vbExclamation + vbYesNo + vbDefaultButton2) Then
                       ' el usuario consiente borrar formulas con errores
                       nErroresWs = 0
                End If
            End If
            If nErroresWs = 0 Then Call FormulasToValuesSheet(ws)
            EliminarFilasColumnasOcultasSheet ws
            ' La siguiente operación requiere que el libro / la hoja de excel esté activo:
            ResetearZoomSheet ws
        End If
    Next
    
    hojaInicial.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ' el siguiente paso requiere que la ventana del libro con las hojas a eliminar esté activa
    wb.Activate
    If ActiveWindow.SelectedSheets.Count = 0 Then
    ElseIf MsgBox("¿Deseas eliminar todas las hojas del libro no seleccionadas?", vbYesNo + vbDefaultButton2) = vbYes Then
        Call EliminarHojasNOSeleccionadas(wb)
    End If
End Sub
' =========================================================
' Función: ContarYListarErroresEnHoja
' Propósito: Recalcula la hoja con máxima garantía y lista
'            todos los errores de fórmula en Debug.Print.
' Parámetro:
'   ws (Worksheet) - hoja a verificar
' Retorna:
'   Long - número de celdas con error de fórmula
' =========================================================
Public Function ContarYListarErroresEnHoja(ws As Worksheet) As Long
Attribute ContarYListarErroresEnHoja.VB_Description = "[modMACROWbkEditableCleaning] Función: ContarYListarErroresEnHoja. Propósito: Recalcula la hoja con máxima garantía y lista. todos los errores de fórmula en Debug.Print. Parámetro:. ws (Worksheet) - hoja a verificar. Retorna:. Long - número de celdas con "
Attribute ContarYListarErroresEnHoja.VB_ProcData.VB_Invoke_Func = " \n21"
    Dim rngErrores As Range
    Dim cell As Range
    Dim nErrores As Long
    
    ' Guardar estado actual — seguro incluso si EnableEvents = False
    Dim prevCalcMode As XlCalculation
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim addr As String
    
    On Error GoTo ErrorHandler
    
    prevCalcMode = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    
    ' === 1. Configurar entorno para recálculo fiable ===
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = False
    
    ' === 2. Recálculo TOTAL con reconstrucción de dependencias ===
    ws.Calculate
    
    ' === 3. Buscar celdas con fórmulas que contengan errores ===
    On Error Resume Next
    Set rngErrores = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo 0
    
    nErrores = 0
    If Not rngErrores Is Nothing Then
        nErrores = rngErrores.Cells.Count
        
        ' === 4. Listar todos los errores en Debug.Print ===
        Debug.Print "=== ERRORES EN HOJA: """ & ws.Name & """ ==="
        Debug.Print "Celda       | Tipo de error"
        Debug.Print String(35, "-")
        
        For Each cell In rngErrores
            Dim v As Variant
            v = cell.Value2   ' USO Value2: más seguro y rápido
            
            Dim errStr As String
            If IsError(v) Then
                Select Case CLng(CVErr(v))
                    Case xlErrDiv0:   errStr = "#¡DIV/0!"
                    Case xlErrNA:     errStr = "#N/A"
                    Case xlErrName:   errStr = "#¿NOMBRE?"
                    Case xlErrNull:   errStr = "#¡NULO!"
                    Case xlErrNum:    errStr = "#¡NUM!"
                    Case xlErrRef:    errStr = "#¡REF!"
                    Case xlErrValue:  errStr = "#¡VALOR!"
                    Case Else:        errStr = "#" & CStr(v)
                End Select
            Else
                errStr = "(no error, pero SpecialCells lo incluyó: valor=" & CStr(v) & ")"
            End If
            
            addr = cell.Address(ReferenceStyle:=xlA1)
            Debug.Print Left$(addr & Space(12), 12) & "| " & errStr
        Next cell
        Debug.Print String(35, "=")
    Else
        Debug.Print "[OK] Hoja """ & ws.Name & """ -> Sin errores de fórmula."
    End If
    
    ' === 5. Restaurar estado original ===
Finish:
    Application.Calculation = prevCalcMode
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    ContarYListarErroresEnHoja = nErrores
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] Excepción en ContarYListarErroresEnHoja: " & Err.Description
    Resume Finish
End Function
'@Description: Convierte todas las fórmulas de una hoja en valores
'@Scope: hoja individual
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Limpieza de datos
Public Sub FormulasToValuesSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.Value = ws.UsedRange.Value
    Debug.Print "[modAPPBudgetQuotesUtilids FormulasToValuesAllSheets] - aplicada a: " & ws.Name
End Sub

'@Description: Resetea el zoom y posiciona el cursor en A1
'@Scope: hoja individual (requiere activación)
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Ajuste visual
Public Sub ResetearZoomSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    Debug.Print "[modAPPBudgetQuotesUtilids ResetearZoomSheet] - aplicada a: " & ws.Name
End Sub

'@Description: Elimina filas y columnas ocultas de una hoja
'@Scope: hoja individual
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Limpieza de datos
Public Sub EliminarFilasColumnasOcultasSheet(ByVal ws As Worksheet)
    Dim i As Long
    
    If ws Is Nothing Then Exit Sub
    
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    
    On Error GoTo ErrorHandler
    
    prevEnableEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    
    ' === 1. Configurar entorno para recálculo fiable ===
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    
    ' Filas (de abajo hacia arriba)
    For i = ws.UsedRange.Rows.Count To 1 Step -1
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' Columnas (de derecha a izquierda)
    For i = ws.UsedRange.Columns.Count To 1 Step -1
        If ws.Columns(i).Hidden Then
            ws.Columns(i).Delete
        End If
    Next i
    Debug.Print "[modAPPBudgetQuotesUtilids EliminarFilasColumnasOcultas] - aplicada a: " & ws.Name
    ' === Restaurar estado original ===
Finish:
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    Exit Sub

ErrorHandler:
    Debug.Print "[ERR] Excepción en EliminarFilasColumnasOcultas: " & Err.Description
    Resume Finish
End Sub

'@Description: Ejecuta el inspector de documento para eliminar metadatos y datos privados
'@Scope: libro completo
'@ArgumentDescriptions: wb: libro a procesar
'@Returns: Nothing
'@Category: Seguridad / Privacidad
Public Sub EjecutarInspectorDeDocumentoVBA(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    
    ' 1. Eliminar propiedades del documento e información personal
    ' Equivale a la opción "Propiedades del documento e información personal"
    wb.RemoveDocumentInformation (xlRDIDocumentProperties)
    wb.RemoveDocumentInformation (xlRDIRemovePersonalInformation)
    
    ' 2. Eliminar comentarios y notas
    ' wb.RemoveDocumentInformation (xlRDIInkAnnotations)
    ' wb.RemoveDocumentInformation (xlRDIComments)
    ' wb.RemoveDocumentInformation (xlRDIDefinedNameComments)
    
    ' 3. Eliminar nombres definidos y rutas de publicación (si existen)
    wb.RemoveDocumentInformation (xlRDIInlineWebExtensions)
    wb.RemoveDocumentInformation (xlRDIDocumentManagementPolicy)
    wb.RemoveDocumentInformation (xlRDIExcelDataModel)
    wb.RemoveDocumentInformation (xlRDIPublishInfo)
    
    Debug.Print "[modAPPBudgetQuotesUtilids EjecutarInspectorDeDocumentoVBA] - Metadatos y datos ocultos eliminados correctamente."
End Sub
Public Sub EliminarHojasNOSeleccionadas(ByVal wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If ws Is Nothing Then
        ElseIf Not HojaEstaSeleccionada(ws.Name) Then
            ws.Delete
        End If
    Next
End Sub

