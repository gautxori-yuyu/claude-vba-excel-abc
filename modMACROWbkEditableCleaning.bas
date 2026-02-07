Attribute VB_Name = "modMACROWbkEditableCleaning"
' ------------------------------------------
' LIMPIEZA Y PREPARACIÓN DE LIBROS / HOJAS
' Convertir un libro de Excel de oferta en
' "editable para enviar a agente comercial"
' ------------------------------------------

'@Folder "6-DOMINIO-Oportunidades y compresores.d-Ofertas.Plantillas"
Option Explicit

Private Const MODULE_NAME As String = "modMACROWbkEditableCleaning"

Public Sub LimpiarLibroActual()
Attribute LimpiarLibroActual.VB_ProcData.VB_Invoke_Func = " \n0"
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
Public Sub LimpiarLibroYHojas(Optional ByVal Wb As Workbook = Nothing, Optional ByVal hojas As Variant = Empty)
Attribute LimpiarLibroYHojas.VB_ProcData.VB_Invoke_Func = " \n0"
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim hojaInicial As Worksheet
    Dim nErroresWs As Long

    ' Validaciones iniciales (ANTES de modificar estado)
    If Wb Is Nothing And IsEmpty(hojas) Then Exit Sub
    If IsEmpty(hojas) Then
        Set hojas = Wb.Worksheets 'Application.Transpose(Application.Transpose(wb.Worksheets))
    End If
    If Wb Is Nothing Then
        For Each ws In hojas
            If Wb Is Nothing Then
                Set Wb = ws.Parent
            ElseIf Not Wb Is ws.Parent Then
                MsgBox "Todas las hojas deben pertenecer al mismo libro de Excel", vbExclamation
                Exit Sub
            End If
        Next
     End If

    Set hojaInicial = ActiveSheet

    ' === GUARDAR ESTADO ORIGINAL ===
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents

    ' === MODIFICAR ESTADO PARA MEJOR RENDIMIENTO ===
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Forzar recálculo completo de *todo el libro*
    FullRecalc

    EjecutarInspectorDeDocumentoVBA Wb

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

    ' El siguiente paso requiere ScreenUpdating = True temporalmente para activar ventana
    Application.ScreenUpdating = True
    Wb.Activate

    If ActiveWindow.SelectedSheets.Count > 0 Then
        If MsgBox("¿Deseas eliminar todas las hojas del libro no seleccionadas?", vbYesNo + vbDefaultButton2) = vbYes Then
            Call EliminarHojasNOSeleccionadas(Wb)
        End If
    End If

CleanUp:
    ' === RESTAURAR ESTADO ORIGINAL ===
    On Error Resume Next
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    Exit Sub

ErrHandler:
    LogCurrentError MODULE_NAME, "[LimpiarLibroYHojas]"
    MsgBox "Error en limpieza de libro: " & Err.Description, vbCritical
    Resume CleanUp
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
Attribute ContarYListarErroresEnHoja.VB_ProcData.VB_Invoke_Func = " \n23"
    Dim rngErrores As Range
    Dim cell As Range
    Dim nErrores As Long
    
    ' Guardar estado actual - seguro incluso si EnableEvents = False
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
        LogInfo MODULE_NAME, "[ContarYListarErroresEnHoja]"
        LogInfo MODULE_NAME, "=== ERRORES EN HOJA: """ & ws.Name & """ ==="
        LogInfo MODULE_NAME, "Celda       | Tipo de error"
        LogInfo MODULE_NAME, String(35, "-")
        
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
            LogInfo MODULE_NAME, Left$(addr & Space(12), 12) & "| " & errStr
        Next cell
        LogInfo MODULE_NAME, String(35, "=")
    Else
        LogInfo MODULE_NAME, "[ContarYListarErroresEnHoja] Hoja """ & ws.Name & """ -> Sin errores de fórmula."
    End If
    
    ' === 5. Restaurar estado original ===
Finish:
    Application.Calculation = prevCalcMode
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    ContarYListarErroresEnHoja = nErrores
    Exit Function

ErrorHandler:
    LogCurrentError MODULE_NAME, "[ContarYListarErroresEnHoja]"
    Resume Finish
End Function
'@Description: Convierte todas las fórmulas de una hoja en valores
'@Scope: hoja individual
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Limpieza de datos
Public Sub FormulasToValuesSheet(ByVal ws As Worksheet)
Attribute FormulasToValuesSheet.VB_ProcData.VB_Invoke_Func = " \n0"
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.Value = ws.UsedRange.Value
    LogInfo MODULE_NAME, "[FormulasToValuesAllSheets] aplicada a: " & ws.Name
End Sub

'@Description: Resetea el zoom y posiciona el cursor en A1
'@Scope: hoja individual (requiere activación)
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Ajuste visual
Public Sub ResetearZoomSheet(ByVal ws As Worksheet)
Attribute ResetearZoomSheet.VB_ProcData.VB_Invoke_Func = " \n0"
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.Zoom = 100
    ws.Range("A1").Select
    LogInfo MODULE_NAME, "[ResetearZoomSheet] aplicada a: " & ws.Name
End Sub

'@Description: Elimina filas y columnas ocultas de una hoja
'@Scope: hoja individual
'@ArgumentDescriptions: ws: hoja a procesar
'@Returns: Nothing
'@Category: Limpieza de datos
Public Sub EliminarFilasColumnasOcultasSheet(ByVal ws As Worksheet)
Attribute EliminarFilasColumnasOcultasSheet.VB_ProcData.VB_Invoke_Func = " \n0"
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
    LogInfo MODULE_NAME, "[EliminarFilasColumnasOcultas] aplicada a: " & ws.Name
    ' === Restaurar estado original ===
Finish:
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    Exit Sub

ErrorHandler:
    LogCurrentError MODULE_NAME, "[EliminarFilasColumnasOcultas]"
    Resume Finish
End Sub

'@Description: Ejecuta el inspector de documento para eliminar metadatos y datos privados
'@Scope: libro completo
'@ArgumentDescriptions: wb: libro a procesar
'@Returns: Nothing
'@Category: Seguridad / Privacidad
Public Sub EjecutarInspectorDeDocumentoVBA(ByVal Wb As Workbook)
Attribute EjecutarInspectorDeDocumentoVBA.VB_ProcData.VB_Invoke_Func = " \n0"
    If Wb Is Nothing Then Exit Sub
    
    ' 1. Eliminar propiedades del documento e información personal
    ' Equivale a la opción "Propiedades del documento e información personal"
    Wb.RemoveDocumentInformation (xlRDIDocumentProperties)
    Wb.RemoveDocumentInformation (xlRDIRemovePersonalInformation)
    
    ' 2. Eliminar comentarios y notas
    ' wb.RemoveDocumentInformation (xlRDIInkAnnotations)
    ' wb.RemoveDocumentInformation (xlRDIComments)
    ' wb.RemoveDocumentInformation (xlRDIDefinedNameComments)
    
    ' 3. Eliminar nombres definidos y rutas de publicación (si existen)
    Wb.RemoveDocumentInformation (xlRDIInlineWebExtensions)
    Wb.RemoveDocumentInformation (xlRDIDocumentManagementPolicy)
    Wb.RemoveDocumentInformation (xlRDIExcelDataModel)
    Wb.RemoveDocumentInformation (xlRDIPublishInfo)
    
    LogInfo MODULE_NAME, "[EjecutarInspectorDeDocumentoVBA] Metadatos y datos ocultos eliminados correctamente."
End Sub
Public Sub EliminarHojasNOSeleccionadas(ByVal Wb As Workbook)
Attribute EliminarHojasNOSeleccionadas.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    For Each ws In Wb.Sheets
        If ws Is Nothing Then
        ElseIf Not HojaEstaSeleccionada(ws.Name) Then
            ws.Delete
        End If
    Next
End Sub

