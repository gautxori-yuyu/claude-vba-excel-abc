Attribute VB_Name = "modMACROUtils"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

'--- PROCEDIMIENTOS DE EJEMPLO PARA USO RÁPIDO ---
Sub InsertarCheckboxConTexto()
Attribute InsertarCheckboxConTexto.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox con texto visible
    Call InsertarCheckbox(MostrarCaption:=True)
End Sub

Sub InsertarCheckboxMarcado()
Attribute InsertarCheckboxMarcado.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox marcado por defecto
    Call InsertarCheckbox(ValorInicial:=True, MostrarCaption:=False)
End Sub

Sub InsertarCheckboxPersonalizado()
Attribute InsertarCheckboxPersonalizado.VB_ProcData.VB_Invoke_Func = " \n0"
    ' Ejemplo: Checkbox con texto personalizado
    Call InsertarCheckbox(TextoPersonalizado:="Opción Personalizada", _
                          MostrarCaption:=True, _
                          HojaDestino:="CONFIG")
End Sub

Sub AplicarDirtyATodasLasHojasConFormulas()
Attribute AplicarDirtyATodasLasHojasConFormulas.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim ws As Worksheet
    Dim rFormulas As Range

    ' Desactivar las alertas para evitar errores si no hay fórmulas en una hoja
    ' On Error Resume Next
    Application.CalculateFullRebuild

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

    ' Reactivar el manejo de errores normal
    On Error GoTo 0
    
    MsgBox "El método Dirty se ha aplicado a todos los rangos con fórmulas en este libro.", vbInformation
End Sub


