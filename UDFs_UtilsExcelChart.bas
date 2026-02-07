Attribute VB_Name = "UDFs_UtilsExcelChart"
'@Folder "4-Servicios.Excel.Charts"
'@IgnoreModule MissingAnnotationArgument
Option Explicit

Private Const MODULE_NAME As String = "UDFs_UtilsExcelChart"

'@UDF
'@Description: Establece el valor mínimo o máximo de un eje de gráfico (primario o secundario)
'@Category: Gráficos
'@ArgumentDescriptions: "Min" o "Max"|"Value" o "Category"|"Primary" o "Secondary"|Valor del límite (numérico o "Auto")|Gráfico a modificar (opcional)
Public Function setChartAxis(MinOrMax As String, _
                             ValueOrCategory As String, _
                             PrimaryOrSecondary As String, _
                             Value As Variant, _
                             Optional cht As Chart = Nothing) As String
Attribute setChartAxis.VB_Description = "[UDFs_UtilsExcelChart] Establece el valor mínimo o máximo de un eje de gráfico (primario o secundario). Aplica a: ActiveSheet|Cells Range"
Attribute setChartAxis.VB_ProcData.VB_Invoke_Func = " \n23"
    
    Dim valueAsText As String
    
    On Error GoTo ErrorHandler
    
    ' Determinar el gráfico a controlar
    If Not cht Is Nothing Then
        ' Gráfico proporcionado por parámetro
    ElseIf ActiveSheet.ChartObjects.Count = 0 Then
        setChartAxis = "No hay gráficos en la hoja"
        Exit Function
    ElseIf Not TypeOf Application.Caller Is Range Then
        Set cht = ActiveSheet.ChartObjects(1).Chart
    Else
        Set cht = Application.Caller.Worksheet.ChartObjects(1).Chart
    End If
    
    ' Aplicar valor según el tipo de eje
    Select Case True
        ' Eje de valores primario
    Case (ValueOrCategory = "Value" Or ValueOrCategory = "Y") And _
         PrimaryOrSecondary = "Primary"
        With cht.Axes(xlValue, xlPrimary)
            If IsNumeric(Value) Then
                If MinOrMax = "Max" Then .MaximumScale = CDbl(Value)
                If MinOrMax = "Min" Then .MinimumScale = CDbl(Value)
            Else
                If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
                If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
            End If
        End With
        
        ' Eje de categorías primario
    Case (ValueOrCategory = "Category" Or ValueOrCategory = "X") And _
         PrimaryOrSecondary = "Primary"
        With cht.Axes(xlCategory, xlPrimary)
            If IsNumeric(Value) Then
                If MinOrMax = "Max" Then .MaximumScale = CDbl(Value)
                If MinOrMax = "Min" Then .MinimumScale = CDbl(Value)
            Else
                If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
                If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
            End If
        End With
        
        ' Eje de valores secundario
    Case (ValueOrCategory = "Value" Or ValueOrCategory = "Y") And _
         PrimaryOrSecondary = "Secondary"
        With cht.Axes(xlValue, xlSecondary)
            If IsNumeric(Value) Then
                If MinOrMax = "Max" Then .MaximumScale = CDbl(Value)
                If MinOrMax = "Min" Then .MinimumScale = CDbl(Value)
            Else
                If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
                If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
            End If
        End With
        
        ' Eje de categorías secundario
    Case (ValueOrCategory = "Category" Or ValueOrCategory = "X") And _
         PrimaryOrSecondary = "Secondary"
        With cht.Axes(xlCategory, xlSecondary)
            If IsNumeric(Value) Then
                If MinOrMax = "Max" Then .MaximumScale = CDbl(Value)
                If MinOrMax = "Min" Then .MinimumScale = CDbl(Value)
            Else
                If MinOrMax = "Max" Then .MaximumScaleIsAuto = True
                If MinOrMax = "Min" Then .MinimumScaleIsAuto = True
            End If
        End With
    End Select
    
    ' Preparar texto de salida
    If IsNumeric(Value) Then
        valueAsText = CStr(Value)
    Else
        valueAsText = "Auto"
    End If
    
    setChartAxis = ValueOrCategory & " " & PrimaryOrSecondary & " " & _
                   MinOrMax & ": " & valueAsText
    
    Exit Function
    
ErrorHandler:
    setChartAxis = "#ERROR: " & Err.Description
End Function


