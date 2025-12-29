Attribute VB_Name = "modMACROAppLifecycle"
Public Function App() As clsAplicacion
Attribute App.VB_Description = "[modMACROAppLifecycle] App (función personalizada). Aplica a: ThisWorkbook"
Attribute App.VB_ProcData.VB_Invoke_Func = " \n21"
    Set App = ThisWorkbook.App
End Function

'@Description: Procedimiento puente para el atajo de teclado
Public Sub ToggleRibbonTab()
Attribute ToggleRibbonTab.VB_ProcData.VB_Invoke_Func = " \n0"
    If Not App() Is Nothing Then
        App().ToggleRibbonMode
    End If
End Sub

