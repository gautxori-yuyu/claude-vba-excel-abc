Attribute VB_Name = "modMACROAppLifecycle"
' ==========================================
' CICLO DE VIDA DE LA APLICACION
' ==========================================
' Funciones publicas para gestion de la aplicacion y Ribbon
' ==========================================

Option Explicit

Public Function App() As clsAplicacion
Attribute App.VB_Description = "[modMACROAppLifecycle] App (función personalizada). Aplica a: ThisWorkbook"
Attribute App.VB_ProcData.VB_Invoke_Func = " \n23"
    Set App = ThisWorkbook.App
End Function

' ==========================================
' GESTION DEL RIBBON
' ==========================================

'@Description: Procedimiento puente para el atajo de teclado
Public Sub ToggleRibbonTab()
    On Error GoTo ErrHandler

    If Not App() Is Nothing Then
        App().ToggleRibbonMode
    End If

    Exit Sub
ErrHandler:
    Debug.Print "[ToggleRibbonTab] Error: " & Err.Description
    MsgBox "Error al cambiar modo del Ribbon: " & Err.Description, vbExclamation
End Sub

'@Description: Recupera el Ribbon si se ha perdido
'@Note: Ejecutar esta macro si el Ribbon desaparece o no responde
Public Sub RecuperarRibbon()
    modRibbonRecovery.RecuperarRibbonManual
End Sub

'@Description: Muestra diagnostico del estado del Ribbon
Public Sub DiagnosticoRibbon()
    modRibbonRecovery.MostrarDiagnosticoRibbon
End Sub

'@Description: Reinicia completamente la aplicacion ABC
Public Sub ReiniciarABC()
    modRibbonRecovery.ReiniciarAplicacion
End Sub

'@Description: Verifica si el Ribbon esta funcionando
'@Returns: Boolean | True si el Ribbon esta operativo
Public Function EstaRibbonDisponible() As Boolean
Attribute EstaRibbonDisponible.VB_Description = "[modMACROAppLifecycle] Verifica si el Ribbon esta funcionando"
Attribute EstaRibbonDisponible.VB_ProcData.VB_Invoke_Func = " \n23"
    EstaRibbonDisponible = modRibbonRecovery.IsRibbonAvailable()
End Function
