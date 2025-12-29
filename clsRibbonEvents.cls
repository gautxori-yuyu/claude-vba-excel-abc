VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public ribbonUI As IRibbonUI
' ==========================================
' EVENTOS DE APLICACION
' ==========================================
Public Event GenerarGraficosDesdeCurvasRto()
Public Event InvertirEjes()
Public Event FormatearCGASING()
Public Event Configurador()
Public Event NuevaOportunidad()
Public Event ReplaceWithNamesInValidations()

Public Sub Init(ByRef ribbonObj As IRibbonUI)
    Set ribbonUI = ribbonObj
End Sub

Public Sub StopEvents()
    Set ribbonUI = Nothing
End Sub

Public Sub OnGenerarGraficosDesdeCurvasRto()
    RaiseEvent GenerarGraficosDesdeCurvasRto
End Sub

Public Sub OnInvertirEjes()
    RaiseEvent InvertirEjes
End Sub

Public Sub OnFormatearCGASING()
    RaiseEvent FormatearCGASING
End Sub

Public Sub OnConfigurador()
    RaiseEvent Configurador
End Sub

Public Sub OnNuevaOportunidad()
    RaiseEvent NuevaOportunidad
End Sub

Public Sub OnReplaceWithNamesInValidations()
    RaiseEvent ReplaceWithNamesInValidations
End Sub

'@Description: Activa una pestaña del Ribbon por id
'@Scope: Friend (uso exclusivo desde clsAplicacion)
'@ArgumentDescriptions: tabId | id del tab a activar
'@Category: UI / Ribbon
Friend Sub ActivarTab(tabId As String)
    On Error Resume Next
    If Not ribbonUI Is Nothing Then
        ribbonUI.ActivateTab tabId
    End If
End Sub

Public Function GetRibbonControlEnabled(control As IRibbonControl) As Boolean
    Dim enabled As Boolean
    Select Case control.id
    Case "btnGenerarGraficos"                    ' GetGraficoEnabled
        enabled = EsFicheroOportunidad() And EsValidoGenerarGrafico()
    Case "btnInvertirSeries"                     ' GetInvertirEjesEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = App.bChartActive ' comprueba que hay un grafico seleccionado
        If enabled Then enabled = EsValidoInvertirEjes()
        ' HAY QUE PASAR A ESTO: el orquestador determina el estado del ribbon:
        ' ... PERO DE MOMENTO, NO ESTA TERMINADO DE IMPLEMENTAR clsExcelFile y clsFileManager
        ' enabled = App.bCanInvertAxes
    Case "btnCGASING"                            ' GetCGASINGEnabled
        enabled = EsFicheroOportunidad()
        If enabled Then enabled = IsDefaultCGasIngSheet()
    Case "btnNuevaOp"                            ' GetNuevaOportunidadEnabled
        enabled = True
    End Select
    GetRibbonControlEnabled = enabled
    'InvalidarControl control.id ' esto ESTA DE SOBRA!!
    Debug.Print "[GetRibbonControlEnabled] - ¿" & control.id & "?: " & CBool(enabled)
End Function

' Función pública para invalidar el ribbon desde otros módulos
Public Sub InvalidarRibbon()
    On Error Resume Next
    Select Case True
    Case TypeName(ribbonUI) = "Empty", ribbonUI Is Nothing
        Debug.Print "[clsRibbonEvents InvalidarRibbon] ERROR!!: ribbonUI es " & IIf(TypeName(ribbonUI) = "Empty", "Empty", "Nothing") & "!!! — posible fallo previo."
        Err.Raise vbObjectError + 573, "clsRibbonEvents.InvalidarRibbon", _
                  "Error en invalidacion de ribbon: " & Err.Description
    Case Else
        ribbonUI.Invalidate
    End Select
    On Error GoTo 0
End Sub

' Función pública para invalidar un control desde otros módulos
Public Sub InvalidarControl(idControl As String)
    On Error GoTo ErrHandler
    Select Case True
    Case TypeName(ribbonUI) = "Empty", ribbonUI Is Nothing
        Debug.Print "[clsRibbonEvents InvalidarControl] ERROR!!: ribbonUI es " & IIf(TypeName(ribbonUI) = "Empty", "Empty", "Nothing") & "!!! — posible fallo previo."
        Err.Raise vbObjectError + 573, "clsRibbonEvents.InvalidarControl", _
                  "Error en invalidacion de control de ribbon: " & Err.Description
    Case Else
        ribbonUI.InvalidateControl idControl
    End Select
    On Error GoTo 0
    
    Exit Sub
ErrHandler:
    Debug.Print "[InvalidarControl] - Error al invalidar control " & idControl & ": " & Err.Description
End Sub


