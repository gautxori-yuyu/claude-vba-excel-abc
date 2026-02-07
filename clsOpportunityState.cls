VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpportunityState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado relacionado con la oportunidad actual
'@Category: State
'@Folder "5-Aplicac (Coord).Estado"
Option Explicit

Private Const MODULE_NAME As String = "clsOpportunityState"

' ==========================================
' ESTADO DE OPORTUNIDAD
' ==========================================
Private mCurrentOpportunity As clsOpportunity
Private mCurrentIndex As Long

' ==========================================
' INICIALIZACIÓN
' ==========================================

Private Sub Class_Initialize()
    Set mCurrentOpportunity = Nothing
    mCurrentIndex = -1
End Sub

Private Sub Class_Terminate()
    Set mCurrentOpportunity = Nothing
End Sub

' ==========================================
' PROPIEDADES PÚBLICAS
' ==========================================

'@Description: Obtiene o establece la oportunidad actual
Public Property Get CurrentOpportunity() As clsOpportunity
    Set CurrentOpportunity = mCurrentOpportunity
End Property

Friend Property Set CurrentOpportunity(op As clsOpportunity)
    Set mCurrentOpportunity = op
    LogDebug MODULE_NAME, "[CurrentOpportunity] Changed to: " & GetCurrentOpportunityLabel()
End Property

'@Description: Obtiene o establece el índice de la oportunidad actual en la colección
Public Property Get CurrentIndex() As Long
    CurrentIndex = mCurrentIndex
End Property

Friend Property Let CurrentIndex(Value As Long)
    mCurrentIndex = Value
    LogDebug MODULE_NAME, "[CurrentIndex] Changed to: " & Value
End Property

' ==========================================
' CONSULTAS DE ESTADO
' ==========================================

'@Description: Indica si hay una oportunidad actualmente seleccionada
Public Function HasCurrentOpportunity() As Boolean
    HasCurrentOpportunity = Not (mCurrentOpportunity Is Nothing)
End Function

'@Description: Obtiene el label/nombre de la oportunidad actual
'@Returns: String con el nombre de la oportunidad o "(ninguna)"
Public Function GetCurrentOpportunityLabel() As String
    If mCurrentOpportunity Is Nothing Then
        GetCurrentOpportunityLabel = "(ninguna)"
    Else
        On Error Resume Next
        GetCurrentOpportunityLabel = mCurrentOpportunity.Label
        If Err.Number <> 0 Then GetCurrentOpportunityLabel = "(error)"
        On Error GoTo 0
    End If
End Function
