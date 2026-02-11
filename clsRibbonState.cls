VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Description: Estado lógico del Ribbon
'@Category: State

'@Folder "4-UI.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbonState"

Private mModoRibbon As eRibbonMode

Public Property Get Modo() As eRibbonMode
    Modo = mModoRibbon
End Property

Friend Property Let Modo(Value As eRibbonMode)
    mModoRibbon = Value
End Property

Public Function Description() As String
    Select Case mModoRibbon
    Case Ribbon_Admin
        Description = "Admin"
    Case Ribbon_Hidden
        Description = "(oculto)"
    Case Ribbon_OpportunityOnly
        Description = "Fich. oportunidad"
    Case Ribbon_User
        Description = "User"
    Case Else
        Description = "Fich. oportunidad"
    End Select
End Function

