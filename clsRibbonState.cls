VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description: Estado lógico del Ribbon
'@Category: State
Option Explicit

Public Event StateChanged()

Private mModoRibbon As eRibbonMode
Private mVisible As Boolean

Public Property Get Modo() As eRibbonMode
    Modo = mModoRibbon
End Property

Public Property Let Modo(Value As eRibbonMode)
    If mModoRibbon <> Value Then
        mModoRibbon = Value
        RaiseEvent StateChanged
    End If
End Property

Public Sub ToggleModo()
    Select Case mModoRibbon
    Case Ribbon_Admin
        mModoRibbon = Ribbon_Hidden
    Case Ribbon_Hidden
        mModoRibbon = Ribbon_User
    Case Ribbon_User
        mModoRibbon = Ribbon_OpportunityOnly
    Case Ribbon_OpportunityOnly
        mModoRibbon = Ribbon_Admin
    Case Else
        mModoRibbon = Ribbon_OpportunityOnly
    End Select
    RaiseEvent StateChanged
End Sub

Public Function RibbonStateDescription() As String
    Select Case mModoRibbon
    Case Ribbon_Admin
        RibbonStateDescription = "Admin"
    Case Ribbon_Hidden
        RibbonStateDescription = "(oculto)"
    Case Ribbon_OpportunityOnly
        RibbonStateDescription = "Fich. oportunidad"
    Case Ribbon_User
        RibbonStateDescription = "User"
    Case Else
        RibbonStateDescription = "Fich. oportunidad"
    End Select
End Function

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Public Property Let Visible(Value As Boolean)
    If mVisible <> Value Then
        mVisible = Value
        RaiseEvent StateChanged
    End If
End Property

'@Description: Indica si la pestaña principal del Ribbon debe ser visible
'@Scope: público (consultado por callbacks)
'@Returns: Boolean
'@Category: UI / Estado
Public Function IsRibbonTabVisible() As Boolean
    Select Case mModoRibbon
    Case Ribbon_Admin
        IsRibbonTabVisible = True

    Case Ribbon_User
        IsRibbonTabVisible = True

    Case Ribbon_OpportunityOnly
        IsRibbonTabVisible = EsFicheroOportunidad()

    Case Else
        IsRibbonTabVisible = False
    End Select
End Function

Public Function IsAdminGroupVisible() As Boolean
    IsAdminGroupVisible = (mModoRibbon = Ribbon_Admin)
End Function

