VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Description: Estado l�gico del Ribbon
'@Category: State

'@Folder "3-UI.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbonState"

Public Event StateChanged()

Private mModoRibbon As eRibbonMode
Private mVisible As Boolean

Public Property Get Modo() As eRibbonMode
    Modo = mModoRibbon
End Property

Public Property Let Modo(Value As eRibbonMode)
    If mModoRibbon = Ribbon_Undefined Then
        ' se considera una INICIALIZACION, no un cambio de estado
        mModoRibbon = Value
    ElseIf mModoRibbon <> Value Then
        mModoRibbon = Value
        RaiseEvent StateChanged
    End If
End Property

Public Sub ToggleModo()
    Select Case mModoRibbon
    Case Ribbon_OpportunityOnly
        mModoRibbon = Ribbon_User
    Case Ribbon_User
        mModoRibbon = Ribbon_Hidden
    Case Ribbon_Hidden
        mModoRibbon = Ribbon_Admin
    Case Ribbon_Admin
        mModoRibbon = Ribbon_OpportunityOnly
    Case Else
        mModoRibbon = Ribbon_OpportunityOnly
    End Select
    RaiseEvent StateChanged
End Sub

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
'@Param isOpportunityFileActive: Indica si hay un archivo de oportunidad activo
'@Scope: público (consultado por callbacks)
'@Returns: Boolean
'@Category: UI / Estado
Public Function IsRibbonTabVisible(isOpportunityFileActive As Boolean) As Boolean
    Select Case mModoRibbon
    Case Ribbon_Admin
        IsRibbonTabVisible = True

    Case Ribbon_User
        IsRibbonTabVisible = True

    Case Ribbon_OpportunityOnly
        IsRibbonTabVisible = isOpportunityFileActive

    Case Else
        IsRibbonTabVisible = False
    End Select
End Function

Public Function IsAdminGroupVisible() As Boolean
    IsAdminGroupVisible = (mModoRibbon = Ribbon_Admin)
End Function

