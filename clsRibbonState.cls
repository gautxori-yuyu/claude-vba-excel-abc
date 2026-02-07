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

'@Folder "3-UI.Excel.Ribbon"
Option Explicit

Private Const MODULE_NAME As String = "clsRibbonState"

Private mModoRibbon As eRibbonMode
Private mVisible As Boolean

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

Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Friend Property Let Visible(Value As Boolean)
    mVisible = Value
End Property

Public Function IsAdminGroupVisible() As Boolean
    IsAdminGroupVisible = (mModoRibbon = Ribbon_Admin)
End Function

' ==================================================================
' METODOS DE UTILIDAD PARA T_RibbonBookState (estado por libro)
' ==================================================================
' Estos métodos operan sobre el UDT sin modificar el estado interno de la clase.
' Permiten evaluar el estado del ribbon para cualquier libro.

'@Description: Obtiene descripción del modo para un estado de libro específico
Public Function GetDescriptionForBook(bookState As T_RibbonBookState) As String
    Select Case bookState.Modo
    Case Ribbon_Admin
        GetDescriptionForBook = "Admin"
    Case Ribbon_Hidden
        GetDescriptionForBook = "(oculto)"
    Case Ribbon_OpportunityOnly
        GetDescriptionForBook = "Fich. oportunidad"
    Case Ribbon_User
        GetDescriptionForBook = "User"
    Case Else
        GetDescriptionForBook = "Fich. oportunidad"
    End Select
End Function

'@Description: Verifica si el grupo admin es visible para un estado de libro específico
Public Function IsAdminGroupVisibleForBook(bookState As T_RibbonBookState) As Boolean
    IsAdminGroupVisibleForBook = (bookState.Modo = Ribbon_Admin)
End Function

'@Description: Verifica si el ribbon debe estar visible para un estado de libro específico
Public Function IsVisibleForBook(bookState As T_RibbonBookState, isOpportunityFile As Boolean) As Boolean
    If Not bookState.Visible Then
        IsVisibleForBook = False
        Exit Function
    End If

    Select Case bookState.Modo
    Case Ribbon_Admin, Ribbon_User
        IsVisibleForBook = True
    Case Ribbon_OpportunityOnly
        IsVisibleForBook = isOpportunityFile
    Case Ribbon_Hidden
        IsVisibleForBook = False
    Case Else
        IsVisibleForBook = isOpportunityFile
    End Select
End Function

'@Description: Sincroniza el estado interno con un T_RibbonBookState
'@Note: Útil cuando se cambia de libro activo
Public Sub SyncFromBookState(bookState As T_RibbonBookState)
    mModoRibbon = bookState.Modo
    mVisible = bookState.Visible
End Sub

'@Description: Obtiene el estado actual como T_RibbonBookState
Public Function ToBookState() As T_RibbonBookState
    ToBookState.Modo = mModoRibbon
    ToBookState.Visible = mVisible
End Function

