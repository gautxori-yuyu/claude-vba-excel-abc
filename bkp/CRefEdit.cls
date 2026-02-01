VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "2-Infraestructura.Excel"
Option Explicit

Private Const MODULE_NAME As String = "CRefEdit"

Private WithEvents oTextBox As MSForms.TextBox
Attribute oTextBox.VB_VarHelpID = -1
Private oUF As Object
Private bEnableKeyBoardInput As Boolean


Event OnTextBoxChange(ByVal CurrentTextBox As MSForms.TextBox)
Private sClassInstanceName As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                          Public Subs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub TextBoxToRefEdit(ByVal ClassInstanceName As String, ByVal TextBox As Object, Optional EnableKeyBoardInput As Boolean)
    sClassInstanceName = ClassInstanceName
    Set oUF = GetUserForm(TextBox)
    oUF.tag = IsFormModal(oUF)
    bEnableKeyBoardInput = EnableKeyBoardInput
    Set oTextBox = TextBox
    TextBox.DropButtonStyle = fmDropButtonStyleReduce
    TextBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
End Sub

Private Sub oTextBox_Change()
    Dim oThis As CRefEdit
    Set oThis = Me
    Call CallByName(oUF, sClassInstanceName, VbSet, oThis)
    Set oThis = Nothing
    RaiseEvent OnTextBoxChange(oTextBox)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                       Object Event Subs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'click the drop button to select the pointer
Private Sub oTextBox_DropButtonClick()
    Call StoreTextboxWidth(oTextBox)
    Call ShowForm(oUF, False)
    Call ShowRefEdit(bEnableKeyBoardInput)
    Call ShowForm(oUF, True)
End Sub

'Hit F2 to get into the box
Private Sub oTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = VBA.vbKeyF2 Then Call oTextBox_DropButtonClick ' Gives the ability to F2 into a function.
    If Not bEnableKeyBoardInput Then KeyCode = 0
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                      Misc Private Subs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetUserForm(ByVal Ctrl As MSForms.control) As Object
    Dim oTmp As Object
    Set oTmp = Ctrl.Parent
    Do While TypeOf oTmp Is MSForms.control
        Set oTmp = oTmp.Parent
    Loop
    Set GetUserForm = oTmp
End Function

Private Sub Class_Terminate()
    'Debug.Print "Class instance treminated and memory released properly related to: " & oTextBox.Name
    Set oTextBox = Nothing:  Set oUF = Nothing
End Sub

