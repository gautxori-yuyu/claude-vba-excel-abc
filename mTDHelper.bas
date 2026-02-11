Attribute VB_Name = "mTDHelper"
'@Folder "4-UI.TaskDialog"
Option Explicit

Private Const MODULE_NAME As String = "mTDHelper"

'mTDHelper: Helper module for cTaskDialog.cls
'Must be included with the class.
Public Sub MagicalTDInitFunction()
        'The trick is a GENIUS!
    'He identified the bug in VBA64 that had been causing the crashing.
    'As if by magic, calling this from Class_Initialize resolves the problem.
End Sub
Public Function TaskDialogCallbackProc(ByVal hwnd As LongPtr, ByVal uNotification As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal lpRefData As cTaskDialog) As LongPtr
Attribute TaskDialogCallbackProc.VB_Description = "[mTDHelper] Task Dialog Callback Proc (función personalizada)"
Attribute TaskDialogCallbackProc.VB_ProcData.VB_Invoke_Func = " \n21"
TaskDialogCallbackProc = lpRefData.zz_ProcessCallback(hwnd, uNotification, wParam, lParam)
End Function
Public Function TaskDialogEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As cTaskDialog) As Long
Attribute TaskDialogEnumChildProc.VB_Description = "[mTDHelper] Task Dialog Enum Child Proc (función personalizada)"
Attribute TaskDialogEnumChildProc.VB_ProcData.VB_Invoke_Func = " \n21"
TaskDialogEnumChildProc = lParam.zz_ProcessEnumCallback(hwnd)
End Function
Public Function TaskDialogSubclassProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As cTaskDialog) As LongPtr
Attribute TaskDialogSubclassProc.VB_Description = "[mTDHelper] Task Dialog Subclass Proc (función personalizada)"
Attribute TaskDialogSubclassProc.VB_ProcData.VB_Invoke_Func = " \n21"
TaskDialogSubclassProc = dwRefData.zz_ProcessSubclass(hwnd, uMsg, wParam, lParam, uIdSubclass)
End Function

'@Description("Muestra un cuadro de diálogo de pregunta (Sí/No) usando TaskDialog.")
Public Function ShowTaskDialogYesNo(ByVal Title As String, ByVal instruction As String, ByVal Content As String) As TDRESULT
Attribute ShowTaskDialogYesNo.VB_Description = "[mTDHelper] Muestra un cuadro de diálogo de pregunta (Sí/No) usando TaskDialog."
Attribute ShowTaskDialogYesNo.VB_ProcData.VB_Invoke_Func = " \n21"
    On Error GoTo ErrHandler

    Dim TaskDlg As cTaskDialog, res As TDRESULT
    Set TaskDlg = New cTaskDialog
    ShowTaskDialogYesNo = TaskDlg.SimpleDialog(instruction, TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON, Title, Content, IDI_QUESTION, Application.hwnd)
    
    Exit Function
ErrHandler:
    ' Fallback to MsgBox if TaskDialog fails
    LogError MODULE_NAME, "[ShowTaskDialogYesNo] Fallback to MsgBox", Err.Number, Err.Description
    Dim result As VbMsgBoxResult
    ShowTaskDialogYesNo = MsgBox(instruction & vbCrLf & Content, vbYesNo + vbQuestion, Title)
End Function


'@Description("Muestra un cuadro de diálogo de error estándar usando TaskDialog.")
Public Sub ShowTaskDialogError(ByVal Title As String, ByVal instruction As String, ByVal Content As String)
    On Error GoTo ErrHandler
    
    Dim TaskDlg As cTaskDialog, res As TDRESULT
    Set TaskDlg = New cTaskDialog
    res = TaskDlg.SimpleDialog(Content, TDCBF_OK_BUTTON, Title, instruction, IDI_ERROR, Application.hwnd)
        
    Exit Sub
ErrHandler:
    ' Fallback to MsgBox if TaskDialog fails for any reason
    LogError MODULE_NAME, "[ShowTaskDialogError] Fallback to MsgBox", Err.Number, Err.Description
    MsgBox instruction & vbCrLf & Content, vbCritical, Title
End Sub

