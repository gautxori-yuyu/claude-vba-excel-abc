Attribute VB_Name = "modUTILSRefEditAPI"
'@Folder "2-Infraestructura.Excel"
Option Explicit

Private Const MODULE_NAME As String = "modUTILSRefEditAPI"

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongPtr
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #End If
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hKBhook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hKBhook As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
    Private Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal fEnable As Long) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As LongPtr

    Private hCBTHook As LongPtr, hKBhook As LongPtr, lPrvWndProc As LongPtr, RefEditHwnd As LongPtr, hwndFrm As LongPtr
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hUf As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hKBhook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hKBhook As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private hCBTHook As Long, hKBhook As Long, lPrvWndProc As Long, RefEditHwnd As Long, hwndFrm As Long
#End If

Private dblTextboxwidth As Double
Private oTextBox As Object
Private sTextBoxText As String
Private bEnableKeyBoardInput As Boolean

'____________________________________________PUBLIC ROUTINES_________________________________________
Public Sub ShowForm(ByVal frm As Object, ByVal Show As Boolean)
Attribute ShowForm.VB_ProcData.VB_Invoke_Func = " \n0"

    Const GWL_EXSTYLE = (-20)
    Const WH_KEYBOARD = 2
    Const WS_EX_LAYERED = &H80000
    Const LWA_ALPHA = &H2&

    Call SetWindowLong(hwndFrm, GWL_EXSTYLE, GetWindowLong(hwndFrm, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwndFrm, 0, IIf(Show = False, 0, 255), LWA_ALPHA)
    Call SetActiveWindow(Application.hwnd)
    Call ShowWindow(hwndFrm, -CLng(Show))
    If frm.tag Then EnableWindow Application.hwnd, 0
    
    If Show = False Then
        If hKBhook = 0 Then
            hKBhook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, _
                                       GetModuleHandle(vbNullString), GetCurrentThreadId)
        End If
        ActiveWindow.RangeSelection.Cells(1).Select
    Else
        Call UnhookWindowsHookEx(hKBhook)
        hKBhook = 0
    End If

End Sub

Public Sub ShowRefEdit(ByVal EnableKeyBoardInput As Boolean)
Attribute ShowRefEdit.VB_ProcData.VB_Invoke_Func = " \n0"
 
    #If Win64 Then
        Dim lVBEhwnd As LongLong
    #Else
        Dim lVBEhwnd As Long
    #End If
    
    Const WH_CBT = 5
    Dim sBuffer As String
    Dim lRet As Long
    
    lVBEhwnd = FindWindow("wndclass_desked_gsk", vbNullString)
    Call ShowWindow(lVBEhwnd, 0)
    sBuffer = VBA.Space(256)
    lRet = GetWindowText(hwndFrm, sBuffer, 256)
    bEnableKeyBoardInput = EnableKeyBoardInput
    If hCBTHook = 0 Then
        hCBTHook = SetWindowsHookEx(WH_CBT, AddressOf HookProc, 0, GetCurrentThreadId)
    End If
    DoEvents
    Call Application.Dialogs(xlDialogGoalSeek).Show
    bEnableKeyBoardInput = False
    Call EnableWindow(Application.hwnd, True)
 
End Sub

Public Sub StoreTextboxWidth(ByVal TextBox As Object)
Attribute StoreTextboxWidth.VB_ProcData.VB_Invoke_Func = " \n0"
    Set oTextBox = TextBox
    dblTextboxwidth = TextBox.Width
End Sub

Public Function IsFormModal(frm As Object) As Boolean
Attribute IsFormModal.VB_Description = "[modUTILSRefEditAPI] Is Form Modal (función personalizada)"
Attribute IsFormModal.VB_ProcData.VB_Invoke_Func = " \n21"
    IsFormModal = Not CBool(SetFocus(Application.hwnd))
    Call IUnknown_GetWindow(frm, VarPtr(hwndFrm))
    Call SetFocus(hwndFrm)
End Function

'____________________________________________PRIVATE ROUTINES_________________________________________
Private Sub TerminateHook()
    Call UnhookWindowsHookEx(hCBTHook)
    Call EnableWindow(Application.hwnd, True)
    hCBTHook = 0
End Sub

#If Win64 Then
Private Function HookProc(ByVal idHook As Long, ByVal wParam As LongLong, ByVal lParam As LongLong) As LongLong
    Dim lp As LongLong
#Else
Private Function HookProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lp As Long
#End If
 
Const HCBT_ACTIVATE = 5
Const GWL_WNDPROC = -4
Const GWL_EXSTYLE = (-20)
Const WS_EX_CONTEXTHELP = &H400
Const SWP_SHOWWINDOW = &H40
Const SWP_NOACTIVATE = &H10
Const GW_CHILD = 5
Const MK_LBUTTON = &H1
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDOWN = &H201
 
Dim tFrmRect As RECT, tRefRect As RECT
Dim p1 As POINTAPI, p2 As POINTAPI
Dim sBuffer As String
Dim PixelPerInch As Single
Dim lRet As Long
    
If idHook = HCBT_ACTIVATE Then
    sBuffer = VBA.Space(256)
    lRet = GetClassName(wParam, sBuffer, 256)
    If VBA.Left(sBuffer, lRet) = "bosa_sdm_XL9" Then
        Call TerminateHook
        RefEditHwnd = GetWindow(wParam, GW_CHILD)
        Call GetWindowRect(hwndFrm, tFrmRect)
        Call GetWindowRect(RefEditHwnd, tRefRect)
        With tRefRect
            p1.x = .Left: p1.y = .Top
            p2.x = .Right + 15: p2.y = .Bottom
        End With
        Call ScreenToClient(wParam, p1)
        Call ScreenToClient(wParam, p2)
        lp = MakeLong_32_64(p2.x, p1.y)
        With tFrmRect
            Call SetWindowPos(wParam, -1, .Left, .Top, _
                              PTtoPX(dblTextboxwidth, False), 0, SWP_SHOWWINDOW + SWP_NOACTIVATE)
        End With
        Call SetWindowLong(wParam, GWL_EXSTYLE, _
                           GetWindowLong(wParam, GWL_EXSTYLE) And Not WS_EX_CONTEXTHELP)
        Call PostMessage(RefEditHwnd, WM_LBUTTONDOWN, MK_LBUTTON, lp)
        Call PostMessage(RefEditHwnd, WM_LBUTTONUP, MK_LBUTTON, lp)
        lPrvWndProc = SetWindowLong(wParam, GWL_WNDPROC, AddressOf CallBack)
    End If
End If
 
HookProc = CallNextHookEx(hCBTHook, idHook, ByVal wParam, ByVal lParam)
 
End Function

#If Win64 Then
Private Function KeyboardProc(ByVal ncode As Long, ByVal wParam As LongLong, ByVal lParam As LongLong) As LongLong
#Else
Private Function KeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Const HC_ACTION = 0
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101

If ncode = HC_ACTION Then
    If wParam = vbKeyEscape Or wParam = vbKeyReturn Then
        Call PostMessage(RefEditHwnd, WM_KEYDOWN, wParam, 0)
        Call PostMessage(RefEditHwnd, WM_KEYUP, wParam, 0)
        Call ShowWindow(RefEditHwnd, False)
        DoEvents
        Exit Function
    End If
    If bEnableKeyBoardInput = False Then
        KeyboardProc = -1
        Exit Function
    End If
End If

KeyboardProc = CallNextHookEx(hKBhook, ncode, wParam, lParam)
    
End Function

#If Win64 Then
Private Function CallBack(ByVal hwnd As LongLong, ByVal msg As Long, ByVal wParam As LongLong, ByVal lParam As LongLong) As LongLong
#Else
Private Function CallBack(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
 
Const GWL_WNDPROC = -4
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_SYSCOMMAND = &H112
Const WM_CLOSE As Long = &H10
Const SC_CLOSE = &HF060&
 
Dim sBuffer1 As String, sBuffer2 As String, sBuffer3 As String
Dim lRet1 As Long, lRet2 As Long, lRet3 As Long
     
sBuffer1 = VBA.Space(256): sBuffer2 = VBA.Space(256)
lRet1 = GetWindowText(RefEditHwnd, sBuffer1, 256): lRet2 = GetWindowText(hwnd, sBuffer2, 256)
   
If InStr(1, VBA.Left(sBuffer1, lRet1), "!") = 0 Then
    Call SetWindowText(RefEditHwnd, ActiveSheet.Name & "!" & VBA.Left(sBuffer1, lRet1))
End If
    
If VBA.Left(sBuffer2, lRet2) <> "RefEditEx" Then
    Call SetWindowText(hwnd, "RefEditEx")
End If

If GetAsyncKeyState(VBA.vbKeyReturn) Or GetAsyncKeyState(VBA.vbKeySeparator) Then
    sBuffer3 = VBA.Space(256)
    lRet3 = GetWindowText(RefEditHwnd, sBuffer3, 256)
    sTextBoxText = VBA.Left(sBuffer3, lRet3)
    Call SetTimer(Application.hwnd, 0, 100, AddressOf SelectRange)
    Call PostMessage(hwnd, WM_CLOSE, 0, 0)
End If
    
If GetAsyncKeyState(VBA.vbKeyEscape) Then
    Call PostMessage(hwnd, WM_CLOSE, 0, 0)
End If

Select Case msg
Case Is = WM_SYSCOMMAND
    If wParam = SC_CLOSE Then
        ShowWindow hwnd, 0
        Call SetActiveWindow(Application.hwnd)
        Call PostMessage(hwnd, WM_CLOSE, 0, 0)
        sTextBoxText = ""                        'VBA.Left(sBuffer1, lRet1)
        Call SetTimer(Application.hwnd, 0, 0, AddressOf SelectRange)
    End If
Case WM_LBUTTONDOWN, WM_LBUTTONUP
    Call SetActiveWindow(Application.hwnd)
    Call PostMessage(hwnd, WM_CLOSE, 0, 0)
    sTextBoxText = VBA.Left(sBuffer1, lRet1)
    Call SetTimer(Application.hwnd, 0, 0, AddressOf SelectRange)
Case Is = WM_CLOSE
    Call SetWindowLong(hwnd, GWL_WNDPROC, lPrvWndProc)
End Select

CallBack = CallWindowProc(lPrvWndProc, hwnd, msg, wParam, ByVal lParam)
 
End Function

Private Sub SelectRange()
    Call KillTimer(Application.hwnd, 0)
    On Error Resume Next
    Range(sTextBoxText).Select
    oTextBox.Text = sTextBoxText
    sTextBoxText = vbNullString
End Sub

#If Win64 Then
Function MakeLong_32_64(ByVal wLow As Long, ByVal wHigh As Long) As LongPtr
Attribute MakeLong_32_64.VB_ProcData.VB_Invoke_Func = " \n0"
    Dim retVal As LongLong, b(3) As Byte
    
    Call MoveMemory(ByVal VarPtr(b(0)), ByVal VarPtr(wLow), 4)
    Call MoveMemory(ByVal VarPtr(b(2)), ByVal VarPtr(wHigh), 4)
    Call MoveMemory(ByVal VarPtr(retVal), ByVal VarPtr(b(0)), 8)
    MakeLong_32_64 = retVal
#Else
Function MakeLong_32_64(ByVal wLow As Integer, ByVal wHigh As Integer) As Long
    Dim retVal As Long, b(3) As Byte
    
    Call MoveMemory(ByVal VarPtr(b(0)), ByVal VarPtr(wLow), 2)
    Call MoveMemory(ByVal VarPtr(b(2)), ByVal VarPtr(wHigh), 2)
    Call MoveMemory(ByVal VarPtr(retVal), ByVal VarPtr(b(0)), 4)
    MakeLong_32_64 = retVal
#End If
End Function

Private Function ScreenDPI(bVert As Boolean) As Long

    Const LOGPIXELSX = 88
    Const LOGPIXELSY = 90

    Static lDPI(1), lDC

    If lDPI(0) = 0 Then
        lDC = GetDC(0)
        lDPI(0) = GetDeviceCaps(lDC, LOGPIXELSX)
        lDPI(1) = GetDeviceCaps(lDC, LOGPIXELSY)
        lDC = ReleaseDC(0, lDC)
    End If
    ScreenDPI = lDPI(Abs(bVert))
    
End Function

Private Function PTtoPX(Points As Double, bVert As Boolean) As Long
    Const POINTS_PER_INCH = 72
    PTtoPX = Points * ScreenDPI(bVert) / POINTS_PER_INCH
End Function


