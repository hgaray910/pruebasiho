Attribute VB_Name = "modScroll"
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
                
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115

Dim PrevProc As Long

Public Sub HookForm(Scroll As Object)
    PrevProc = SetWindowLong(Scroll.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm(Scroll As Object)
    SetWindowLong Scroll.hWnd, GWL_WNDPROC, PrevProc
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
        
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            SendMessage hWnd, WM_VSCROLL, ByVal 1, ByVal 0
        Else
            SendMessage hWnd, WM_VSCROLL, ByVal 0, ByVal 0
        End If
    End If
End Function
