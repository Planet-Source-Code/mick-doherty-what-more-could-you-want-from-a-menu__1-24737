Attribute VB_Name = "APITimer"
Option Explicit

Declare Function SetTimer Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
        
Declare Function KillTimer Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal nIDEvent As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Integer

Const ID_TIMER = 50000

Public Function StartTimer(hwnd As Long, milliSecond As Single) As Long
    SetTimer hwnd, ID_TIMER, milliSecond, AddressOf TimerProc
End Function

Public Sub EndTimer(hwnd As Long)
    KillTimer hwnd, ID_TIMER
End Sub

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    Static bClick As Boolean
    
    On Error Resume Next
    'This is how we will check for clicks in the disabled menus.
    'The most significant bit of  the GetAsyncKeyState return is set
    'if the current state of the button is pressed. The least significant
    'bit is set if it has been pressed since the last call. We only want to
    'know if it is pressed while over the menuitem.
    If (GetAsyncKeyState(vbLeftButton) \ 100) And &HFF& <> 0 Then
         'This procedure will test for button down every time it times-out, so
         'we will only call MenuClickNoDismiss() the first time we see the button down.
         If bClick Then
            MenuClickNoDismiss 'PopMenu MenuClickNoDismiss()
            'set our button down flag
            bClick = False
        End If
    Else
        'and reset when we release the button.
        bClick = True
    End If

End Sub
