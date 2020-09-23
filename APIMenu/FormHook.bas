Attribute VB_Name = "FormHook"
Option Explicit

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
        ByVal hwnd As Long, _
        ByVal Msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

Const GWL_WNDPROC = -4

Const WM_COMMAND = &H111
Public Const WM_DRAWITEM = &H2B
Const WM_INITMENU = &H116
Const WM_INITMENUPOPUP = &H117
Const WM_MEASUREITEM = &H2C
Const WM_MENUSELECT = &H11F
Const WM_PAINT = &HF
Const WM_SYSCOMMAND = &H112

Global lpPrevWndProc As Long
Global ghw As Long

Public AppForm As Form

Public Sub Hook(FRM As Form)
    Set AppForm = FRM
    ghw = FRM.hwnd
    lpPrevWndProc = SetWindowLong(ghw, GWL_WNDPROC, AddressOf WindowProc)
    MenuCreate       'PopMenu MenuCreate()
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(ghw, GWL_WNDPROC, lpPrevWndProc)
    DestroyMenu hFileMenu
End Sub

Function WindowProc(ByVal hwnd As Long, _
            ByVal uMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long

    On Error Resume Next
    
    Select Case uMsg
        
        Case WM_MENUSELECT
            MenuSel = CLng(wParam And &HFFFF&) 'LoWord of wParam here is MenuID
            MenuSelect MenuSel   'PopMenu MenuSelect()
            WindowProc = 0
        
        Case WM_MEASUREITEM 'lParam here is a pointer to a MeasureItemStruct
            MeasureMenu lParam  'PopMenu MeasureMenu()
            WindowProc = 0
        
        Case WM_DRAWITEM    'lParam here is a pointer to a DrawItemStruct
            DrawMenu lParam     'PopMenu DrawMenu()
            WindowProc = 0
            
        Case WM_PAINT
            'set the forms menu
            SetMenu ghw, hFormMenu
            WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
            
        Case WM_INITMENU
            If CLng(wParam And &HFFFF&) = hFormMenu Then
                'Set the up and down arrow colours
                ColourArrows    'Graphics ColourArrows
            End If
            WindowProc = 0

        Case WM_COMMAND
            'When we activate the PopupMenu through the forms menu we catch the
            'clicks here
            MenuClick (wParam)    'PopMenu MenuClick()
            WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
            
        Case WM_SYSCOMMAND
            'Menu Items in the system menu are caught here
            MenuClick (wParam)    'PopMenu MenuClick()
            WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
        
        Case Else
            WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    
    End Select

End Function
