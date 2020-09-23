Attribute VB_Name = "Graphics"
Option Explicit

Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Declare Function TransparentBlt Lib "msimg32.dll" _
        (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal nSrcWidth As Long, _
        ByVal nSrcHeight As Long, _
        ByVal crTransparent As Long) As Boolean

Const COLOR_HIGHLIGHTTEXT = 14
Const COLOR_MENUTEXT = 7
    
Public Sub ColourArrows()

    Dim x As Single, y As Single, txtClr As Long, hltClr As Long
    
    'First get the colours to use
    txtClr = GetSysColor(COLOR_MENUTEXT)
    hltClr = GetSysColor(COLOR_HIGHLIGHTTEXT)
    
    'don't forget that the maskcolour may be the same as the
    'Menutext or Highlight colour.
    If txtClr <> &HFF00FF And hltClr <> &HFF00FF Then
        MaskColour = &HFF00FF
    Else
        MaskColour = txtClr And hltClr / 100
    End If
    
    'picArrows is made up of 3 9X9 squares
    'we will use the first one as a template
    With AppForm.picArrows
        For x = 0 To 8
            For y = 0 To 8
                'The black pixels are part of the arrow
                If GetPixel(.hdc, x, y) = vbBlack Then
                    SetPixel .hdc, x + 9, y, txtClr
                    SetPixel .hdc, x + 18, y, hltClr
                Else
                    'the white pixels are the mask
                    SetPixel .hdc, x + 9, y, MaskColour
                    SetPixel .hdc, x + 18, y, MaskColour
                End If
            Next y
        Next x
    End With

End Sub

Public Sub DrawArrow(DestDC As Long, rctDest As RECT, Up As Boolean, Highlight As Boolean)

    Dim DestX As Integer, DestY As Integer, SrcX As Integer

    'Set Draw Coordinates
    DestX = rctDest.Left + ((rctDest.Right - rctDest.Left) / 2) - 6
    DestY = rctDest.Top + 6

    'Set Forecolor
    If Highlight Then
        SrcX = 18
    Else
        SrcX = 9
    End If

    Select Case Up
        Case True
            'draw up arrow
            TransparentBlt DestDC, DestX, DestY, 9, 5, AppForm.picArrows.hdc, SrcX, 0, 9, 5, MaskColour
        Case Else
            'draw down arrow
            TransparentBlt DestDC, DestX, DestY, 9, 5, AppForm.picArrows.hdc, SrcX, 4, 9, 5, MaskColour
    End Select

End Sub

