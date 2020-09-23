Attribute VB_Name = "PopMenu"
Option Explicit

Declare Function CreateMenu Lib "user32" () As Long

Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function SetMenu Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hMenu As Long) As Long

Declare Function DrawMenuBar Lib "user32" _
        (ByVal hwnd As Long) As Long

Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpNewItem As Any) As Long

Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, _
        ByVal lpString As Any) As Long

Declare Function TrackPopupMenu Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal wFlags As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nReserved As Long, _
        ByVal hwnd As Long, _
        ByVal lprc As Any) As Long
        
Public Declare Function DestroyMenu Lib "user32" _
        (ByVal hMenu As Long) As Long
        
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function SetRect Lib "user32" (lpRect As RECT, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Declare Function OffsetRect Lib "user32" _
        (lpRect As RECT, _
        ByVal x As Long, _
        ByVal y As Long) As Long

Declare Function DrawCaption Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal hdc As Long, _
        pcRect As RECT, _
        ByVal un As Long) As Long
        
Declare Function GetMenuItemRect Lib "user32" _
        (ByVal hwnd As Long, ByVal hMenu As Long, _
        ByVal uItem As Long, _
        lprcItem As RECT) As Long

Declare Function CheckMenuRadioItem Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal IDFirst As Long, _
        ByVal IDLast As Long, _
        ByVal IDSelected As Long, _
        ByVal uFlags As Long) As Long
        
Declare Function SetMenuItemBitmaps Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long, _
        ByVal hBitmapUnchecked As Long, _
        ByVal hBitmapChecked As Long) As Long

Declare Function SetMenuDefaultItem Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal uItem As Long, _
        ByVal fByPos As Long) As Long
        
Declare Function GetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long) As Long
        
Declare Function SetPixel Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal crColor As Long) As Long

Declare Function FillRect Lib "user32" _
        (ByVal hdc As Long, _
        lpRect As RECT, _
        ByVal hBrush As Long) As Long
        
Declare Function StretchBlt Lib "gdi32" _
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
        ByVal dwRop As Long) As Long

Declare Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long

Declare Function WindowFromDC Lib "user32" _
        (ByVal hdc As Long) As Long

Declare Function GetDC Lib "user32" _
        (ByVal hwnd As Long) As Long
        
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, _
        ByVal uParam As Long, _
        ByVal lpvParam As Any, _
        ByVal fuWinIni As Long) As Long

Declare Function GetSystemMenu Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal bRevert As Long) As Long
        
Declare Function GetMenuItemID Lib "user32" _
        (ByVal hMenu As Long, _
        ByVal nPos As Long) As Long

Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" _
        (ByVal hMenu As Long, _
        ByVal un As Long, _
        ByVal bool As Boolean, _
        lpcMenuItemInfo As MENUITEMINFO) As Long

Declare Function DrawEdge Lib "user32" _
        (ByVal hdc As Long, _
        qrc As RECT, _
        ByVal edge As Long, _
        ByVal grfFlags As Long) As Long

Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hdc As Long, _
        ByVal lpStr As String, _
        ByVal nCount As Long, _
        lpRect As RECT, _
        ByVal wFormat As Long) As Long

Declare Function SetTextColor Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal crColor As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type

Public Type PointAPI
    x As Long
    y As Long
End Type

Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hStyleMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Const MIIM_ID = &H2
Const MIIM_STATE = &H1
Const MIIM_TYPE = &H10

Const MF_CHECKED = &H8&
Const MF_DISABLED = &H2&
Const MF_ENABLED = &H0&
Const MF_MENUBREAK = &H40&
Const MF_OWNERDRAW = &H100&
Const MF_POPUP = &H10&
Const MF_SEPARATOR = &H800&
Const MF_STRING = &H0&

Const MFS_GRAYED = &H3&

Const TPM_RETURNCMD = &H100&

Const ODA_DRAWENTIRE = &H1

Const ODT_MENU = 1

Const ODS_DISABLED = &H4
Const ODS_SELECTED = &H1
Const CUSTOM_SELECTED = &H5

Const DC_GRADIENT = &H20
Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_SMALLCAP = &H2
Const DC_TEXT = &H8

Const COLOR_HIGHLIGHT = 13
Const COLOR_MENU = 4

Const HKEY_CURRENT_USER = &H80000001
Const REG_SZ = 1

Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1

Const DT_CENTER = &H1
Const DT_SINGLELINE = &H20
Const DT_VCENTER = &H4

Const BDR_SUNKENOUTER = &H2
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public hFileMenu As Long, hStyleMenu As Long, hFormMenu As Long, hHelpMenu, hSysMenu As Long
Public MP As PointAPI, sMenu As Long
Public mnuHeight As Single
Public hwndMenu As Long
Public MenuSel As Long
Public bmpFolder As StdPicture
Public bmpCheck As StdPicture
Public bmpExit As StdPicture
Public bmpAbout As StdPicture
Public bmpStyle As StdPicture
Public FolderPics() As String
Public picMenuRect As RECT
Public MaskColour As Long, sPath As String * 260
Public CurrentPaper As String, CurrentFile As String

Public Sub MeasureMenu(ByRef lP As Long)
    
    'It would appear that you cannot actually get measurements here,
    'you can only set them. There are no measurements until after the
    'Menu is drawn, but you only get a WM_MEASUREITEM message before the
    'initial WM_DRAWITEM, so maybe this message should have been named
    'WM_SETINITIALITEMSIZE.
    
    Dim MIS As MEASUREITEMSTRUCT
    'Load MIS with structure in memory
    CopyMemory MIS, ByVal lP, Len(MIS)
        
    Select Case MIS.itemID
        Case 1000
            MIS.itemHeight = 180
            MIS.itemWidth = 3   '(18 - 3) - 12. I don't know where the 12 comes
            'from, but there always seems to be 12 pixels more than I want.
            '18 - 3 is Small Titlebar height offset by 3 pixels to bring it closer
            'to the Menu Items.
        Case 1001
            MIS.itemWidth = 3
            
        Case 1402
            MIS.itemHeight = 126
            MIS.itemWidth = 136
        Case Else
            MIS.itemHeight = 17
    End Select
    'Return the updated structure
    CopyMemory ByVal lP, MIS, Len(MIS)
    
End Sub

Public Sub DrawMenu(ByRef lP As Long)

    Dim DIS As DRAWITEMSTRUCT, rct As RECT, lRslt As Long

    CopyMemory DIS, ByVal lP, Len(DIS)
    
    Select Case DIS.itemID
    '----SideBar--------------------------------------------------------------------
        Case 1000
            With frmAbout
                'since we can't measure in the MeasureMenu sub we'll do it here.
                'we cannot just get the sidebar height as it will only return
                'the height of an empty menu item. (i.e. 13). Maybe we can get the
                'height of the whole menu with some other API call that I don't know
                'about. GetWindowRect() combined with WindowFromDC() works after
                'we have drawn the menu, which is a bit too late.
                Dim i As Integer
                If mnuHeight = 0 Then
                    For i = 1 To 10
                        GetMenuItemRect ghw, hFileMenu, i, rct
                        mnuHeight = mnuHeight + (rct.Bottom - rct.Top)
                    Next i
                End If
                
                'set the size of our sidebar
                SetRect rct, 0, 0, mnuHeight + 1, 18
                'change the caption of the hidden frmAbout so that we can change
                'the sidebar text.
                .Caption = CurrentFile
                'This is a bit of a copout, but it works.
                'You could always use GradientFillRect and then draw rotated text
                'straight onto the sidebar, but this is much easier.
                'Draw a form caption onto our hidden frmAbout, the length of our
                'menu height. frmAbout must be set to autoredraw.
                DrawCaption .hwnd, .hdc, rct, DC_SMALLCAP Or DC_ACTIVE Or DC_TEXT Or DC_GRADIENT

                Dim x As Single, y As Single
                Dim ncolor As Long

                'rotate our caption through 270 degrees
                'and paint onto menu
                For x = 0 To mnuHeight
                    For y = 0 To 17
                        ncolor = GetPixel(.hdc, x, y)
                        SetPixel DIS.hdc, y, mnuHeight - x, ncolor
                    Next y
                Next x
                'that rotation was simple.
                'I don't know why the msdn article was so complex.

                'remove the caption picture from frmAbout
                .Cls
            
            End With
    '-------------------------------------------------------------------------------
    
    '----System Menu SideBar--------------------------------------------------------
        Case 1001
            'I'm not doing any fancy work here just filling the sidebar by making
            'the rect longer than the system menu to show basic Drawing to SysMenu
            Dim tRect As RECT
            tRect = DIS.rcItem
            tRect.Right = tRect.Right + 3
            tRect.Bottom = tRect.Bottom + 200
            FillRect DIS.hdc, tRect, 28
    '-------------------------------------------------------------------------------
            
    '----ICQ Style Seperator--------------------------------------------------------
        Case 1002
            'Use GetTextExtent
            Dim rctLine As RECT, strSep As String
            strSep = " System Menu "
            rctLine.Left = DIS.rcItem.Left + 1
            rctLine.Right = DIS.rcItem.Right - 1
            rctLine.Top = DIS.rcItem.Top + ((DIS.rcItem.Bottom - DIS.rcItem.Top) / 2)
            rctLine.Bottom = rctLine.Top + 2
'            SetTextColor DIS.hdc, &HFF0000
            DrawEdge DIS.hdc, rctLine, BDR_SUNKENOUTER, BF_RECT
            DrawText DIS.hdc, strSep, Len(strSep), DIS.rcItem, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    '-------------------------------------------------------------------------------
    
    '----Up Menu--------------------------------------------------------------------
        Case 1401
            'we have to draw the highlight state as well as the normal state
            If DIS.itemState = CUSTOM_SELECTED Then
                'The hBrush paramater in FillRect does not need to
                'call GetSysColorBrush. Just add 1 to the COLOR_CONSTANT
                FillRect DIS.hdc, DIS.rcItem, COLOR_HIGHLIGHT + 1
                DrawArrow DIS.hdc, DIS.rcItem, True, True   'Graphics DrawArrow()
            Else
                FillRect DIS.hdc, DIS.rcItem, COLOR_MENU + 1
                DrawArrow DIS.hdc, DIS.rcItem, True, False  'Graphics DrawArrow()
            End If
            
            'Get handle to menuwindow so we can find its DC when we need it.
            hwndMenu = WindowFromDC(DIS.hdc)
    '-------------------------------------------------------------------------------
    
    '----Picture Menu---------------------------------------------------------------
        Case 1402
            picMenuRect = DIS.rcItem
            With picMenuRect
                StretchBlt DIS.hdc, .Left + 3, .Top + 3, 160, 120, AppForm.picSelected.hdc, 0, 0, AppForm.picSelected.Width, AppForm.picSelected.Height, vbSrcCopy
            End With
    '-------------------------------------------------------------------------------
    
    '----Down Menu------------------------------------------------------------------
        Case 1403
            If DIS.itemState = CUSTOM_SELECTED Then
                FillRect DIS.hdc, DIS.rcItem, COLOR_HIGHLIGHT + 1
                DrawArrow DIS.hdc, DIS.rcItem, False, True  'Graphics DrawArrow()
            Else
                FillRect DIS.hdc, DIS.rcItem, COLOR_MENU + 1
                DrawArrow DIS.hdc, DIS.rcItem, False, False 'Graphics DrawArrow()
            End If
            
            hwndMenu = WindowFromDC(DIS.hdc)
    '-------------------------------------------------------------------------------
    
        Case Else
            'do nothing
    End Select
    
    CopyMemory ByVal lP, DIS, Len(DIS)

End Sub

Public Sub MenuCreate()
    
    'Get the system menu
    hSysMenu = GetSystemMenu(ghw, False)
    
    'create the menu bases
    hFormMenu = CreateMenu()
    hFileMenu = CreatePopupMenu()
    hStyleMenu = CreatePopupMenu()
    hHelpMenu = CreatePopupMenu()
    
    '----Form Menu-----------------------------------------------------------------
    AppendMenu hFormMenu, MF_STRING Or MF_POPUP, hFileMenu, "File"
    AppendMenu hFormMenu, MF_STRING Or MF_POPUP, hHelpMenu, "Help"
    '------------------------------------------------------------------------------
    
    '----PopUp\File-Sub Menu-------------------------------------------------------
    AppendMenu hFileMenu, MF_OWNERDRAW Or MF_DISABLED, 1000, 0& 'SideBar
    'we don't want the menu to close when we click up or down so we will disable
    'those menuitems
    AppendMenu hFileMenu, MF_OWNERDRAW Or MF_DISABLED Or MF_MENUBREAK, 1401, 0&  'up
    AppendMenu hFileMenu, MF_OWNERDRAW, 1402, 0&  'picture menu
    AppendMenu hFileMenu, MF_OWNERDRAW Or MF_DISABLED, 1403, 0&  'down
    AppendMenu hFileMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hFileMenu, MF_STRING, 1500, "Set Wallpaper Folder"
    AppendMenu hFileMenu, MF_POPUP, hStyleMenu, "Set Wallpaper Style"
    AppendMenu hFileMenu, 0&, 1700, "About"
    AppendMenu hFileMenu, MF_SEPARATOR, 0&, 0&
    AppendMenu hFileMenu, 0&, 2000, "Exit"
    '------------------------------------------------------------------------------
    
    '----Style Sub Menu------------------------------------------------------------
    AppendMenu hStyleMenu, MF_STRING, 1601, "Centre"
    AppendMenu hStyleMenu, MF_STRING, 1602, "Tile"
    AppendMenu hStyleMenu, MF_STRING, 1603, "Stretch"
    '------------------------------------------------------------------------------
    
    '----Help Sub Menu-------------------------------------------------------------
    AppendMenu hHelpMenu, 0&, 11000, "Not a Lot of help Here"
    '------------------------------------------------------------------------------
    
    '----Sytem Menu----------------------------------------------------------------
    'I don't know what MS have lined up the Shortcut label with on the Close Menu,
    'but it's obviously not the Right hand edge as Alt+F4 gets cut short
    Dim MII As MENUITEMINFO, lpMII As MENUITEMINFO, sysMenuID As Long
    sysMenuID = GetMenuItemID(hSysMenu, 0)
    
    lpMII.fMask = MIIM_ID Or MIIM_TYPE
    lpMII.cbSize = Len(MII)
    lpMII.wID = 1002
    lpMII.fType = MF_MENUBREAK Or MF_OWNERDRAW
    'bool set to False so un is MenuID
    InsertMenuItem hSysMenu, sysMenuID, False, lpMII
    
    MII.fMask = MIIM_ID Or MIIM_TYPE Or MIIM_STATE
    MII.cbSize = Len(MII)
    MII.wID = 1001
    MII.fType = MF_OWNERDRAW
    MII.fState = MFS_GRAYED 'This is the only way to disable a system menuitem.
    'bool set to True so un is Menu position
    InsertMenuItem hSysMenu, 0, True, MII
    '------------------------------------------------------------------------------
    
    '----Set Menu Item pictures----------------------------------------------------
    Set bmpFolder = LoadResPicture(102, 0)
    Set bmpCheck = LoadResPicture(103, 0)
    Set bmpExit = LoadResPicture(104, 0)
    Set bmpAbout = LoadResPicture(105, 0)
    Set bmpStyle = LoadResPicture(106, 0)
    '------------------------------------------------------------------------------
    
    '----Assign Menu Item Pictures---------------------------------------------------------
    'Set MenuItemBitmaps is really meant for setting custom checkmarks and should
    'use monochrome pictures, but if you're not too bothered about the look of the
    'pictures, then this method is much simpler than OwnerDraw.
    SetMenuItemBitmaps hFileMenu, 1500, 0&, bmpFolder, bmpFolder
    SetMenuItemBitmaps hFileMenu, hStyleMenu, 0&, bmpStyle, bmpStyle
    SetMenuItemBitmaps hFileMenu, 1700, 0&, bmpAbout, bmpAbout
    SetMenuItemBitmaps hFileMenu, 2000, 0&, bmpExit, bmpExit
    '------------------------------------------------------------------------------
    
    '----Custom CheckMarks---------------------------------------------------------
    Dim i As Integer
    'If you omit this loop the menu will use Dots for these
    'menuitems as we will be selecting them via CheckMenuRadioItem.
    'The black parts of this picture will automatically become COLOR_MENUTEXT
    'and the White parts COLOR_MENU.
    For i = 1601 To 1603
        SetMenuItemBitmaps hFileMenu, i, 0&, 0&, bmpCheck
    Next i
    '------------------------------------------------------------------------------
    
    '----Set a Default Item--------------------------------------------------------
    'The text in this menu will be Bold. I do not know why I could not set this
    'Menu bold with the MF_DEFAULT flag when I created it.
    SetMenuDefaultItem hFileMenu, 1700, 0&
    '------------------------------------------------------------------------------
    
    '----Get and Set Current Style setting-----------------------------------------
    Dim iTile As Integer, iStyle As Integer
    iTile = ReadRegString(HKCU, "Control Panel\Desktop", "TileWallpaper") 'Registry ReadRegString()
    iStyle = ReadRegString(HKCU, "Control Panel\Desktop", "WallpaperStyle")
    CheckMenuRadioItem hFileMenu, 1601, 1603, 1601 + iTile + iStyle, 0&
    '----and Current Wallpaper-----------------------------------------------------
    CurrentPaper = ReadRegString(HKCU, "Control Panel\Desktop", "Wallpaper") 'Registry ReadRegString()
    If UCase(CurrentPaper) = "NONE" Or CurrentPaper = "" Then Exit Sub
    AppForm.picSelected.Picture = LoadPicture(CurrentPaper)
    CurrentFile = CurrentPaper
    PathStripPath CurrentFile
    CurrentFile = Left(CurrentFile, InStr(CurrentFile, vbNullChar) - 1)
    frmAbout.Caption = CurrentFile
    StartFolder = Left(CurrentPaper, Len(CurrentPaper) - Len(CurrentFile))
    FilterPath StartFolder, "*.bmp"
    '------------------------------------------------------------------------------
    
 End Sub

Public Sub MenuTrack()
    
    '----Set the Arrow Colours before displaying the menu--------------------------
    ColourArrows    'Graphics ColourArrows
    '------------------------------------------------------------------------------
    
    GetCursorPos MP
    
    '----display and monitor the menu----------------------------------------------
    sMenu = TrackPopupMenu(hFileMenu, TPM_RETURNCMD, MP.x, MP.y, 0, ghw, 0&)
        
    If sMenu <> 0 Then
        MenuClick (sMenu)
    End If

End Sub

Public Sub MenuClick(mnuID)
    
    Select Case mnuID
        Case 2000
            '----Exit Menu
            EndTimer ghw
            UnHook
            Unload AppForm
            Exit Sub
            
        Case 1402
            '----Picture Menu
            SetPaper CurrentPaper
                        
        Case 1500
            '----Folder Select Menu
            FolderBrowse 'BrowseFolders FolderBrowse()
            
        Case 1700
            '----About Menu
            frmAbout.Show 1
            
        Case 1601 To 1603
            '----Option Menus
            CheckMenuRadioItem hFileMenu, 1601, 1603, mnuID, 0&
            SetStyle CLng(mnuID)
        
        Case 11000
            '----Help Submenu
            MsgBox "Sorry! You have found a problem that help can't help you with." _
            & vbCrLf & "Please contact your program vendor for more help."
            
        Case Else
            Exit Sub
            
    End Select

End Sub

Public Sub MenuSelect(mnuID As Long)
    'called from FormHook WindowProc()
    Select Case mnuID
        Case 1401, 1403
            StartTimer ghw, 10 'APITimer StartTimer()
        
        Case Else
            EndTimer ghw   'APITimer EndTimer()
    End Select

End Sub

Public Sub MenuClickNoDismiss()
    'called from APITimer TimerProc()
    'note that these menuitems have been set MF_DISABLED.
    Static i As Integer

    If UBound(FolderPics) < 0 Then Exit Sub
    
    Select Case MenuSel
        Case 1401
            If i < UBound(FolderPics) Then
                i = i + 1
            End If
        
        Case 1403
            If i > 0 Then
                i = i - 1
            End If
        
        Case Else
            Exit Sub
            
    End Select
    
    '----Force a Redraw of the SideBar----------------------------------------------
    'This was a bit of a Headache.  Once windows has drawn to a DC
    'it makes it available for other processes. When we send a WM_DRAWITEM message,
    'we need to send the DC to draw to along with it. Have you ever tried to get the
    'DC of a menu? I was under the impression that the handle returned by CreateMenu
    'was the menu's Windowhandle. As it turned out it wasn't. In the DrawMenu
    'routine I got the windowhandle of the menu from it's DC (which we know at the
    'time of drawing) and stored it in hwndMenu. From this handle we can get the DC.
    Dim DIS As DRAWITEMSTRUCT
    
    CurrentFile = FolderPics(i)
    
    DIS.itemID = 1000
    DIS.hdc = GetDC(hwndMenu)
    
    SendMessage ghw, WM_DRAWITEM, ByVal hFileMenu, DIS
    
    '----and Picture Menu-----------------------------------------------------------
    sPath = StartFolder & vbNullChar 'Add a Null terminator to the path so that
    PathAppend sPath, FolderPics(i)  'Pathappend can replace the Null Char with a
    CurrentPaper = sPath             'backslash (if needed) and filename.
    
    AppForm.picSelected.Picture = LoadPicture(sPath)
    DIS.itemID = 1402
    DIS.rcItem = picMenuRect
    DIS.hdc = GetDC(hwndMenu)
    
    SendMessage ghw, WM_DRAWITEM, ByVal hFileMenu, DIS
    '-------------------------------------------------------------------------------
    
End Sub

Sub SetStyle(ID As Long)
    
    Dim iTile As String, iStyle As String
    
    Select Case ID - 1601
        Case 0
            'Centre
            iTile = "0"
            iStyle = "0"
        Case 1
            'Tile
            iTile = "1"
            iStyle = "0"
        Case 2
            'Stretch
            iTile = "0"
            iStyle = "2"
    End Select

    WriteRegString HKCU, "Control Panel\Desktop", "TileWallpaper", iTile 'Registry WriteRegString()
    WriteRegString HKCU, "Control Panel\Desktop", "WallpaperStyle", iStyle
    SetPaper CurrentPaper
    
End Sub

Sub SetPaper(sFileName As String)
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0&, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE
End Sub
