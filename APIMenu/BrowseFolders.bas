Attribute VB_Name = "BrowseFolders"
Option Explicit

Declare Function SHBrowseForFolder Lib "shell32" _
        (lpbi As BrowseInfo) As Long

Declare Function SHGetPathFromIDList Lib "shell32" _
        (ByVal pidList As Long, _
        ByVal lpBuffer As String) As Long

Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
        (ByVal hwnd As Long, _
        ByVal Msg As Long, _
        wParam As Any, _
        lParam As Any) As Long
        
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Declare Function PathAppend Lib "Shlwapi" Alias "PathAppendA" _
        (ByVal sPath As String, _
        ByVal sMore As String) As Boolean
        
Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" _
        (ByVal pszPath As String)

Const BIF_RETURNONLYFSDIRS = &H1
Const BIF_STATUSTEXT = &H4

Const BFFM_ENABLEOK = &H465
Const BFFM_SETSELECTION = &H466

Const BFFM_INITIALIZED = 1
Const BFFM_SELCHANGED = 2

Const MAX_PATH = 260

Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallBack As Long
    lParam As Long
    iImage As Integer
End Type

Public StartFolder As String
Public CurrentSelection As String * MAX_PATH

Public Function FolderBrowse() As String

    Dim BI As BrowseInfo
    Dim lRslt As Long
    Dim strReturn As String

    With BI
        .hWndOwner = ghw
        .lpszTitle = "Choose a folder with Bitmaps."
        .ulFlags = BIF_RETURNONLYFSDIRS
        .lpfnCallBack = BrowseProc(AddressOf BrowseCallbackProc)
    End With

    lRslt = SHBrowseForFolder(BI)

    If lRslt Then
        lRslt = SHGetPathFromIDList(lRslt, CurrentSelection)
        StartFolder = Left(CurrentSelection, InStr(CurrentSelection, vbNullChar) - 1)
        FolderBrowse = StartFolder
        CurrentFile = FolderPics(0)
        CurrentPaper = StartFolder & "\" & CurrentFile
        AppForm.picSelected.Picture = LoadPicture(CurrentPaper)
    Else
        'Cancel pressed
        FolderBrowse = StartFolder
        FilterPath StartFolder, "*.bmp"
    End If
        
    CoTaskMemFree lRslt

End Function

Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    
    On Error Resume Next
    
    Dim lRslt As Long, bRslt As Boolean, lbuffer As String * MAX_PATH
    
    'Initially disable the OK button
    SendMessage hwnd, BFFM_ENABLEOK, 0, ByVal False
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            'set the start directory
            If StartFolder = "" Then
                StartFolder = Left(CurrentPaper, Len(CurrentPaper) - Len(CurrentFile))
            End If
        
            SendMessage hwnd, BFFM_SETSELECTION, 0, ByVal StartFolder
        
        Case BFFM_SELCHANGED
            lRslt = SHGetPathFromIDList(lParam, CurrentSelection)
            If lRslt <> 0 Then
                bRslt = FilterPath(CurrentSelection, "*.bmp")
                'enable the OK button if the folder contains bitmaps
                If bRslt Then
                    SendMessage hwnd, BFFM_ENABLEOK, 0, ByVal True
                End If
            Else
                BrowseCallbackProc = 1
                CoTaskMemFree lRslt
                Exit Function
            End If
            CoTaskMemFree lRslt
        
    End Select

    BrowseCallbackProc = 0

End Function

Function BrowseProc(ByVal lParam As Long) As Long
    BrowseProc = lParam
End Function

Function FilterPath(ByVal sPath As String, sExt As String) As Boolean
    
    Dim sRslt As String
    sPath = sPath & (vbNullChar + Space(5)) 'pad sPath so we can append
    PathAppend sPath, sExt 'it with extension
    
    sRslt = Dir(sPath)
    
    If Len(sRslt) > 0 Then
        FilterPath = True
        GetPictures sPath
    End If
    
End Function

Function GetPictures(sPath As String)
    'fill FolderPics() array with bmp filenames in selected folder
    Dim sFile As String, i As Integer
    
    Erase FolderPics
    
    sFile = Dir(sPath)
    
    While sFile > ""
        DoEvents
        ReDim Preserve FolderPics(i)
        FolderPics(i) = sFile
        i = i + 1
        sFile = Dir
    Wend

End Function
