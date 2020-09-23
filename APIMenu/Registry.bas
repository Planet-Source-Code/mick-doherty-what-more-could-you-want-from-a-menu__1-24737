Attribute VB_Name = "Registry"
Option Explicit

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, _
        lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal lpData As String, _
        ByVal cbData As Long) As Long
        
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        phkResult As Long) As Long

Public Const HKCU = &H80000001
Const REG_SZ = 1

Public Sub WriteRegString(Group As Long, Section As String, Key As String, NewString As String)

    Dim lKeyID As Long
    
    RegCreateKey Group, Section, lKeyID
    If lKeyID = 0 Then
        Exit Sub
    End If
    If Len(NewString) = 0 Then
        RegSetValueEx lKeyID, Key, 0&, REG_SZ, 0&, 0&
    Else
        RegSetValueEx lKeyID, Key, 0&, REG_SZ, NewString, Len(NewString)
    End If
    
End Sub

Public Function ReadRegString(Group As Long, Section As String, Key As String) As String
    
    Dim lRslt As Long, lKeyID As Long, lBufferSize As Long
    Dim sKeyValue As String
    
    RegCreateKey Group, Section, lKeyID
    If lKeyID = 0 Then
        ReadRegString = Empty
        Exit Function
    End If
    lRslt = RegQueryValueEx(lKeyID, Key, 0&, REG_SZ, 0&, lBufferSize)
    If lBufferSize < 2 Then
        ReadRegString = Empty
        Exit Function
    End If
    
    sKeyValue = Space(lBufferSize + 1)
    lRslt = RegQueryValueEx(lKeyID, Key, 0&, REG_SZ, ByVal sKeyValue, lBufferSize)
    ReadRegString = Left(sKeyValue, lBufferSize - 1)
    
End Function
