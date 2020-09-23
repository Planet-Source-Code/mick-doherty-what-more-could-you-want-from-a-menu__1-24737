VERSION 5.00
Begin VB.Form frmPopupMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "Titled Menu Demo"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   3375
   Icon            =   "frmPopupMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picArrows 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   1
      Top             =   0
      Width           =   405
   End
   Begin VB.PictureBox picSelected 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   60
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   12000
   End
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.ScaleMode = vbPixels ' API works in Pixels
    picArrows.Picture = LoadResPicture(101, 0)
    Hook Me    'FormHook Hook()
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        MenuTrack  'PopMenu MenuTrack()
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim FRM As Form
    UnHook     'FormHook UnHook()
    DestroyMenu hFileMenu
    For Each FRM In Forms
        Unload FRM
    Next
End Sub
