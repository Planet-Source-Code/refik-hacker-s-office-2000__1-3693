VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hacker`s Office 2000"
   ClientHeight    =   720
   ClientLeft      =   3225
   ClientTop       =   480
   ClientWidth     =   6360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image10 
      Height          =   480
      Left            =   5760
      Picture         =   "Form1.frx":08CA
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":1194
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   4440
      Picture         =   "Form1.frx":1A5E
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   3720
      Picture         =   "Form1.frx":2328
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   3000
      Picture         =   "Form1.frx":2BF2
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2280
      Picture         =   "Form1.frx":34BC
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1560
      Picture         =   "Form1.frx":3D86
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   840
      Picture         =   "Form1.frx":4650
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":4F1A
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuTrayMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuTraySize 
         Caption         =   "&Size"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Mi&nimize"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Public Sub SetTrayMenuItems(window_state As Integer)
    Select Case window_state
        Case vbMinimized
            mnuTrayMinimize.Enabled = False
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        Case vbNormal
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = True
            mnuTrayRestore.Enabled = False
            mnuTraySize.Enabled = True
    End Select
End Sub

Private Sub Form_Load()
    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    AddToTray Me, mnuTray
    
    SetTrayTip "VB Helper tray icon program"
End Sub

' Enable the correct tray menu items.
Private Sub Form_Resize()
    SetTrayMenuItems WindowState
    
    If WindowState <> vbMinimized Then _
        LastState = WindowState
End Sub
' Important! Remove the tray icon.
Private Sub Form_Unload(Cancel As Integer)
    RemoveFromTray
End Sub

Private Sub Image1_Click()
Show
frmMail.Show
End Sub



Private Sub Image10_Click()
Show
frmProgramCloser.Show
End Sub

Private Sub Image2_Click()
Show
frmArchiv.Show
End Sub

Private Sub Image3_Click()
Show
frmPortScanner.Show
End Sub

Private Sub Image4_Click()
Show
frmCDplayer.Show
End Sub

Private Sub Image5_Click()
Show
frmWeb.Show
End Sub

Private Sub Image6_Click()
Show
frmViewer.Show
End Sub

Private Sub Image7_Click()
Show
frmSpeed.Show
End Sub

Private Sub Image8_Click()
Show
frmHtml.Show
End Sub

Private Sub mnuTrayClose_Click()
    Unload Me
End Sub

Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
End Sub


Private Sub mnuTrayMove_Click()
    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_MOVE, 0&
End Sub


Private Sub mnuTrayRestore_Click()
    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_RESTORE, 0&
End Sub


Private Sub mnuTraySize_Click()
    SendMessage hwnd, WM_SYSCOMMAND, _
        SC_SIZE, 0&
End Sub
