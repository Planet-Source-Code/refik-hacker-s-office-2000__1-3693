VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmWeb 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Web Browser"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Home"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Forward"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8415
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   12135
      ExtentX         =   21405
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go"
      Height          =   255
      Left            =   9000
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8655
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.Navigate = Text1.Text
End Sub

Private Sub Command2_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command3_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command4_Click()
WebBrowser1.GoHome
End Sub
