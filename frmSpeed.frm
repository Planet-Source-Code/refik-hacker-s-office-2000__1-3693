VERSION 5.00
Begin VB.Form frmSpeed 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Speed Test"
   ClientHeight    =   4515
   ClientLeft      =   4605
   ClientTop       =   3135
   ClientWidth     =   5205
   Icon            =   "frmSpeed.frx":0000
   LinkTopic       =   "frm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmSpeed.frx":000C
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Test 5"
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   1215
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Test 4"
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   915
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Test 3"
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   615
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   2040
      Left            =   1230
      ScaleHeight     =   134
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   4
      Top             =   315
      Width           =   3825
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Test 2"
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   315
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   3120
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Test 1"
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   1230
      TabIndex        =   3
      ToolTipText     =   "Speed Time"
      Top             =   15
      Width           =   3825
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Clocks the Time to Run a Process"
      Top             =   1515
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Countx As Single
Private Sub cmd_Click(Index As Integer)
Dim i As Integer, h As Integer
On Error Resume Next
Screen.MousePointer = vbHourglass
Countx = 0
pic.Cls
lbl1.Caption = 0
lbl.Caption = 0
Timer1.Enabled = True
Select Case Index
Case 0 'Counting
For i = 0 To 1000
  lbl1.Caption = i
  DoEvents
Next
Case 1 'Drawing
For i = 0 To 255
  pic.Line (i, 0)-(i, pic.ScaleHeight), RGB(0, i, i)
  DoEvents
Next
Case 2 'Psets
For i = 0 To 255
  For h = 0 To pic.ScaleHeight
    pic.PSet (i, h), RGB(0, i, h)
  Next
  DoEvents
Next
Case 3 'Print Text
For i = 0 To pic.ScaleHeight
  For h = 0 To pic.ScaleWidth
    pic.CurrentX = h
    pic.CurrentY = i
    pic.ForeColor = RGB(h, 0, i)
    pic.Print "Hello"
  Next
  DoEvents
Next
Case 4 'Print Slower
For i = 0 To pic.ScaleHeight
  For h = 0 To pic.ScaleWidth
    pic.CurrentX = h
    pic.CurrentY = i
    pic.ForeColor = RGB(h, 0, i)
    pic.Print "Hello"
  Next
  DoEvents
Next
For i = 0 To pic.ScaleHeight
  For h = 0 To pic.ScaleWidth
    pic.CurrentX = h
    pic.CurrentY = i
    pic.Circle (h, i), 5, Rnd * vbRed
  Next
  DoEvents
Next
End Select
lbl.Caption = Countx * 0.1
Timer1.Enabled = False
Screen.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
Countx = Countx + 1
lbl.Caption = Countx * 0.1
End Sub
