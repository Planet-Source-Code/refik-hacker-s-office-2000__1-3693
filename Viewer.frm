VERSION 5.00
Begin VB.Form frmViewer 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Viewer"
   ClientHeight    =   6135
   ClientLeft      =   1125
   ClientTop       =   1455
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6135
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Pis 
      BackColor       =   &H00000000&
      Caption         =   " Picture  "
      ForeColor       =   &H0000FF00&
      Height          =   5895
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   10455
      Begin VB.Image FileImage 
         Height          =   855
         Left            =   120
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox PatternCombo 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5760
      Width           =   2175
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.DirListBox DirList 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.FileListBox FileList 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   2175
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
    'On Error GoTo DriveError
    DirList.Path = DriveList.Drive
    Exit Sub

DriveError:
    DriveList.Drive = DirList.Path
    Exit Sub
End Sub


Private Sub FileList_Click()
Dim fname As String

    On Error GoTo LoadPictureError

    fname = FileList.Path + "\" + FileList.FileName
    Caption = "Viewer [" & fname & "]"
    
    MousePointer = vbHourglass
    DoEvents
    FileImage.Picture = LoadPicture(fname)
    MousePointer = vbDefault
    
    Exit Sub

LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "Viewer [Invalid picture]"
    Exit Sub
End Sub

Private Sub Form_Load()
    PatternCombo.AddItem "Graphic (*.gif;*.jpg;*.ico;*.bmp;*.wmf;*.dib)"
    PatternCombo.AddItem "Bitmaps (*.bmp)"
    PatternCombo.AddItem "GIF (*.gif)"
    PatternCombo.AddItem "JPEG (*.jpg)"
    PatternCombo.AddItem "Icons (*.ico)"
    PatternCombo.AddItem "Matafiles (*.wmf)"
    PatternCombo.AddItem "DIBs (*.dib)"
    PatternCombo.AddItem "All Files (*.*)"
    PatternCombo.ListIndex = 0
End Sub

Private Sub Form_Resize()
Const GAP = 60

Dim wid As Integer
Dim hgt As Integer

    If WindowState = vbMinimized Then Exit Sub

    wid = DriveList.Width
    DriveList.Move GAP, GAP, wid
    PatternCombo.Move GAP, ScaleHeight - PatternCombo.Height, wid
    
    hgt = (PatternCombo.Top - DriveList.Top - DriveList.Height - 3 * GAP) / 2
    If hgt < 100 Then hgt = 100
    DirList.Move GAP, DriveList.Top + DriveList.Height + GAP, wid, hgt
    FileList.Move GAP, DirList.Top + DirList.Height + GAP, wid, hgt
End Sub


Private Sub PatternCombo_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer

    pat = PatternCombo.List(PatternCombo.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    FileList.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub


