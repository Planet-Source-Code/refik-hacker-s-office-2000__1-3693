VERSION 5.00
Begin VB.Form frmProgramCloser 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program Closer"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6360
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&End Selected Task"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "frmClose.frx":0000
      Left            =   120
      List            =   "frmClose.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Show Running Tasks"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   7
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmProgramCloser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindow Lib "user32" _
(ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib _
"user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" _
Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal _
lpString As String, ByVal cch As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Sub LoadTaskList()
Dim CurrWnd As Long
Dim Length As Long
Dim TaskName As String
Dim parent As Long

List1.Clear
CurrWnd = GetWindow(frmProgramCloser.hwnd, GW_HWNDFIRST)

While CurrWnd <> 0
parent = GetParent(CurrWnd)
Length = GetWindowTextLength(CurrWnd)
TaskName = Space$(Length + 1)
Length = GetWindowText(CurrWnd, TaskName, Length + 1)
TaskName = Left$(TaskName, Len(TaskName) - 1)

If Length > 0 Then
    If TaskName <> Me.Caption Then
        If TaskName <> "taskmon" Then
            List1.AddItem TaskName
        End If
    End If
End If
CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
DoEvents

Wend

End Sub

Private Sub Command1_Click()
LoadTaskList
End Sub

Private Sub Command2_Click()
On Error GoTo erlevel
Dim winHwnd As Long
Dim RetVal As Long
winHwnd = FindWindow(vbNullString, List1.Text)
Debug.Print winHwnd
If winHwnd <> 0 Then
RetVal = PostMessage(winHwnd, &H10, 0&, 0&)
If RetVal = 0 Then
MsgBox "Error posting message."
End If
Else: MsgBox List1.Text + " is not open."
End If
erlevel:
LoadTaskList
End Sub





Private Sub Form_Load()
stayontop Me
End Sub

Private Sub Timer1_Timer()
If List1.Text = "" Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If
End Sub

