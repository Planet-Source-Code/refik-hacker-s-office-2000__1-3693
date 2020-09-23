VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPortScanner 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Port Scanner"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Stop"
      Height          =   255
      Left            =   5520
      TabIndex        =   42
      Top             =   840
      Width           =   2895
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   2880
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Port Watcher"
      Height          =   1815
      Left            =   -15
      TabIndex        =   37
      Top             =   9840
      Width           =   8295
      Begin VB.CommandButton Command12 
         Caption         =   "Save to Log"
         Height          =   285
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Clear"
         Height          =   285
         Left            =   1200
         TabIndex        =   38
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtPortWatch 
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   5
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtIncoming 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   1455
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Release"
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Watch"
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPortWatch 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   5
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   2400
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Scan &All Possible Ports!"
      Height          =   255
      Left            =   5505
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save to &Log"
      Height          =   255
      Left            =   7080
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   5520
      TabIndex        =   35
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   7680
   End
   Begin VB.Frame Frame2 
      Caption         =   "Timed Scans"
      Height          =   1695
      Left            =   -15
      TabIndex        =   32
      Top             =   8040
      Width           =   3495
      Begin VB.OptionButton Option1 
         Caption         =   "Selected Ports Only"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Multiple Range"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Timer Off"
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Timer On"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtSeconds 
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "60"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Scan Ports Every"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Selected 
      Caption         =   "Selected Ports Only"
      Height          =   1695
      Left            =   3585
      TabIndex        =   29
      Top             =   8040
      Width           =   4695
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   11
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   10
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   9
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   8
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   7
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   6
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   4
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   3
         Left            =   240
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   5
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   5
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Scan Now"
         Height          =   615
         Left            =   2400
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtSelPort 
         Height          =   285
         Index           =   1
         Left            =   240
         MaxLength       =   5
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Multiple Range"
      Height          =   1935
      Left            =   4680
      TabIndex        =   28
      Top             =   7680
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtStop 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtStart 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Stop Port"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Start Port"
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   7095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   $"PortScan.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   6135
      Left            =   5520
      TabIndex        =   43
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port Control    Coded by: Dustin Davis    Bootleg Software Inc.    Http://www.warpnet.org/bsi"
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   105
      TabIndex        =   40
      Top             =   9600
      Width           =   8295
   End
End
Attribute VB_Name = "frmPortScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'Port Control
'Coded by Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'Please use this code to learn from, do not steal it. It took several hours to
'complete this damn thing. Do not just compile it and label it as your own.
'What kind of programmer would you be? You wouldnt! You'd be a thief. If you
'Use any of this code, please give credit where its due, Me! I hope you learn from this
'
' - Dustin Davis -
'
'**************************************************************************************

Public ScanOnOff As Boolean
Public TimerOnOff As Boolean

Private Sub Command1_Click()
ScanOnOff = True
If txtStart.Text = "" Then
    MsgBox "You Must Enter a Starting Port", vbExclamation, "HEY!"
    Exit Sub
ElseIf txtStop.Text = "" Then
    MsgBox "You Must Enter a Stoping Port", vbExclamation, "HEY!"
    Exit Sub
Else
    Call scan_multiple
End If
End Sub

Private Sub Command10_Click()
Winsock2.Close
Winsock3.Close
AddText txtIncoming, vbCrLf & Time & " Released Port: " & Winsock2.LocalPort
AddText txtIncoming, vbCrLf & Time & " Released Port: " & Winsock3.LocalPort
End Sub

Private Sub Command11_Click()
txtIncoming.Text = ""
End Sub

Private Sub Command12_Click()
Open "received.log" For Output As #1
    Write #1, txtIncoming.Text
Close #1
AddText txtIncoming, vbCrLf & Time & " Log Saved!"
End Sub

Private Sub Command2_Click()
If txtSelPort(0).Text = "" Then
    MsgBox "Must Enter at least ONE (1) port to scan", vbExclamation, "HEY!"
Else
    Call scan_selected
End If
End Sub

Private Sub Command3_Click()
TimerOnOff = True
Command4.Enabled = True
Timer1.interval = txtSeconds.Text * 1000
AddText Text1, vbCrLf & Time & "Timer Activated"
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
TimerOnOff = False
Command3.Enabled = True
Timer1.interval = txtSeconds.Text * 1000
AddText Text1, vbCrLf & Time & "Timer Deactivated"
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
ScanOnOff = False
AddText Text1, vbCrLf & Time & " Scan Stoped by User"
End Sub

Private Sub Command6_Click()
Text1.Text = ""
End Sub

Private Sub Command7_Click()
Open "portcontrol.log" For Output As #1
    Write #1, Text1.Text
Close #1
AddText Text1, vbCrLf & Time & " Log Saved!"
End Sub

Private Sub Command8_Click()
ScanOnOff = True
Call scan_all
End Sub

Private Sub Command9_Click()
On Error GoTo errors
If Not txtPortWatch(0).Text = "" Then
    Winsock2.LocalPort = txtPortWatch(0).Text
    Winsock2.Listen
    AddText txtIncoming, vbCrLf & Time & " Watching Port " & txtPortWatch(0).Text
Else
    txtPortWatch(0).Text = ""
End If

If Not txtPortWatch(1).Text = "" Then
    Winsock3.LocalPort = txtPortWatch(1).Text
    Winsock3.Listen
    AddText txtIncoming, vbCrLf & Time & " Watching Port " & txtPortWatch(1).Text
Else
    txtPortWatch(1).Text = ""
End If


errors:
    If Err.Number = 10048 Then
        AddText txtIncoming, vbCrLf & Time & " Port(s) already in use!"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
Text1.Text = Time & " Port Scanner " & "Hacker`s Office 2000" & vbCrLf & vbCrLf
Text1.FontBold = True
Text1.ForeColor = &HFFFFF
txtIncoming.FontBold = True
Option1(0).Value = True
Command4.Enabled = False
Label6.FontBold = True
Label6.FontSize = 8
TimerOnOff = False
End Sub

Private Sub Timer1_Timer()
If TimerOnOff = True Then
    If Option1(0).Value = True Then
        ScanOnOff = True
            If txtStart.Text = "" Then
                MsgBox "You Must Enter a Starting Port", vbExclamation, "HEY!"
            ElseIf txtStop.Text = "" Then
                MsgBox "You Must Enter a Stoping Port", vbExclamation, "HEY!"
            Else
            Call scan_multiple
            End If
        Exit Sub
    ElseIf Option1(0).Value = False Then
        scan_selected
        Exit Sub
    End If
ElseIf TimerOnOff = False Then
    Exit Sub
End If
End Sub

Function AddText(textcontrol As Object, text2add As String)

'This function was obtained from Planet-source-code.com

    On Error GoTo errhandlr
    tmptxt$ = textcontrol.Text 'just in Case of an accident
    textcontrol.SelStart = Len(textcontrol.Text) ' move the "cursor" to the End of the text file
    textcontrol.SelLength = 0 ' highlight nothing (this becomes the selected text)
    textcontrol.SelText = text2add ' set the selected text ot text2add
    AddText = 1
    GoTo quitt ' goto the End of the Sub
    'error handlers
errhandlr:


    If Err.Number <> 438 Then 'check the Error number and restore the
        textcontrol.Text = tmptxt$ 'original text If the control supports it
    End If

    AddText = 0
    GoTo quitt
quitt:
    tmptxt$ = ""
End Function

Public Function scan_multiple()
Dim intStart As Long
Dim intStop As Long
intStart = txtStart.Text
intStop = txtStop.Text

On Error GoTo errors

AddText Text1, vbCrLf & Time & " Starting Scan from " & txtStart.Text & " to " & txtStop.Text
intStop = intStop + 1
intStart = intStart - 1

Do
    DoEvents
     intStart = intStart + 1
    If ScanOnOff = True Then
     Winsock1.Close
     DoEvents
     Label5.Caption = "Scanning: " & intStart
     DoEvents
     Winsock1.LocalPort = intStart
     DoEvents
     Winsock1.Listen
     DoEvents
    ElseIf ScanOnOff = False Then
        Exit Function
    ElseIf intStart >= intStop Then
        Exit Function
    End If
    DoEvents
Loop Until intStart >= intStop
AddText Text1, vbCrLf & Time & " Scan Done!"

errors:
    If Err.Number = 10048 Then
        AddText Text1, vbCrLf & Time & " Port " & Winsock1.LocalPort & " is in Use!"
        DoEvents
        Resume Next
    End If
End Function

Public Function scan_selected()

On Error GoTo errors

If txtSelPort(0).Text = "" Then
    txtSelPort(0).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(0).Text
    Winsock1.Listen
End If

If txtSelPort(1).Text = "" Then
    txtSelPort(1).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(1).Text
    Winsock1.Listen
End If

If txtSelPort(2).Text = "" Then
    txtSelPort(2).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(2).Text
    Winsock1.Listen
End If

If txtSelPort(3).Text = "" Then
    txtSelPort(3).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(3).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(4).Text = "" Then
    txtSelPort(4).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(4).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(5).Text = "" Then
    txtSelPort(5).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(5).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(6).Text = "" Then
    txtSelPort(6).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(6).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(7).Text = "" Then
    txtSelPort(7).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(7).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(8).Text = "" Then
    txtSelPort(8).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(8).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(9).Text = "" Then
    txtSelPort(9).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(9).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(10).Text = "" Then
    txtSelPort(10).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(10).Text
    Winsock1.Listen
    DoEvents
End If

If txtSelPort(11).Text = "" Then
    txtSelPort(11).Text = ""
Else
    Winsock1.Close
    Winsock1.LocalPort = txtSelPort(11).Text
    Winsock1.Listen
    DoEvents
End If

errors:
    If Err.Number = 10048 Then
        AddText Text1, vbCrLf & Time & " Port " & Winsock1.LocalPort & " is in Use!"
        DoEvents
        Resume Next
    End If
End Function

Public Function scan_all()
Dim inStart As Long
On Error GoTo errors

AddText Text1, vbCrLf & Time & " Scanning All Possible Ports" 'From 1 - 65530

Do
    DoEvents
     intStart = intStart + 1
    If ScanOnOff = True Then
        Winsock1.Close
        DoEvents
        Label5.Caption = "Scanning: " & intStart
        DoEvents
        Winsock1.LocalPort = intStart
        DoEvents
        Winsock1.Listen
        DoEvents
    ElseIf ScanOnOff = False Then
        Exit Function
    End If
    DoEvents
Loop Until inStart >= 65530

AddText Text1, vbCrLf & Time & "Scan Done!"

errors:
    If Err.Number = 10048 Then
        AddText Text1, vbCrLf & Time & " Port " & Winsock1.LocalPort & " is in Use!"
        DoEvents
        Resume Next
    End If
End Function

Private Sub Winsock2_Close()
AddText Text1, vbCrLf & Time & " Port " & Winsock2.RemotePort & " Is no longer blocked"
End Sub

Private Sub Winsock2_Connect()
AddText Text1, vbCrLf & Time & " Blocking Port " & Winsock2.RemotePort
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept requestID
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock2.GetData Data
AddText txtIncoming, vbCrLf & Time & " " & Data
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
AddText Text1, vbCrLf & Time & " Socket Error: " & Number & vbCrLf & "Error Description: " & Description
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Winsock3.Close
Winsock3.Accept requestID
End Sub

