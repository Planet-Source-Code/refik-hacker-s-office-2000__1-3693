VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "E-Mail "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Status:"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   960
      TabIndex        =   15
      Top             =   3480
      Width           =   5175
      Begin VB.Label StatusTxt 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.TextBox txtEmailServer 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox ToNametxt 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtEmailBodyOfMessage 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1920
      Width           =   6855
   End
   Begin VB.TextBox txtEmailSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtToEmailAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtFromEmailAddress 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send E-Mail"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "E-Mail Server"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "There Name"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Your Name"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Subject"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "To"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "From (e-mail address)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single



Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.Protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.Connect ' Start connection
    
    WaitFor ("220")
    
    StatusTxt.Caption = "Connecting...."
    StatusTxt.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    StatusTxt.Caption = "Connected"
    StatusTxt.Refresh

    Winsock1.SendData (first)

    StatusTxt.Caption = "Sending Message"
    StatusTxt.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    StatusTxt.Caption = "Disconnecting"
    StatusTxt.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = "" ' Sent response code to blank **IMPORTANT**
End Sub


Private Sub Command1_Click()
    SendEmail txtEmailServer.Text, txtFromName.Text, txtFromEmailAddress.Text, txtToEmailAddress.Text, txtToEmailAddress.Text, txtEmailSubject.Text, txtEmailBodyOfMessage.Text
    'MsgBox ("Mail Sent")
    StatusTxt.Caption = "Mail Sent"
    StatusTxt.Refresh
    Beep
    
    Close
End Sub

Private Sub Command2_Click()
    
Unload Me
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData Response ' Check for incoming response *IMPORTANT*

End Sub
