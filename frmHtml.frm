VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHtml 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Download and Display HTML"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6615
   Icon            =   "frmHtml.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox URL 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "http://www.sympadmin.8m.com"
      Top             =   3360
      Width           =   4695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmHtml.frx":0442
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   $"frmHtml.frx":050B
      ForeColor       =   &H0000FF00&
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "URL:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3390
      Width           =   375
   End
End
Attribute VB_Name = "frmHtml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written exclusively for VB Center by Marco Cordero.

Private Sub Command1_Click()
    
    Dim txt As String
    Dim b() As Byte
    
    On Error GoTo ErrorHandler
    
    
    Command1.Enabled = False
    
    ' This opens the file specified in the URL text box
    b() = Inet1.OpenURL(URL.Text, 1)
    
    txt = ""
    
    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    ' This loads the opened file into the RichTextBox control
    RichTextBox1.Text = txt
    
    Command1.Enabled = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "The document you requested could not be found.", vbCritical

    Exit Sub

End Sub
