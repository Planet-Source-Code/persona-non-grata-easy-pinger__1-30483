VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Ping Prg ©2001 Danne R"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Net Send >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox image1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'Kein
      FillColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      ScaleHeight     =   48
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   48
      TabIndex        =   5
      Top             =   720
      Width           =   720
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ping"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   315
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Enter the message here:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Host name, IP or URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Status: unknown"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Host IP:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rolled As Boolean
Private Sub Command1_Click()
Dim echo As ICMP_ECHO_REPLY, tmp As String
If Len(Text1.Text) > 0 Then
If Not IsNumeric(Text1.Text) Then
   tmp = AddrByName(Text1.Text)
End If
Call Ping(tmp, echo)
Label1.Caption = "Host IP: " & tmp
g = echo.RoundTripTime
If g = 589844 Then
Beep
Label2.Caption = Estado_host
image1.Cls
image1.Picture = LoadPicture(App.Path & "\10.gif")
Else

Label2.Caption = "Accesstime: " & echo.RoundTripTime & " ms"
image1.Cls
image1.Picture = LoadPicture(App.Path & "\11.gif")
Dim tp As Integer
If InStrB(1, Text1.Text, ".", vbTextCompare) < 1 Then
Command2.Enabled = True
End If
End If
End If
End Sub

Private Sub Command2_Click()
If rolled = False Then
Command2.Caption = "Net Send <<"
rolled = True
Text2.SetFocus
Me.Height = 3960
Else
rolled = False
Command2.Caption = "Net Send >>"
Text2.Text = ""
Text1.SetFocus
Me.Height = 1920
End If
End Sub

Private Sub Command3_Click()
If Len(Text2.Text) > 0 Then
msg = Text2.Text
Shell "net send " & Text1.Text & " " & msg, vbHide
Text2.Text = ""
Me.Height = 1920
End If

End Sub

Private Sub Form_Activate()
Call SetWindowPos(Me.hwnd, -1, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW)
rolled = False
End Sub

Private Sub Form_Load()
IP_Initialize

End Sub
Private Sub Form_Unload(Cancel As Integer)
WSACleanup
End Sub

Private Sub Text1_Change()
Command2.Enabled = False
rolled = False
Command2.Caption = "Net Send >>"
Text2.Text = ""
Text1.SetFocus
Me.Height = 1920
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Command1_Click
End If
End Sub

