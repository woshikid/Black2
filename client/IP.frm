VERSION 5.00
Begin VB.Form IP 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "       Connect to"
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   3120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Connect to"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Remote Host:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "IP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Text1.Text = Client.TCP.LocalIP(3)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    IP.Hide
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    Dim i As Long
    i = InStr(1, Text1.Text, ":")
    Client.TCP.CloseSck 0
    If i > 0 Then
        Client.TCP.RemotePort(0) = Val(Mid(Text1.Text, i + 1))
        Client.TCP.RemoteHost(0) = Left(Text1.Text, i - 1)
    Else
        Client.TCP.RemotePort(0) = 24 'server port
        Client.TCP.RemoteHost(0) = Text1.Text 'server ip
    End If
    'show state
    Client.Label2.Caption = "Black Client - Connecting to " & Text1.Text & "..."
    Client.TCP.Connect 0
    IP.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        Label4_Click
    End If
End Sub
