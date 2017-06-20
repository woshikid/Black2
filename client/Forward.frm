VERSION 5.00
Begin VB.Form Forward 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Forward"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
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
   ScaleHeight     =   1455
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Remote to Local"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Local to Remote"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3120
      MaxLength       =   6
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Enable"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Port:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Destination:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Source Port:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Port Forwarding"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Forward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label5_Click()
    On Error Resume Next
    If Label5.Caption = "Enable" Then
        Client.TCP.CloseSck 1
        Client.tunnelConnected = False
        Client.tunnelData = vbNullString
        Err.Clear
        If Option1.Value = True Then
            Client.tunnelListen = True
            Client.TCP.LocalPort(1) = Val(Text1.Text)
            Client.TCP.Listen 1
            Client.TCP.CloseSck 1
            If Err Then
                MessageBox Forward.hwnd, Err.Description, "Forward Faild", vbOKOnly
                Exit Sub
            End If
            Client.SendData ChrB(47) + ChrB(1) + ChrB(Val(Text3.Text) \ 256) + ChrB(Val(Text3.Text) Mod 256), Text2.Text
        Else
            Client.tunnelListen = False
            Client.TCP.RemoteHost(1) = Text2.Text
            Client.TCP.RemotePort(1) = Val(Text3.Text)
            Client.SendData ChrB(47) + ChrB(2) + ChrB(Val(Text1.Text) \ 256) + ChrB(Val(Text1.Text) Mod 256)
        End If
        Text1.Locked = True
        Text2.Locked = True
        Text3.Locked = True
        Option1.Enabled = False
        Option2.Enabled = False
        Label5.Visible = False
        Shape1.Visible = False
    Else
        Client.tunnelDisable
        Label5.Caption = "Enable"
        Text1.Locked = False
        Text2.Locked = False
        Text3.Locked = False
        Option1.Enabled = True
        Option2.Enabled = True
    End If
End Sub

Private Sub Label6_Click()
    On Error Resume Next
    Forward.Hide
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text1.Text)
    If i > 65535 Then i = 65535
    If i < 0 Then i = 0
    Text1.Text = i
    detecChange
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    detecChange
End Sub

Private Sub Text3_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text3.Text)
    If i > 65535 Then i = 65535
    If i < 0 Then i = 0
    Text3.Text = i
    detecChange
End Sub

Private Sub detecChange()
    On Error Resume Next
    If Text1.Text <> vbNullString And Text2.Text <> vbNullString And Text3.Text <> vbNullString Then
        Label5.Visible = True
        Shape1.Visible = True
    Else
        Label5.Visible = False
        Shape1.Visible = False
    End If
End Sub
