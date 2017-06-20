VERSION 5.00
Begin VB.Form MidServer 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Explorer"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
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
   ScaleHeight     =   2775
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      ItemData        =   "Mid.frx":0000
      Left            =   2520
      List            =   "Mid.frx":0002
      TabIndex        =   4
      Top             =   600
      Width           =   1590
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      ItemData        =   "Mid.frx":0004
      Left            =   960
      List            =   "Mid.frx":0006
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      ItemData        =   "Mid.frx":0008
      Left            =   120
      List            =   "Mid.frx":000A
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   240
      Left            =   3840
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   240
      Left            =   2880
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Mid Server:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4200
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " MidServer"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "MidServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()
    On Error Resume Next
    If Label2.Caption = "Connect" Then
        Dim i As Long
        i = InStr(1, Text1.Text, ":")
        Client.TCP.CloseSck 2
        If i > 0 Then
            Client.TCP.RemotePort(2) = Val(Mid(Text1.Text, i + 1))
            Client.TCP.RemoteHost(2) = Left(Text1.Text, i - 1)
        Else
            Client.TCP.RemotePort(2) = 25 'mid server port
            Client.TCP.RemoteHost(2) = Text1.Text 'mid server ip
        End If
        Client.TCP.Connect 2
    Else
        Client.TCP_CloseSck 2
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    MidServer.Hide
End Sub

Private Sub Label5_Click()
    On Error Resume Next
    List1.Clear
    List2.Clear
    List3.Clear
    Client.SendDataA ChrB(2)
End Sub

Private Sub List1_Click()
    On Error Resume Next
    List2.ListIndex = List1.ListIndex
    List3.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    Dim id As Long
    id = Val(List1.List(List1.ListIndex))
    If id > 0 Then Client.SendDataA ChrB(1) + ChrB(id \ 256) + ChrB(id Mod 256)
    MidServer.Hide
End Sub

Private Sub List2_Click()
    On Error Resume Next
    List1.ListIndex = List2.ListIndex
    List3.ListIndex = List2.ListIndex
End Sub

Private Sub List2_DblClick()
    On Error Resume Next
    Dim id As Long
    id = Val(List1.List(List2.ListIndex))
    If id > 0 Then Client.SendDataA ChrB(1) + ChrB(id \ 256) + ChrB(id Mod 256)
    MidServer.Hide
End Sub

Private Sub List3_Click()
    On Error Resume Next
    List1.ListIndex = List3.ListIndex
    List2.ListIndex = List3.ListIndex
End Sub

Private Sub List3_DblClick()
    On Error Resume Next
    Dim id As Long
    id = Val(List1.List(List3.ListIndex))
    If id > 0 Then Client.SendDataA ChrB(1) + ChrB(id \ 256) + ChrB(id Mod 256)
    MidServer.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        Label2_Click
    End If
End Sub
