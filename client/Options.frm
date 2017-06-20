VERSION 5.00
Begin VB.Form Options 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " Options"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
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
   ScaleHeight     =   1890
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   12
      Text            =   "1000"
      Top             =   1080
      Width           =   480
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "1000"
      Top             =   720
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "1024"
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Camera Interval ms:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " WaveIn:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1485
      Width           =   735
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      Top             =   1485
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Get"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1485
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Desktop Interval ms:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   1890
      Left            =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3495
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
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Package Length:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    On Error Resume Next
    Client.SendData ChrB(74), Combo1.List(Combo1.ListIndex)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    Dim i As Long
    i = Val(Text1.Text)
    Client.SendData ChrB(20) + ChrB(i \ 256) + ChrB(i Mod 256)
    Client.packageLen = i
    Options.Hide
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    Options.Hide
End Sub

Private Sub Label6_Click()
    On Error Resume Next
    Dim i As Long
    i = Val(Text2.Text)
    Client.DesktopDelay.Interval = i
    Options.Hide
End Sub

Private Sub Label10_Click()
    On Error Resume Next
    Dim i As Long
    i = Val(Text3.Text)
    Client.CameraDelay.Interval = i
    Options.Hide
End Sub

Private Sub Label7_Click()
    On Error Resume Next
    Client.SendData ChrB(71)
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text1.Text)
    If i > 8192 Then i = 8192
    If i < 1 Then i = 1
    Text1.Text = i
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text2.Text)
    If i > 9999 Then i = 9999
    If i < 1 Then i = 1
    Text2.Text = i
End Sub

Private Sub Text3_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text3.Text)
    If i > 9999 Then i = 9999
    If i < 1 Then i = 1
    Text3.Text = i
End Sub
