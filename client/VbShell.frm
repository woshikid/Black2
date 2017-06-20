VERSION 5.00
Begin VB.Form VbShell 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Shell"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
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
   ScaleHeight     =   765
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   390
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      Top             =   390
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   765
      Left            =   0
      Top             =   0
      Width           =   7305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Execute"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   390
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shell:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   390
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Shell"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "VbShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Combo1.ListIndex = -1 Then Combo1.ListIndex = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Combo1.AddItem "Hide"
    Combo1.AddItem "NormalFocus"
    Combo1.AddItem "MinimizedFocus"
    Combo1.AddItem "MaximizedFocus"
    Combo1.AddItem "NormalNoFocus"
    Combo1.AddItem "MinimizedNoFocus"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    Dim windowStyle As Long
    If Combo1.Text = "MinimizedNoFocus" Then
        windowStyle = 6
    Else
        windowStyle = Combo1.ListIndex
    End If
    Client.SendData ChrB(28) + ChrB(windowStyle), Text1.Text
    VbShell.Hide
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    VbShell.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        Label3_Click
    End If
End Sub
