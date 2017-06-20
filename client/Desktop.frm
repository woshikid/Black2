VERSION 5.00
Begin VB.Form Desktop 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Desktop"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Width           =   255
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "P"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   900
         TabIndex        =   8
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "S"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "M"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   " Desktop"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   255
      Width           =   2535
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   150
         Left            =   -300
         Shape           =   3  'Circle
         Top             =   -300
         Visible         =   0   'False
         Width           =   150
      End
   End
End
Attribute VB_Name = "Desktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private ifdeal As Boolean
Private lastButton As Long
Private nowX As Long
Private nowY As Long

Public Sub Form_Show()
    On Error Resume Next
    Desktop.Picture1.Left = 0
    Desktop.Picture1.Top = 255
    Desktop.Picture1.Cls
    Desktop.Show
    SetWindowPos Desktop.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    If Desktop.WindowState <> 0 Then Exit Sub
    Desktop.width = IIf(Desktop.Picture1.width > Screen.width, Screen.width, Desktop.Picture1.width)
    If Desktop.width < 1800 Then Desktop.width = 1800
    Desktop.Height = IIf(Desktop.Picture1.Height + 255 > Screen.Height, Screen.Height, Desktop.Picture1.Height + 255)
    Desktop.Picture2.width = Desktop.width - Desktop.Picture3.width
    Desktop.Picture3.Left = Desktop.Picture2.width
    Desktop.Text1.width = Desktop.width - 1800
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    Client.globalDesktoping = False
    Client.desktoping = False
    Client.Focuser.Enabled = False
    Client.BMPShow.Enabled = False
    Client.DesktopDelay.Enabled = False
    UnHook
    Client.SendData ChrB(46)
    Text1.ToolTipText = ""
    Picture2.ToolTipText = ""
    Desktop.Hide
    Client.desktopMenuShow
    Client.eraseDesktop
End Sub

Private Sub Label2_Click()
    On Error Resume Next
    If Client.sendEvent = True Then Client.Meventoff_Click
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Client.sendEvent = True Then Exit Sub
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    If Label3.Caption = "M" Then
        Client.SendData ChrB(62)
        Label3.Caption = "U"
        Label3.ForeColor = &HFFFFFF
        Label3.BackColor = &HFF&
    ElseIf Label3.Caption = "U" Then
        Client.SendData ChrB(63)
        Label3.Caption = "M"
        Label3.ForeColor = &HFF&
        Label3.BackColor = &H0&
    ElseIf Label3.Caption = "H" Then
        Client.Mkidoff_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 104 Then 'h
        If Label3.Visible = False Then Client.Mhookon_Click
    ElseIf KeyAscii = 115 Then 's
        Label4_Click
    ElseIf KeyAscii = 112 Then 'p
        Label5_Click
    ElseIf KeyAscii = 117 Then 'u
        If Label3.Visible = True Then
            If Label3.Caption = "M" Then
                Client.Mhookoff_Click
            Else
                Label3_Click
            End If
        End If
    End If
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    saveJPG Picture1.Image, "Desktop" & Client.getTimeString & ".jpg", 2
End Sub

Private Sub Label5_Click()
    On Error Resume Next
    Client.SendData ChrB(85)
End Sub

Private Sub Picture1_Click()
    On Error Resume Next
    If Client.desktoping = False And Client.sendEvent = False Then Client.DesktopDelay_Timer
End Sub

Private Sub Picture2_DblClick()
    On Error Resume Next
    Desktop.WindowState = 1
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Client.sendEvent = True Then Client.SendData ChrB(40) + ChrB(Button) + getPos + ChrB(0)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Client.sendEvent = True Then Client.SendData ChrB(40) + ChrB(Button) + getPos + ChrB(1)
    lastButton = Button
End Sub

Private Sub Picture1_DblClick()
    On Error Resume Next
    If Client.sendEvent = True Then Client.SendData ChrB(40) + ChrB(lastButton) + getPos + ChrB(0)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    nowX = x
    nowY = y
    If Client.sendEvent = True Then Client.MouseSender.Enabled = True
    If ifdeal = False Then
        ifdeal = True
        Exit Sub
    End If
    Dim reactWidth As Long
    reactWidth = 300
    Dim pLeft As Single
    Dim pTop As Single
    Dim pWidth As Single
    Dim pHeight As Single
    Dim fWidth As Single
    Dim fHeight As Single
    pLeft = Picture1.Left
    pTop = Picture1.Top
    pWidth = Picture1.width
    pHeight = Picture1.Height
    fWidth = Desktop.width
    fHeight = Desktop.Height
    If x + pLeft < reactWidth And pLeft < 0 Then
        Picture1.Left = IIf(pLeft + 300 > 0, 0, pLeft + 300)
    ElseIf fWidth - x - pLeft < reactWidth And pLeft + pWidth > fWidth Then
        Picture1.Left = IIf(pLeft - 300 + pWidth < fWidth, fWidth - pWidth, pLeft - 300)
    ElseIf y + pTop < 255 + reactWidth And pTop < 255 Then
        Picture1.Top = IIf(pTop + 300 > 255, 255, pTop + 300)
    ElseIf fHeight - y - pTop < reactWidth And pTop + pHeight > fHeight Then
        Picture1.Top = IIf(pTop - 300 + pHeight < fHeight, fHeight - pHeight, pTop - 300)
    End If
    ifdeal = False
End Sub

Public Function getPos() As String
    On Error Resume Next
    Dim head(3) As Byte
    Dim pixelX As Long
    Dim pixelY As Long
    pixelX = nowX \ Screen.TwipsPerPixelX
    pixelY = nowY \ Screen.TwipsPerPixelY
    If Client.miniDesktop = False Then
        head(0) = pixelX \ 256
        head(1) = pixelX Mod 256
        head(2) = pixelY \ 256
        head(3) = pixelY Mod 256
    Else
        head(0) = pixelX * 4 \ 256
        head(1) = pixelX * 4 Mod 256
        head(2) = pixelY * 4 \ 256
        head(3) = pixelY * 4 Mod 256
    End If
    getPos = CStr(head)
End Function
