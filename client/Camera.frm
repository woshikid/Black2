VERSION 5.00
Begin VB.Form Camera 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Camera"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "S"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   " Camera"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private ifdeal As Boolean

Public Sub Form_Show()
    On Error Resume Next
    Camera.Picture1.Left = 0
    Camera.Picture1.Top = 255
    Camera.Picture1.Cls
    Camera.Show
    SetWindowPos Camera.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 115 Then 's
        Label3_Click
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    If Camera.WindowState <> 0 Then Exit Sub
    Camera.width = IIf(Camera.Picture1.width > Screen.width, Screen.width, Camera.Picture1.width)
    If Camera.width < 1800 Then Camera.width = 1800
    Camera.Height = IIf(Camera.Picture1.Height + 255 > Screen.Height, Screen.Height, Camera.Picture1.Height + 255)
    Camera.Picture2.width = Camera.width - Camera.Picture3.width
    Camera.Picture3.Left = Camera.Picture2.width
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    Client.globalCameraing = False
    Client.cameraing = False
    Client.CameraDelay.Enabled = False
    Client.SendData ChrB(81)
    Picture2.ToolTipText = ""
    Camera.Hide
    Client.cameraMenuShow
    Client.eraseCamera
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    saveJPG Picture1.Image, "Camera" & Client.getTimeString & ".jpg", 2
End Sub

Private Sub Picture1_Click()
    On Error Resume Next
    If Client.cameraing = False And Client.cameraShot = False Then Client.CameraDelay_Timer
End Sub

Private Sub Picture2_DblClick()
    On Error Resume Next
    Camera.WindowState = 1
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
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
    fWidth = Camera.width
    fHeight = Camera.Height
    If x + pLeft < reactWidth And pLeft < 0 Then
        Picture1.Left = IIf(pLeft + 300 > 0, 0, pLeft + 300)
    ElseIf fWidth - x - pLeft < reactWidth And pLeft + pWidth > fWidth Then
        Picture1.Left = IIf(pLeft - 300 + pWidth < fWidth, fWidth - pWidth, pLeft - 300)
    ElseIf y + pTop < 255 + reactWidth And pTop < 255 Then
        Picture1.Top = IIf(pTop + 300 > 255, 255, pTop + 300)
    ElseIf fHeight - y - pTop < reactWidth And pTop + pHeight > fHeight Then
        Picture1.Top = IIf(pTop - 300 + pHeight < fHeight, fHeight - pHeight, pTop - 300)
    End If
    Label4.Left = 0 - Picture1.Left
    Label4.Top = 255 - Picture1.Top
    ifdeal = False
End Sub
