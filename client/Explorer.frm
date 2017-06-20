VERSION 5.00
Begin VB.Form Explorer 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Explorer"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
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
   ScaleHeight     =   5895
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4800
      Width           =   3975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1710
      ItemData        =   "Explorer.frx":0000
      Left            =   120
      List            =   "Explorer.frx":0002
      TabIndex        =   2
      Top             =   2760
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1710
      ItemData        =   "Explorer.frx":0004
      Left            =   120
      List            =   "Explorer.frx":0006
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1710
      Hidden          =   -1  'True
      Left            =   4440
      System          =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2010
      Left            =   4440
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FFFFFF&
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   5520
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   405
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8280
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
      Left            =   8040
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Explorer"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Explorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    On Error Resume Next
    Dim driver As String
    driver = Combo1.List(Combo1.ListIndex)
    Client.SendData ChrB(35), Left(driver, InStr(driver, ":"))
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.Path = Dir1.Path
    Text2.Text = Dir1.Path
    If Right(Text2.Text, 1) <> "\" Then Text2.Text = Text2.Text + "\"
    Text2.SelStart = Len(Text2.Text)
    hideUploadButton
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    On Error Resume Next
    showUploadButton
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    Text2.Text = Dir1.Path
    If Right(Text2.Text, 1) <> "\" Then Text2.Text = Text2.Text + "\"
    Text2.SelStart = Len(Text2.Text)
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    If Client.downloading = True Then Exit Sub
    Dim filePath As String
    filePath = Text2.Text + List2.List(List2.ListIndex)
    'incase file is readonly
    SetAttr filePath, vbNormal
    'open file
    Open filePath For Binary As #1
    'check if readable
    If Dir(filePath) = vbNullString Then
        Close #1
        MessageBox Explorer.hwnd, "Write error!", "Download Failed", vbOKOnly
        Exit Sub
    End If
    Dim Length As Long
    Length = LOF(1)
    'continue from breakpoint
    Seek #1, Length + 1
    'make the head
    Dim head(3) As Byte
    Dim i As Long
    For i = 3 To 0 Step -1
        'divide the long data into bytes
        head(i) = Length Mod 256
        Length = Length \ 256
    Next i
    'send length
    Client.SendData ChrB(4) + CStr(head), Text1.Text + List2.List(List2.ListIndex)
    Client.downloading = True
    hideDownloadButton
End Sub

Private Sub Label2_Click()
    On Error Resume Next
    If Client.uploading = True Then Exit Sub
    Dim filePath As String
    filePath = Text2.Text + File1.List(File1.ListIndex)
    'check if the file is existed
    If Dir(filePath, vbHidden + vbSystem) = vbNullString Then
        MessageBox Explorer.hwnd, "File not found!", "Upload Failed", vbOKOnly
        Exit Sub
    End If
    'open file
    Open filePath For Binary As #2
    Client.SendData ChrB(10), Text1.Text + File1.List(File1.ListIndex)
    Client.uploading = True
    hideUploadButton
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label4_Click()
    On Error Resume Next
    Explorer.Hide
End Sub

Private Sub Label5_Click()
    On Error Resume Next
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
    hideUploadButton
    Client.SendData ChrB(29)
End Sub

Private Sub Label6_Click()
    On Error Resume Next
    Close #1
    Client.downloading = False
    Client.SendData ChrB(9)
    If Client.uploading = False Then Client.Speed.Enabled = False
    hideDownloadProgress
    showDownloadButton
End Sub

Private Sub Label7_Click()
    On Error Resume Next
    Close #2
    Client.uploading = False
    Client.SendData ChrB(14)
    If Client.downloading = False Then Client.Speed.Enabled = False
    hideUploadProgress
    showUploadButton
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    Client.SendData ChrB(32), List1.List(List1.ListIndex)
End Sub

Private Sub List2_Click()
    On Error Resume Next
    showDownloadButton
End Sub

Public Sub showDownloadProgress()
    On Error Resume Next
    Shape4.Visible = True
    Shape5.Visible = True
    Shape6.Visible = True
    Label6.Visible = True
    Shape6.width = 0
    Label6.ToolTipText = ""
End Sub

Public Sub hideDownloadProgress()
    On Error Resume Next
    Shape4.Visible = False
    Shape5.Visible = False
    Shape6.Visible = False
    Label6.Visible = False
End Sub

Public Sub showUploadProgress()
    On Error Resume Next
    Shape7.Visible = True
    Shape8.Visible = True
    Shape9.Visible = True
    Label7.Visible = True
    Shape8.width = Shape7.width
    Label7.ToolTipText = ""
End Sub

Public Sub hideUploadProgress()
    On Error Resume Next
    Shape7.Visible = False
    Shape8.Visible = False
    Shape9.Visible = False
    Label7.Visible = False
End Sub

Public Sub showDownloadButton()
    On Error Resume Next
    If Client.downloading = False And List2.ListIndex >= 0 Then
        Label1.Visible = True
        Shape1.Visible = True
    End If
End Sub

Public Sub showUploadButton()
    On Error Resume Next
    If Client.uploading = False And Text1.Text <> "" And File1.ListIndex >= 0 Then
        Label2.Visible = True
        Shape2.Visible = True
    End If
End Sub

Public Sub hideDownloadButton()
    On Error Resume Next
    Label1.Visible = False
    Shape1.Visible = False
End Sub

Public Sub hideUploadButton()
    On Error Resume Next
    Label2.Visible = False
    Shape2.Visible = False
End Sub
