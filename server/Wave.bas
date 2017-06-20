Attribute VB_Name = "Wave"
Option Explicit
Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function WaveInStart Lib "winmm.dll" Alias "waveInStart" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type
Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type
Private prevWndProc As Long
Private Const GWL_WNDPROC = (-4)
Private Const WAVE_FORMAT_PCM = 1
Private Const WAVE_MAPPER = -1&
Private Const CALLBACK_WINDOW = &H10000
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_SHARE = &H2000
Private Const GMEM_ZEROINIT = &H40
Private Const MM_WIM_DATA = &H3C0
Public Const BUF_SIZE_ONE As Long = 11025
Public Const CHANNEL = 1
Public Const SAMPLES = 11025&
Public BUF_SIZE As Long
Public bufsec As Long
Private WavInFmt As WAVEFORMAT
Private hWaveIn As Long
Private inHdr As WAVEHDR
Public hMemIn() As Byte
Public waveReady As Boolean
Public waveData() As Byte
Public wavePos As Long
Public waveSendFinish As Boolean
Public waveRecFinish As Boolean

Public Sub WaveInInit(ByVal nCh As Long, ByVal Sample As Long, Optional ByVal nBits As Long = 16)
    On Error Resume Next
    If waveReady = True Then Exit Sub
    prevWndProc = GetWindowLong(Server.hwnd, GWL_WNDPROC)
    SetWindowLong Server.hwnd, GWL_WNDPROC, AddressOf WaveInProc
    Dim ret As Long
    WavInFmt.wFormatTag = WAVE_FORMAT_PCM
    WavInFmt.nChannels = nCh
    WavInFmt.nSamplesPerSec = Sample
    WavInFmt.nBlockAlign = nBits * nCh \ 8
    WavInFmt.wBitsPerSample = nBits
    WavInFmt.cbSize = 0
    WavInFmt.nAvgBytesPerSec = nBits * Sample * nCh \ 8
    ret = waveInOpen(hWaveIn, WAVE_MAPPER, WavInFmt, Server.hwnd, 0, CALLBACK_WINDOW)
    If ret <> 0 Then
        waveInClose hWaveIn
        SetWindowLong Server.hwnd, GWL_WNDPROC, prevWndProc
        Exit Sub
    End If
    ReDim hMemIn(BUF_SIZE - 1)
    inHdr.lpData = VarPtr(hMemIn(0))
    inHdr.dwBufferLength = BUF_SIZE
    inHdr.dwFlags = 0
    inHdr.dwLoops = 0
    inHdr.dwUser = 0
    ret = waveInPrepareHeader(hWaveIn, inHdr, Len(inHdr))
    If ret <> 0 Then
        waveInUnprepareHeader hWaveIn, inHdr, Len(inHdr)
        Erase hMemIn
        waveInClose hWaveIn
        SetWindowLong Server.hwnd, GWL_WNDPROC, prevWndProc
        Exit Sub
    End If
    waveReady = True
    ReDim waveData(BUF_SIZE - 1)
End Sub

Public Function WaveInProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If waveReady = True And Msg = MM_WIM_DATA Then
        If waveSendFinish = True Then
            waveData = hMemIn
            WaveInRecord
            wavePos = 0
            waveSendFinish = False
            Server.sendWave
        Else
            waveRecFinish = True
        End If
    End If
    WaveInProc = CallWindowProc(prevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub WaveInRecord()
    On Error Resume Next
    If waveReady = False Then Exit Sub
    waveRecFinish = False
    Dim ret As Long
    ret = waveInAddBuffer(hWaveIn, inHdr, Len(inHdr))
    If ret = 0 Then WaveInStart (hWaveIn)
End Sub

Public Sub WaveInDeinit()
    On Error Resume Next
    If waveReady = False Then Exit Sub
    waveInStop hWaveIn
    waveInReset hWaveIn
    waveInUnprepareHeader hWaveIn, inHdr, Len(inHdr)
    waveInClose hWaveIn
    SetWindowLong Server.hwnd, GWL_WNDPROC, prevWndProc
    Erase hMemIn
    Erase waveData
    waveReady = False
End Sub
