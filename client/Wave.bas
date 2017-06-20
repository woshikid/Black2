Attribute VB_Name = "Wave"
Option Explicit
Private Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
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
        Reserved As Long
End Type
Private Const GWL_WNDPROC = (-4)
Private Const WAVE_FORMAT_PCM = 1
Private Const WAVE_MAPPER = -1&
Private Const CALLBACK_WINDOW = &H10000
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_SHARE = &H2000
Private Const GMEM_ZEROINIT = &H40
Private Const MM_WOM_DONE = &H3BD
Private WavOutFmt As WAVEFORMAT
Private hWaveOut As Long
Private hMemOut() As Byte
Public Const CHANNEL = 1
Public Const SAMPLES = 11025&
Public Const BUF_SIZE_ONE As Long = 11025
Public BUF_SIZE As Long
Public bufsec As Long
Public extraBuf() As Byte
Public extraUsed As Boolean
Public outHdr As WAVEHDR
Public waveReady As Boolean
Public waveData() As Byte
Public wavePos As Long
Public waveSendFinish As Boolean
Public wavePlayFinish As Boolean
Private prevWndProc As Long

Public Sub WaveOutInit(ByVal nCh As Long, ByVal Sample As Long, Optional ByVal nBits As Long = 16)
    On Error Resume Next
    If waveReady = True Then Exit Sub
    prevWndProc = GetWindowLong(Client.hwnd, GWL_WNDPROC)
    SetWindowLong Client.hwnd, GWL_WNDPROC, AddressOf WaveOutProc
    Dim ret As Long
    WavOutFmt.wFormatTag = WAVE_FORMAT_PCM
    WavOutFmt.nChannels = nCh
    WavOutFmt.nSamplesPerSec = Sample
    WavOutFmt.nBlockAlign = nBits * nCh \ 8
    WavOutFmt.wBitsPerSample = nBits
    WavOutFmt.cbSize = 0
    WavOutFmt.nAvgBytesPerSec = nBits * Sample * nCh \ 8
    ret = waveOutOpen(hWaveOut, WAVE_MAPPER, WavOutFmt, Client.hwnd, 0, CALLBACK_WINDOW)
    If ret <> 0 Then
        waveOutClose hWaveOut
        SetWindowLong Client.hwnd, GWL_WNDPROC, prevWndProc
        Exit Sub
    End If
    ReDim hMemOut(BUF_SIZE - 1)
    outHdr.lpData = VarPtr(hMemOut(0))
    outHdr.dwBufferLength = BUF_SIZE
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0
    outHdr.dwUser = 0
    ret = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If ret <> 0 Then
        waveOutUnprepareHeader hWaveOut, outHdr, Len(outHdr)
        Erase hMemOut
        waveOutClose hWaveOut
        SetWindowLong Client.hwnd, GWL_WNDPROC, prevWndProc
        Exit Sub
    End If
    waveReady = True
    ReDim waveData(BUF_SIZE - 1)
    wavePos = 0
    waveSendFinish = False
    wavePlayFinish = True
    extraUsed = False
End Sub

Public Function WaveOutProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If waveReady = True And Msg = MM_WOM_DONE Then
        If waveSendFinish = True Then
            CopyMemory ByVal outHdr.lpData, waveData(0), BUF_SIZE
            WaveOutPlayback
            wavePos = 0
            waveSendFinish = False
            If extraUsed = True Then
                CopyMemory waveData(wavePos), extraBuf(0), UBound(extraBuf) + 1
                wavePos = wavePos + UBound(extraBuf) + 1
                extraUsed = False
                Erase extraBuf
                Client.SendData ChrB(69)
            End If
        Else
            wavePlayFinish = True
        End If
    End If
    WaveOutProc = CallWindowProc(prevWndProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub WaveOutPlayback()
    On Error Resume Next
    If waveReady = False Then Exit Sub
    wavePlayFinish = False
    waveOutWrite hWaveOut, outHdr, Len(outHdr)
End Sub

Public Sub WaveOutDeinit()
    On Error Resume Next
    If waveReady = False Then Exit Sub
    waveOutReset hWaveOut
    waveOutUnprepareHeader hWaveOut, outHdr, Len(outHdr)
    waveOutClose hWaveOut
    SetWindowLong Client.hwnd, GWL_WNDPROC, prevWndProc
    Erase hMemOut
    Erase waveData
    Erase extraBuf
    waveReady = False
End Sub
