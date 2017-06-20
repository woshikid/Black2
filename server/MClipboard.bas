Attribute VB_Name = "MClipboard"
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Const CF_UNICODETEXT = 13
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private clipData() As Long
Private clipFormat() As Long
Public clipLength As Long

Public Function getClipboardText() As String
    On Error Resume Next
    If OpenClipboard(0) = 0 Then Exit Function
    Dim hMem As Long
    Dim lpMem As Long
    Dim Length As Long
    Dim data() As Byte
    hMem = GetClipboardData(CF_UNICODETEXT)
    If hMem <> 0 Then
        Length = GlobalSize(hMem)
        If Length > 0 Then
            lpMem = GlobalLock(hMem)
            If lpMem <> 0 Then
                ReDim data(Length - 1)
                CopyMemory data(0), ByVal lpMem, Length
                GlobalUnlock hMem
                getClipboardText = CStr(data)
                Length = InStr(getClipboardText, Chr(0)) - 1
                If Length < 0 Then Length = 0
                getClipboardText = Left(getClipboardText, Length)
            End If
        End If
    End If
    CloseClipboard
End Function

Public Sub saveClipboard()
    On Error Resume Next
    freeClipData
    clipLength = -1
    If OpenClipboard(0) = 0 Then Exit Sub
    clipLength = 0
    Dim wFormat As Long
    Dim hMem As Long
    Dim lpMem As Long
    Dim Length As Long
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim overMem As Boolean
    overMem = False
    wFormat = EnumClipboardFormats(0)
    Do While wFormat <> 0
        hMem = GetClipboardData(wFormat)
        If hMem <> 0 Then
            Length = GlobalSize(hMem)
            If Length > 0 Then
                If clipLength + Length < 20971520 Then 'no more than 20M
                    lpMem = GlobalLock(hMem)
                    If lpMem <> 0 Then
                        hMemory = GlobalAlloc(&H2, Length)
                        If hMemory <> 0 Then
                            lpMemory = GlobalLock(hMemory)
                            If lpMemory <> 0 Then
                                Dim clip As Long
                                clip = UBound(clipFormat)
                                clipFormat(clip) = wFormat
                                ReDim Preserve clipFormat(clip + 1)
                                ReDim Preserve clipData(clip)
                                clipData(clip) = hMemory
                                CopyMemory ByVal lpMemory, ByVal lpMem, Length
                                clipLength = clipLength + Length
                                GlobalUnlock hMemory
                            Else
                                GlobalFree hMemory
                                overMem = True
                            End If
                        Else
                            overMem = True
                        End If
                        GlobalUnlock hMem
                    Else
                        overMem = True
                    End If
                Else
                    overMem = True
                End If
            End If
        End If
        If overMem = True Then
            clipLength = -1
            Exit Do
        End If
        wFormat = EnumClipboardFormats(wFormat)
    Loop
    CloseClipboard
End Sub

Public Sub restoreClipboard()
    On Error Resume Next
    If OpenClipboard(0) = 0 Then
        freeClipData
        Clipboard.Clear
        Exit Sub
    End If
    EmptyClipboard
    Dim clipMax As Long
    Dim i As Long
    clipMax = UBound(clipFormat) - 1
    For i = 0 To clipMax
        SetClipboardData clipFormat(i), clipData(i)
    Next i
    CloseClipboard
    freeClipData
End Sub

Public Sub freeClipData()
    On Error Resume Next
    Dim clipMax As Long
    Dim i As Long
    clipMax = UBound(clipFormat) - 1
    For i = 0 To clipMax
        GlobalFree clipData(i)
    Next i
    Erase clipData
    ReDim clipFormat(0)
End Sub

Public Sub initClipboard()
    On Error Resume Next
    ReDim clipFormat(0)
End Sub
