Attribute VB_Name = "Hook"
Option Explicit
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Type tagKBDLLHOOKSTRUCT
    vkCode As Integer
    scanCode As Integer
    flags As Integer
    time As Integer
    dwExtraInfo As Integer
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private hook1 As Long
Private hook2 As Long
Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE_LL = 14
Public hooked As Boolean
Public freezed As Boolean
Private hkey() As Byte
Public mouseStatus As Byte
Public mouseHooked As Boolean

Public Sub EnableHook()
    On Error Resume Next
    If freezed = True Or hooked = True Then Exit Sub
    ReDim hkey(0)
    hook1 = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyHook, App.hInstance, 0)
    hooked = True
End Sub

Public Sub UnHook()
    On Error Resume Next
    If hooked = False Then Exit Sub
    UnMouseHook
    UnhookWindowsHookEx hook1
    Erase hkey
    hooked = False
End Sub

Public Sub EnableMouseHook()
    On Error Resume Next
    If freezed = True Or hooked = False Or mouseHooked = True Then Exit Sub
    hook2 = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
    mouseHooked = True
End Sub

Public Sub UnMouseHook()
    On Error Resume Next
    If mouseHooked = False Then Exit Sub
    UnhookWindowsHookEx hook2
    mouseHooked = False
    mouseStatus = 2
End Sub

Public Function KeyHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    KeyHook = 0
    Dim ks As tagKBDLLHOOKSTRUCT
    CopyMemory ks, ByVal lParam, LenB(ks)
    If ks.dwExtraInfo = 0 Or ks.dwExtraInfo = 1 Or ks.dwExtraInfo = 32 Or ks.dwExtraInfo = 33 Then
        Dim i As Long
        i = UBound(hkey)
        hkey(i) = ks.vkCode
        ReDim Preserve hkey(i + 1)
    End If
    CallNextHookEx hook1, vkCode, wParam, ByVal lParam
End Function

Public Function MouseHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    MouseHook = 0
    If wParam = 513 Then
        mouseStatus = 3
    ElseIf wParam = 516 Then
        mouseStatus = 6
    ElseIf wParam = 519 Then
        mouseStatus = 9
    ElseIf wParam = 514 Or wParam = 517 Or wParam = 520 Then
        mouseStatus = 2
    End If
    CallNextHookEx hook2, vkCode, wParam, ByVal lParam
End Function

Public Function ReadKey() As Long
    On Error Resume Next
    ReadKey = hkey(0)
    Dim i As Long
    i = UBound(hkey)
    If i > 0 Then
        CopyMemory hkey(0), hkey(1), i
        ReDim Preserve hkey(i - 1)
    End If
End Function

Public Sub Freeze()
    On Error Resume Next
    If freezed = True Then Exit Sub
    UnHook
    hook1 = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf FreezeHook, App.hInstance, 0)
    hook2 = SetWindowsHookEx(WH_MOUSE_LL, AddressOf FreezeHook, App.hInstance, 0)
    freezed = True
End Sub

Public Sub UnFreeze()
    On Error Resume Next
    If freezed = False Then Exit Sub
    UnhookWindowsHookEx hook1
    UnhookWindowsHookEx hook2
    freezed = False
End Sub

Public Function FreezeHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    FreezeHook = 1
End Function
