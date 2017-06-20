Attribute VB_Name = "Hook"
Option Explicit
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Type tagKBDLLHOOKSTRUCT
    vkCode As Integer
    scanCode As Integer
    flags As Integer
    time As Integer
    dwExtraInfo As Integer
End Type
Private hook1 As Long
Private Const WH_KEYBOARD_LL = 13
Public hooked As Boolean

Public Sub EnableHook()
    On Error Resume Next
    If hooked = True Then Exit Sub
    hook1 = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyHook, App.hInstance, 0)
    hooked = True
End Sub

Public Sub UnHook()
    On Error Resume Next
    If hooked = False Then Exit Sub
    UnhookWindowsHookEx hook1
    hooked = False
End Sub

Public Function KeyHook(ByVal vkCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    KeyHook = 1
    Dim ks As tagKBDLLHOOKSTRUCT
    CopyMemory ks, ByVal lParam, LenB(ks)
    If ks.dwExtraInfo = 0 Or ks.dwExtraInfo = 1 Or ks.dwExtraInfo = 32 Or ks.dwExtraInfo = 33 Then
        Client.SendData ChrB(52) + ChrB(ks.vkCode) + ChrB(0)
    ElseIf ks.dwExtraInfo = 128 Or ks.dwExtraInfo = 129 Then
        Client.SendData ChrB(52) + ChrB(ks.vkCode) + ChrB(2)
    End If
End Function
