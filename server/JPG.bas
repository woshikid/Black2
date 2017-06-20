Attribute VB_Name = "JPG"
Option Explicit
Private Const ClsidJPEG As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const ClsidGIF As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Public Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
    GUID As GUID
    NumberOfvalues As Long
    type As Long
    Value As Long
End Type
Private Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public jpgData() As Byte
Public jpgQuality As Long

Public Sub makeJPG(ByVal pict As StdPicture, ByVal imageType As Long)
    On Error Resume Next
    If jpgQuality < 1 Then jpgQuality = 1
    If jpgQuality > 100 Then jpgQuality = 100
    Dim tSI As GdiplusStartupInput
    Dim lRes As Long
    Dim lGDIP As Long
    Dim lBitmap As Long
    ' Initialize GDI+
    tSI.GdiplusVersion = 1
    GdiplusStartup lGDIP, tSI
    ' Create the GDI+ bitmap from the image handle
    GdipCreateBitmapFromHBITMAP pict.Handle, 0, lBitmap
    Dim tJpgEncoder As GUID
    Dim tParams As EncoderParameters
    If imageType = 1 Then
        CLSIDFromString StrPtr(ClsidGIF), tJpgEncoder
    Else
        ' Initialize the encoder GUID
        CLSIDFromString StrPtr(ClsidJPEG), tJpgEncoder
        ' Initialize the encoder parameters
        tParams.Count = 1
        With tParams.Parameter ' Quality
            ' Set the Quality GUID
            CLSIDFromString StrPtr(EncoderQuality), .GUID
            .NumberOfvalues = 1
            .type = 4
            .Value = VarPtr(jpgQuality)
        End With
    End If
    Dim Stream As IUnknown
    CreateStreamOnHGlobal 0, 1, Stream
    GdipSaveImageToStream lBitmap, Stream, tJpgEncoder, tParams
    ReDim jpgData(0)
    StreamToArray Stream, jpgData
    'Destroy the bitmap
    GdipDisposeImage lBitmap
    'Shutdown GDI+
    GdiplusShutdown lGDIP
End Sub

Private Sub StreamToArray(Stream As IUnknown, arrayBytes() As Byte)
    On Error Resume Next
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    If Stream Is Nothing Then Exit Sub
    If GetHGlobalFromStream(ObjPtr(Stream), o_hMem) = 0 Then
        If o_hMem <> 0 Then
            o_lngByteCount = GlobalSize(o_hMem)
            If o_lngByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(o_lngByteCount - 1)
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lngByteCount
                    GlobalUnlock o_hMem
                End If
            End If
        End If
    End If
End Sub
