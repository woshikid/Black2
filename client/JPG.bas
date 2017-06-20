Attribute VB_Name = "JPG"
Option Explicit
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "oleaut32.dll" (ByVal lpStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, lpIPicture As IPicture) As Long
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
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal fileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long

Public Function ArrayToPicture(PictureData() As Byte) As IPicture
    On Error Resume Next
    Dim aGUID(0 To 3) As Long
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    Dim Stream As IUnknown
    Set Stream = ArrayToStream(PictureData)
    If Not Stream Is Nothing Then OleLoadPicture Stream, 0, 0, aGUID(0), ArrayToPicture
End Function

Private Function ArrayToStream(PictureData() As Byte) As IUnknown
    On Error Resume Next
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim size As Long
    size = UBound(PictureData) + 1
    If size > 0 Then
        o_hMem = GlobalAlloc(&H2, size)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, PictureData(0), size
                GlobalUnlock o_hMem
                CreateStreamOnHGlobal o_hMem, 1, ArrayToStream
            Else
                GlobalFree o_hMem
            End If
        End If
    End If
End Function

Public Sub saveJPG(ByVal pict As StdPicture, ByVal fileName As String, ByVal imageType As Long, Optional ByVal jpgQuality As Long = 100)
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
    GdipSaveImageToFile lBitmap, StrPtr(fileName), tJpgEncoder, tParams
    'Destroy the bitmap
    GdipDisposeImage lBitmap
    'Shutdown GDI+
    GdiplusShutdown lGDIP
End Sub
