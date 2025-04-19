Attribute VB_Name = "Graphics"
Option Explicit

' Siehe http://foren.activevb.de/forum/vb-classic/thread-417060/beitrag-417069/Bsp-per-CreateDIBSection-und-Se/.

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Public Type RGB
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pbmi As BITMAPINFO, ByVal usage As Long, ByVal ppvBits As Long, ByVal hSection As Long, ByVal Offset As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPictDesc As PICTDESC, ByRef riid As IID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private tBITMAPINFO As BITMAPINFO
Private tPictDesc As PICTDESC
Private IID_IPicture As IID

Public FrameBuffer() As RGB
Public RocketBitmap() As RGB
Public Background() As RGB

Public Function GetRGB(Col As Long) As RGB

GetRGB.Red = (Col And &HFF&)
GetRGB.Green = (Col And &HFF00&) \ &H100
GetRGB.Blue = (Col And &HFF0000) \ &H10000
  
End Function

Private Function HandleToPicture(ByVal hGDIHandle As Long) As StdPicture

Call OleCreatePictureIndirect(tPictDesc, IID_IPicture, 1&, HandleToPicture)
    
End Function

Public Sub Init(PB As PictureBox)

'PB.ScaleMode = vbPixels
'PB.Appearance = 0
'PB.AutoRedraw = True

ReDim FrameBuffer(PB.ScaleWidth * PB.ScaleHeight)
Background = FrameBuffer

With tBITMAPINFO
    .bmiHeader.biSize = Len(tBITMAPINFO)
    .bmiHeader.biWidth = CLng(PB.ScaleWidth)
    .bmiHeader.biHeight = -CLng(PB.ScaleHeight)
    .bmiHeader.biPlanes = 1
    .bmiHeader.biBitCount = 24
End With

With IID_IPicture
    .Data1 = &H7BF80981
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(3) = &HAA
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
End With

End Sub

Public Sub Draw(PB As PictureBox)

Dim hBitmap As Long

hBitmap = CreateDIBSection(0&, tBITMAPINFO, 0&, 0&, 0&, 0&)

With tPictDesc
    .cbSizeOfStruct = Len(tPictDesc)
    .picType = vbPicTypeBitmap
    .hgdiObj = hBitmap
End With

Call SetDIBits(0&, hBitmap, 0&, Abs(tBITMAPINFO.bmiHeader.biHeight), FrameBuffer(0), tBITMAPINFO, DIB_RGB_COLORS)

PB.Picture = HandleToPicture(hBitmap)
PB.Refresh

Call DeleteObject(hBitmap)

End Sub
