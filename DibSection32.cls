VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DibSection32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declares
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
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
    ' Note no palette entry here, not needed
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const OBJ_BITMAP As Long = 7
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
    
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long


Private hDC As Long
Private hDIB As Long
Private hBmpOld As Long
Private bi As BITMAPINFO
Private SizeW As Long, SizeH As Long
Private BitsPtr As Long

Public Function Create(ByVal NewW As Long, ByVal NewH As Long) As Boolean
    ' Save the actual WxH that we want
    SizeW = NewW * 4&
    SizeH = NewH
    ' Adjust to next 64x64 multiple
    NewW = 64& * (1& + NewW \ 64&)
    NewH = 64& * (1& + NewH \ 64&)
    ' Check if we need to create a new bitmap
    With bi.bmiHeader
        If .biSizeImage > 0 Then
            If NewW <= .biWidth And NewH <= Abs(.biHeight) Then
                ' Erase the bits anyway
                ZeroBits
                ' We're happy
                Create = True
                Exit Function
            End If
        End If
    End With
    ' We do need to create a new one :
    
    ' Clear up old references
    Destroy
    
    ' Make new stuff
    hDC = CreateCompatibleDC(0)
    If (hDC <> 0) Then
        With bi.bmiHeader
            .biSize = Len(bi.bmiHeader)
            .biWidth = NewW
            .biHeight = -NewH  ' Negative for top-down bitmap
            .biPlanes = 1
            .biBitCount = 32
            .biCompression = BI_RGB
            .biSizeImage = Abs(.biHeight) * BytesPerScanline
        End With
        hDIB = CreateDIBSection(hDC, bi, DIB_RGB_COLORS, BitsPtr, 0, 0)
        If hDIB Then
            hBmpOld = SelectObject(hDC, hDIB)
            ZeroBits
            Create = True
        Else
            DeleteObject hDC
            hDC = 0
        End If
    End If
End Function

Private Sub Destroy()
    If (hDC <> 0) Then
        If (hDIB <> 0) Then
            SelectObject hDC, hBmpOld
            DeleteObject hDIB
        End If
        DeleteObject hDC
    End If
    hDC = 0: hDIB = 0: hBmpOld = 0: BitsPtr = 0
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub ZeroBits()
    ' Quick check to see if we have any data
    If hDC = 0 Then Exit Sub
    
    ZeroMemory ByVal BitsPtr, bi.bmiHeader.biSizeImage
End Sub

Public Property Get DeviceContext() As Long
    DeviceContext = hDC
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = 4& * bi.bmiHeader.biWidth
End Property

Public Function GetBits(ByRef ArrBits() As Byte, ByVal psa As Long, ByRef W As Long, ByRef H As Long) As Boolean
    ' Quick check to see if we have any data
    If hDC = 0 Then Exit Function
    
    ' Return the Width and Height (the actual ones)
    W = SizeW
    H = SizeH
    
    ' Update 08/08/06 : Removed this check, it's more trouble than it's worth
    '                   Just make sure the array is not VB-allocated before you call this function
    ' Erase the array (in case it's already got data)
    'Erase ArrBits
    
    ' Create a temp SAFEARRAY2D structure
    Dim sa As SAFEARRAY2D
    
    ' Initialize the array
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Abs(bi.bmiHeader.biHeight)
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bi.bmiHeader.biWidth * 4 ' Bytes per scanline
        .pvData = BitsPtr
    End With
    
    ' Copy the data
    CopyMemory ByVal psa, sa, LenB(sa)
    
    ' Copy the descriptor
    CopyMemory ByVal VarPtrArray(ArrBits()), psa, 4
End Function

Public Sub FreeBits(ByRef ArrBits() As Byte)
    ' Probably not necessary, but we'll do it anyway
    ' Note that there is no check here that the array actually had our
    ' bits in it, and so this will free ANY array
    ' DO NOT USE IT unless the array was allocated by us, or the memory
    ' will not be released
    CopyMemory ByVal VarPtrArray(ArrBits()), 0&, 4
End Sub
