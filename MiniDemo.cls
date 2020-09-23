VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MiniDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private hWnd As Long    ' Window we are drawing in

Private Type STARINFO
    x As Double
    y As Double
    dx As Double
    dy As Double
    brightness As Byte
End Type

Private Const CountStars As Long = 500&

Private DemoStars(0 To CountStars - 1) As STARINFO

Private DemoDibText As DibSection32
Private DemoDibStars As DibSection32
Private DemoDibTemp As DibSection32

Private DemoDibBackBuf As DibSection32

Private DemoDibWork As DibSection32

Private ArrWork() As Byte, saWork As SAFEARRAY2D
Private ArrDraw() As Byte, saDraw As SAFEARRAY2D
Private ArrTemp() As Byte, saTemp As SAFEARRAY2D
Private ArrText() As Byte, saText As SAFEARRAY2D
Private ArrStars() As Byte, saStars As SAFEARRAY2D

Private DstW As Long, DstH As Long
Private SrcW As Long, SrcH As Long


Private DemoLookup() As RECT
Private DemoText() As String

Private AlphaMulLUT(0 To 255, 0 To 255) As Byte ' Lookup table (scaled byte multiplication)
Private AlphaAddLUT(0 To 255, 0 To 255) As Byte ' Lookup table (capped byte addition)
Private AlphaSubLUT(0 To 255, 0 To 255) As Byte ' Lookup table (capped byte subtraction)
Private FadeInLUT(0 To 255, 0 To 255) As Byte   ' Lookup table (fade-in)
Private FadeOutLUT(0 To 255, 0 To 255) As Byte  ' Lookup table (fade-out)

Private frd As FastRenderData       ' Setup information for fast render object
Private Renderer As IFastRender     ' Fast render object pointer

Private DemoPosition As Long                    ' Current position

' Initialisation
Public Sub Init(ByVal hWndAttach As Long)
    hWnd = hWndAttach
    If hWnd = 0 Then Exit Sub
    Dim rcView As RECT
    GetClientRect hWnd, rcView
    
    Set Renderer = FastRender.InitFastRender(frd)
    
    'Randomize Timer
    
    Set DemoDibBackBuf = New DibSection32
    DemoDibBackBuf.Create rcView.Right, rcView.Bottom
    
    Set DemoDibTemp = New DibSection32
    DemoDibTemp.Create rcView.Right, rcView.Bottom
    
    Set DemoDibStars = New DibSection32
    DemoDibStars.Create rcView.Right, rcView.Bottom
    
    Set DemoDibText = New DibSection32
    DemoDibText.Create rcView.Right, rcView.Bottom

    On Error GoTo NO_TEXT_FILE
    Dim nFile As Integer
    Dim sText As String
    nFile = FreeFile
    Open App.Path & "\Scrolltext.txt" For Input As nFile
    sText = Input$(LOF(nFile), nFile)
    Close nFile
    GoTo SKIP_ERROR_HANDLER
    
NO_TEXT_FILE:
    'ODS "Error loading file (handle=" & nFile & "): " & vbCrLf & "> " & Err.Description & vbCrLf
    sText = Replace$("come closer and see|see into the trees|find the girl|while you can|come closer and see|see into the dark|just follow your eyes|just follow your eyes||i hear her voice|calling my name|the sound is deep|in the dark|i hear her voice|and start to run|into the trees|into the trees||into the trees||suddenly i stop|but i know it's too late|i'm lost in a forest|all alone|the girl was never there|it's always the same|i'm running towards nothing|again and again and again|", "|", vbCrLf)
    On Error GoTo 0
SKIP_ERROR_HANDLER:

    Dim hFont As Long, hFontOld As Long, lf As LOGFONT
    Dim hPenOld As Long, hBrushOld As Long
    
    hFontOld = GetCurrentObject(DemoDibBackBuf.DeviceContext, OBJ_FONT)
    GetObject hFontOld, LenB(lf), lf
    
    Dim FontFace() As Byte
    FontFace = StrConv("Comic Sans MS" & vbNullChar, vbFromUnicode)
    
    With lf
        Call CopyMemory(.lfFaceName(1), FontFace(0), UBound(FontFace) + 1)
        .lfQuality = NONANTIALIASED_QUALITY
        .lfHeight = (.lfHeight * 2)
        
    End With
    
    hFont = CreateFontIndirect(lf)

    Dim rc As RECT
    rc.Right = rcView.Right
    rc.Bottom = rcView.Bottom
    ReDim DemoLookup(0 To rc.Right, 0 To rc.Bottom)
    hFontOld = SelectObject(DemoDibBackBuf.DeviceContext, hFont)
    DrawText DemoDibBackBuf.DeviceContext, sText, Len(sText), rc, DT_CALCRECT Or DT_MULTILINE Or DT_CENTER Or DT_TOP
    SelectObject DemoDibBackBuf.DeviceContext, hFontOld
    
    Set DemoDibWork = New DibSection32
    rc.Bottom = rc.Bottom + rcView.Bottom
    DemoDibWork.Create rc.Right, rc.Bottom
    
    ODS "Total height : %i\n", rc.Bottom
    
    rc.Top = rc.Top + rcView.Bottom
    
    Dim i As Long, j As Long, k As Long
    
'' Grid (useful for working out the displacement transform)
'    Dim hBrRed As Long, hBrBlue As Long, hBrOld As Long
'    hBrRed = CreateSolidBrush(vbRed)
'    hBrBlue = CreateSolidBrush(vbBlue)
'
'    hBrOld = SelectObject(DemoDibWork.DeviceContext, hBrRed)
'
'    For j = 0 To rc.Bottom Step 16
'        For i = 0 To rc.Right Step 16
'            If (i \ 16 + j \ 16) Mod 2 = 0 Then
'                SelectObject DemoDibWork.DeviceContext, hBrRed
'            Else
'                SelectObject DemoDibWork.DeviceContext, hBrBlue
'            End If
'            Rectangle DemoDibWork.DeviceContext, i, j, i + 16, j + 16
'        Next
'    Next
'
'    SelectObject DemoDibWork.DeviceContext, hBrOld
'
'    DeleteObject hBrBlue
'    DeleteObject hBrRed
    
    SetTextColor DemoDibWork.DeviceContext, vbWhite
    SetBkMode DemoDibWork.DeviceContext, TRANSPARENT
    hFontOld = SelectObject(DemoDibWork.DeviceContext, hFont)
    DrawText DemoDibWork.DeviceContext, sText, Len(sText), rc, DT_CENTER Or DT_MULTILINE Or DT_TOP
    SelectObject DemoDibWork.DeviceContext, hFontOld
    
    DeleteObject hFont
    
    ' Build the lookup table
    Dim dz As Double
    Dim W As Double, minW As Double, maxW As Double
    Dim x As Double, y As Double, t As Double, tmp As Double
    Dim ang As Double, dist As Double
        
    For j = 0 To rcView.Bottom
        For i = 0 To rcView.Right
            ' Both coords from -1 to 1
            x = 2# * (CDbl(i) / CDbl(rcView.Right) - 0.5)
            y = 2# * (CDbl(j) / CDbl(rcView.Bottom) - 0.5)
            dist = Sqr(x * x + y * y)
            ang = ATan2(y, x)
            
            ang = ang + Atn(1) * 0.25 * Sin((1 - (y + 1) * 0.5) * Atn(1) * 8)

            'ang = ang + dist * Atn(1) * 0.5
            If y < 0 Then
                'ang = ang + Atn(1) * y * (1 - x)
            End If

            x = Cos(ang) * dist
            y = Sin(ang) * dist

            t = ((x * y * y) + (dist * 0.5 + 0.5)) * 0.5 ' 0 <= t <= 1
            t = 1 - t
            t = t * Atn(1) * 4 * 2      ' 0 <= t <= 2pi
            tmp = x + Sin(t * 5) * 0.02

            t = ((x * x * y) + (dist * 0.5 + 0.5)) * 0.5 ' 0 <= t <= 1
            t = 1 - t
            t = t * Atn(1) * 4 * 2      ' 0 <= t <= 2pi
            y = y + Sin(t * 4) * 0.03
            
            x = tmp


            ' Convert from (-1:1) to (0:1)
            y = (y + 1) * 0.5

            x = x / (0.3 + y * (1# - 0.3))
            y = y / (0.7 + y * (1# - 0.7))

            x = (x + 1) * 0.5


            ' Now convert back, and scale the x to match...
            DemoLookup(i, j).Left = 256# * rcView.Right * x * CDbl(rc.Right) / CDbl(rcView.Right)
            DemoLookup(i, j).Top = 256# * rcView.Bottom * y * CDbl(rc.Right) / CDbl(rcView.Right)

            ' Bound the values
            If DemoLookup(i, j).Left \ 256& < 0 Then DemoLookup(i, j).Left = 0
            If DemoLookup(i, j).Left \ 256& >= rc.Right Then DemoLookup(i, j).Left = 0
            If DemoLookup(i, j).Top \ 256& < 0 Then DemoLookup(i, j).Top = 0
            If DemoLookup(i, j).Top \ 256& >= rc.Bottom Then DemoLookup(i, j).Top = 0
        Next
    Next
    
    ' Initialize the stars
    If DemoPosition = 0 Then
        For i = 0 To CountStars - 1
            With DemoStars(i)
                .x = rcView.Right * Rnd
                .y = rcView.Bottom * Rnd
                
                .dx = .x - rcView.Right * 0.5
                .dy = .y - rcView.Bottom * 0.5
                x = Sqr(.dx * .dx + .dy * .dy)
                W = Rnd * 0.75 + 0.25
                .brightness = CLng(W * 255)
                .dx = (3# * W) * .dx / x
                .dy = (3# * W) * .dy / x
            End With
        Next
    End If
    
    ' Set up the Alpha calcs lookups
    For i = 0 To 255
        For j = 0 To 255
            ' Scaled Multiplication
            AlphaMulLUT(i, j) = (CLng(i) * CLng(j)) \ 255&
            ' Capped addition
            k = CLng(i) + CLng(j)
            If k > 255 Then k = 255
            AlphaAddLUT(i, j) = k
            ' Capped subtraction
            k = CLng(i) - CLng(j)
            If k < 0 Then k = 0
            AlphaSubLUT(i, j) = k
            ' Fade in : linear
            FadeInLUT(i, j) = AlphaMulLUT(i, j)
            ' Fade out : sine
            FadeOutLUT(i, j) = (CDbl(i) * Sin((CDbl(j) / 255#) * Atn(1) * 2#))
        Next
    Next
    
    ' Text styles for FPS
    
    Dim hDC As Long
    hDC = GetDC(hWnd)
    hFont = GetCurrentObject(hDC, OBJ_FONT)
    GetObject hFont, Len(lf), lf
    ReleaseDC hWnd, hDC
    hFont = CreateFontIndirect(lf)
    hFontOld = SelectObject(DemoDibBackBuf.DeviceContext, hFont)
    DeleteObject hFontOld
    
    SetBkMode DemoDibBackBuf.DeviceContext, TRANSPARENT
    SetTextColor DemoDibBackBuf.DeviceContext, vbGreen
    
    
    DemoPosition = 0

End Sub

' Move one step forward and update display
Public Sub Step()
    Static InProc As Boolean
    If InProc Then Exit Sub
    InProc = True
    
    DemoDibWork.GetBits ArrWork, VarPtr(saWork), SrcW, SrcH     ' Source buffer with text
    DemoDibBackBuf.GetBits ArrDraw, VarPtr(saDraw), DstW, DstH  ' Main drawing buffer
    DemoDibText.GetBits ArrText, VarPtr(saText), DstW, DstH     ' Buffer for drawing text
    DemoDibStars.GetBits ArrStars, VarPtr(saStars), DstW, DstH  ' Buffer for drawing stars
    DemoDibTemp.GetBits ArrTemp, VarPtr(saTemp), DstW, DstH     ' General temporary buffer


    Dim i As Long, j As Long, k As Long, l As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim x0 As Long, y0 As Long
    Dim X1 As Long, Y1 As Long
    Dim tmpf As Double, tmpg As Double
    
    Dim StartTime As Long
    Dim TextTime As Long
    Dim StarsTime As Long
    Dim EndTime As Long
    
    'ODS "\nBeginning Timer Procedure\n\n"
    
    'Do
    
    StartTime = timeGetTime

    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            x0 = DemoLookup(i \ 4, j).Left \ 256
            y0 = DemoLookup(i \ 4, j).Top \ 256

            X1 = DemoLookup(i \ 4, j).Left And 255
            Y1 = DemoLookup(i \ 4, j).Top And 255

            y0 = (y0 + DemoPosition) Mod SrcH
            x0 = x0 * 4&

            If (x0 \ 4& < SrcW - 1 And y0 < SrcH - 1) Then
                ' Calculate merged at x
                r = AlphaMulLUT(ArrWork(x0 + 2, y0), 255 - X1) + AlphaMulLUT(ArrWork(x0 + 2 + 4, y0), X1)
                g = AlphaMulLUT(ArrWork(x0 + 1, y0), 255 - X1) + AlphaMulLUT(ArrWork(x0 + 1 + 4, y0), X1)
                b = AlphaMulLUT(ArrWork(x0 + 0, y0), 255 - X1) + AlphaMulLUT(ArrWork(x0 + 0 + 4, y0), X1)
                ' Add merged at y
                r = r + AlphaMulLUT(ArrWork(x0 + 2, y0), 255 - Y1) + AlphaMulLUT(ArrWork(x0 + 2, y0 + 1), Y1)
                g = g + AlphaMulLUT(ArrWork(x0 + 1, y0), 255 - Y1) + AlphaMulLUT(ArrWork(x0 + 1, y0 + 1), Y1)
                b = b + AlphaMulLUT(ArrWork(x0 + 0, y0), 255 - Y1) + AlphaMulLUT(ArrWork(x0 + 0, y0 + 1), Y1)
                ' Average
                r = r \ 2
                g = g \ 2
                b = b \ 2
                ' Maximum is alpha
                a = r
                If b > a Then a = b
                If g > a Then a = g
                ' Store
                ArrText(i + 3, j) = a
                ArrText(i + 2, j) = r
                ArrText(i + 1, j) = g
                ArrText(i + 0, j) = b
            Else
                r = ArrWork(x0 + 2, y0)
                g = ArrWork(x0 + 1, y0)
                b = ArrWork(x0 + 0, y0)
                ' Alpha
                a = r
                If b > a Then a = b
                If g > a Then a = g
                ' Store
                ArrText(i + 3, j) = a
                ArrText(i + 2, j) = r
                ArrText(i + 1, j) = g
                ArrText(i + 0, j) = b
            End If
        Next
    Next
    'ODS "Blurring Text\n"
    Renderer.Blur3x3 ArrText(0, 0), ArrTemp(0, 0), DstW - 1, DstH - 1, DemoDibBackBuf.BytesPerScanline
    ' Overlay the new text on the old text with alpha-blending
    'ODS "Overlaying text"
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            k = ArrTemp(i + 3, j)
            If k > 0 Then
                k = 255 - AlphaMulLUT(255 - k, 255 - k)
                ArrText(i + 2, j) = AlphaMulLUT(ArrText(i + 2, j), k) + AlphaMulLUT(ArrTemp(i + 2, j), 255 - k)
                ArrText(i + 1, j) = AlphaMulLUT(ArrText(i + 1, j), k) + AlphaMulLUT(ArrTemp(i + 1, j), 255 - k)
                ArrText(i + 0, j) = AlphaMulLUT(ArrText(i + 0, j), k) + AlphaMulLUT(ArrTemp(i + 0, j), 255 - k)
                ArrText(i + 3, j) = k
            End If
        Next
    Next

    ' Now handle the stars
    TextTime = timeGetTime

    'ODS "Fading stars\n"
    ' Fade the old stars slightly (to the temp)
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            'ArrTemp(i + 2, j) = AlphaMulLUT(ArrStars(i + 2, j), 240)
            'ArrTemp(i + 1, j) = AlphaMulLUT(ArrStars(i + 1, j), 240)
            'ArrTemp(i + 0, j) = AlphaMulLUT(ArrStars(i + 0, j), 240)
            ArrTemp(i + 3, j) = AlphaMulLUT(ArrStars(i + 3, j), 240)
        Next
    Next
    'ODS "Blurring stars\n"
    ' Blur from the temp buffer back to the stars buffer
'    Renderer.Blur3x3 ArrTemp(0, 0), ArrStars(0, 0), DstW - 1, DstH - 1, DemoDibBackBuf.BytesPerScanline
'    Renderer.Blur3x3 ArrStars(0, 0), ArrTemp(0, 0), DstW - 1, DstH - 1, DemoDibBackBuf.BytesPerScanline
    Renderer.Blur3x3 ArrTemp(0, 0), ArrStars(0, 0), DstW - 1, DstH - 1, DemoDibBackBuf.BytesPerScanline

    'ODS "Updating stars\n"
    ' Update the stars and add them to the temp buffer
    For i = 0 To CountStars - 1
        With DemoStars(i)
            .x = .x + .dx
            .y = .y + .dy
            .dx = .dx * 1.075
            .dy = .dy * 1.075 '+ 0.15

            x0 = CLng(.x + 0.5) * 4&
            y0 = CLng(.y + 0.5)
            If x0 < 0 Or y0 < 0 Or x0 >= DstW Or y0 >= DstH Then
                ' reset
                .x = CSng(DstW) * 0.5 * 0.25
                .y = CSng(DstH) * 0.5

                tmpf = Rnd * Atn(1) * 4 * 2
                tmpg = Rnd * 0.75 + 0.25
                .brightness = CLng(tmpg * 255)
                .dx = Cos(tmpf) * (3#) * tmpg
                .dy = Sin(tmpf) * (3#) * tmpg

                .x = .x + .dx * tmpg * 3
                .x = .x + .dy * tmpg * 3
            Else
                ArrTemp(x0 + 3, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 3, y0))
                'ArrTemp(x0 + 2, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 2, y0))
                'ArrTemp(x0 + 1, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 1, y0))
                'ArrTemp(x0 + 0, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 0, y0))
            End If
        End With
    Next
    ' Blur from the temp buffer back to the stars buffer
    Renderer.Blur3x3 ArrTemp(0, 0), ArrStars(0, 0), DstW - 1, DstH - 1, DemoDibBackBuf.BytesPerScanline
    ' Add the stars again
    For i = 0 To CountStars - 1
        With DemoStars(i)
            x0 = CLng(.x + 0.5) * 4&
            y0 = CLng(.y + 0.5)
            If x0 >= 0 And y0 >= 0 And x0 < DstW And y0 < DstH Then
                ArrStars(x0 + 3, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 3, y0))
                'ArrStars(x0 + 2, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 2, y0))
                'ArrStars(x0 + 1, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 1, y0))
                'ArrStars(x0 + 0, y0) = AlphaAddLUT(.brightness, ArrStars(x0 + 0, y0))
            End If
        End With
    Next
    ' Calculate star colour based on alpha (todo: convert to palette?)
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            k = ArrStars(i + 3, j)
            If k < 128 Then
                ArrStars(i + 2, j) = k * 2
                ArrStars(i + 1, j) = 0
                ArrStars(i + 0, j) = 0
            Else
                ArrStars(i + 2, j) = 255
                ArrStars(i + 1, j) = (k - 128) * 2
                ArrStars(i + 0, j) = 0
            End If
        Next
    Next


    StarsTime = timeGetTime
    'ODS "Merging\n"

    ' Merge the stars and text buffers into the draw buffer [fade stars a bit]
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            k = ArrText(i + 3, j)
            ArrDraw(i + 2, j) = AlphaMulLUT(ArrText(i + 2, j), k) + AlphaMulLUT(ArrStars(i + 2, j), (255 - k))
            ArrDraw(i + 1, j) = AlphaMulLUT(ArrText(i + 1, j), k) + AlphaMulLUT(ArrStars(i + 1, j), (255 - k))
            ArrDraw(i + 0, j) = AlphaMulLUT(ArrText(i + 0, j), k) + AlphaMulLUT(ArrStars(i + 0, j), (255 - k))
        Next
    Next
    
    'ODS "Applying time-based fade\n"

    If DemoPosition < 65 Then
        k = CLng(Sqr(255& * DemoPosition * 4&)) Mod 256
        For j = 0 To DstH - 1
            For i = 0 To DstW - 1 Step 4
                ArrDraw(i + 3, j) = FadeInLUT(ArrDraw(i + 3, j), k)
                ArrDraw(i + 2, j) = FadeInLUT(ArrDraw(i + 2, j), k)
                ArrDraw(i + 1, j) = FadeInLUT(ArrDraw(i + 1, j), k)
                ArrDraw(i + 0, j) = FadeInLUT(ArrDraw(i + 0, j), k)
            Next
        Next
    End If
    If DemoPosition > SrcH - 64 Then
        k = (((SrcH - DemoPosition) * 4& * (SrcH - DemoPosition) * 4&) \ 256&) Mod 256
        For j = 0 To DstH - 1
            For i = 0 To DstW - 1 Step 4
                ArrDraw(i + 3, j) = FadeOutLUT(ArrDraw(i + 3, j), k)
                ArrDraw(i + 2, j) = FadeOutLUT(ArrDraw(i + 2, j), k)
                ArrDraw(i + 1, j) = FadeOutLUT(ArrDraw(i + 1, j), k)
                ArrDraw(i + 0, j) = FadeOutLUT(ArrDraw(i + 0, j), k)
            Next
        Next
    End If
    
    EndTime = timeGetTime
    Dim s As String
        
    On Error Resume Next
    s = "FPS: "
    s = s & CLng(1000# / CSng(EndTime - StartTime) + 0.5)
    TextOut DemoDibBackBuf.DeviceContext, 2, 2, s, Len(s)
'
'    s = "Text: " & vbTab
'    s = s & (TextTime - StartTime)
'    TextOut DemoDibBackBuf.DeviceContext, 0, 16, s, Len(s)
'
'
'    s = "Stars: " & vbTab
'    s = s & (StarsTime - TextTime)
'    TextOut DemoDibBackBuf.DeviceContext, 0, 32, s, Len(s)

    
    ' Finally, output
    Dim hDC As Long
    hDC = GetDC(hWnd)
    BitBlt hDC, 0, -2, DstW, DstH, DemoDibBackBuf.DeviceContext, 0, 0, vbSrcCopy
    ReleaseDC hWnd, hDC
    
    DemoPosition = DemoPosition + 1
    

    'ODS "All done!\n"
    DemoDibWork.FreeBits ArrWork
    DemoDibBackBuf.FreeBits ArrDraw
    DemoDibText.FreeBits ArrText
    DemoDibStars.FreeBits ArrStars
    DemoDibTemp.FreeBits ArrTemp
    
    If DemoPosition > SrcH Then
        ' Initialize (same window)
        Init hWnd
    End If
    
    InProc = False
    
    'ODS "Exiting procedure\n"
End Sub

Private Sub Class_Terminate()
    ' Free arrays
    If hWnd = 0 Then Exit Sub
    On Error Resume Next

End Sub
