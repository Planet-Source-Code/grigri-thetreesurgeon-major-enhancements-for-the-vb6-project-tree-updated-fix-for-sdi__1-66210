VERSION 5.00
Begin VB.Form SettingsDlg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TheTreeSurgeon - Settings"
   ClientHeight    =   4170
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPane 
      Caption         =   "About The Tree Surgeon"
      Height          =   3615
      Index           =   3
      Left            =   2760
      TabIndex        =   15
      Top             =   0
      Width           =   4215
      Begin VB.PictureBox picDemo 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   32
         Top             =   1800
         Width           =   3975
         Begin VB.Timer tmrDemo 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   2400
            Top             =   240
         End
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "grigri@shinyhappypixels.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Tag             =   "mailto:grigri@shinyhappypixels.com"
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "Written by grigri, 2006"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "for VB6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "The Tree Surgeon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame fraPane 
      Caption         =   "General Settings"
      Height          =   3615
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox chkSetting 
         Caption         =   "Right-click on icon shows shell context menu"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Tag             =   "7"
         Top             =   360
         Width           =   3495
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Show tree lines"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   13
         Tag             =   "8"
         Top             =   720
         Width           =   3495
      End
   End
   Begin VB.ListBox lstPanes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      IntegralHeight  =   0   'False
      ItemData        =   "Config.frx":000C
      Left            =   120
      List            =   "Config.frx":001C
      TabIndex        =   2
      Top             =   75
      Width           =   2535
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraPane 
      Caption         =   "Icon Settings"
      Height          =   3615
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton btnCustomIcons 
         Caption         =   "View / Edit per-component icons"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   2280
         Width           =   3735
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Allow per-component icons"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Tag             =   "13"
         Top             =   1920
         Width           =   3975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Display shell icon for all components"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   28
         Tag             =   "12"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Display form icon if available and same size"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Tag             =   "9"
         Top             =   960
         Width           =   3975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Display shell icon for related documents"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Tag             =   "2"
         Top             =   240
         Width           =   3975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Display shell icon for resource files"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Tag             =   "3"
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "(useful if you have custom IconHandlers)"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1560
         Width           =   3615
      End
   End
   Begin VB.Frame fraPane 
      Caption         =   "Text Settings"
      Height          =   3615
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.OptionButton opAntiAlias 
         Caption         =   "Natural ClearType"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   25
         Tag             =   "6"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.OptionButton opAntiAlias 
         Caption         =   "ClearType"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   24
         Tag             =   "5"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.OptionButton opAntiAlias 
         Caption         =   "Standard"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   23
         Tag             =   "4"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.OptionButton opAntiAlias 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   22
         Tag             =   "3"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkAntiAlias 
         Caption         =   "Force Anti-Aliasing"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   3975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Show unsaved items in italics"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Tag             =   "0"
         Top             =   240
         Width           =   3135
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Show startup items in bold"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Tag             =   "1"
         Top             =   600
         Width           =   3135
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Show filenames in tree"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Tag             =   "4"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Show path relative to project folder"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   9
         Tag             =   "5"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Filenames in grey"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   8
         Tag             =   "6"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Font"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   975
      End
   End
End
Attribute VB_Name = "SettingsDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ts As TreeSurgeon

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Const IDC_HAND As Long = 32649&

Private hLinkCursor As Long
Private hOldCursor As Long

'Private OldSettings() As Boolean
Private TmpSettings() As String
Private Validated As Boolean

Private Type STARINFO
    x As Double
    y As Double
    dx As Double
    dy As Double
    brightness As Byte
End Type

Private Const CountStars As Long = 250&

Private DemoStars(0 To CountStars - 1) As STARINFO

Private DemoDibBackBuf As DibSection32
Private DemoDibTempBuf As DibSection32
Private DemoDibWork As DibSection32
Private DemoLookup() As RECT
Private DemoText() As String
Private DemoPosition As Long

Public Sub EditSettings(Settings() As String, ts2 As TreeSurgeon)
    Set ts = ts2
    Dim i As Long
    
    Load Me
    
    ReDim TmpSettings(LBound(Settings) To UBound(Settings))
    For i = LBound(TmpSettings) To UBound(TmpSettings)
        TmpSettings(i) = Settings(i)
    Next
    
    UpdateControls
    
    Me.Show vbModal
    
    If Validated Then
        For i = LBound(TmpSettings) To UBound(TmpSettings)
            Settings(i) = TmpSettings(i)
        Next
    End If
    
    Unload Me
    
    Set ts = Nothing
End Sub

Private Sub btnCancel_Click()
    Validated = False
    Me.Visible = False
End Sub

Private Sub btnOK_Click()
    Validated = True
    Me.Visible = False
End Sub

Private Sub cboFont_Change()
    Static Inside As Boolean
    If Inside Then Exit Sub
    Inside = True
    
    '
    Dim iPos As Long
    Dim i As Long, j As Long
    iPos = FindInCombo(cboFont.hWnd, cboFont.Text)
    If iPos = -1 Then
        
    Else
        i = Len(cboFont.Text)
        cboFont.Text = cboFont.List(iPos)
        cboFont.SelStart = i
        cboFont.SelLength = Len(cboFont.Text) - i + 1
    End If
    '
    
    Inside = False
End Sub

Private Sub cboFont_GotFocus()
    If Len(cboFont.Tag) > 0 Then
        cboFont.Text = cboFont.Tag
        cboFont.SelStart = 1
        cboFont.SelLength = Len(cboFont.Text)
    End If
End Sub

Private Sub cboFont_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 9
End Sub

Private Sub cboFont_LostFocus()
    Dim iPos As Long
    iPos = FindInCombo(cboFont.hWnd, cboFont.Text, True)
    If iPos = -1 Then
        cboFont.Tag = cboFont.Text
        cboFont.Text = "(default)"
        TmpSettings([Tree Font]) = ""
    Else
        cboFont.Tag = ""
        TmpSettings([Tree Font]) = cboFont.Text
    End If
End Sub

Private Sub chkAntiAlias_Click()
    Dim i As Long
    If chkAntiAlias.Value = 0 Then
        For i = opAntiAlias.LBound To opAntiAlias.ubound
            opAntiAlias(i).Enabled = False
        Next
        TmpSettings([AntiAlias Method]) = 0
    Else
        For i = opAntiAlias.LBound To opAntiAlias.ubound
            opAntiAlias(i).Enabled = True
        Next
    End If
End Sub

Private Sub chkSetting_Click(Index As Integer)
    Dim Which As SettingsEntries
    Which = CLng(chkSetting(Index).Tag)
    
    If Which = FileNames Then
        If chkSetting(Index).Value = 1 Then
            chkSetting(SettingsEntries.[FileNames in Grey]).Enabled = True
            chkSetting(SettingsEntries.[FileNames Relative Paths]).Enabled = True
        Else
            chkSetting(SettingsEntries.[FileNames in Grey]).Enabled = False
            chkSetting(SettingsEntries.[FileNames Relative Paths]).Enabled = False
        End If
    End If
    
    TmpSettings(Which) = IIf(chkSetting(Index).Value = 1, "1", "0")
End Sub

Private Sub UpdateControls()
    Dim chkBox As CheckBox
    For Each chkBox In chkSetting
        If TmpSettings(chkBox.Tag) Then
            chkBox.Value = 1
        Else
            chkBox.Value = 0
        End If
    Next
    
    Dim i As Long
    
    If TmpSettings([AntiAlias Method]) = 0 Then
        chkAntiAlias.Value = 0
        For i = opAntiAlias.LBound To opAntiAlias.ubound
            opAntiAlias(i).Value = False
            opAntiAlias(i).Enabled = False
        Next
    Else
        chkAntiAlias.Value = 1
        For i = opAntiAlias.LBound To opAntiAlias.ubound
            opAntiAlias(i).Value = (TmpSettings([AntiAlias Method]) = opAntiAlias(i).Tag)
            opAntiAlias(i).Enabled = True
        Next
    End If
    
    cboFont.Text = TmpSettings([Tree Font])
End Sub

Private Sub btnCustomIcons_Click()
    Dim dlg As New PerComponentIcons
    
    dlg.ManageComponentIcons ts
End Sub

Private Sub Form_Load()
    lstPanes.ListIndex = 0
    hLinkCursor = LoadCursor(0, IDC_HAND)
    
    InitDemo
    
    GetFonts
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrDemo.Enabled = False
    Set DemoDibBackBuf = Nothing
    Set DemoDibWork = Nothing
    Set DemoDibTempBuf = Nothing
End Sub

Private Sub lblEmail_Click()
    ShellExecute hWnd, "open", lblEmail.Tag, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor hLinkCursor
End Sub

Private Sub lstPanes_Click()
    If lstPanes.ListIndex = -1 Then Exit Sub
    On Error Resume Next
    fraPane(lstPanes.ListIndex).ZOrder
    If lstPanes.ListIndex = 3 Then
        tmrDemo.Enabled = True
    Else
        tmrDemo.Enabled = False
    End If
    On Error GoTo 0
End Sub

Private Sub opAntiAlias_Click(Index As Integer)
    If chkAntiAlias.Value = 0 Then Exit Sub
    Dim i As Long
    For i = opAntiAlias.LBound To opAntiAlias.ubound
        If opAntiAlias(i).Value Then
            TmpSettings([AntiAlias Method]) = opAntiAlias(i).Tag
            Exit Sub
        End If
    Next
End Sub

Private Sub GetFonts()
    cboFont.Clear
    Dim arrFonts() As String
    arrFonts = ListFonts(hDC)
    
    Dim i As Long
    
    For i = 0 To UBound(arrFonts)
        cboFont.AddItem arrFonts(i)
    Next
    
    cboFont.AddItem "(Default)", 0
End Sub

Private Sub InitDemo()
    Randomize Timer
    
    Set DemoDibBackBuf = New DibSection32
    DemoDibBackBuf.Create picDemo.ScaleWidth, picDemo.ScaleHeight
    Set DemoDibTempBuf = New DibSection32
    DemoDibTempBuf.Create picDemo.ScaleWidth, picDemo.ScaleHeight

    On Error GoTo NO_TEXT_FILE
    Dim nFile As Integer
    Dim sText As String
    nFile = FreeFile
    Open App.Path & "\Scrolltext2.txt" For Input As nFile
    sText = Input$(LOF(nFile), nFile)
    Close nFile
    GoTo SKIP_ERROR_HANDLER
    
NO_TEXT_FILE:
    ODS "Error loading file (handle=" & nFile & ": " & vbCrLf & "> " & Err.Description & vbCrLf
    sText = Replace$("                                  |come closer and see|see into the trees|find the girl|while you can|come closer and see|see into the dark|just follow your eyes|just follow your eyes||i hear her voice|calling my name|the sound is deep|in the dark|i hear her voice|and start to run|into the trees|into the trees||into the trees||suddenly i stop|but i know it's too late|i'm lost in a forest|all alone|the girl was never there|it's always the same|i'm running towards nothing|again and again and again|", "|", vbCrLf)
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
    rc.Right = picDemo.ScaleWidth
    rc.Bottom = picDemo.ScaleHeight
    ReDim DemoLookup(0 To rc.Right, 0 To rc.Bottom)
    hFontOld = SelectObject(DemoDibBackBuf.DeviceContext, hFont)
    DrawText DemoDibBackBuf.DeviceContext, sText, Len(sText), rc, DT_CALCRECT Or DT_MULTILINE Or DT_CENTER Or DT_TOP
    SelectObject DemoDibBackBuf.DeviceContext, hFontOld
    
    Set DemoDibWork = New DibSection32
    rc.Bottom = rc.Bottom + picDemo.ScaleHeight
    DemoDibWork.Create rc.Right, rc.Bottom
    
    ODS "Total height : " & rc.Bottom
    
    rc.Top = rc.Top + picDemo.ScaleHeight
    SetTextColor DemoDibWork.DeviceContext, vbGreen
    SetBkMode DemoDibWork.DeviceContext, TRANSPARENT
    hFontOld = SelectObject(DemoDibWork.DeviceContext, hFont)
    DrawText DemoDibWork.DeviceContext, sText, Len(sText), rc, DT_CENTER Or DT_MULTILINE Or DT_TOP
    SelectObject DemoDibWork.DeviceContext, hFontOld
'    hPenOld = SelectObject(DemoDibWork.DeviceContext, GetStockObject(WHITE_PEN))
'    hBrushOld = SelectObject(DemoDibWork.DeviceContext, GetStockObject(NULL_BRUSH))
'    Rectangle DemoDibWork.DeviceContext, rc.Left, rc.Top, rc.Right - 1, rc.Bottom - 1
'    Rectangle DemoDibWork.DeviceContext, rc.Left + 3, rc.Top + 3, rc.Right - 4, rc.Bottom - 4
'    SelectObject DemoDibWork.DeviceContext, hPenOld
'    SelectObject DemoDibWork.DeviceContext, hBrushOld
    
    DeleteObject hFont
    
    ' Build the lookup table
    Dim i As Long, j As Long
    Dim dz As Double
    Dim w As Double, minW As Double, maxW As Double
    Dim x As Double
    
    minW = 1.3
    maxW = 0.5
    For j = 0 To picDemo.ScaleHeight
        w = minW + (CDbl(j) / picDemo.ScaleHeight) ^ 0.9 * (maxW - minW)
        For i = 0 To picDemo.ScaleWidth
            dz = CDbl(rc.Right) / CDbl(picDemo.ScaleWidth)
            
            ' Make X go from -1 to 1
            x = 2# * (CDbl(i) / CDbl(picDemo.ScaleWidth) - 0.5)
            ' Adjust by scale
            x = x * w
            
            DemoLookup(i, j).Left = CLng(0.5 * rc.Right + x * rc.Right * 0.5 + 5# * Sin(2# * (CDbl(j) / picDemo.ScaleHeight) * Atn(1) * 4))
            DemoLookup(i, j).Top = CLng(CDbl(j) + 5# * Sin(2# * (CDbl(i) / picDemo.ScaleWidth) * (CDbl(j) / picDemo.ScaleHeight) * Atn(1) * 4))
            If DemoLookup(i, j).Left < 0 Then DemoLookup(i, j).Left = 0
            If DemoLookup(i, j).Left >= rc.Right Then DemoLookup(i, j).Left = 0
            If DemoLookup(i, j).Top < 0 Then DemoLookup(i, j).Top = 0
            If DemoLookup(i, j).Top >= rc.Bottom Then DemoLookup(i, j).Top = 0
            
            ' Calculate alpha
            x = (CDbl(j) / picDemo.ScaleHeight) ' From 0 to 1
            x = 2# * (x - 0.5)                  ' From -1 to 1
            x = x * Sin(6# * (CDbl(i) / picDemo.ScaleWidth) * (CDbl(j) / picDemo.ScaleHeight) * Atn(1) * 4)
            x = x * x                           ' Extremes ( from 0 to 1 )
            DemoLookup(i, j).Right = 255& - CLng(x * 255#)
        Next
    Next
    
    If DemoPosition = 0 Then
        For i = 0 To CountStars - 1
            With DemoStars(i)
                .x = picDemo.ScaleWidth * Rnd
                .y = picDemo.ScaleHeight * Rnd
                
                .dx = .x - picDemo.ScaleWidth * 0.5
                .dy = .y - picDemo.ScaleHeight * 0.5
                x = Sqr(.dx * .dx + .dy * .dy)
                w = Rnd * 0.75 + 0.25
                .brightness = CLng(w * 255)
                .dx = (3# * w) * .dx / x
                .dy = (3# * w) * .dy / x
            End With
        Next
    End If
    
    DemoPosition = 0
End Sub

Private Sub tmrDemo_Timer()
    ' Just copy for now
    Dim ArrSrc() As Byte, ArrDst() As Byte, ArrTmp() As Byte
    Dim sa1 As SAFEARRAY2D, sa2 As SAFEARRAY2D, sa3 As SAFEARRAY2D

    On Error Resume Next
    
    Dim SrcW As Long, SrcH As Long
    DemoDibWork.GetBits ArrSrc, sa1, SrcW, SrcH
    
    Dim DstW As Long, DstH As Long
    'DemoDibBackBuf.ZeroBits
    DemoDibTempBuf.GetBits ArrTmp, sa3, DstW, DstH
    DemoDibBackBuf.GetBits ArrDst, sa2, DstW, DstH
    
    Dim i As Long, j As Long, k As Long
    Dim x0 As Long, y0 As Long
    Dim tmpf As Single, tmpg As Single
    
'    For j = 0 To DstH - 1
'        For i = 0 To DstW - 1 Step 4
'            ArrDst(i + 3, j) = ArrDst(i + 3, j) \ 2
'            ArrDst(i + 2, j) = ArrDst(i + 2, j) \ 2
'            ArrDst(i + 1, j) = ArrDst(i + 1, j) \ 2
'            ArrDst(i + 0, j) = ArrDst(i + 0, j) \ 2
'        Next
'    Next


    'ODS "Copying..." & vbCrLf & "----------" & vbCrLf
    On Error GoTo SKIP
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            x0 = DemoLookup(i \ 4, j).Left
            y0 = (DemoLookup(i \ 4, j).Top + DemoPosition) Mod SrcH
            x0 = x0 * 4&
            ArrDst(i + 3, j) = CLng(ArrSrc(x0 + 3, y0)) \ 2 + ArrDst(i + 3, j) \ 2
            ArrDst(i + 2, j) = CLng(ArrSrc(x0 + 2, y0)) \ 2 + ArrDst(i + 2, j) \ 2
            ArrDst(i + 1, j) = CLng(ArrSrc(x0 + 1, y0)) \ 2 + ArrDst(i + 1, j) \ 2
            ArrDst(i + 0, j) = CLng(ArrSrc(x0 + 0, y0)) \ 2 + ArrDst(i + 0, j) \ 2
        Next
    Next
    
    
    
    For i = 0 To CountStars - 1
        With DemoStars(i)
            .x = .x + .dx
            .y = .y + .dy
'            tmpf = .dx * 0.999657324975557 - .dy * 2.61769483078731E-02
'            .dy = .dy * 0.999657324975557 + .dx * 2.61769483078731E-02
'            .dx = tmpf
.dx = .dx * 1.05
.dy = .dy * 1.05
            
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
                ArrDst(x0 + 3, y0) = 255
                ArrDst(x0 + 2, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 2, y0)), 255)
                ArrDst(x0 + 1, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 1, y0)), 255)
                ArrDst(x0 + 0, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 0, y0)), 255)
            End If
        End With
    Next
    
    ' Blur horizontal
    For j = 0 To DstH - 1
        For i = 4 To DstW - 5 Step 4
            ArrTmp(i + 3, j) = (2& * CLng(ArrDst(i + 3, j)) + CLng(ArrDst(i + 3 - 4, j)) + CLng(ArrDst(i + 3 + 4, j))) \ 4
            ArrTmp(i + 2, j) = (2& * CLng(ArrDst(i + 2, j)) + CLng(ArrDst(i + 2 - 4, j)) + CLng(ArrDst(i + 2 + 4, j))) \ 4
            ArrTmp(i + 1, j) = (2& * CLng(ArrDst(i + 1, j)) + CLng(ArrDst(i + 1 - 4, j)) + CLng(ArrDst(i + 1 + 4, j))) \ 4
            ArrTmp(i + 0, j) = (2& * CLng(ArrDst(i + 0, j)) + CLng(ArrDst(i + 0 - 4, j)) + CLng(ArrDst(i + 0 + 4, j))) \ 4
        Next
    Next
    ' Blur vertical
    For j = 1 To DstH - 2
        For i = 0 To DstW - 1 Step 4
            ArrDst(i + 3, j) = (2& * CLng(ArrTmp(i + 3, j)) + CLng(ArrTmp(i + 3, j - 1)) + CLng(ArrTmp(i + 3, j + 1))) \ 4
            ArrDst(i + 2, j) = (2& * CLng(ArrTmp(i + 2, j)) + CLng(ArrTmp(i + 2, j - 1)) + CLng(ArrTmp(i + 2, j + 1))) \ 4
            ArrDst(i + 1, j) = (2& * CLng(ArrTmp(i + 1, j)) + CLng(ArrTmp(i + 1, j - 1)) + CLng(ArrTmp(i + 1, j + 1))) \ 4
            ArrDst(i + 0, j) = (2& * CLng(ArrTmp(i + 0, j)) + CLng(ArrTmp(i + 0, j - 1)) + CLng(ArrTmp(i + 0, j + 1))) \ 4
        Next
    Next
    ' Blur horizontal
    For j = 0 To DstH - 1
        For i = 4 To DstW - 5 Step 4
            ArrTmp(i + 3, j) = (2& * CLng(ArrDst(i + 3, j)) + CLng(ArrDst(i + 3 - 4, j)) + CLng(ArrDst(i + 3 + 4, j))) \ 4
            ArrTmp(i + 2, j) = (2& * CLng(ArrDst(i + 2, j)) + CLng(ArrDst(i + 2 - 4, j)) + CLng(ArrDst(i + 2 + 4, j))) \ 4
            ArrTmp(i + 1, j) = (2& * CLng(ArrDst(i + 1, j)) + CLng(ArrDst(i + 1 - 4, j)) + CLng(ArrDst(i + 1 + 4, j))) \ 4
            ArrTmp(i + 0, j) = (2& * CLng(ArrDst(i + 0, j)) + CLng(ArrDst(i + 0 - 4, j)) + CLng(ArrDst(i + 0 + 4, j))) \ 4
        Next
    Next
    ' Blur vertical
    For j = 1 To DstH - 2
        For i = 0 To DstW - 1 Step 4
            ArrDst(i + 3, j) = (2& * CLng(ArrTmp(i + 3, j)) + CLng(ArrTmp(i + 3, j - 1)) + CLng(ArrTmp(i + 3, j + 1))) \ 4
            ArrDst(i + 2, j) = (2& * CLng(ArrTmp(i + 2, j)) + CLng(ArrTmp(i + 2, j - 1)) + CLng(ArrTmp(i + 2, j + 1))) \ 4
            ArrDst(i + 1, j) = (2& * CLng(ArrTmp(i + 1, j)) + CLng(ArrTmp(i + 1, j - 1)) + CLng(ArrTmp(i + 1, j + 1))) \ 4
            ArrDst(i + 0, j) = (2& * CLng(ArrTmp(i + 0, j)) + CLng(ArrTmp(i + 0, j - 1)) + CLng(ArrTmp(i + 0, j + 1))) \ 4
        Next
    Next
    
    For i = 0 To CountStars - 1
        With DemoStars(i)
            x0 = CLng(.x + 0.5) * 4&
            y0 = CLng(.y + 0.5)
            If x0 < 0 Or y0 < 0 Or x0 >= DstW Or y0 >= DstH Then
            Else
                ArrDst(x0 + 3, y0) = 255
                ArrDst(x0 + 2, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 2, y0)), 255)
                ArrDst(x0 + 1, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 1, y0)), 255)
                ArrDst(x0 + 0, y0) = Max(CLng(.brightness) + CLng(ArrDst(x0 + 0, y0)), 255)
            End If
        End With
    Next
    
    For j = 0 To DstH - 1
        For i = 0 To DstW - 1 Step 4
            x0 = DemoLookup(i \ 4, j).Left
            y0 = (DemoLookup(i \ 4, j).Top + DemoPosition) Mod SrcH
            x0 = x0 * 4&
            If ArrSrc(x0 + 2, y0) + ArrSrc(x0 + 1, y0) + ArrSrc(x0 + 0, y0) > 0 Then
                ArrDst(i + 3, j) = (CLng(ArrSrc(x0 + 3, y0) * 2&) + CLng(ArrDst(i + 3, j))) \ 3
                ArrDst(i + 2, j) = (CLng(ArrSrc(x0 + 2, y0) * 2&) + CLng(ArrDst(i + 2, j))) \ 3
                ArrDst(i + 1, j) = (CLng(ArrSrc(x0 + 1, y0) * 2&) + CLng(ArrDst(i + 1, j))) \ 3
                ArrDst(i + 0, j) = (CLng(ArrSrc(x0 + 0, y0) * 2&) + CLng(ArrDst(i + 0, j))) \ 3
            End If
        Next
    Next
    
    If DemoPosition < 65 Then
        k = CLng(Sqr(256& * DemoPosition * 4&))
        'ODS vbCrLf & DemoPosition & " => " & k & vbCrLf
        For j = 0 To DstH - 1
            For i = 0 To DstW - 1 Step 4
                ArrDst(i + 3, j) = (CLng(ArrDst(i + 3, j)) * k) \ 256&
                ArrDst(i + 2, j) = (CLng(ArrDst(i + 2, j)) * k) \ 256&
                ArrDst(i + 1, j) = (CLng(ArrDst(i + 1, j)) * k) \ 256&
                ArrDst(i + 0, j) = (CLng(ArrDst(i + 0, j)) * k) \ 256&
                'ArrDst(i + 0, j) = CLng(255) * (63 - DemoPosition) * 4 \ 252&
            Next
        Next
    End If
    If DemoPosition > SrcH - 65 Then
        k = ((SrcH - DemoPosition) * 4& * (SrcH - DemoPosition) * 4&) \ 256&
        For j = 0 To DstH - 1
            For i = 0 To DstW - 1 Step 4
                ArrDst(i + 3, j) = CLng(ArrDst(i + 3, j)) * k \ 256&
                ArrDst(i + 2, j) = CLng(ArrDst(i + 2, j)) * k \ 256&
                ArrDst(i + 1, j) = CLng(ArrDst(i + 1, j)) * k \ 256&
                ArrDst(i + 0, j) = CLng(ArrDst(i + 0, j)) * k \ 256&
                'ArrDst(i + 0, j) = CLng(255) * (SrcH - DemoPosition) * 4 \ 252&
            Next
        Next
    End If


SKIP:
    If Err Then
        ODS "Error : " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0

    DemoDibWork.FreeBits ArrDst
    DemoDibBackBuf.FreeBits ArrSrc
    DemoDibTempBuf.FreeBits ArrTmp

    BitBlt picDemo.hDC, 0, 0, DstW, DstH, DemoDibBackBuf.DeviceContext, 0, 0, vbSrcCopy
'    BitBlt picDemo.hDC, 0, 0, SrcW, SrcH, DemoDibWork.DeviceContext, 0, DemoPosition, vbSrcCopy


'picDemo.Cls
'picDemo.Print DemoPosition
    DemoPosition = DemoPosition + 1
    
    If DemoPosition > SrcH Then
        InitDemo
    End If
End Sub

Private Function Max(ByVal A As Long, ByVal B As Long) As Long
    If A > B Then Max = A Else Max = B
End Function
