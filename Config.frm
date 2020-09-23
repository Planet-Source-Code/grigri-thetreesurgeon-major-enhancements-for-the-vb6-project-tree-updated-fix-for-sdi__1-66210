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
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Interval        =   20
            Left            =   2400
            Top             =   240
         End
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "(Version x.x)"
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
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   2415
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
         Left            =   120
         TabIndex        =   19
         Tag             =   "mailto:grigri@shinyhappypixels.com"
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "Written by grigri, 2006"
         Height          =   375
         Left            =   120
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
            Name            =   "Comic Sans MS"
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
         Top             =   240
         Width           =   3855
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
   Begin VB.Frame fraPane 
      Caption         =   "General Settings"
      Height          =   3615
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox chkSetting 
         Caption         =   "Double Buffer"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   33
         Tag             =   "14"
         Top             =   1080
         Width           =   3495
      End
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
End
Attribute VB_Name = "SettingsDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private ts As TreeSurgeon

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Const IDC_HAND As Long = 32649&

Private hLinkCursor As Long
Private hOldCursor As Long

'Private OldSettings() As Boolean
Private TmpSettings() As String
Private Validated As Boolean

Private Demo As New MiniDemo

Public Sub EditSettings(Settings() As String, ts2 As TreeSurgeon)
    Set ts = ts2
    Dim i As Long
    
    Load Me
    
    ReDim TmpSettings(LBound(Settings) To UBound(Settings))
    For i = LBound(TmpSettings) To UBound(TmpSettings)
        TmpSettings(i) = Settings(i)
    Next
    
    UpdateControls
    
    i = GetSetting(App.Title, "SettingsDialog", "LastActivePane", 3)
    lstPanes.ListIndex = i
    
    Me.Show vbModal
    
    If Validated Then
        For i = LBound(TmpSettings) To UBound(TmpSettings)
            Settings(i) = TmpSettings(i)
        Next
    End If
    
    'Unload Me
    
    Set ts = Nothing
End Sub

Private Sub btnCancel_Click()
    Validated = False
    tmrDemo.Enabled = False
    Unload Me
    'Me.Visible = False
End Sub

Private Sub btnOK_Click()
    Validated = True
    tmrDemo.Enabled = False
    Unload Me
    'Me.Visible = False
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
        For i = opAntiAlias.LBound To opAntiAlias.UBound
            opAntiAlias(i).Enabled = False
        Next
        TmpSettings([AntiAlias Method]) = 0
    Else
        For i = opAntiAlias.LBound To opAntiAlias.UBound
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
        For i = opAntiAlias.LBound To opAntiAlias.UBound
            opAntiAlias(i).Value = False
            opAntiAlias(i).Enabled = False
        Next
    Else
        chkAntiAlias.Value = 1
        For i = opAntiAlias.LBound To opAntiAlias.UBound
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
    
    lblVersion.Caption = "(Version " & App.Major & "." & App.Minor & ")"
    
    Demo.Init picDemo.hWnd
    
    GetFonts
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrDemo.Enabled = False
    
    SaveSetting App.Title, "SettingsDialog", "LastActivePane", lstPanes.ListIndex
End Sub

Private Sub lblEmail_Click()
    SetCursor hLinkCursor
    ShellExecute hWnd, "open", lblEmail.Tag, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetCursor hLinkCursor
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
    For i = opAntiAlias.LBound To opAntiAlias.UBound
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

Private Sub tmrDemo_Timer()
    Demo.Step
End Sub
