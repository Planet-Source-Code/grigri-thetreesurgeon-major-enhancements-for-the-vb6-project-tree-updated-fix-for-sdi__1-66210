VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeSurgeon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const DEBUG_SORTING = False

Public Enum SettingsEntries
    [Unsaved In Italics] = 0
    [Startup In Bold] = 1
    [ShellIcon For RelDocs] = 2
    [ShellIcon For ResFiles] = 3
    [FileNames] = 4
    [FileNames Relative Paths] = 5
    [FileNames in Grey] = 6
    [ShellMenu] = 7
    [TreeLines] = 8
    [Form Icon if Present] = 9
    [AntiAlias Method] = 10
    [Tree Font] = 11
    [All ShellIcons] = 12
    [PerComponentIcons] = 13
    [DoubleBuffer] = 14
    [SETTINGS_LAST]
End Enum

Private Const TREE_TIMER As Long = &H12345678
Private Const TREE_TIMER2 As Long = &H12345679

Public VBInstance As VBE
Private hWndProject As Long, hWndTree As Long
Private iTimerID As Long
Private iTimerID2 As Long

Public IconCache As New CachedImageList

Private bNeedSortTree As Boolean, NoRefreshCount As Long

Private ProjSubClass As SubClassData
Private TreeSubClass As SubClassData
Private CompareThunk As ThunkData

Private WithEvents PrjWatcher As VBProjectsEvents
Attribute PrjWatcher.VB_VarHelpID = -1
Private WithEvents CmpWatcher As VBComponentsEvents
Attribute CmpWatcher.VB_VarHelpID = -1
Private WithEvents FileWatcher As FileControlEvents
Attribute FileWatcher.VB_VarHelpID = -1

Private WithEvents Toolbar As MyToolbar
Attribute Toolbar.VB_VarHelpID = -1

Private Settings(0 To SETTINGS_LAST - 1) As String

' -- Custom Draw Variables
' ** Update 08/08/06 : changed these from being Static inside ProjectWndProc to being Private module-level
Private hFontDefault As Long
Private hFontBold As Long
Private hFontItalic As Long
Private hFontBoldItalic As Long
Private hPenLines As Long
Private hImlTree As Long
Private hDrawDC As Long
Private BmpBackBuffer As Bitmap


' Basic Setup
Public Sub Init(vbInst As VBE)
    ' Store the VB instance
    Set VBInstance = vbInst
    
    ' Load previous settings
    LoadSettings
    
    ' Setup thunks
    InitThunk CompareThunk, ObjPtr(Me), AddressOf RedirectCompareProc
    
    ' Get the project window handle and its treeview
    
    ' ** Update 08/08/2006 : Fixed for the SDI environment
    'hWndProject = FindWindowEx(VBInstance.MainWindow.hWnd, 0, "PROJECT", vbNullString)
    hWndProject = FindThreadWindow("PROJECT", vbNullString)
    
    hWndTree = FindWindowEx(hWndProject, 0, "SysTreeView32", vbNullString)
    
    ' ** Update 08/08/2006 : Paranoid check
    If hWndProject = 0 Or hWndTree = 0 Then
        MsgBox "Could not find Project Treeview; The Tree Surgeon cannot function", vbCritical Or vbOKOnly
        Exit Sub
    End If
    
    ' Subclass 'em both
    SubClass ProjSubClass, hWndProject, ObjPtr(Me), AddressOf Redirects.RedirectProjectWndProc
    SubClass TreeSubClass, hWndTree, ObjPtr(Me), AddressOf Redirects.RedirectTreeWndProc
    
    ' Sink standard events
    Set PrjWatcher = VBInstance.Events.VBProjectsEvents
    Set CmpWatcher = VBInstance.Events.VBComponentsEvents(Nothing)
    
    ' Set timers
    iTimerID = SetTimer(hWndTree, TREE_TIMER, 500, 0)       ' Refresh tree twice a second
    iTimerID2 = SetTimer(hWndTree, TREE_TIMER2, 60000, 0)   ' Rebuild icon cache every minute
    
    ' Create the toolbar form
    Set Toolbar = New MyToolbar
    Load Toolbar
    Toolbar.Visible = True
    SetParent Toolbar.hWnd, hWndProject
    PositionToolbar
    
    ' Refresh
    SortProjectTree Nothing
    RefreshTree True
End Sub

' Clearup
Private Sub Class_Terminate()
    ' Save settings
    SaveSettings
    
    ' Remove our refresh timer
    KillTimer hWndTree, iTimerID
    ' ** UPDATE 08/08/2006 : Must release BOTH timers!
    KillTimer hWndTree, iTimerID2
    
    ' Remove subclass
    UnSubClass ProjSubClass, hWndProject
    UnSubClass TreeSubClass, hWndTree
    
    ' Kill toolbar
    SetParent Toolbar.hWnd, 0&
    Unload Toolbar
    Set Toolbar = Nothing
    RestoreOldToolbarPosition
    
    ' Refresh
    RefreshTree True
End Sub

Friend Function ProjectWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_SYSCOLORCHANGE
        ODS "*** Tree : WM_SYSCOLORCHANGE\n"
    Case WM_NOTIFY
        'ODS "Project window, WM_NOTIFY" & vbCrLf
        
        Dim nm As NMHDR
        CopyMemory nm, ByVal lParam, LenB(nm)
        
        If nm.hwndFrom = hWndTree And nm.Code = NM_RCLICK Then
            If HandleTreeContextMenu Then Exit Function
        End If
        
        If nm.hwndFrom = hWndTree And nm.Code = NM_DBLCLK Then
            If HandleDoubleClick Then Exit Function
        End If
        
        If nm.hwndFrom = hWndTree And nm.Code = NM_CUSTOMDRAW Then
            ' Custom draw
            ' ** Update 08/08/06 : Removed Static declarations, now set at module-level
            
            Dim rcText As RECT, rcItem As RECT, ptTemp As POINTAPI, ptTemp2 As POINTAPI
            Dim rcUpdate As RECT
            Dim dx As Long, dy As Long
            Dim x As Long, y As Long
            Dim sNodeText As String, sFileText As String
            Dim hItem1 As Long, hItem2 As Long
            Dim hOldPen As Long, hOldBrush As Long
            Dim flags As Long
            Dim hIml As Long, iImageIndex As Long, iTempImageIndex As Long
            Dim hImlShell As Long
            Dim picIcon As StdPicture
            'Dim sfi As SHFILEINFO
            
            Dim obj As Object, prj As VBProject, cmp As VBComponent
            Dim tv As NMTVCUSTOMDRAW
            CopyMemory tv, ByVal lParam, LenB(tv)
            Select Case tv.dwDrawStage
            Case CDDS_PREPAINT
            
                
                
                ' Prepare backbuffer
                If Settings(DoubleBuffer) Then
                    'rcUpdate = tv.rc
                    GetClientRect hWndTree, rcUpdate
                    If rcUpdate.Left = rcUpdate.Right And rcUpdate.Top = rcUpdate.Bottom Then
                        ProjectWndProc = CDRF_DODEFAULT
                        Exit Function
                    End If
                    Set BmpBackBuffer = New Bitmap
                    With BmpBackBuffer
                        'ODS "CDDS_PREPAINT, rc = (%i,%i)-(%i,%i)\n", rcUpdate.Left, rcUpdate.Top, rcUpdate.Right, rcUpdate.Bottom
                        .Create rcUpdate.Right - rcUpdate.Left, rcUpdate.Bottom - rcUpdate.Top, tv.hDC
                        FillRect .handle, rcUpdate, GetSysColorBrush(COLOR_WINDOW)
                    End With
                    hDrawDC = BmpBackBuffer.handle
                Else
                    hDrawDC = tv.hDC
                End If
                
                ' Setup fonts depending on settings
                hFontDefault = SendMessage(hWndTree, WM_GETFONT, 0, ByVal 0&)
                Dim lf As LOGFONT
                GetObject hFontDefault, LenB(lf), lf
            
                
                If Len(Settings([Tree Font])) > 0 Then
                    Dim FontFace() As Byte
                    FontFace = StrConv(Settings([Tree Font]) & vbNullChar, vbFromUnicode)
                    CopyMemory lf.lfFaceName(1), FontFace(0), UBound(FontFace) + 1
                End If
                
                If Settings([AntiAlias Method]) <> 0 Then
                    lf.lfQuality = Settings([AntiAlias Method])
                End If
                
                lf.lfItalic = 0: lf.lfWeight = 0
                hFontDefault = CreateFontIndirect(lf)
                
                lf.lfItalic = 0: lf.lfWeight = IIf(Settings(SettingsEntries.[Startup In Bold]), 700, 0)
                hFontBold = CreateFontIndirect(lf)
                
                lf.lfItalic = IIf(Settings(SettingsEntries.[Unsaved In Italics]), 1, 0): lf.lfWeight = 0
                hFontItalic = CreateFontIndirect(lf)
                
                lf.lfItalic = IIf(Settings(SettingsEntries.[Unsaved In Italics]), 1, 0): lf.lfWeight = IIf(Settings(SettingsEntries.[Startup In Bold]), 700, 0)
                hFontBoldItalic = CreateFontIndirect(lf)
                
                ' Pen for drawing the lines
                Dim lb As LOGBRUSH
                lb.lbColor = RGB(128, 128, 128)
                lb.lbStyle = BS_SOLID
                hPenLines = ExtCreatePen(PS_ALTERNATE, 1, lb, 0, ByVal 0&)
                
                ' The tree imagelist
                hImlTree = SendMessage(hWndTree, TVM_GETIMAGELIST, TVSIL_NORMAL, ByVal 0&)
                
                ' We want per-item notifications, and postpaint for clearup
                ProjectWndProc = CDRF_NOTIFYITEMDRAW Or CDRF_NOTIFYPOSTPAINT
                Exit Function
            Case CDDS_POSTPAINT
                ' Clearup our GDI objects
                DeleteObject hFontDefault
                DeleteObject hFontBold
                DeleteObject hFontItalic
                DeleteObject hFontBoldItalic
                DeleteObject hPenLines
                
                ' Blit the backbuffer
                If Settings(DoubleBuffer) Then
                    'rcUpdate = tv.rc   ' Just the region to update
                    GetClientRect hWndTree, rcUpdate    ' Whole view (preparing for static background images)
                    BmpBackBuffer.Blit tv.hDC, rcUpdate.Left, rcUpdate.Top
                    ' And delete it
                    Set BmpBackBuffer = Nothing
                End If
                
            Case CDDS_ITEMPREPAINT
                '
                ' Default to tree icon list and standard tree icon
                hIml = hImlTree
                iImageIndex = GetNodeIcon(hWndTree, tv.dwItemSpec)
                
                Set obj = GetNodeObject(tv.dwItemSpec)
                '-------------------------------------- DUNNO NODE -----------------------
                If obj Is Nothing Then
                    ' Retrieve default text, default font
                    ' This should only occur for the Folder nodes, if they are active
                    sNodeText = GetNodeText(hWndTree, tv.dwItemSpec)
                    sFileText = vbNullString
                    SelectObject hDrawDC, hFontDefault
                Else
                    '------------------------------------- PROJECT NODE -----------------------
                    If TypeOf obj Is VBProject Then
                        ' It's a project!
                        Set prj = obj
                        Set cmp = Nothing
                        If prj Is prj.Collection.StartProject Then
                            If prj.IsDirty Then
                                SelectObject hDrawDC, hFontBoldItalic
                            Else
                                SelectObject hDrawDC, hFontBold
                            End If
                        Else
                            If prj.IsDirty Then
                                SelectObject hDrawDC, hFontItalic
                            Else
                                SelectObject hDrawDC, hFontDefault
                            End If
                        End If
                        sNodeText = prj.Name
                        If Settings(SettingsEntries.FileNames) Then
                            If prj.FileName = "" Then
                                sFileText = "Not Saved"
                            Else
                                sFileText = FileNameFromPath(prj.FileName)
                            End If
                        End If
                    End If
                    '----------------------------------- COMPONENT NODE -----------------------
                    If TypeOf obj Is VBComponent Then
                        Set cmp = obj
                        Set prj = cmp.Collection.Parent
                        
                        If cmp Is GetProjectStartupComponent(prj) Then
                            If cmp.IsDirty Then
                                SelectObject hDrawDC, hFontBoldItalic
                            Else
                                SelectObject hDrawDC, hFontBold
                            End If
                        Else
                            If cmp.IsDirty Then
                                SelectObject hDrawDC, hFontItalic
                            Else
                                SelectObject hDrawDC, hFontDefault
                            End If
                        End If
                        
                        If Settings(SettingsEntries.FileNames) Then
                            If cmp.FileNames(1) = vbNullString Then
                                sFileText = "Not Saved"
                            Else
                                If Settings(SettingsEntries.[FileNames Relative Paths]) And (prj.FileName <> vbNullString) Then
                                    sFileText = GetRelativePath(prj.FileName, cmp.FileNames(1))
                                Else
                                    sFileText = FileNameFromPath(cmp.FileNames(1))
                                End If
                            End If
                        End If
                        
                        Select Case cmp.Type
                        Case vbext_ct_RelatedDocument
                            sNodeText = FileNameFromPath(cmp.FileNames(1))
                            sFileText = ""
                            ' Get the real file icon if requested
                            If Settings(SettingsEntries.[ShellIcon For RelDocs]) Then
                                iTempImageIndex = IconCache.AddShellIcon(cmp.FileNames(1))
                                If iTempImageIndex > -1 Then
                                    hIml = IconCache.ImageListHandle
                                    iImageIndex = iTempImageIndex
                                End If
                            End If
                        Case vbext_ct_ResFile
                            sNodeText = FileNameFromPath(cmp.FileNames(1))
                            sFileText = ""
                            ' Get the real file icon if requested
                            If Settings(SettingsEntries.[ShellIcon For ResFiles]) Then
                                iTempImageIndex = IconCache.AddShellIcon(cmp.FileNames(1))
                                If iTempImageIndex > -1 Then
                                    hIml = IconCache.ImageListHandle
                                    iImageIndex = iTempImageIndex
                                End If
                            End If
                        Case Else
                            sNodeText = cmp.Name
                            ' Check for per-component icon
                            If Settings(PerComponentIcons) Then
                                Dim sIconFile As String
                                On Error Resume Next
                                sIconFile = GetCustomIcon(cmp)
                                If Len(sIconFile) > 0 Then
                                    iTempImageIndex = IconCache.AddIcon(sIconFile)
                                    If iTempImageIndex > -1 Then
                                        hIml = IconCache.ImageListHandle
                                        iImageIndex = iTempImageIndex
                                        GoTo GOT_COMPONENT_ICON
                                    End If
                                End If
                            End If
                            ' Check for shell icons
                            On Error GoTo SKIP_SHELL_ICONS
                            If CBool(Settings([All ShellIcons])) Then
                                iTempImageIndex = IconCache.AddShellIcon(cmp.FileNames(1))
                                If iTempImageIndex > -1 Then
                                    hIml = IconCache.ImageListHandle
                                    iImageIndex = iTempImageIndex
                                End If
                            End If
SKIP_SHELL_ICONS:
                            ' Custom form icon
                            ' Note that for some stupid reason, if the icon specified in the "Icon"
                            ' property contains several sizes, we can't specify which one we want.
                            If cmp.Type = vbext_ct_VBForm Then
                                If Settings([Form Icon if Present]) Then
                                    On Error Resume Next
                                    Set picIcon = cmp.Properties("Icon").Object
                                    If Err Then
                                        Set picIcon = Nothing
                                    Else
                                        If picIcon.Type <> vbPicTypeIcon Then
                                            ' Only allow icons
                                            Set picIcon = Nothing
                                        Else
                                            If picIcon.Width <> picIcon.Height Then
                                                ' Must have same width and height
                                                Set picIcon = Nothing
                                            Else
                                                If Abs(Toolbar.ScaleX(picIcon.Width, vbHimetric, vbPixels) - GetSystemMetrics(SM_CXSMICON)) > 1 Then
                                                    ' Must be correct size (looks funny otherwise)
                                                    Set picIcon = Nothing
                                                End If
                                            End If
                                        End If
                                    End If
                                    On Error GoTo 0
                                End If
                            End If
GOT_COMPONENT_ICON:
                        End Select
                    End If
                End If
                
                ' Retrieve placement rectangles and metrics
                rcItem.Left = tv.dwItemSpec
                SendMessage hWndTree, TVM_GETITEMRECT, 0, rcItem
                rcText.Left = tv.dwItemSpec
                SendMessage hWndTree, TVM_GETITEMRECT, 1, rcText
                
                dx = SendMessage(hWndTree, TVM_GETINDENT, 0, ByVal 0&)
                dy = SendMessage(hWndTree, TVM_GETITEMHEIGHT, 0, ByVal 0&)
                
                ' ------------ Lines
                If Settings(SettingsEntries.TreeLines) Then
                    hOldPen = SelectObject(hDrawDC, hPenLines)
                    hItem1 = tv.dwItemSpec
                    x = rcText.Left - dx - 1
                    y = rcText.Top
                    
                    ' Draw Horizontal
                    MoveToEx hDrawDC, x - dx + dx \ 2, y + dy \ 2, ByVal 0&
                    LineTo hDrawDC, x - 1, y + dy \ 2
                    ' Draw vertical
                    hItem2 = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem1)
                    If hItem2 <> 0 Then
                        MoveToEx hDrawDC, x - dx + dx \ 2, y, ByVal 0&
                        LineTo hDrawDC, x - dx + dx \ 2, y + dy \ 2
                    End If
    
                    
                    While hItem1 <> 0
                        
                        'If hItem2 <> 0 Then
                            x = x - dx
                            If SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem1) <> 0 Then
                                MoveToEx hDrawDC, x + dx \ 2, y, ByVal 0&
                                LineTo hDrawDC, x + dx \ 2, y + dy
                            End If
                        'End If
                        hItem2 = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem1)
                        hItem1 = hItem2
                    Wend
                    Call SelectObject(hDrawDC, hOldPen)
                End If
                
                ' ------------ Plus / Minus
                If SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal tv.dwItemSpec) <> 0 Then
                    ' Position of column before icon
                    x = rcText.Left - dx * 2 - 1
                    y = rcText.Top
                    ' Centre of cell
                    x = x + dx \ 2
                    y = y + dy \ 2
                    
                    ' Draw box
                    Rectangle hDrawDC, x - 4, y - 4, x + 5, y + 5
                    ' Horizontal line (Minus)
                    MoveToEx hDrawDC, x - 2, y, ByVal 0&
                    LineTo hDrawDC, x + 3, y
                    If (GetNodeState(hWndTree, tv.dwItemSpec) And TVIS_EXPANDED) = 0 Then
                        ' Vertical line for plus
                        MoveToEx hDrawDC, x, y - 2, ByVal 0&
                        LineTo hDrawDC, x, y + 3
                    End If
                End If
                
                ' Icon
                x = rcText.Left - dx
                y = rcText.Top
                
                ImageList_GetIconSize hIml, dx, dy
                hOldBrush = SelectObject(hDrawDC, GetSysColorBrush(COLOR_WINDOW))
                hOldPen = SelectObject(hDrawDC, GetStockObject(NULL_PEN))
                Rectangle hDrawDC, x, y, x + dx, y + dy
                SelectObject hDrawDC, hOldBrush
                SelectObject hDrawDC, hOldPen
                
                
                If picIcon Is Nothing Then
                    flags = ILD_TRANSPARENT
                    'If CBool(tv.uItemState And CDIS_SELECTED) Then
                    '    flags = flags Or ILD_BLEND50
                    'End If
                    ImageList_Draw hIml, iImageIndex, hDrawDC, x, y, flags
                Else
                    DrawIconEx hDrawDC, x, y, picIcon.handle, 0, 0, 0, 0, DI_NORMAL
                End If

                
                ' Text (finally!)
                'GetTextExtentPoint32 hDrawDC, sNodeText, Len(sNodeText), ptTemp
                'MeasureTextExact hDrawDC, sNodeText, ptTemp
                DrawText hDrawDC, sNodeText, Len(sNodeText), rcText, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                ptTemp.x = rcText.Right - rcText.Left
                
                If Settings(FileNames) And (Len(sFileText) > 0) Then
                    ' Prefix space
                    sFileText = " (" & sFileText & ")"
                    ' Calc width
                    'GetTextExtentPoint32 hDrawDC, sFileText, Len(sFileText), ptTemp2
                    'MeasureTextExact hDrawDC, sFileText, ptTemp2
                    DrawText hDrawDC, sFileText, Len(sFileText), rcText, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                    ptTemp2.x = rcText.Right - rcText.Left
                Else
                    ptTemp2.x = 0
                End If
               
                If CBool(tv.uItemState And CDIS_SELECTED) Then
                    SetTextColor hDrawDC, GetSysColor(COLOR_HIGHLIGHTTEXT)
                    SetBkColor hDrawDC, GetSysColor(COLOR_HIGHLIGHT)
                    hOldBrush = SelectObject(hDrawDC, GetSysColorBrush(COLOR_HIGHLIGHT))
                Else
                    SetTextColor hDrawDC, GetSysColor(COLOR_WINDOWTEXT)
                    SetBkColor hDrawDC, GetSysColor(COLOR_WINDOW)
                    hOldBrush = SelectObject(hDrawDC, GetSysColorBrush(COLOR_WINDOW))
                End If
                SetBkMode hDrawDC, TRANSPARENT
                hOldPen = SelectObject(hDrawDC, GetStockObject(NULL_PEN))
                
                rcText.Top = rcItem.Top
                rcText.Bottom = rcItem.Bottom
                rcText.Right = rcText.Left + ptTemp.x + ptTemp2.x + 4
                
                Rectangle hDrawDC, rcText.Left, rcText.Top, rcText.Right, rcText.Bottom
                rcText.Left = rcText.Left + 2
                DrawText hDrawDC, sNodeText, Len(sNodeText), rcText, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                
                If Settings(FileNames) And (Len(sFileText) > 0) Then
                    If CBool(tv.uItemState And CDIS_SELECTED) = False Then
                        If Settings([FileNames in Grey]) Then
                            SetTextColor hDrawDC, RGB(96, 96, 96)
                        End If
                    End If
                    rcText.Left = rcText.Left + ptTemp.x
                    DrawText hDrawDC, sFileText, Len(sFileText), rcText, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
                End If
                
                SelectObject hDrawDC, hOldBrush
                SelectObject hDrawDC, hOldPen
                
'                tv.rc.Left = rcItem.Left
'                tv.rc.Top = rcItem.Top
'                tv.rc.Right = rcText.Right
'                tv.rc.Bottom = rcItem.Bottom
'
'                CopyMemory ByVal lParam, tv, LenB(tv)
                
                ProjectWndProc = CDRF_SKIPDEFAULT
                Exit Function
            End Select
        End If
    Case WM_SIZE
        ProjectWndProc = CallWindowProc(ProjSubClass.pfnWndProcNext, hWnd, uMsg, wParam, lParam)
        PositionToolbar
        Exit Function
    End Select
    ProjectWndProc = CallWindowProc(ProjSubClass.pfnWndProcNext, hWnd, uMsg, wParam, lParam)
End Function

Friend Function TreeWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hItem As Long
    Select Case uMsg
    Case WM_TIMER
        If wParam = TREE_TIMER Then
            '
            If NoRefreshCount = 0 Then
                If bNeedSortTree Then
                    SortProjectTree Nothing
                    bNeedSortTree = False
                Else
                    RefreshTree
                End If
            End If
            '
            TreeWndProc = 1
            Exit Function
        ElseIf wParam = TREE_TIMER2 Then
            If Not IconCache Is Nothing Then
                IconCache.Clear
            End If
            RefreshTree
        End If
    Case WM_ERASEBKGND
        If Settings(DoubleBuffer) Then
            TreeWndProc = 1
            Exit Function
        End If
'    Case WM_HSCROLL, WM_VSCROLL
'        LockWindowUpdate hwnd
'        TreeWndProc = CallWindowProc(TreeSubClass.pfnWndProcNext, hwnd, uMsg, wParam, lParam)
'        LockWindowUpdate 0
'        RefreshTree
'        Exit Function
    Case TVM_INSERTITEMA
        TreeWndProc = CallWindowProc(TreeSubClass.pfnWndProcNext, hWnd, uMsg, wParam, lParam)
        bNeedSortTree = True
        Exit Function
    Case TVM_SORTCHILDREN
        ' Transform this into a TVM_SORTCHILDRENCB message with a pointer
        ' to our callback function (thunked)
        
        If lParam = 0 Then GoTo Default_Processing
        
        NoRefreshCount = NoRefreshCount + 1
            
        Dim obj As Object, prj As VBProject, cmp As VBComponent
        
        'MsgBox "Getting Node Object for item '" & GetNodeText(hWnd, lParam) & "'..."
        'DebugBreak
        
        Set obj = GetNodeObject(lParam)
        'MsgBox "It's a " & TypeName(obj)
        
        If obj Is Nothing Then NoRefreshCount = NoRefreshCount - 1: GoTo Default_Processing
        
        If TypeOf obj Is VBProject Then
            
            'MsgBox "Project.."
            Set prj = obj
            ' Build a lookup mapping item lParams into components
            ' (the callback function receives the lParams of the items)
            Dim colItems As New Collection
            Dim sOrderString As String

            'MsgBox "Getting order string..."
            sOrderString = GetProjectOrderString(prj)

            'MsgBox "Getting items..."
            hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_CHILD, ByVal lParam)
            If hItem = 0 Then NoRefreshCount = NoRefreshCount - 1: GoTo Default_Processing
            While hItem <> 0
                Set obj = GetNodeObject(hItem)
                If Not (obj Is Nothing) Then
                    If TypeOf obj Is VBComponent Then
                        Set cmp = obj
                        colItems.Add cmp, "lp" & GetNodeParam(hWndTree, hItem)
                    End If
                End If
                hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
            Wend
            
            'MsgBox "Testing if less than 2..."
            If colItems.Count < 2 Then NoRefreshCount = NoRefreshCount - 1: GoTo Default_Processing
            'MsgBox "2 or more"

            ' Rebuild the project order string inside the callback function (sentinel)
            CompareTreeItems lParam, 0, Nothing
            
            ' Build the sort structure
            Dim sort As TVSORTCB
            With sort
                .lParam = VarPtr(colItems)
                .lpFnCompare = CompareThunk.pFn
                .hParentItem = lParam
            End With
            ' Send the actual sorting message
            TreeWndProc = SendMessage(hWnd, TVM_SORTCHILDRENCB, 0, sort)
            ' And exit
            NoRefreshCount = NoRefreshCount - 1
            Exit Function
        End If
        NoRefreshCount = NoRefreshCount - 1
    Case WM_SYSCOLORCHANGE
        ODS "*** Tree : WM_SYSCOLORCHANGE\n"
    End Select
    
Default_Processing:
    TreeWndProc = CallWindowProc(TreeSubClass.pfnWndProcNext, hWnd, uMsg, wParam, lParam)
End Function

Private Function HandleDoubleClick() As Boolean
    Dim hit As TVHITTESTINFO
    ' Retrieve cursor position
    GetCursorPos hit.pt
    ' Offset for this window
    ScreenToClient hWndTree, hit.pt
    ' Where does it lie?
    SendMessage hWndTree, TVM_HITTEST, 0, hit
    
    If hit.flags And TVHT_ONITEM Then
        Dim cmp As VBComponent
        On Error GoTo SKIP_COMPONENT
        Set cmp = GetNodeObject(hit.hItem)
        
        If cmp.Type = vbext_ct_RelatedDocument Then
'            If LCase$(Right$(cmp.FileNames(1), 4)) = ".txt" Then
'                MsgBox "Trapped execution of text document"
'                HandleDoubleClick = True
'                Exit Function
'            End If
        End If
        
SKIP_COMPONENT:
        On Error GoTo 0
    End If

End Function

Private Function HandleTreeContextMenu() As Boolean
    ' If we don't want the context menu then return straight away
    If Settings(SettingsEntries.ShellMenu) = False Then Exit Function


    Dim hit As TVHITTESTINFO
    ' Retrieve cursor position
    GetCursorPos hit.pt
    ' Offset for this window
    ScreenToClient hWndTree, hit.pt
    ' Where does it lie?
    SendMessage hWndTree, TVM_HITTEST, 0, hit
    'ODS "> Hit Flags : %s\n" , Hex$(hit.flags)
    
    Dim hPrevItem As Long
    If hit.flags And TVHT_ONITEMICON Then
        'ODS "> On Icon\n"
        ' Right-click on the icon!
        Dim obj As Object, prj As VBProject, cmp As VBComponent
        Dim Folder As String, File As String, ShowMenu As Boolean
        
        ShowMenu = False
        Set obj = GetNodeObject(hit.hItem)
        If TypeOf obj Is VBComponent Then
            Set cmp = obj
            Set prj = cmp.Collection.Parent
            If Len(cmp.FileNames(1)) > 0 Then
                Folder = FolderFromPath(cmp.FileNames(1))
                File = FileNameFromPath(cmp.FileNames(1))
                ShowMenu = True
            End If
        ElseIf TypeOf obj Is VBProject Then
            Set prj = obj
            If Len(prj.FileName) > 0 Then
                Folder = FolderFromPath(prj.FileName)
                File = FileNameFromPath(prj.FileName)
                ShowMenu = True
            End If
        End If
        If ShowMenu Then
            hPrevItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
            ' Hilite the item
            SendMessage hWndTree, TVM_SELECTITEM, TVGN_CARET, ByVal hit.hItem
            ' Display a menu
            ShowShellMenu hWndTree, Folder, File
            ' Unhilite the item
            SendMessage hWndTree, TVM_SELECTITEM, TVGN_CARET, ByVal hPrevItem
            RefreshTree
            ' Return True to say we've processed the message
            HandleTreeContextMenu = True
            Exit Function
        End If
    End If
End Function

Private Sub CmpWatcher_ItemAdded(ByVal cmp As VBIDE.VBComponent)
    Dim prj As VBProject
    Set prj = cmp.Collection.Parent
'    ' Sort the tree first
'    SortProjectTree prj
    ' Now rebuild the sort order with the new component
    
    'SetProjectOrderString prj, BuildDefaultProjectOrderString(prj)
End Sub

Private Sub CmpWatcher_ItemRemoved(ByVal cmp As VBIDE.VBComponent)
    ' If the component is in the order string then
    ' we need to remove it
    Dim sOrderString As String, iPos As Long
    Dim sEntry As String
    
    ' Get the project orderstring
    sOrderString = GetProjectOrderString(cmp.Collection.Parent)
    If LenB(sOrderString) = 0 Then Exit Sub
    
    ' Get the text for the entry
    sEntry = GetComponentOrderEntry(cmp)
    
    ' See if it's in the orderstring
    iPos = InStr(sOrderString, ":" & sEntry & ":")
    If iPos = 0 Then Exit Sub
    
    ' Remove it
    sOrderString = Left$(sOrderString, iPos - 1) & Mid$(sOrderString, iPos + Len(sEntry) + 1)
    SetProjectOrderString cmp.Collection.Parent, sOrderString
End Sub

Private Sub CmpWatcher_ItemRenamed(ByVal cmp As VBIDE.VBComponent, ByVal OldName As String)
    ' If the component is in the order string then
    ' we need to replace it
    Debug.Assert cmp.Type <> vbext_ct_RelatedDocument And cmp.Type <> vbext_ct_ResFile
    Dim sOrderString As String, iPos As Long
    sOrderString = GetProjectOrderString(cmp.Collection.Parent)
    iPos = InStr(sOrderString, ":" & OldName & ":")
    If iPos > 0 Then
        sOrderString = Left$(sOrderString, iPos) & cmp.Name & Mid$(sOrderString, iPos + Len(OldName) + 1)
        SetProjectOrderString cmp.Collection.Parent, sOrderString
    End If
End Sub

Private Sub CmpWatcher_ItemSelected(ByVal cmp As VBIDE.VBComponent)
    If cmp Is Nothing Then
        ' Disable moving
        Toolbar.SetStatus False
    Else
        ' Enable moving
        Toolbar.SetStatus True
    End If
End Sub

Public Sub DoEditSettings()
    Dim dlg As New SettingsDlg
    dlg.EditSettings Settings, Me
    
    RefreshTree True
End Sub

Private Sub Toolbar_ItemClicked(ByVal Index As ToolbarItems)
    Dim prj As VBProject, cmp As VBComponent
    Set cmp = VBInstance.SelectedVBComponent
    If cmp Is Nothing Then Exit Sub
    
    Set prj = cmp.Collection.Parent
    
    Dim sOrderString As String
    Dim arrOrder() As String
    Dim sEntry As String
    Dim i As Long
    Dim First As Long, Last As Long
    
    sOrderString = GetProjectOrderString(prj)
    
    sEntry = GetComponentOrderEntry(cmp)
    
    If (LenB(sOrderString) = 0) Or (InStr(sOrderString, ":" & sEntry & ":") = 0) Then
        'MsgBox "Item '" & sEntry & "' not in string '" & sOrderString & "' - Rebulding"
        sOrderString = BuildDefaultProjectOrderString(prj)
        'MsgBox "New String : " & sOrderString
    End If
    
    arrOrder = Split(sOrderString, ":")
    
    First = LBound(arrOrder) + 1
    Last = UBound(arrOrder) - 1
    If Index = [Move Up] Then
        First = First + 1 ' We cannot move it up if it's the first one
    Else
        Last = Last - 1   ' We cannot move it down if it's the last one
    End If
    
    For i = First To Last
        If arrOrder(i) = sEntry Then
            ' This is it!
            If Index = [Move Down] Then
                arrOrder(i) = arrOrder(i + 1)
                arrOrder(i + 1) = sEntry
            Else
                arrOrder(i) = arrOrder(i - 1)
                arrOrder(i - 1) = sEntry
            End If
            
            Exit For
        End If
    Next
    
    ' Rebuild the modified order string
    
    #If DEBUG_SORTING Then
    ODS "Old OrderString = %s\n", sOrderString
    #End If
    
    sOrderString = Join(arrOrder, ":")
    
    #If DEBUG_SORTING Then
    ODS "New OrderString = %s\n", sOrderString
    #End If
    
    SetProjectOrderString prj, sOrderString
    
    ' And sort the tree
    bNeedSortTree = True
    SortProjectTree prj
    RefreshTree
    
    Exit Sub
End Sub

' Reposition the toolbar properly
Private Function PositionToolbar()
    Dim rc As RECT, hWndTemp As Long
    Dim rc2 As RECT
    GetClientRect hWndProject, rc
    hWndTemp = FindWindowEx(hWndProject, 0, "MsoCommandBarDock", "MsoDockTop")
    If hWndTemp <> 0 Then
        GetClientRect hWndTemp, rc2
        SetWindowPos hWndTemp, 0, 0, 0, rc.Right - Toolbar.Width \ Screen.TwipsPerPixelX, rc2.Bottom, SWP_NOOWNERZORDER
        Toolbar.Move rc.Right * Screen.TwipsPerPixelX - Toolbar.Width, rc.Top * Screen.TwipsPerPixelX, Toolbar.Width, rc2.Bottom * Screen.TwipsPerPixelY
    Else
        ' Should never happen
        Toolbar.Move rc.Right * Screen.TwipsPerPixelX - Toolbar.Width, 0
    End If
End Function

' Restore the old toolbar to its proper position
Private Function RestoreOldToolbarPosition()
    Dim rc As RECT, rc2 As RECT, hWndTemp As Long
    hWndTemp = FindWindowEx(hWndProject, 0, "MsoCommandBarDock", "MsoDockTop")
    GetClientRect hWndProject, rc
    If hWndTemp <> 0 Then
        GetClientRect hWndTemp, rc2
        SetWindowPos hWndTemp, 0, 0, 0, rc.Right, rc2.Bottom, SWP_NOOWNERZORDER
    End If
End Function

' For a node in the project tree, retrieve the associated VBProject [Recursive]
Private Function GetNodeProject(ByVal hItem As Long) As VBProject
    If hItem = 0 Then Exit Function
    On Error GoTo HANDLE_ANY_ERROR
    
    Dim Icon As Long
    Dim txt As String
    Dim txt_name As String
    Dim i As Long
    
    Icon = GetNodeIcon(hWndTree, hItem)
    Select Case Icon
    Case 4, 19, 20, 21 ' Project
        txt = GetNodeText(hWndTree, hItem)
        i = InStr(txt, " (")
        If i > 0 Then
            txt_name = Left$(txt, i - 1)
        Else
            txt_name = txt
        End If
        Set GetNodeProject = VBInstance.VBProjects(txt_name)
    Case Else
        Set GetNodeProject = GetNodeProject(SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem))
    End Select
    
    Exit Function
HANDLE_ANY_ERROR:
    Set GetNodeProject = Nothing
    Exit Function
End Function

' For a node in the project tree, retrieve the associated object - VBProject or VBComponent, or Nothing
Private Function GetNodeObject(ByVal hItem As Long) As Object
    Dim Icon As Long
    Dim prj As VBProject
    Dim cmp As VBComponent
    Dim txt As String
    Dim txt_name As String
    Dim txt_file As String
    Dim i As Long
    
    txt = GetNodeText(hWndTree, hItem)
    i = InStr(txt, " (")
    If i > 0 Then
        txt_name = Left$(txt, i - 1)
        txt_file = Mid$(txt, i + 2, Len(txt) - i - 2)
    Else
        txt_name = txt
        txt_file = ""
    End If
    Icon = GetNodeIcon(hWndTree, hItem)
    
    On Error Resume Next
        
    Select Case Icon
    Case 12 ' Related Doc
        Set prj = Nothing
        Set cmp = Nothing
        Set prj = GetNodeProject(hItem)
        If prj Is Nothing Then Exit Function
        For Each cmp In prj.VBComponents
            If cmp.Type = vbext_ct_RelatedDocument Then
                ' Could potentially be a problem here if there are 2 components with the same
                ' file name but different paths (eg "C:\SomeDir\ReadMe.txt" and "C:\SomeOtherDir\ReadMe.txt"
                ' Not sure how to handle that yet...
                If FileNameFromPath(cmp.FileNames(1)) = txt_file Then
                    Set GetNodeObject = cmp
                    Exit Function
                End If
            End If
        Next
    Case 10 ' Resource
        Set prj = Nothing
        Set cmp = Nothing
        Set prj = GetNodeProject(hItem)
        If prj Is Nothing Then Exit Function
        For Each cmp In prj.VBComponents
            If cmp.Type = vbext_ct_ResFile Then
                ' Could potentially be a problem here if there are 2 components with the same
                ' file name but different paths (eg "C:\SomeDir\ReadMe.txt" and "C:\SomeOtherDir\ReadMe.txt"
                ' Not sure how to handle that yet...
                If FileNameFromPath(cmp.FileNames(1)) = txt_file Then
                    Set GetNodeObject = cmp
                    Exit Function
                End If
            End If
        Next
    Case 16, 17 ' Folder
        Set GetNodeObject = Nothing
    Case 4, 19, 20, 21 ' Project
        Set GetNodeObject = GetNodeProject(hItem)
    Case Else ' VBComponent
        Set prj = Nothing
        Set cmp = Nothing
        Set prj = GetNodeProject(hItem)
        If prj Is Nothing Then Exit Function
        Set cmp = prj.VBComponents(txt_name)
        If cmp Is Nothing Then Exit Function
        Set GetNodeObject = cmp
    End Select
End Function

' Retrieve the startup component for a project - either the startup form, the module containing Sub Main or Nothing
Private Function GetProjectStartupComponent(prj As VBProject) As VBComponent
    On Error Resume Next
    
    ' Form ?
    Set GetProjectStartupComponent = prj.VBComponents.StartUpObject
    If Err.Number = 0 Then Exit Function
    
    ' None?
    If CLng(prj.VBComponents.StartUpObject) = 1 Then Exit Function
    
    ' Must be "Sub Main"
    Dim cmp As VBComponent, cm As CodeModule, mem As Member
    For Each cmp In prj.VBComponents
        If cmp.Type = vbext_ct_StdModule Then
            Set cm = cmp.CodeModule
            Set mem = Nothing
            Set mem = cm.Members("Main")
            If Not (mem Is Nothing) Then
                If mem.Type = vbext_mt_Method Then
                    If mem.Scope = vbext_Public Then
                        Set GetProjectStartupComponent = cmp
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

' Retrieve the hItem for a project
Private Function FindProjectNode(ByVal prj As VBProject) As Long
    Dim hItem As Long
    
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
    While hItem <> 0
        '
        If GetNodeProject(hItem) Is prj Then
            FindProjectNode = hItem
            Exit Function
        End If
        '
        hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Wend
End Function

' Retrieve the hItem for a component [can be recursive]
Private Function FindComponentNode(ByVal cmp As VBComponent, Optional ByVal hItemRoot As Long = 0) As Long
    If hItemRoot = 0 Then
        hItemRoot = FindProjectNode(cmp.Collection.Parent)
        If hItemRoot = 0 Then Exit Function
    End If
    
    Dim hItem As Long
    Dim obj As Object
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItemRoot)
    While hItem <> 0
        '
        Set obj = GetNodeObject(hItem)
        If obj Is cmp Then
            ' Found it!
            FindComponentNode = hItem
            Exit Function
        ElseIf obj Is Nothing Then
            ' Probably a folder item (I hate that mode)
            ' Recurse on its children
            FindComponentNode = FindComponentNode(cmp, hItem)
            ' Exit if we've found it
            If FindComponentNode <> 0 Then Exit Function
        End If
        '
        hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Wend
End Function

Private Sub RefreshTree(Optional ByVal Clear As Boolean = False)
    'RedrawWindow hWndTree, ByVal 0&, 0, IIf(Clear, 5, 1)
    InvalidateRect hWndTree, ByVal 0&, IIf(Clear, 1, 0)
    UpdateWindow hWndTree
    
    'MyToolbar.ZOrder
End Sub

Friend Function CompareTreeItems(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal colItems As Collection) As Long
    Static OrderString As String

    ' Sentinel to reset orderstring
    If colItems Is Nothing Then
        OrderString = GetProjectOrderString(GetNodeProject(lParam1))
        Exit Function
    End If

    On Error Resume Next
    Dim cmp1 As VBComponent, cmp2 As VBComponent
    Set cmp1 = colItems("lp" & lParam1)
    Set cmp2 = colItems("lp" & lParam2)
    If Err Then Exit Function
    Err.Clear

    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(OrderString, ":" & GetComponentOrderEntry(cmp1) & ":")
    pos2 = InStr(OrderString, ":" & GetComponentOrderEntry(cmp2) & ":")

    If (pos1 <> 0) And (pos2 <> 0) Then
        ' Both items are listed, so the difference of the positions defines the order
        CompareTreeItems = pos1 - pos2
        Exit Function
    End If

    ' One or both of the items are not in the list, so we'll compare by string
    CompareTreeItems = StrComp(GetComponentOrderEntry(cmp1), GetComponentOrderEntry(cmp2), vbTextCompare)
End Function

Private Function GetProjectOrderString(ByVal prj As VBProject) As String
    On Error Resume Next
    GetProjectOrderString = prj.ReadProperty("TheTreeSurgeon", "ComponentOrder")
    
    ' Check for items that aren't in the list...
    Dim cmp As VBComponent, sEntry As String
    For Each cmp In prj.VBComponents
        sEntry = GetComponentOrderEntry(cmp)
        If InStr(GetProjectOrderString, ":" & sEntry & ":") = 0 Then
            GetProjectOrderString = GetProjectOrderString & (sEntry & ":")
        End If
    Next
End Function

Private Sub SetProjectOrderString(ByVal prj As VBProject, ByVal order As String)
    prj.WriteProperty "TheTreeSurgeon", "ComponentOrder", order
End Sub

Private Sub SortProjectTree(ByVal prj As VBProject)
    #If DEBUG_SORTING Then
    ODS "-> SortProjectTree\n"
    #End If
    
    Dim hItemRoot As Long, hItem As Long
    
    If prj Is Nothing Then
        hItemRoot = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
    Else
        hItemRoot = FindProjectNode(prj)
    End If
    
    SortTreeItems hWndTree, hItemRoot
    
    #If DEBUG_SORTING Then
    ODS "<- SortProjectTree\n"
    #End If
End Sub

Private Sub SortTreeItems(ByVal hWnd As Long, ByVal hItem As Long)
    #If False And DEBUG_SORTING Then
    Static Level As Long
    Level = Level + 1
    ODS Space$(Level) & "-> SortTreeItems\n"
    #End If
    
    ' Sort this item with no recursion
    SendMessage hWnd, TVM_SORTCHILDREN, 0, ByVal hItem
    ' Do recursion manually
    hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItem)
    While hItem <> 0
        SortTreeItems hWnd, hItem
        hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Wend
    
    #If False And DEBUG_SORTING Then
    ODS Space$(Level) & "<- SortTreeItems\n"
    Level = Level - 1
    #End If
End Sub

Private Function BuildDefaultProjectOrderString(ByVal prj As VBProject) As String
    'ODS "** BuildDefaultProjectOrderString **\n"
    Dim hItem As Long
    
    hItem = FindProjectNode(prj)
    If hItem = 0 Then Exit Function
    
    BuildDefaultProjectOrderString = ""
    AddComponentsToString hItem, BuildDefaultProjectOrderString
    
    BuildDefaultProjectOrderString = BuildDefaultProjectOrderString & ":"
End Function

Private Sub AddComponentsToString(ByVal hItemParent As Long, ByRef Entries As String)
    Dim hItem As Long
    Dim obj As Object, cmp As VBComponent
    'ODS "Adding components of item '" & GetNodeText(hWndTree, hItemParent) & "'\n"
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItemParent)
    While hItem <> 0
        '
        'ODS "Item '" & GetNodeText(hWndTree, hItem) & "' is a "
        Set obj = GetNodeObject(hItem)
        If obj Is Nothing Then
            'ODS "Folder?\n"
            ' Probably a folder node - recurse
            AddComponentsToString hItem, Entries
        ElseIf TypeOf obj Is VBComponent Then
            Set cmp = obj
            'ODS "Component! => " & GetComponentOrderEntry(cmp) & vbcrlf
            Entries = Entries & ":" & GetComponentOrderEntry(cmp)
        Else
            'ODS "Unknown\n"
        End If
        '
        hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    Wend
End Sub

Private Function GetComponentOrderEntry(ByVal cmp As VBComponent) As String
    If (cmp.Type = vbext_ct_ResFile) Or (cmp.Type = vbext_ct_RelatedDocument) Then
        GetComponentOrderEntry = FileNameFromPath(cmp.FileNames(1))
    Else
        GetComponentOrderEntry = cmp.Name
    End If
End Function

Private Sub LoadSettings()
    On Error GoTo ERR_HANDLER
    
    Dim SettingsString  As String, i As Long
    SettingsString = GetSetting(App.Title, "Settings", "SettingsString", "1,1,1,1,1,1,1,1,1,0,4,,0,0,1")
    
    Dim ArrTokens() As String
    ArrTokens = Split(SettingsString, ",")
    For i = LBound(ArrTokens) To UBound(ArrTokens)
        If i > UBound(Settings) Then Exit For
        Settings(i) = ArrTokens(i)
    Next
    
    Exit Sub
    
ERR_HANDLER:
    MsgBox "There was an error loading the previous settings"
End Sub

Private Sub SaveSettings()
    Dim SettingsString  As String, i As Long
    For i = LBound(Settings) To UBound(Settings)
        SettingsString = SettingsString & (CStr(Settings(i)) & ",")
    Next
    SettingsString = Mid$(SettingsString, 1, Len(SettingsString) - 1)
    SaveSetting App.Title, "Settings", "SettingsString", SettingsString
End Sub

Public Function GetCustomIcon(cmp As VBComponent) As String
    If cmp.Type = vbext_ct_RelatedDocument Or cmp.Type = vbext_ct_ResFile Then Exit Function
    On Error Resume Next
    GetCustomIcon = cmp.ReadProperty("TreeSurgeonIcon")
    On Error GoTo 0
End Function

Public Sub SetCustomIcon(cmp As VBComponent, IconFile As String)
    If cmp.Type = vbext_ct_RelatedDocument Or cmp.Type = vbext_ct_ResFile Then Exit Sub
    cmp.WriteProperty "TreeSurgeonIcon", IconFile
End Sub
