Attribute VB_Name = "Treedom"
Option Explicit

Private Declare Sub ODS Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TVITEM
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

Private Type TVITEMEX
    mask As Long
    hItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
End Type

Private Type TVSORTCB
    hParentItem As Long
    lpFnCompare As Long
    lParam As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Private Type NMTVCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemParam As Long
    clrText As Long
    clrBack As Long
    iLevel As Long
End Type

Private Const OBJ_FONT As Long = 6

Private Const LF_FACESIZE As Long = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_DEFERERASE As Long = &H2000
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOCOPYBITS As Long = &H100
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const WM_SIZE As Long = &H5

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long


Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Const GWL_WNDPROC As Long = -4

Private Const CCM_FIRST As Long = &H2000    ' First Common Control Message Index
Private Const CCM_GETUNICODEFORMAT As Long = (CCM_FIRST + 6)
Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)

'Private Const NMTVDISPINFOA As Long = TV_DISPINFOA
'Private Const NMTVDISPINFOW As Long = TV_DISPINFOW
'Private Const NMTVKEYDOWN As Long = TV_KEYDOWN
'Private Const TV_DISPINFO As Long = NMTVDISPINFO
'Private Const TV_DISPINFOA As Long = NMTVDISPINFOA
'Private Const TV_DISPINFOW As Long = NMTVDISPINFOW
Private Const TV_FIRST As Long = &H1100
Private Const TVN_FIRST As Long = (-400)
'Private Const TV_KEYDOWN As Long = NMTVKEYDOWN
Private Const TVC_BYKEYBOARD As Long = &H2
Private Const TVC_BYMOUSE As Long = &H1
Private Const TVC_UNKNOWN As Long = &H0
Private Const TVCDRF_NOIMAGES As Long = &H10000
Private Const TVE_COLLAPSE As Long = &H1
Private Const TVE_COLLAPSERESET As Long = &H8000
Private Const TVE_EXPAND As Long = &H2
Private Const TVE_EXPANDPARTIAL As Long = &H4000
Private Const TVE_TOGGLE As Long = &H3
Private Const TVGN_CARET As Long = &H9
Private Const TVGN_CHILD As Long = &H4
Private Const TVGN_DROPHILITE As Long = &H8
Private Const TVGN_FIRSTVISIBLE As Long = &H5
Private Const TVGN_LASTVISIBLE As Long = &HA
Private Const TVGN_NEXT As Long = &H1
Private Const TVGN_NEXTVISIBLE As Long = &H6
Private Const TVGN_PARENT As Long = &H3
Private Const TVGN_PREVIOUS As Long = &H2
Private Const TVGN_PREVIOUSVISIBLE As Long = &H7
Private Const TVGN_ROOT As Long = &H0
Private Const TVHT_ABOVE As Long = &H100
Private Const TVHT_BELOW As Long = &H200
Private Const TVHT_NOWHERE As Long = &H1
Private Const TVHT_ONITEMBUTTON As Long = &H10
Private Const TVHT_ONITEMICON As Long = &H2
Private Const TVHT_ONITEMINDENT As Long = &H8
Private Const TVHT_ONITEMLABEL As Long = &H4
Private Const TVHT_ONITEMRIGHT As Long = &H20
Private Const TVHT_ONITEMSTATEICON As Long = &H40
Private Const TVHT_ONITEM As Long = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
Private Const TVHT_TOLEFT As Long = &H800
Private Const TVHT_TORIGHT As Long = &H400
Private Const TVI_FIRST As Long = -&HFFFF&
Private Const TVI_LAST As Long = -&HFFFE&
Private Const TVI_ROOT As Long = -&H10000
Private Const TVI_SORT As Long = -&HFFFD&
Private Const TVIF_CHILDREN As Long = &H40
Private Const TVIF_DI_SETITEM As Long = &H1000
Private Const TVIF_HANDLE As Long = &H10
Private Const TVIF_IMAGE As Long = &H2
Private Const TVIF_INTEGRAL As Long = &H80
Private Const TVIF_PARAM As Long = &H4
Private Const TVIF_SELECTEDIMAGE As Long = &H20
Private Const TVIF_STATE As Long = &H8
Private Const TVIF_TEXT As Long = &H1
'Private Const TVINSERTSTRUCT_V1_SIZE As Long = TVINSERTSTRUCTW_V1_SIZE
'Private Const TVINSERTSTRUCTA As Long = TV_INSERTSTRUCTA
'Private Const TVINSERTSTRUCTW As Long = TV_INSERTSTRUCTW
Private Const TVIS_BOLD As Long = &H10
Private Const TVIS_CUT As Long = &H4
Private Const TVIS_DROPHILITED As Long = &H8
Private Const TVIS_EXPANDED As Long = &H20
Private Const TVIS_EXPANDEDONCE As Long = &H40
Private Const TVIS_EXPANDPARTIAL As Long = &H80
Private Const TVIS_OVERLAYMASK As Long = &HF00&
Private Const TVIS_SELECTED As Long = &H2
Private Const TVIS_STATEIMAGEMASK As Long = &HF000&
Private Const TVIS_USERMASK As Long = &HF000&
Private Const TVM_CREATEDRAGIMAGE As Long = (TV_FIRST + 18)
Private Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Private Const TVM_EDITLABELA As Long = (TV_FIRST + 14)
Private Const TVM_EDITLABELW As Long = (TV_FIRST + 65)
Private Const TVM_ENDEDITLABELNOW As Long = (TV_FIRST + 22)
Private Const TVM_ENSUREVISIBLE As Long = (TV_FIRST + 20)
Private Const TVM_EXPAND As Long = (TV_FIRST + 2)
Private Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Private Const TVM_GETCOUNT As Long = (TV_FIRST + 5)
Private Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
Private Const TVM_GETIMAGELIST As Long = (TV_FIRST + 8)
Private Const TVM_GETINDENT As Long = (TV_FIRST + 6)
Private Const TVM_GETINSERTMARKCOLOR As Long = (TV_FIRST + 38)
Private Const TVM_GETISEARCHSTRINGA As Long = (TV_FIRST + 23)
Private Const TVM_GETISEARCHSTRINGW As Long = (TV_FIRST + 64)
Private Const TVM_GETITEMA As Long = (TV_FIRST + 12)
Private Const TVM_GETITEMHEIGHT As Long = (TV_FIRST + 28)
Private Const TVM_GETITEMRECT As Long = (TV_FIRST + 4)
Private Const TVM_GETITEMSTATE As Long = (TV_FIRST + 39)
Private Const TVM_GETITEMW As Long = (TV_FIRST + 62)
Private Const TVM_GETLINECOLOR As Long = (TV_FIRST + 41)
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_GETSCROLLTIME As Long = (TV_FIRST + 34)
Private Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Private Const TVM_GETTOOLTIPS As Long = (TV_FIRST + 25)
Private Const TVM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
Private Const TVM_GETVISIBLECOUNT As Long = (TV_FIRST + 16)
Private Const TVM_HITTEST As Long = (TV_FIRST + 17)
Private Const TVM_INSERTITEMA As Long = (TV_FIRST + 0)
Private Const TVM_INSERTITEMW As Long = (TV_FIRST + 50)
Private Const TVM_MAPACCIDTOHTREEITEM As Long = (TV_FIRST + 42)
Private Const TVM_MAPHTREEITEMTOACCID As Long = (TV_FIRST + 43)
Private Const TVM_SELECTITEM As Long = (TV_FIRST + 11)
Private Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Private Const TVM_SETIMAGELIST As Long = (TV_FIRST + 9)
Private Const TVM_SETINDENT As Long = (TV_FIRST + 7)
Private Const TVM_SETINSERTMARK As Long = (TV_FIRST + 26)
Private Const TVM_SETINSERTMARKCOLOR As Long = (TV_FIRST + 37)
Private Const TVM_SETITEMA As Long = (TV_FIRST + 13)
Private Const TVM_SETITEMHEIGHT As Long = (TV_FIRST + 27)
Private Const TVM_SETITEMW As Long = (TV_FIRST + 63)
Private Const TVM_SETLINECOLOR As Long = (TV_FIRST + 40)
Private Const TVM_SETSCROLLTIME As Long = (TV_FIRST + 33)
Private Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Private Const TVM_SETTOOLTIPS As Long = (TV_FIRST + 24)
Private Const TVM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Private Const TVM_SORTCHILDREN As Long = (TV_FIRST + 19)
Private Const TVM_SORTCHILDRENCB As Long = (TV_FIRST + 21)
Private Const TVN_BEGINDRAGA As Long = (TVN_FIRST - 7)
Private Const TVN_BEGINDRAGW As Long = (TVN_FIRST - 56)
Private Const TVN_BEGINLABELEDITA As Long = (TVN_FIRST - 10)
Private Const TVN_BEGINLABELEDITW As Long = (TVN_FIRST - 59)
Private Const TVN_BEGINRDRAGA As Long = (TVN_FIRST - 8)
Private Const TVN_BEGINRDRAGW As Long = (TVN_FIRST - 57)
Private Const TVN_DELETEITEMA As Long = (TVN_FIRST - 9)
Private Const TVN_DELETEITEMW As Long = (TVN_FIRST - 58)
Private Const TVN_ENDLABELEDITA As Long = (TVN_FIRST - 11)
Private Const TVN_ENDLABELEDITW As Long = (TVN_FIRST - 60)
Private Const TVN_GETDISPINFOA As Long = (TVN_FIRST - 3)
Private Const TVN_GETDISPINFOW As Long = (TVN_FIRST - 52)
Private Const TVN_GETINFOTIPA As Long = (TVN_FIRST - 13)
Private Const TVN_GETINFOTIPW As Long = (TVN_FIRST - 14)
Private Const TVN_ITEMEXPANDEDA As Long = (TVN_FIRST - 6)
Private Const TVN_ITEMEXPANDEDW As Long = (TVN_FIRST - 55)
Private Const TVN_ITEMEXPANDINGA As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDINGW As Long = (TVN_FIRST - 54)
Private Const TVN_KEYDOWN As Long = (TVN_FIRST - 12)
Private Const TVN_LAST As Long = (-499)
Private Const TVN_SELCHANGEDA As Long = (TVN_FIRST - 2)
Private Const TVN_SELCHANGEDW As Long = (TVN_FIRST - 51)
Private Const TVN_SELCHANGINGA As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGINGW As Long = (TVN_FIRST - 50)
Private Const TVN_SETDISPINFOA As Long = (TVN_FIRST - 4)
Private Const TVN_SETDISPINFOW As Long = (TVN_FIRST - 53)
Private Const TVN_SINGLEEXPAND As Long = (TVN_FIRST - 15)
Private Const TVNRET_DEFAULT As Long = 0
Private Const TVNRET_SKIPNEW As Long = 2
Private Const TVNRET_SKIPOLD As Long = 1
Private Const TVS_CHECKBOXES As Long = &H100
Private Const TVS_DISABLEDRAGDROP As Long = &H10
Private Const TVS_EDITLABELS As Long = &H8
Private Const TVS_FULLROWSELECT As Long = &H1000
Private Const TVS_HASBUTTONS As Long = &H1
Private Const TVS_HASLINES As Long = &H2
Private Const TVS_INFOTIP As Long = &H800
Private Const TVS_LINESATROOT As Long = &H4
Private Const TVS_NOHSCROLL As Long = &H8000
Private Const TVS_NONEVENHEIGHT As Long = &H4000
Private Const TVS_NOSCROLL As Long = &H2000
Private Const TVS_NOTOOLTIPS As Long = &H80
Private Const TVS_RTLREADING As Long = &H40
Private Const TVS_SHOWSELALWAYS As Long = &H20
Private Const TVS_SINGLEEXPAND As Long = &H400
Private Const TVS_TRACKSELECT As Long = &H200
Private Const TVSIF_NOSINGLEEXPAND As Long = &H8000
Private Const TVSIL_NORMAL As Long = 0
Private Const TVSIL_STATE As Long = 2

Private Const WM_NOTIFY As Long = &H4E
Private Const NM_FIRST As Long = 0
Private Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)

Private Const CDRF_DODEFAULT As Long = &H0
Private Const CDRF_NEWFONT As Long = &H2
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20
Private Const CDRF_NOTIFYPOSTERASE As Long = &H40
Private Const CDRF_NOTIFYPOSTPAINT As Long = &H10
Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20
Private Const CDRF_SKIPDEFAULT As Long = &H4
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_POSTERASE As Long = &H4
Private Const CDDS_POSTPAINT As Long = &H2
Private Const CDDS_PREPAINT As Long = &H1
Private Const CDDS_PREERASE As Long = &H3
Private Const CDDS_SUBITEM As Long = &H20000
Private Const CDDS_ITEMPOSTERASE As Long = (CDDS_ITEM Or CDDS_POSTERASE)
Private Const CDDS_ITEMPOSTPAINT As Long = (CDDS_ITEM Or CDDS_POSTPAINT)
Private Const CDDS_ITEMPREERASE As Long = (CDDS_ITEM Or CDDS_PREERASE)
Private Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM Or CDDS_PREPAINT)

Private Const CDIS_CHECKED As Long = &H8
Private Const CDIS_DEFAULT As Long = &H20
Private Const CDIS_DISABLED As Long = &H4
Private Const CDIS_FOCUS As Long = &H10
Private Const CDIS_GRAYED As Long = &H2
Private Const CDIS_HOT As Long = &H40
Private Const CDIS_INDETERMINATE As Long = &H100
Private Const CDIS_MARKED As Long = &H80
Private Const CDIS_SELECTED As Long = &H1
Private Const CDIS_SHOWKEYBOARDCUES As Long = &H200

'USER32
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GCL_WNDPROC As Long = -24
Private Const WM_PAINT As Long = &HF&
Private Const WM_ERASEBKGND As Long = &H14

'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Const DEFAULT_GUI_FONT As Long = 17

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long

Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long


'
Private VBInstance As VBE
Private lpOldWndProcTree As Long, lpOldWndProcProject As Long
Private lpDefTreeWndProc As Long
Private hWndProject As Long, hWndTree As Long
Private iTimerID As Long

Public Function BeginHook(vbInst As VBE) As Boolean
    Set VBInstance = vbInst
    
    hWndProject = FindWindowEx(vbInst.MainWindow.hWnd, 0, "PROJECT", vbNullString)
    If hWndProject = 0 Then GoTo HANDLE_ERROR
    
    hWndTree = FindWindowEx(hWndProject, 0, "SysTreeView32", vbNullString)
    If hWndTree = 0 Then GoTo HANDLE_ERROR
    
    lpOldWndProcProject = SetWindowLong(hWndProject, GWL_WNDPROC, AddressOf ProjectWndProc)
    If lpOldWndProcProject = 0 Then GoTo HANDLE_ERROR
    
    lpOldWndProcTree = SetWindowLong(hWndTree, GWL_WNDPROC, AddressOf TreeWndProc)
    If lpOldWndProcTree = 0 Then GoTo HANDLE_ERROR
    
    lpDefTreeWndProc = GetClassLong(hWndTree, GCL_WNDPROC)
    
    iTimerID = SetTimer(0, 0, 250, AddressOf RefreshTree)
    
    SetParent MyToolbar.hWnd, hWndProject
    MyToolbar.Move 0, 0
    MyToolbar.Visible = True
    MyToolbar.ZOrder
    ReposToolbar
    
    'MsgBox "Hooked oK"
    BeginHook = True
    Exit Function
HANDLE_ERROR:
    MsgBox "Error, unhooking"
    EndHook
End Function

Public Sub EndHook()
    SetParent MyToolbar.hWnd, 0
    Unload MyToolbar
    
    If iTimerID <> 0 Then
        KillTimer 0, iTimerID
    End If
    If hWndProject <> 0 Then
        If lpOldWndProcProject <> 0 Then
            SetWindowLong hWndProject, GWL_WNDPROC, lpOldWndProcProject
            lpOldWndProcProject = 0
        End If
        hWndProject = 0
    End If
    If hWndTree <> 0 Then
        If lpOldWndProcTree <> 0 Then
            SetWindowLong hWndTree, GWL_WNDPROC, lpOldWndProcTree
            lpOldWndProcTree = 0
        End If
        RedrawWindow hWndTree, ByVal 0&, 0, 1
        UpdateWindow hWndTree
        hWndTree = 0
    End If
End Sub

Private Function TreeWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = TVM_SORTCHILDREN Then
        Dim obj As Object, prj As VBProject, cmp As VBComponent
        On Error Resume Next
        Set prj = GetNodeProject(lParam)
        If prj Is Nothing Then GoTo DO_DEFAULT
        
        Dim col As New Collection, hItem As Long
        hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_CHILD, ByVal lParam)
        While hItem <> 0
            col.Add GetNodeObject(hItem), "k" & Hex$(GetNodeParam(hWnd, hItem))
            hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
        Wend
        Dim sort As TVSORTCB
        sort.hParentItem = lParam
        sort.lParam = VarPtr(col)
        sort.lpFnCompare = GetFunctionAddress(AddressOf TreeCompareItems)
        TreeWndProc = SendMessage(hWnd, TVM_SORTCHILDRENCB, 0, sort)
        
        ' Now it's sorted, refresh the project sort string
        Dim s As String
        s = ":"
        hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_CHILD, ByVal lParam)
        While hItem <> 0
            s = s & GetNodeText(hWnd, hItem) & ":"
            hItem = SendMessage(hWnd, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
        Wend
        prj.WriteProperty "Treedom4VB", "ItemsOrder", s
        
        Exit Function
    End If
    
DO_DEFAULT:
    TreeWndProc = CallWindowProc(lpOldWndProcTree, hWnd, uMsg, wParam, lParam)
End Function

Private Function ProjectWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_NOTIFY Then
        'If wParam = 1 Then
            Dim nm As NMHDR
            CopyMemory nm, ByVal lParam, LenB(nm)
            If nm.Code = NM_CUSTOMDRAW Then
                Dim tv As NMTVCUSTOMDRAW
                Static hFontDefault As Long
                Static hFontBold As Long
                Static hFontItalic As Long
                Static hFontBoldItalic As Long
                
                CopyMemory tv, ByVal lParam, LenB(tv)
                
                Select Case tv.dwDrawStage
                Case CDDS_PREPAINT
                    ' Construct fonts
                    hFontDefault = GetStockObject(DEFAULT_GUI_FONT) 'GetCurrentObject(tv.hdc, OBJ_FONT)
                    Dim lf As LOGFONT
                    GetObject hFontDefault, LenB(lf), lf
                    lf.lfWeight = 700 ' Bold
                    lf.lfItalic = 0 ' No Italic
                    hFontBold = CreateFontIndirect(lf)
                    lf.lfWeight = 700 ' Bold
                    lf.lfItalic = 1 ' Italic
                    hFontBoldItalic = CreateFontIndirect(lf)
                    lf.lfWeight = 0  ' Not Bold
                    lf.lfItalic = 1 ' Italic
                    hFontItalic = CreateFontIndirect(lf)
                    ' We want per-item notifications, and postpaint for clearup
                    ProjectWndProc = CDRF_NOTIFYITEMDRAW Or CDRF_NOTIFYPOSTPAINT
                    Exit Function
                Case CDDS_POSTPAINT
                    ' Delete our fonts
                    If hFontBold <> 0 Then
                        DeleteObject hFontBold
                        hFontBold = 0
                    End If
                    If hFontBoldItalic <> 0 Then
                        DeleteObject hFontBoldItalic
                        hFontBoldItalic = 0
                    End If
                    If hFontItalic <> 0 Then
                        DeleteObject hFontItalic
                        hFontItalic = 0
                    End If
                Case CDDS_ITEMPREPAINT
                    On Error Resume Next
                    Dim prj As VBProject, cmp As VBComponent, obj As Object
                    Dim IsStartup As Boolean
                    Set obj = GetNodeObject(tv.dwItemSpec)
                    If obj Is Nothing Then
                        ' Do default and return
                        ProjectWndProc = CDRF_DODEFAULT
                        Exit Function
                    End If
                    If TypeOf obj Is VBProject Then
                        Set prj = obj
                        If prj Is prj.Collection.StartProject Then
                            If prj.IsDirty Then
                                SelectObject tv.hdc, hFontBoldItalic
                            Else
                                SelectObject tv.hdc, hFontBold
                            End If
                        Else
                            If prj.IsDirty Then
                                SelectObject tv.hdc, hFontItalic
                            Else
                                SelectObject tv.hdc, hFontDefault
                            End If
                        End If
                    ElseIf TypeOf obj Is VBComponent Then
                        Set cmp = obj
                        Set prj = cmp.Collection.Parent
                        IsStartup = False
                        If VarType(prj.VBComponents.StartUpObject) = vbLong Then
                            'ODS "Startup: Zilch or Sub Main" & vbCrLf
                            If CLng(prj.VBComponents.StartUpObject) = 0 Then
                                'ODS "Sub Main" & vbCrLf
                                ' Sub Main startup
                                If cmp.Type = vbext_ct_StdModule Then
                                    'ODS "> Module '" & cmp.Name & "'" & vbCrLf
                                    Dim mem As Member
                                    Set mem = cmp.CodeModule.Members("Main")
                                    If Not (mem Is Nothing) Then
                                        'ODS ">> Got a member called 'Main'" & vbCrLf
                                        If mem.Type = vbext_mt_Method Then
                                            'ODS ">>> It's a method!" & vbCrLf
                                            IsStartup = True
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            ' There is a startup component
                            If prj.VBComponents.StartUpObject Is cmp Then
                                IsStartup = True
                            End If
                        End If
                        
                        If IsStartup Then
                            If cmp.IsDirty Then
                                SelectObject tv.hdc, hFontBoldItalic
                            Else
                                SelectObject tv.hdc, hFontBold
                            End If
                        Else
                            If cmp.IsDirty Then
                                SelectObject tv.hdc, hFontItalic
                            Else
                                SelectObject tv.hdc, hFontDefault
                            End If
                        End If
                    Else
                        ' Should never happen, but handle anyway
                        ProjectWndProc = CDRF_DODEFAULT
                        Exit Function
                    End If
                    
                    CopyMemory ByVal lParam, tv, LenB(tv)
                    ProjectWndProc = CDRF_NEWFONT
                    Exit Function
                End Select
                
            End If
        'End If
    End If
    If uMsg = WM_SIZE Then
        ProjectWndProc = CallWindowProc(lpOldWndProcProject, hWnd, uMsg, wParam, lParam)
        ReposToolbar
        Exit Function
    End If
    ProjectWndProc = CallWindowProc(lpOldWndProcProject, hWnd, uMsg, wParam, lParam)
End Function

Private Function ReposToolbar()
    Dim rc As RECT, hWndTemp As Long
    Dim rc2 As RECT
    GetClientRect hWndProject, rc
    hWndTemp = FindWindowEx(hWndProject, 0, "MsoCommandBarDock", "MsoDockTop")
    If hWndTemp <> 0 Then
        GetClientRect hWndTemp, rc2
        SetWindowPos hWndTemp, 0, 0, 0, rc.Right - MyToolbar.Width \ Screen.TwipsPerPixelX, rc2.Bottom, SWP_NOOWNERZORDER
        MyToolbar.Move rc.Right * Screen.TwipsPerPixelX - MyToolbar.Width, 0, MyToolbar.Width, rc2.Bottom * Screen.TwipsPerPixelY
    End If
End Function

Private Function GetNodeText(ByVal hTree As Long, ByVal hItem As Long) As String
    Dim item As TVITEM
    item.mask = TVIF_TEXT
    item.hItem = hItem
    item.pszText = String$(256, 0)
    item.cchTextMax = 255
    If SendMessage(hTree, TVM_GETITEMA, 0, item) Then
        GetNodeText = Left$(item.pszText, InStr(item.pszText, vbNullChar) - 1)
    End If
End Function

Private Function GetNodeIcon(ByVal hTree As Long, ByVal hItem As Long) As Long
    Dim item As TVITEM
    item.mask = TVIF_IMAGE
    item.hItem = hItem
    If SendMessage(hTree, TVM_GETITEMA, 0, item) Then
        GetNodeIcon = item.iImage
    End If
End Function

Private Function GetNodeParam(ByVal hTree As Long, ByVal hItem As Long) As Long
    Dim item As TVITEM
    item.mask = TVIF_PARAM
    item.hItem = hItem
    If SendMessage(hTree, TVM_GETITEMA, 0, item) Then
        GetNodeParam = item.lParam
    End If
End Function

Public Function DebugTree() As String
    DebugTree = DebugTreeItem(0, 0)
End Function

Public Function DebugTreeItem(ByVal hItem As Long, ByVal Level As Long) As String
    If hItem = 0 Then
        hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
        Level = 0
    End If
    If hItem = 0 Then Exit Function
    Dim hItemTemp As Long
    ' Handle Main Item
    DebugTreeItem = Space$(Level) & GetNodeText(hWndTree, hItem) & " : " & GetNodeIcon(hWndTree, hItem) & " - lParam = " & GetNodeParam(hWndTree, hItem) & vbCrLf
    ' Handle First Child
    hItemTemp = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItem)
    If hItemTemp <> 0 Then DebugTreeItem = DebugTreeItem & DebugTreeItem(hItemTemp, Level + 1)
    ' Handle siblings
    hItemTemp = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    If hItemTemp <> 0 Then DebugTreeItem = DebugTreeItem & DebugTreeItem(hItemTemp, Level)
End Function

Private Function GetNodeProject(ByVal hItem As Long) As VBProject
    If hItem = 0 Then Exit Function
    On Error GoTo HANDLE_ANY_ERROR
    
    Dim Icon As Long
    Dim Text As String
    Icon = GetNodeIcon(hWndTree, hItem)
    Select Case Icon
    Case 4, 19, 20, 21 ' Project
        Text = GetNodeText(hWndTree, hItem)
        Text = Left$(Text, InStr(Text, " ") - 1)
        Set GetNodeProject = VBInstance.VBProjects(Text)
    Case Else
        Set GetNodeProject = GetNodeProject(SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem))
    End Select
    
    Exit Function
HANDLE_ANY_ERROR:
    Set GetNodeProject = Nothing
    Exit Function
End Function

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
                If cmp.FileNames(1) = txt_file Then
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
                If cmp.FileNames(1) = txt_file Then
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

Private Sub RefreshTree(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    RedrawWindow hWndTree, ByVal 0&, 0, 1
    UpdateWindow hWndTree
    
    'MyToolbar.ZOrder
End Sub

Private Function GetFunctionAddress(ByVal lpFn As Long) As Long
    GetFunctionAddress = lpFn
End Function

' Extract the filename part from a full path
Public Function FileNameFromPath(ByVal sPath As String) As String
    Dim i As Integer
    i = InStrRev(sPath, "\")
    If i < 1 Then
        FileNameFromPath = sPath
    Else
        FileNameFromPath = Mid$(sPath, i + 1)
    End If
End Function

' Returns the text that the project tree node should have for a given component
' Related Docs and ResFiles don't have Names
' if the component is not saved its displayed filename will be the same as its name
Private Function ComponentNodeText(ByVal cmp As VBComponent) As String
    If cmp.Type <> vbext_ct_RelatedDocument And cmp.Type <> vbext_ct_ResFile Then
        ComponentNodeText = cmp.Name
    End If
    If cmp.FileNames(1) = "" Then
        ComponentNodeText = ComponentNodeText & (" (" & cmp.Name & ")")
    Else
        ComponentNodeText = ComponentNodeText & (" (" & FileNameFromPath(cmp.FileNames(1)) & ")")
    End If
End Function

Private Function TreeCompareItems(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParamSort As Long) As Long
    Dim col As Collection
    CopyMemory col, ByVal lParamSort, 4
    Dim cmp1 As VBComponent, cmp2 As VBComponent
    Dim prj As VBProject, sOrder As String
    Dim s1 As String, s2 As String
    Dim i1 As Long, i2 As Long
    On Error Resume Next
    Set cmp1 = col("k" & Hex$(lParam1))
    Set cmp2 = col("k" & Hex$(lParam2))
    If (cmp1 Is Nothing) Or (cmp2 Is Nothing) Then
        TreeCompareItems = 0
    Else
        Debug.Assert cmp1.Collection.Parent Is cmp2.Collection.Parent
        Set prj = cmp1.Collection.Parent
        s1 = ComponentNodeText(cmp1)
        s2 = ComponentNodeText(cmp2)
        
        sOrder = ""
        sOrder = prj.ReadProperty("Treedom4VB", "ItemsOrder")
        If sOrder = "" Then sOrder = ":"
        
        i1 = InStr(sOrder, ":" & s1 & ":")
        i2 = InStr(sOrder, ":" & s2 & ":")
        If i1 > 0 And i2 > 0 Then
            TreeCompareItems = i1 - i2
            GoTo ALL_DONE
        End If
        If cmp1.Type = cmp2.Type Then
            If cmp1.Type = vbext_ct_ResFile Or cmp1.Type = vbext_ct_RelatedDocument Then
                TreeCompareItems = StrComp(cmp1.FileNames(1), cmp2.FileNames(1))
            Else
                TreeCompareItems = StrComp(cmp1.Name, cmp2.Name)
            End If
        Else
            TreeCompareItems = cmp1.Type - cmp2.Type
        End If
    End If
ALL_DONE:
    CopyMemory col, 0&, 4
End Function

Public Sub MoveItemUp()
    Dim s As String, ts As String
    Dim arr() As String
    Dim i As Long
    
    If VBInstance.SelectedVBComponent Is Nothing Then Exit Sub
    s = ComponentNodeText(VBInstance.SelectedVBComponent)
    arr = Split(VBInstance.SelectedVBComponent.Collection.Parent.ReadProperty("Treedom4VB", "ItemsOrder"), ":")
    
    If arr(1) = s Then
        ' Can't move up
        Exit Sub
    End If
    
    For i = 2 To UBound(arr)
        If arr(i) = s Then
            arr(i) = arr(i - 1)
            arr(i - 1) = s
            Exit For
        End If
    Next
    
    Call VBInstance.SelectedVBComponent.Collection.Parent.WriteProperty("Treedom4VB", "ItemsOrder", Join(arr, ":"))
    
    Dim hItem As Long
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
    SendMessage hWndTree, TVM_SORTCHILDREN, 0, ByVal hItem
End Sub

Public Sub MoveItemDown()
    Dim s As String, ts As String
    Dim arr() As String
    Dim i As Long
    
    If VBInstance.SelectedVBComponent Is Nothing Then Exit Sub
    s = ComponentNodeText(VBInstance.SelectedVBComponent)
    arr = Split(VBInstance.SelectedVBComponent.Collection.Parent.ReadProperty("Treedom4VB", "ItemsOrder"), ":")
    
    For i = 1 To UBound(arr) - 1
        If arr(i) = s Then
            arr(i) = arr(i + 1)
            arr(i + 1) = s
            Exit For
        End If
    Next
    
    Call VBInstance.SelectedVBComponent.Collection.Parent.WriteProperty("Treedom4VB", "ItemsOrder", Join(arr, ":"))
    
    Dim hItem As Long
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    hItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
    SendMessage hWndTree, TVM_SORTCHILDREN, 0, ByVal hItem
End Sub


