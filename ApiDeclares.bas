Attribute VB_Name = "ApiDeclares"
Option Explicit

Public Declare Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Public Declare Sub DebugBreak Lib "kernel32.dll" ()

Public Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long

Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type TVITEM
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

Public Type TVITEMEX
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

Public Type TVSORTCB
    hParentItem As Long
    lpFnCompare As Long
    lParam As Long
End Type

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Public Type NMTVCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemParam As Long
    clrText As Long
    clrBack As Long
    iLevel As Long
End Type

Public Type TVHITTESTINFO
    pt As POINTAPI
    flags As Long
    hItem As Long
End Type

Public Const OBJ_FONT As Long = 6

Public Const LF_FACESIZE As Long = 32
Public Const LF_FULLFACESIZE As Long = 64

Public Type LOGFONT
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

Public Const NONANTIALIASED_QUALITY As Long = 3&
Public Const ANTIALIASED_QUALITY As Long = 4&
Public Const CLEARTYPE_QUALITY As Long = 5&
Public Const CLEARTYPE_NATURAL_QUALITY As Long = 6&

Public Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hDC As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Public Type ENUMLOGFONTEX
        elfLogFont As LOGFONT
        elfFullName(LF_FULLFACESIZE) As Byte
        elfStyle(LF_FACESIZE) As Byte
        elfScript(LF_FACESIZE) As Byte
End Type
Public Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type
Public Type FONTSIGNATURE
        fsUsb(4) As Long
        fsCsb(2) As Long
End Type
Public Type NEWTEXTMETRICEX
        ntmTm As NEWTEXTMETRIC
        ntmFontSig As FONTSIGNATURE
End Type


Public Const SWP_ASYNCWINDOWPOS As Long = &H4000
Public Const SWP_DEFERERASE As Long = &H2000
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW As Long = &H80
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOCOPYBITS As Long = &H100
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOREDRAW As Long = &H8
Public Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
Public Const SWP_NOSENDCHANGING As Long = &H400
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_SHOWWINDOW As Long = &H40

Public Const WM_SIZE As Long = &H5
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_DRAWITEM As Long = &H2B
Public Const WM_INITMENUPOPUP As Long = &H117
Public Const WM_MEASUREITEM As Long = &H2C
Public Const WM_MENUSELECT As Long = &H11F

Public Const CB_FINDSTRING  As Long = &H14C
Public Const CB_FINDSTRINGEXACT  As Long = &H158

Public Const SW_SHOWNORMAL As Long = &H1&

Public Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function AbortPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function FlattenPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPath Lib "gdi32" (ByVal hDC As Long, lpPoint As Any, lpTypes As Any, ByVal nSize As Long) As Long


Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

Public Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long


Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long


Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Const DI_NORMAL = &H3
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type


Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CXSMICON As Long = 49&

Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const GWL_WNDPROC As Long = -4

Public Const CCM_FIRST As Long = &H2000    ' First Common Control Message Index
Public Const CCM_GETUNICODEFORMAT As Long = (CCM_FIRST + 6)
Public Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)

'Public Const NMTVDISPINFOA As Long = TV_DISPINFOA
'Public Const NMTVDISPINFOW As Long = TV_DISPINFOW
'Public Const NMTVKEYDOWN As Long = TV_KEYDOWN
'Public Const TV_DISPINFO As Long = NMTVDISPINFO
'Public Const TV_DISPINFOA As Long = NMTVDISPINFOA
'Public Const TV_DISPINFOW As Long = NMTVDISPINFOW
Public Const TV_FIRST As Long = &H1100
Public Const TVN_FIRST As Long = (-400)
'Public Const TV_KEYDOWN As Long = NMTVKEYDOWN
Public Const TVC_BYKEYBOARD As Long = &H2
Public Const TVC_BYMOUSE As Long = &H1
Public Const TVC_UNKNOWN As Long = &H0
Public Const TVCDRF_NOIMAGES As Long = &H10000
Public Const TVE_COLLAPSE As Long = &H1
Public Const TVE_COLLAPSERESET As Long = &H8000
Public Const TVE_EXPAND As Long = &H2
Public Const TVE_EXPANDPARTIAL As Long = &H4000
Public Const TVE_TOGGLE As Long = &H3
Public Const TVGN_CARET As Long = &H9
Public Const TVGN_CHILD As Long = &H4
Public Const TVGN_DROPHILITE As Long = &H8
Public Const TVGN_FIRSTVISIBLE As Long = &H5
Public Const TVGN_LASTVISIBLE As Long = &HA
Public Const TVGN_NEXT As Long = &H1
Public Const TVGN_NEXTVISIBLE As Long = &H6
Public Const TVGN_PARENT As Long = &H3
Public Const TVGN_PREVIOUS As Long = &H2
Public Const TVGN_PREVIOUSVISIBLE As Long = &H7
Public Const TVGN_ROOT As Long = &H0
Public Const TVHT_ABOVE As Long = &H100
Public Const TVHT_BELOW As Long = &H200
Public Const TVHT_NOWHERE As Long = &H1
Public Const TVHT_ONITEMBUTTON As Long = &H10
Public Const TVHT_ONITEMICON As Long = &H2
Public Const TVHT_ONITEMINDENT As Long = &H8
Public Const TVHT_ONITEMLABEL As Long = &H4
Public Const TVHT_ONITEMRIGHT As Long = &H20
Public Const TVHT_ONITEMSTATEICON As Long = &H40
Public Const TVHT_ONITEM As Long = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
Public Const TVHT_TOLEFT As Long = &H800
Public Const TVHT_TORIGHT As Long = &H400
Public Const TVI_FIRST As Long = -&HFFFF&
Public Const TVI_LAST As Long = -&HFFFE&
Public Const TVI_ROOT As Long = -&H10000
Public Const TVI_SORT As Long = -&HFFFD&
Public Const TVIF_CHILDREN As Long = &H40
Public Const TVIF_DI_SETITEM As Long = &H1000
Public Const TVIF_HANDLE As Long = &H10
Public Const TVIF_IMAGE As Long = &H2
Public Const TVIF_INTEGRAL As Long = &H80
Public Const TVIF_PARAM As Long = &H4
Public Const TVIF_SELECTEDIMAGE As Long = &H20
Public Const TVIF_STATE As Long = &H8
Public Const TVIF_TEXT As Long = &H1
'Public Const TVINSERTSTRUCT_V1_SIZE As Long = TVINSERTSTRUCTW_V1_SIZE
'Public Const TVINSERTSTRUCTA As Long = TV_INSERTSTRUCTA
'Public Const TVINSERTSTRUCTW As Long = TV_INSERTSTRUCTW
Public Const TVIS_BOLD As Long = &H10
Public Const TVIS_CUT As Long = &H4
Public Const TVIS_DROPHILITED As Long = &H8
Public Const TVIS_EXPANDED As Long = &H20
Public Const TVIS_EXPANDEDONCE As Long = &H40
Public Const TVIS_EXPANDPARTIAL As Long = &H80
Public Const TVIS_OVERLAYMASK As Long = &HF00&
Public Const TVIS_SELECTED As Long = &H2
Public Const TVIS_STATEIMAGEMASK As Long = &HF000&
Public Const TVIS_USERMASK As Long = &HF000&
Public Const TVM_CREATEDRAGIMAGE As Long = (TV_FIRST + 18)
Public Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Public Const TVM_EDITLABELA As Long = (TV_FIRST + 14)
Public Const TVM_EDITLABELW As Long = (TV_FIRST + 65)
Public Const TVM_ENDEDITLABELNOW As Long = (TV_FIRST + 22)
Public Const TVM_ENSUREVISIBLE As Long = (TV_FIRST + 20)
Public Const TVM_EXPAND As Long = (TV_FIRST + 2)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETCOUNT As Long = (TV_FIRST + 5)
Public Const TVM_GETEDITCONTROL As Long = (TV_FIRST + 15)
Public Const TVM_GETIMAGELIST As Long = (TV_FIRST + 8)
Public Const TVM_GETINDENT As Long = (TV_FIRST + 6)
Public Const TVM_GETINSERTMARKCOLOR As Long = (TV_FIRST + 38)
Public Const TVM_GETISEARCHSTRINGA As Long = (TV_FIRST + 23)
Public Const TVM_GETISEARCHSTRINGW As Long = (TV_FIRST + 64)
Public Const TVM_GETITEMA As Long = (TV_FIRST + 12)
Public Const TVM_GETITEMHEIGHT As Long = (TV_FIRST + 28)
Public Const TVM_GETITEMRECT As Long = (TV_FIRST + 4)
Public Const TVM_GETITEMSTATE As Long = (TV_FIRST + 39)
Public Const TVM_GETITEMW As Long = (TV_FIRST + 62)
Public Const TVM_GETLINECOLOR As Long = (TV_FIRST + 41)
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETSCROLLTIME As Long = (TV_FIRST + 34)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Public Const TVM_GETTOOLTIPS As Long = (TV_FIRST + 25)
Public Const TVM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
Public Const TVM_GETVISIBLECOUNT As Long = (TV_FIRST + 16)
Public Const TVM_HITTEST As Long = (TV_FIRST + 17)
Public Const TVM_INSERTITEMA As Long = (TV_FIRST + 0)
Public Const TVM_INSERTITEMW As Long = (TV_FIRST + 50)
Public Const TVM_MAPACCIDTOHTREEITEM As Long = (TV_FIRST + 42)
Public Const TVM_MAPHTREEITEMTOACCID As Long = (TV_FIRST + 43)
Public Const TVM_SELECTITEM As Long = (TV_FIRST + 11)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETIMAGELIST As Long = (TV_FIRST + 9)
Public Const TVM_SETINDENT As Long = (TV_FIRST + 7)
Public Const TVM_SETINSERTMARK As Long = (TV_FIRST + 26)
Public Const TVM_SETINSERTMARKCOLOR As Long = (TV_FIRST + 37)
Public Const TVM_SETITEMA As Long = (TV_FIRST + 13)
Public Const TVM_SETITEMHEIGHT As Long = (TV_FIRST + 27)
Public Const TVM_SETITEMW As Long = (TV_FIRST + 63)
Public Const TVM_SETLINECOLOR As Long = (TV_FIRST + 40)
Public Const TVM_SETSCROLLTIME As Long = (TV_FIRST + 33)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_SETTOOLTIPS As Long = (TV_FIRST + 24)
Public Const TVM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT
Public Const TVM_SORTCHILDREN As Long = (TV_FIRST + 19)
Public Const TVM_SORTCHILDRENCB As Long = (TV_FIRST + 21)
Public Const TVN_BEGINDRAGA As Long = (TVN_FIRST - 7)
Public Const TVN_BEGINDRAGW As Long = (TVN_FIRST - 56)
Public Const TVN_BEGINLABELEDITA As Long = (TVN_FIRST - 10)
Public Const TVN_BEGINLABELEDITW As Long = (TVN_FIRST - 59)
Public Const TVN_BEGINRDRAGA As Long = (TVN_FIRST - 8)
Public Const TVN_BEGINRDRAGW As Long = (TVN_FIRST - 57)
Public Const TVN_DELETEITEMA As Long = (TVN_FIRST - 9)
Public Const TVN_DELETEITEMW As Long = (TVN_FIRST - 58)
Public Const TVN_ENDLABELEDITA As Long = (TVN_FIRST - 11)
Public Const TVN_ENDLABELEDITW As Long = (TVN_FIRST - 60)
Public Const TVN_GETDISPINFOA As Long = (TVN_FIRST - 3)
Public Const TVN_GETDISPINFOW As Long = (TVN_FIRST - 52)
Public Const TVN_GETINFOTIPA As Long = (TVN_FIRST - 13)
Public Const TVN_GETINFOTIPW As Long = (TVN_FIRST - 14)
Public Const TVN_ITEMEXPANDEDA As Long = (TVN_FIRST - 6)
Public Const TVN_ITEMEXPANDEDW As Long = (TVN_FIRST - 55)
Public Const TVN_ITEMEXPANDINGA As Long = (TVN_FIRST - 5)
Public Const TVN_ITEMEXPANDINGW As Long = (TVN_FIRST - 54)
Public Const TVN_KEYDOWN As Long = (TVN_FIRST - 12)
Public Const TVN_LAST As Long = (-499)
Public Const TVN_SELCHANGEDA As Long = (TVN_FIRST - 2)
Public Const TVN_SELCHANGEDW As Long = (TVN_FIRST - 51)
Public Const TVN_SELCHANGINGA As Long = (TVN_FIRST - 1)
Public Const TVN_SELCHANGINGW As Long = (TVN_FIRST - 50)
Public Const TVN_SETDISPINFOA As Long = (TVN_FIRST - 4)
Public Const TVN_SETDISPINFOW As Long = (TVN_FIRST - 53)
Public Const TVN_SINGLEEXPAND As Long = (TVN_FIRST - 15)
Public Const TVNRET_DEFAULT As Long = 0
Public Const TVNRET_SKIPNEW As Long = 2
Public Const TVNRET_SKIPOLD As Long = 1
Public Const TVS_CHECKBOXES As Long = &H100
Public Const TVS_DISABLEDRAGDROP As Long = &H10
Public Const TVS_EDITLABELS As Long = &H8
Public Const TVS_FULLROWSELECT As Long = &H1000
Public Const TVS_HASBUTTONS As Long = &H1
Public Const TVS_HASLINES As Long = &H2
Public Const TVS_INFOTIP As Long = &H800
Public Const TVS_LINESATROOT As Long = &H4
Public Const TVS_NOHSCROLL As Long = &H8000
Public Const TVS_NONEVENHEIGHT As Long = &H4000
Public Const TVS_NOSCROLL As Long = &H2000
Public Const TVS_NOTOOLTIPS As Long = &H80
Public Const TVS_RTLREADING As Long = &H40
Public Const TVS_SHOWSELALWAYS As Long = &H20
Public Const TVS_SINGLEEXPAND As Long = &H400
Public Const TVS_TRACKSELECT As Long = &H200
Public Const TVSIF_NOSINGLEEXPAND As Long = &H8000
Public Const TVSIL_NORMAL As Long = 0
Public Const TVSIL_STATE As Long = 2

Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long


' ****************************************************************************************************
' * Declares for: Shell API                                                                          *
' ****************************************************************************************************
Private Declare Function SHChangeIconDialog Lib "Shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal dwMaxFile As Long, lpIconIndex As Long) As Long
Public Declare Function PathParseIconLocation Lib "shlwapi.dll" Alias "PathParseIconLocationA" (ByVal pszIconFile As String) As Long
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
' ****************************************************************************************************
' * Declares for: ImageList API                                                                      *
' ****************************************************************************************************
Public Declare Function ImageList_Add Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Public Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Function ImageList_Copy Lib "comctl32.dll" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'Public Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (ByRef pimldp As IMAGELISTDRAWPARAMS) As Long
Public Declare Function ImageList_Duplicate Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Public Declare Function ImageList_GetBkColor Lib "comctl32.dll" (ByVal hIml As Long) As Long
'Public Declare Function ImageList_GetDragImage Lib "comctl32.dll" (ByRef ppt As Point, ByRef pptHotspot As Point) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_ExtractIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal hInstance As Long, ByVal i As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIml As Long, ByRef cx As Long, ByRef cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
'Public Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByRef pImageInfo As IMAGEINFO) As Long
Public Declare Function ImageList_LoadImage Lib "comctl32.dll" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Merge Lib "comctl32.dll" (ByVal himl1 As Long, ByVal i1 As Long, ByVal hIml2 As Long, ByVal i2 As Long, ByVal dx As Long, ByVal dy As Long) As Long
Public Declare Function ImageList_Read Lib "comctl32.dll" (ByRef pstm As Long) As Long
Public Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long) As Long
Public Declare Function ImageList_Replace Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal hIml As Long, ByVal clrBk As Long) As Long
Public Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Function ImageList_SetIconSize Lib "comctl32.dll" (ByVal hIml As Long, ByVal cx As Long, ByVal cy As Long) As Long
Public Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal hIml As Long, ByVal uNewCount As Long) As Long
Public Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" (ByVal hIml As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Long
Public Declare Function ImageList_Write Lib "comctl32.dll" (ByVal hIml As Long, ByRef pstm As Long) As Long
Public Const ILC_COLOR As Long = &H0
Public Const ILC_COLOR16 As Long = &H10
Public Const ILC_COLOR24 As Long = &H18
Public Const ILC_COLOR32 As Long = &H20
Public Const ILC_COLOR4 As Long = &H4
Public Const ILC_COLOR8 As Long = &H8
Public Const ILC_COLORDDB As Long = &HFE
Public Const ILC_MASK As Long = &H1
Public Const ILC_PALETTE As Long = &H800
Public Const ILD_BLEND25 As Long = &H2
Public Const ILD_BLEND50 As Long = &H4
Public Const ILD_BLEND As Long = ILD_BLEND50
Public Const ILD_FOCUS As Long = ILD_BLEND25
Public Const ILD_IMAGE As Long = &H20
Public Const ILD_MASK As Long = &H10
Public Const ILD_NORMAL As Long = &H0
Public Const ILD_OVERLAYMASK As Long = &HF00
Public Const ILD_ROP As Long = &H40
Public Const ILD_SELECTED As Long = ILD_BLEND50
Public Const ILD_TRANSPARENT As Long = &H1

Public Const WM_NOTIFY As Long = &H4E
Public Const NM_FIRST As Long = 0
Public Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12)
Public Const NM_RCLICK As Long = (NM_FIRST - 5)
Public Const NM_DBLCLK As Long = (NM_FIRST - 3)


Public Const CDRF_DODEFAULT As Long = &H0
Public Const CDRF_NEWFONT As Long = &H2
Public Const CDRF_NOTIFYITEMDRAW As Long = &H20
Public Const CDRF_NOTIFYPOSTERASE As Long = &H40
Public Const CDRF_NOTIFYPOSTPAINT As Long = &H10
Public Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20
Public Const CDRF_SKIPDEFAULT As Long = &H4

Public Const CDDS_ITEM As Long = &H10000
Public Const CDDS_POSTERASE As Long = &H4
Public Const CDDS_POSTPAINT As Long = &H2
Public Const CDDS_PREPAINT As Long = &H1
Public Const CDDS_PREERASE As Long = &H3
Public Const CDDS_SUBITEM As Long = &H20000
Public Const CDDS_ITEMPOSTERASE As Long = (CDDS_ITEM Or CDDS_POSTERASE)
Public Const CDDS_ITEMPOSTPAINT As Long = (CDDS_ITEM Or CDDS_POSTPAINT)
Public Const CDDS_ITEMPREERASE As Long = (CDDS_ITEM Or CDDS_PREERASE)
Public Const CDDS_ITEMPREPAINT As Long = (CDDS_ITEM Or CDDS_PREPAINT)

Public Const CDIS_CHECKED As Long = &H8
Public Const CDIS_DEFAULT As Long = &H20
Public Const CDIS_DISABLED As Long = &H4
Public Const CDIS_FOCUS As Long = &H10
Public Const CDIS_GRAYED As Long = &H2
Public Const CDIS_HOT As Long = &H40
Public Const CDIS_INDETERMINATE As Long = &H100
Public Const CDIS_MARKED As Long = &H80
Public Const CDIS_SELECTED As Long = &H1
Public Const CDIS_SHOWKEYBOARDCUES As Long = &H200

Public Const MF_SEPARATOR As Long = &H800&
Public Const MF_STRING As Long = &H0&
Public Const MF_BYPOSITION  As Long = &H400&

Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByRef lprc As Any) As Long
Public Const TPM_NONOTIFY As Long = &H80&
Public Const TPM_RETURNCMD As Long = &H100&



'USER32
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Public Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GCL_WNDPROC As Long = -24
Public Const WM_PAINT As Long = &HF&
Public Const WM_ERASEBKGND As Long = &H14
Public Const WM_GETFONT As Long = &H31
Public Const WM_TIMER As Long = &H113
Public Const WM_CONTEXTMENU As Long = &H7B
Public Const WM_SYSCOLORCHANGE As Long = &H15
Public Const WM_HSCROLL As Long = &H114
Public Const WM_VSCROLL  As Long = &H115

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum SystemColours
COLOR_HIGHLIGHT = 13
COLOR_HIGHLIGHTTEXT = 14
COLOR_WINDOW = 5
COLOR_WINDOWTEXT = 8
End Enum

Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As SystemColours) As Long
Public Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As SystemColours) As Long



Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long



Public Enum DrawTextParams
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_EDITCONTROL = &H2000
    DT_END_ELLIPSIS = &H8000
    DT_EXPANDTABS = &H40
    DT_LEFT = &H0
    DT_MODIFYSTRING = &H10000
    DT_MULTILINE = (&H1)
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_PATH_ELLIPSIS = &H4000
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORD_ELLIPSIS = &H40000
    DT_WORDBREAK = &H10
End Enum

Public Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As DrawTextParams) As Long

Public Const OPAQUE As Long = 2
Public Const TRANSPARENT As Long = 1

Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Const PS_DASH As Long = 1
Public Const PS_DASHDOT As Long = 3
Public Const PS_DASHDOTDOT As Long = 4
Public Const PS_DOT As Long = 2
Public Const PS_SOLID As Long = 0
Public Const PS_ALTERNATE As Long = 8
Public Const PS_COSMETIC As Long = &H0
Public Const PS_GEOMETRIC As Long = &H10000

Public Const BS_SOLID As Long = 0

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long


'KERNEL32
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Const DEFAULT_GUI_FONT As Long = 17
Public Const NULL_BRUSH As Long = 5
Public Const NULL_PEN As Long = 8
Public Const WHITE_PEN As Long = 6
Public Const BLACK_PEN = 7


Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long

Public Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

Public Const SHGFI_ICON As Long = &H100
Public Const SHGFI_OVERLAYINDEX As Long = &H40
Public Const SHGFI_SHELLICONSIZE As Long = &H4
Public Const SHGFI_SMALLICON As Long = &H1
Public Const SHGFI_SYSICONINDEX As Long = &H4000

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const MAX_PATH As Long = 260

Public Type SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

' Private variables for enumeration functions and such -- don't touch them
Private TempFontList() As String
Private FontCount As Long
Private WantedWindowClass As String, WantedWindowCaption As String


' ----------- Helper Functions
Public Function GetNodeText(ByVal hTree As Long, ByVal hItem As Long) As String
    Dim Item As TVITEM
    Item.mask = TVIF_TEXT
    Item.hItem = hItem
    Item.pszText = String$(256, 0)
    Item.cchTextMax = 255
    If SendMessage(hTree, TVM_GETITEMA, 0, Item) Then
        GetNodeText = Left$(Item.pszText, InStr(Item.pszText, vbNullChar) - 1)
    End If
End Function

Public Function GetNodeIcon(ByVal hTree As Long, ByVal hItem As Long) As Long
    Dim Item As TVITEM
    Item.mask = TVIF_IMAGE
    Item.hItem = hItem
    If SendMessage(hTree, TVM_GETITEMA, 0, Item) Then
        GetNodeIcon = Item.iImage
    End If
End Function

Public Function GetNodeState(ByVal hTree As Long, ByVal hItem As Long) As Long
    Dim Item As TVITEM
    Item.mask = TVIF_STATE
    Item.stateMask = 255&
    Item.hItem = hItem
    If SendMessage(hTree, TVM_GETITEMA, 0, Item) Then
        GetNodeState = Item.state
    End If
End Function

Public Function GetNodeParam(ByVal hTree As Long, ByVal hItem As Long) As Long
    Dim Item As TVITEM
    Item.mask = TVIF_PARAM
    Item.hItem = hItem
    If SendMessage(hTree, TVM_GETITEMA, 0, Item) Then
        GetNodeParam = Item.lParam
    End If
End Function

Public Function GetNodeDepth(ByVal hTree As Long, ByVal hItem As Long) As Long
    GetNodeDepth = 0
    While hItem <> 0
        hItem = SendMessage(hTree, TVM_GETNEXTITEM, TVGN_PARENT, ByVal hItem)
        GetNodeDepth = GetNodeDepth + 1
    Wend
End Function

' Extract the folder part from a full path (without the trailing \)
Public Function FolderFromPath(ByVal sPath As String) As String
    Dim i As Integer
    i = InStrRev(sPath, "\")
    If i < 1 Then
        FolderFromPath = ""
    Else
        FolderFromPath = Left$(sPath, i - 1)
    End If
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

Public Function GetRelativePath(ByVal FromPath As String, ByVal ToPath As String) As String
    Dim i As Long
    GetRelativePath = Space$(260)
    If PathRelativePathTo(GetRelativePath, FromPath, 0, ToPath, 0) Then
        i = InStr(GetRelativePath, vbNullChar)
        If i > 1 Then
            GetRelativePath = Left$(GetRelativePath, i - 1)
            Exit Function
        End If
    End If
    
    ' Error, just return absolute path
    GetRelativePath = ToPath
End Function

' The GetTextExtent32() function doesn't work very well, especially with italic fonts,
' so this is an alternative (a bit computationally expensive though)
Public Sub MeasureTextExact(ByVal hDC As Long, ByVal sText As String, ByRef tSize As POINTAPI)
    Dim Points() As POINTAPI
    Dim Types() As Byte
    Dim Count As Long
    Dim rcBounds As RECT
    Dim i As Long
    Dim OldMode As Long
    
    OldMode = SetBkMode(hDC, OPAQUE)
    
    BeginPath hDC
    TextOut hDC, 0, 0, sText, Len(sText)
    EndPath hDC
    
    SetBkMode hDC, OldMode
    
    FlattenPath hDC
    
    ' Get the number of line segments drawn
    Count = GetPath(hDC, 0&, 0&, 0)
    ' Create the buffers
    ReDim Points(0 To Count - 1)
    ReDim Types(0 To Count - 1)
    ' Retrieve the line segments
    GetPath hDC, Points(0), Types(0), Count
    
    ' Delete the path object
    AbortPath hDC
    
    'ODS "Points in path: " & Count & vbCrLf
    
    For i = 0 To Count - 1
        With Points(i)
            If .x < rcBounds.Left Then rcBounds.Left = .x
            If .x > rcBounds.Right Then rcBounds.Right = .x
            If .y < rcBounds.Top Then rcBounds.Top = .y
            If .y > rcBounds.Bottom Then rcBounds.Bottom = .y
        End With
    Next
    
    tSize.x = 2 + rcBounds.Right - rcBounds.Left
    tSize.y = 2 + rcBounds.Bottom - rcBounds.Top
    
End Sub

Public Function ListFonts(ByVal hDC As Long) As String()
    FontCount = 0
    ReDim TempFontList(0 To 100) As String
    
    Dim lf As LOGFONT
    
    EnumFontFamiliesEx hDC, lf, AddressOf EnumFontFamExProc, 0, 0
    
    ReDim Preserve TempFontList(0 To FontCount - 1)
    ListFonts = TempFontList
End Function

Private Function EnumFontFamExProc(ByRef elfe As ENUMLOGFONTEX, ByRef nmtme As NEWTEXTMETRICEX, ByVal FontType As Long, ByVal lParam As Long) As Long
    Dim sFont As String
    
    sFont = StrConv(elfe.elfLogFont.lfFaceName, vbUnicode)
    If FontCount = UBound(TempFontList) Then
        ReDim Preserve TempFontList(0 To FontCount + 100)
    End If
    TempFontList(FontCount) = sFont
    FontCount = FontCount + 1
    
    EnumFontFamExProc = 1 ' Continue
End Function

Public Function FindInCombo(ByVal hWndCombo As Long, ByVal sString As String, Optional ByVal Exact As Boolean = False) As Long
    Dim arrANSI() As Byte
    arrANSI = StrConv(sString & vbNullChar, vbFromUnicode)
    
    If Exact Then
        FindInCombo = SendMessage(hWndCombo, CB_FINDSTRINGEXACT, -1, arrANSI(0))
    Else
        FindInCombo = SendMessage(hWndCombo, CB_FINDSTRING, -1, arrANSI(0))
    End If
End Function

' Returns the string before first null char encountered (if any) from an ANSI string.
Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Remove the ",index" from an icon path
Public Function RemoveIconIdx(ByVal sPath As String)
    Dim i As Integer
    i = InStrRev(sPath, ",")
    If i < 1 Then
        RemoveIconIdx = sPath
    Else
        RemoveIconIdx = Left$(sPath, i - 1)
    End If
End Function


' A quick note: SHChangeIconDialog is an undocumented windows function, and as such
' it only accepts ANSI on 9x/Me and UNICODE on NT/2k/XP, so this function handles both
Public Function BrowseForIcon(ByVal hWndOwner As Long, ByVal sDir As String) As String
    Dim sFileName As String, iIndex As Long
    sDir = RemoveIconIdx(sDir)
    sFileName = sDir & String$(260 - Len(sDir), 0)
    If IsWindowsNT Then sFileName = StrConv(sFileName, vbUnicode)
    If SHChangeIconDialog(hWndOwner, sFileName, 260, iIndex) = 0 Then
        BrowseForIcon = ""
    Else
        If IsWindowsNT Then sFileName = StrConv(sFileName, vbFromUnicode)
        BrowseForIcon = GetStrFromBufferA(sFileName) & "," & iIndex
    End If
End Function

Public Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Public Sub ODS(ByVal Format As String, ParamArray Values())
    Format = Replace$(Format, "\r", vbCr)
    Format = Replace$(Format, "\n", vbLf)
    Format = Replace$(Format, "\t", vbTab)
    
    Dim Tokens() As String
    Dim Output As String
    Dim i As Long, j As Long
    Tokens = Split(Format, "%")
    Output = Tokens(LBound(Tokens))
    For i = LBound(Tokens) + 1 To UBound(Tokens)
        Select Case CLng(Asc(Tokens(i)))
        Case 105& ' i
            Output = Output & CStr(CLng(Values(j))) & Mid$(Tokens(i), 2)
            j = j + 1
        Case 120& ' x
            Output = Output & "&h" & Hex$(CLng(Values(j))) & Mid$(Tokens(i), 2)
            j = j + 1
        Case 115& ' s
            Output = Output & CStr(Values(j)) & Mid$(Tokens(i), 2)
            j = j + 1
        End Select
    Next
    OutputDebugString Output
End Sub

' ** Update 08/08/06 : Added the following 3 functions (and associated API declares) for compatibility with the SDI mode
' 2 main things to notice:
'   * The function prototypes for the callbacks on both EnumThreadWindows() and EnumChildWindows() are identical.
'     Since we want them to do the same thing, we can use the same callback for both.
'   * In the callback function, the lParam is declared ByRef and when passing to it we use VarPtr()
'     This enables the value to be modified in the calling procedure.

Public Function FindThreadWindow(ByVal WindowClass As String, ByVal WindowCaption As String) As Long
    WantedWindowClass = WindowClass
    WantedWindowCaption = WindowCaption
    Dim hWndResult As Long
    EnumThreadWindows App.ThreadID, AddressOf EnumWindowCallback, VarPtr(hWndResult)
    FindThreadWindow = hWndResult
End Function

Private Function DrillForWindow(ByVal hWnd As Long) As Long
    ' Check if the passed window has a direct child which meets our criteria
    DrillForWindow = FindWindowEx(hWnd, 0, WantedWindowClass, WantedWindowCaption)
    If DrillForWindow <> 0 Then Exit Function
    
    ' It doesn't, so loop all child windows and call again until found
    Dim hWndChild As Long
    EnumChildWindows hWnd, AddressOf EnumWindowCallback, VarPtr(hWndChild)
End Function

Private Function EnumWindowCallback(ByVal hWnd As Long, ByRef lParam As Long) As Long
    Dim hWndTemp As Long
    hWndTemp = DrillForWindow(hWnd)
    If hWndTemp <> 0 Then
        lParam = hWndTemp       ' Found the window so return it
        EnumWindowCallback = 0  ' Stop enumeration
    Else
        EnumWindowCallback = 1  ' Continue enumeration
    End If
End Function

' Update 08/08/06 : Moved this function fom SettingsDlg.frm
' Return the angle with tangent opp/hyp. The returned
' value is between PI and -PI.
Public Function ATan2(ByVal opp As Single, ByVal adj As Single) As Single
Dim angle As Single

    ' Get the basic angle.
    If Abs(adj) < 0.0001 Then
        angle = Atn(1) * 2#
    Else
        angle = Abs(Atn(opp / adj))
    End If

    ' See if we are in quadrant 2 or 3.
    If adj < 0 Then
        ' angle > PI/2 or angle < -PI/2.
        angle = Atn(1) * 4# - angle
    End If

    ' See if we are in quadrant 3 or 4.
    If opp < 0 Then
        angle = -angle
    End If

    ' Return the result.
    ATan2 = angle
End Function


