Attribute VB_Name = "ShellContextMenu"
Option Explicit

' ***********************************************************
' *  ShellContextMenu.bas                                   *
' *---------------------------------------------------------*
' *  By grigri, 2005                                        *
' *  Uses : API declares in "ApiDeclares.bas"               *
' *         Helper Typelib  "TreedomHelper.tlb"             *
' *  v1.0                                                   *
' *---------------------------------------------------------*
' *  Displays a fully-functional shell context menu for     *
' *  any folder/file combination.                           *
' ***********************************************************

' The range of menu item identifiers the shell can use
Private Const MIN_SHELL_ID As Long = 1&
Private Const MAX_SHELL_ID As Long = 30000&

' Our menu item identifiers
Private Enum MenuItemIDs
    MENU_OPEN_FOLDER = MAX_SHELL_ID + 1 ' Open Containing Folder
End Enum

' Pointer to the context menu handler
Private pCM2 As IContextMenu2
' Previous window procedure (subclass needed to handle menu events)
Private pOldWndProc As Long


Public Sub ShowShellMenu(ByVal hwnd As Long, ByVal sPath As String, ByVal sFile As String)
    Dim pDesk As IShellFolder   ' Pointer to Desktop IShellFolder (root)
    Dim pSF As IShellFolder     ' Pointer to IShellFolder of the sPath
    Dim mem As IMalloc          ' The Shell's memory allocator (for disposal of the pidl)
    Dim hMenu As Long           ' Menu handle
    Dim tmp As Long             ' A dummy value, passed by reference to several functions
    Dim pt As POINTAPI          ' Poition of the mouse cursor in screen coords
    Dim pidl As Long            ' pidl (used twice)
    Dim pCM As IContextMenu     ' Pointer to the base context menu handler
    Dim idCmd As Long           ' Returned command identifier offset from the menu
    Dim indexMenu As Long       ' Menu Items Index
    
    ' Some basic checks
    If IsWindow(hwnd) = 0 Then
        ' We MUST have a valid window handle
        Exit Sub
    End If
    If GetWindowThreadProcessId(hwnd, ByVal 0&) <> App.ThreadID Then
        ' And it MUST be in the same thread as the dll (else we can't subclass it)
        Exit Sub
    End If
    
    ' Retrieve the memory allocator
    Set mem = SHGetMalloc
    ' Get the shell root pointer
    Set pDesk = SHGetDesktopFolder
    
    ' Declare the IIDs for IContextMenu and IShellFolder
    Dim iidCM As TreedomHelper.Guid
    Dim iidSF As TreedomHelper.Guid
    With iidCM
        .Data1 = &H214E4
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With iidSF
        .Data1 = &H214E6
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    ' Make the strings null-terminated so we can pass them
    ' Note that the functions we use these in (ISF::ParseDisplayName and ISF::BindToObject)
    ' take UNICODE strings so we use StrPtr to pass them as Long
    sPath = sPath & vbNullChar
    sFile = sFile & vbNullChar
    
    ' Given the folder path, get its pidl relative the the desktop
    pDesk.ParseDisplayName 0, ByVal 0&, StrPtr(sPath), tmp, pidl, 0
    ' And get the folder interface
    Set pSF = pDesk.BindToObject(pidl, tmp, VarPtr(iidSF))
    ' Free the pidl using the shell's memory allocator, it's not needed any more
    mem.Free pidl
    If pSF Is Nothing Then
        ' This should never happen, but you never know
        Exit Sub
    End If
    
    ' Get the pidl for the file name, relative to the folder path
    tmp = 0
    pSF.ParseDisplayName 0, ByVal 0&, StrPtr(sFile), tmp, pidl, 0
    ' Retrieve the IContextMenu interface pointer for the file
    tmp = 0
    Set pCM = pSF.GetUIObjectOf(0, 1, pidl, VarPtr(iidCM), tmp)
    ' Free the pidl using the shell's memory allocator, it's not needed any more
    mem.Free pidl
    If pCM Is Nothing Then
        ' This should never happen, but you never know
        Exit Sub
    End If
    
    ' Try and get the extended context menu interface pointer (IContextMenu2)
    On Error Resume Next
    Set pCM2 = pCM
    On Error GoTo 0
    
    ' Create a blank menu
    hMenu = CreatePopupMenu
    
    indexMenu = 0   ' Number of items - offset for the shell
    
    If pCM2 Is Nothing Then
        ' There is no extended interface, so just use the old one
        pCM.QueryContextMenu hMenu, indexMenu, MIN_SHELL_ID, MAX_SHELL_ID, CMF_EXPLORE
    Else
        ' We have an extended interface, so use it to fill the menu
        pCM2.QueryContextMenu hMenu, indexMenu, MIN_SHELL_ID, MAX_SHELL_ID, CMF_EXPLORE
        ' And temporarily subclass the owner window so we can handle the menu messages
        pOldWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf TempWindowProc)
    End If
    
    ' Add our custom items to the top of the menu
    InsertMenu hMenu, 0, MF_BYPOSITION Or MF_SEPARATOR, 0, 0&
    InsertMenu hMenu, 0, MF_BYPOSITION Or MF_STRING, MENU_OPEN_FOLDER, "Open Containing Folder"
    
    ' Get the screen cursor position
    GetCursorPos pt
    ' Display the menu and return the result (0 is returned if nothing was selected)
    idCmd = TrackPopupMenu(hMenu, TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, ByVal 0&)
    
    ' Unsubclass if needed
    If pOldWndProc <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, pOldWndProc
        pOldWndProc = 0
    End If
    
    If idCmd <> 0 Then
        If idCmd >= MIN_SHELL_ID And idCmd <= MAX_SHELL_ID Then
            Dim ici As CMINVOKECOMMANDINFO
            With ici
                .cbSize = LenB(ici)
                .lpVerb = (idCmd - MIN_SHELL_ID) And &HFFFF
                .nShow = SW_SHOWNORMAL
            End With
            
            If pCM2 Is Nothing Then
                pCM.InvokeCommand ici
            Else
                pCM2.InvokeCommand ici
            End If
        Else
            Select Case idCmd
            Case MENU_OPEN_FOLDER
                ' Open an explorer window to the folder
                ShellExecute hwnd, "explore", sPath, vbNullString, vbNullString, SW_SHOWNORMAL
            End Select
        End If
    End If
    
    ' Free the IContextMenu2 pointer (avoid accidental use)
    Set pCM2 = Nothing
End Sub

Private Function TempWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo Default_Processing
    
    If pCM2 Is Nothing Then GoTo Default_Processing
    
    Select Case uMsg
    Case WM_MEASUREITEM, WM_DRAWITEM
        If wParam <> 0 Then GoTo Default_Processing
        pCM2.HandleMenuMsg uMsg, wParam, lParam
        TempWindowProc = -1
        Exit Function
    Case WM_INITMENUPOPUP
        pCM2.HandleMenuMsg uMsg, wParam, lParam
        TempWindowProc = 0
        Exit Function
    End Select
    
Default_Processing:
    TempWindowProc = CallWindowProc(pOldWndProc, hwnd, uMsg, wParam, lParam)
End Function

