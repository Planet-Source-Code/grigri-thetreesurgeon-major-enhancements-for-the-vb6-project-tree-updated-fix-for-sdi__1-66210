VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CachedImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private hIml As Long
Private col As Collection


Private Sub Class_Initialize()
    Set col = New Collection
    hIml = ImageList_Create(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CXSMICON), ILC_COLOR32 Or ILC_MASK, 10, 10)
    ImageList_SetBkColor hIml, vbWhite
End Sub

Private Sub Class_Terminate()
    ImageList_Destroy hIml
End Sub

Public Property Get ImageListHandle() As Long
    ImageListHandle = hIml
End Property

Public Function AddIcon(ByVal IconPath As String) As Long
    Dim Idx As Long
    Idx = GetIconIndex(IconPath)
    If Idx <> -1 Then
        AddIcon = Idx
        Exit Function
    End If
    
    Dim sIconPath As String, iIndex As Long, iIconIndex As Long, hIcon As Long
    Dim sBuf As String
    
    sIconPath = IconPath
    iIconIndex = PathParseIconLocation(sIconPath)
    sIconPath = Left$(sIconPath, lstrlen(sIconPath))
    sBuf = String$(260, 0)
    ExpandEnvironmentStrings sIconPath, sBuf, 260
    sIconPath = Left$(sBuf, lstrlen(sBuf))
    
    If ExtractIconEx(sIconPath, iIconIndex, ByVal 0, hIcon, 1) = 0 Then
        ' No icons here...
        AddIcon = -1
        Exit Function
    End If
    
    AddIcon = ImageList_ReplaceIcon(hIml, -1, hIcon)
    DestroyIcon hIcon
    
    col.Add AddIcon, IconPath
End Function

Public Function AddShellIcon(ByVal FilePath As String) As Long
    If Len(FilePath) = 0 Then AddShellIcon = -1: Exit Function
    
    Dim Idx As Long
    Idx = GetIconIndex(FilePath)
    If Idx <> -1 Then
        AddShellIcon = Idx
        Exit Function
    End If

    Dim sfi As SHFILEINFO
    If SHGetFileInfo(FilePath, 0, sfi, LenB(sfi), SHGFI_ICON Or SHGFI_SMALLICON) <> 0 Then
        If sfi.hIcon <> 0 Then
            AddShellIcon = ImageList_ReplaceIcon(hIml, -1, sfi.hIcon)
            DestroyIcon sfi.hIcon
            col.Add AddShellIcon, FilePath
        Else
            AddShellIcon = -1
        End If
    Else
        AddShellIcon = -1
    End If
End Function

Public Function GetIconIndex(ByVal IconPath As String) As Long
    On Error Resume Next
    Err.Clear
    GetIconIndex = col(IconPath)
    If Err Then
        GetIconIndex = -1
    End If
    On Error GoTo 0
End Function

Public Function RemoveIcon(ByVal IconPath As String) As Boolean
    Dim Idx As Long
    Idx = GetIconIndex(IconPath)
    If Idx > -1 Then
        ImageList_Remove hIml, Idx
        col.Remove IconPath
    End If
End Function

Public Sub Clear()
    ImageList_Remove hIml, -1
    Set col = New Collection
End Sub
