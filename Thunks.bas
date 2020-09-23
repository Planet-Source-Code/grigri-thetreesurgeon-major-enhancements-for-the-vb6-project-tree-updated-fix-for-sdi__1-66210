Attribute VB_Name = "Thunks"
Option Explicit

Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long

Private Const GWL_WNDPROC As Long = -4

Public Type ThunkData
    pfn As Long
    Code(0 To 5) As Long
End Type

Public Type SubClassData
    pfnWndProcNext As Long
    Thunk As ThunkData
End Type

Public Sub InitThunk(Thunk As ThunkData, ByVal ParamValue As Long, ByVal pfnDest As Long)
    With Thunk
        .Code(0) = &HB82434FF
        .Code(1) = ParamValue
        .Code(2) = &H4244489
        .Code(3) = &HB8909090
        .Code(4) = pfnDest
        .Code(5) = &H9090E0FF
        .pfn = VarPtr(.Code(0))
    End With
End Sub

Public Sub SubClass(Data As SubClassData, ByVal hWnd As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long)
    With Data
        If .pfnWndProcNext Then
            Call DoSetWindowLong(hWnd, GWL_WNDPROC, .pfnWndProcNext)
            .pfnWndProcNext = 0
        End If
        InitThunk .Thunk, ThisPtr, pfnRedirect
        .pfnWndProcNext = DoSetWindowLong(hWnd, GWL_WNDPROC, .Thunk.pfn)
    End With
End Sub
Public Sub UnSubClass(Data As SubClassData, ByVal hWnd As Long)
    With Data
        If .pfnWndProcNext Then
            Call DoSetWindowLong(hWnd, GWL_WNDPROC, .pfnWndProcNext)
            .pfnWndProcNext = 0
        End If
    End With
End Sub

Private Function DoSetWindowLong(ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwValue As Long) As Long
    If IsWindowUnicode(hWnd) Then
        DoSetWindowLong = SetWindowLongW(hWnd, nIndex, dwValue)
    Else
        DoSetWindowLong = SetWindowLongA(hWnd, nIndex, dwValue)
    End If
End Function

Private Function DoGetWindowLong(ByVal hWnd As Long, ByVal nIndex As Long) As Long
    If IsWindowUnicode(hWnd) Then
        DoGetWindowLong = GetWindowLongW(hWnd, nIndex)
    Else
        DoGetWindowLong = GetWindowLongA(hWnd, nIndex)
    End If
End Function

