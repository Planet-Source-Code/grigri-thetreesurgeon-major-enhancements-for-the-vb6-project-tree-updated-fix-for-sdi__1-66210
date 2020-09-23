Attribute VB_Name = "Redirects"
Option Explicit

Public Function GetFunctionAddress(ByVal pFn As Long) As Long
    GetFunctionAddress = pFn
End Function

Public Function RedirectProjectWndProc(ByVal This As TreeSurgeon, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    RedirectProjectWndProc = This.ProjectWndProc(hwnd, uMsg, wParam, lParam)
End Function

Public Function RedirectTreeWndProc(ByVal This As TreeSurgeon, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    RedirectTreeWndProc = This.TreeWndProc(hwnd, uMsg, wParam, lParam)
End Function

Public Function RedirectCompareProc(ByVal This As TreeSurgeon, ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal pcolItems As Long) As Long
    Dim col As Collection
    ' Dereference pointer
    CopyMemory col, ByVal pcolItems, 4
    ' Forward function call with collection object
    RedirectCompareProc = This.CompareTreeItems(lParam1, lParam2, col)
    ' Zero the object pointer (note how the 0& parameter is ByRef - ByVal will of course crash it)
    CopyMemory col, 0&, 4
End Function

