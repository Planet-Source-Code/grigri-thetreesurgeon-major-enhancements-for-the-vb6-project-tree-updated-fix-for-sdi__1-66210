VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private hDC As Long
Private hBmp As Long
Private hBmpOld As Long
Private W As Long, H As Long

Public Sub Create(ByVal NewW As Long, ByVal NewH As Long, ByVal hCompatibleDC As Long)
    Destroy
    
    W = NewW
    H = NewH
    
    hBmp = CreateCompatibleBitmap(hCompatibleDC, W, H)
    hDC = CreateCompatibleDC(hCompatibleDC)
    hBmpOld = SelectObject(hDC, hBmp)
    
    'ODS "Bitmap::Create(%i,%i) => hDC = %x, hBmp = %x, hBmpOld = %x\n", W, H, hDC, hBmp, hBmpOld
End Sub

Private Sub Destroy()
    'ODS "Bitmap::Destroy()"
    If hDC Then
        SelectObject hDC, hBmpOld
        DeleteObject hBmp
        DeleteDC hDC
    End If
    hDC = 0: hBmp = 0: hBmpOld = 0
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Blit(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long)
    BitBlt hDestDC, x, y, W, H, hDC, 0, 0, vbSrcCopy
End Sub

Public Property Get handle() As Long
    handle = hDC
End Property
