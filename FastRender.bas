Attribute VB_Name = "FastRender"
Option Explicit

Private Const Blur3x3HexString As String = "558BEC53565733C983C104034D1C8B751883EE028B7D1483EF0251BB04000000034D0C33C033D28A118A410403D08A41FC03D02B4D1C8A41FC03D08A0103D08A410403D0034D1C034D1C8A41FC03D08A0103D08A410403D02B4D1CB8398EE338F7E2D1EA2B4D0C034D1088112B4D10414B75AD4F75A559034D1C4E75975F5E33C05B5DC218"
'Private Const Blur3x3HexString As String = "558BEC53565733C983C104034D1C8B751883EE028B7D1483EF0251BB04000000034D0C33C033D28A11D1E28A410403D08A41FC03D02B4D1C8A41FC03D08A0103D08A410403D0034D1C034D1C8A41FC03D08A0103D08A410403D02B4D1CB8CDCCCCCCF7E2C1EA032B4D0C034D1088112B4D10414B75AA4F75A259034D1C4E75945F5E33C05B5DC21800"


Public Type FastRenderData
    pVTable As Long
    VTable(0 To 3) As Long
End Type

Private Blur3x3Bytes() As Byte

Public Function InitFastRender(Data As FastRenderData) As IFastRender
    With Data
        .VTable(0) = GetFunctionAddress(AddressOf FastRender_QueryInterface)
        .VTable(1) = GetFunctionAddress(AddressOf FastRender_AddRefRelease)
        .VTable(2) = GetFunctionAddress(AddressOf FastRender_AddRefRelease)
        Blur3x3Bytes = HexToBytes(Blur3x3HexString)
        .VTable(3) = VarPtr(Blur3x3Bytes(0))
        
        .pVTable = VarPtr(.VTable(0))
        
    End With
    CopyMemory InitFastRender, VarPtr(Data), 4
End Function

Private Function FastRender_QueryInterface(This As FastRenderData, riid As Long, pvObj As Long) As Long
    ' Not needed
End Function

Private Function FastRender_AddRefRelease(This As FastRenderData) As Long
    ' Not needed
End Function

' Boring function to convert a hex string to a byte array.
' Not optimized, it's not called often enough to warrant it.
Private Function HexToBytes(sHex As String) As Byte()
    On Error GoTo ERR_HANDLER
    Dim Bytes() As Byte
    ReDim Bytes(0 To Len(sHex) \ 2 - 1)
    Dim i As Long, j As Long
    j = 0
    For i = 1 To Len(sHex) - 1 Step 2
        Bytes(j) = CByte("&h" & Mid$(sHex, i, 2))
        j = j + 1
    Next
    HexToBytes = Bytes
ERR_HANDLER:
    If Err Then
        ODS "> Error " & Err.Description & vbCrLf
    End If
End Function
