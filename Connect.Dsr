VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8115
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14610
   _ExtentX        =   25770
   _ExtentY        =   14314
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "The Tree Surgeon"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VBInstance As VBIDE.VBE

Private Thingy As TreeSurgeon

Private MenuButton As CommandBarControl
Private WithEvents MenuWatcher As CommandBarEvents
Attribute MenuWatcher.VB_VarHelpID = -1

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo ERROR_HANDLER
    
    'save the vb instance
    Set VBInstance = Application
    
    Set Thingy = New TreeSurgeon
    
    Thingy.Init VBInstance
    
    AddMenuItem
      
    Exit Sub
    
ERROR_HANDLER:
    
    MsgBox "AddinInstance_OnConnection(): " & vbCrLf & Err.Description

End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    MenuButton.Delete
    Set MenuButton = Nothing
    Set MenuWatcher = Nothing
    
    Set Thingy = Nothing
    
End Sub

Private Sub AddMenuItem()
    On Error GoTo ERROR_HANDLER
    Dim cb As CommandBar
    
    Set cb = VBInstance.CommandBars("Add-Ins")
    If cb Is Nothing Then Exit Sub
    
    Set MenuButton = cb.Controls.Add(1)
    MenuButton.Caption = "TheTreeSurgeon Settings"
    
    If MenuWatcher Is Nothing Then
        Set MenuWatcher = VBInstance.Events.CommandBarEvents(MenuButton)
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    
End Sub

Private Sub MenuWatcher_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Thingy Is Nothing Then Exit Sub
    
    Thingy.DoEditSettings
End Sub
