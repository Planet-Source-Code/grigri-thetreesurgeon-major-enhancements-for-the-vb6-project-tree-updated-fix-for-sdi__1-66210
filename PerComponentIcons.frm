VERSION 5.00
Begin VB.Form PerComponentIcons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Tree Surgeon - Per-component icons"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PerComponentIcons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRemoveAll 
      Caption         =   "Remove All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox txtIconPath 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
   End
   Begin VB.ListBox lstComponents 
      Height          =   1860
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.ComboBox cboProjects 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "PerComponentIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ts As TreeSurgeon
Private SelPrj As VBProject
Private SelCmp As VBComponent

Public Function ManageComponentIcons(ts2 As TreeSurgeon)
    Set ts = ts2
    
    Load Me
    
    UpdateProjects
    
    Me.Show vbModal
    
    Unload Me
    
    ts.IconCache.Clear
    
    Set ts = Nothing
End Function

Private Sub btnBrowse_Click()
    Dim sTempIcon As String
    sTempIcon = BrowseForIcon(hwnd, txtIconPath.Text)
    If sTempIcon <> "" Then txtIconPath.Text = sTempIcon
    txtIconPath_Validate False
End Sub

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub btnRemoveAll_Click()
    If SelPrj Is Nothing Then Exit Sub
    
    If MsgBox("Are you sure you want to delete all per-component icons?", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
    
    Dim cmp As VBComponent
    On Error Resume Next
    For Each cmp In SelPrj.VBComponents
        ts.SetCustomIcon cmp, vbNullString
    Next
    UpdateComponents
End Sub

Private Sub cboProjects_Click()
    If cboProjects.ListIndex = -1 Then
        Set SelPrj = Nothing
        Set SelCmp = Nothing
        lstComponents.Clear
    End If
    
    Set SelPrj = ts.VBInstance.VBProjects(cboProjects.List(cboProjects.ListIndex))
    
    UpdateComponents
End Sub

Private Sub lstComponents_Click()
    If lstComponents.ListIndex = -1 Then
        btnBrowse.Enabled = False
        txtIconPath.Enabled = False
        Exit Sub
    End If
    
    Set SelCmp = SelPrj.VBComponents(lstComponents.List(lstComponents.ListIndex))
    
    btnBrowse.Enabled = True
    txtIconPath.Enabled = True
    txtIconPath.Text = ts.GetCustomIcon(SelCmp)
End Sub

Private Sub lstComponents_ItemCheck(Item As Integer)
    Dim cmp As VBComponent
    Set cmp = SelPrj.VBComponents(lstComponents.List(Item))
    If Len(ts.GetCustomIcon(cmp)) > 0 Then
        lstComponents.Selected(Item) = True
    Else
        lstComponents.Selected(Item) = False
    End If
End Sub

Private Sub txtIconPath_GotFocus()
    txtIconPath.SelStart = 1
    txtIconPath.SelLength = Len(txtIconPath.Text)
End Sub

Private Sub UpdateProjects()
    cboProjects.Clear
    lstComponents.Clear
    Dim prj As VBProject
    For Each prj In ts.VBInstance.VBProjects
        cboProjects.AddItem prj.Name
        If prj Is ts.VBInstance.ActiveVBProject Then
            cboProjects.ListIndex = cboProjects.NewIndex
        End If
    Next
    UpdateComponents
End Sub

Private Sub UpdateComponents()
    lstComponents.Clear
    If SelPrj Is Nothing Then Exit Sub
    Dim cmp As VBComponent
    For Each cmp In SelPrj.VBComponents
        If cmp.Type <> vbext_ct_ResFile And cmp.Type <> vbext_ct_RelatedDocument Then
            lstComponents.AddItem cmp.Name
            If Len(ts.GetCustomIcon(cmp)) > 0 Then
                lstComponents.Selected(lstComponents.NewIndex) = True
            Else
                lstComponents.Selected(lstComponents.NewIndex) = False
            End If
            If cmp Is ts.VBInstance.SelectedVBComponent Then
                lstComponents.ListIndex = lstComponents.NewIndex
            End If
        End If
    Next
End Sub

Private Sub txtIconPath_Validate(Cancel As Boolean)
    If SelCmp Is Nothing Then Exit Sub
    txtIconPath.Text = Trim$(txtIconPath.Text)
    ts.SetCustomIcon SelCmp, txtIconPath.Text
    If Len(txtIconPath.Text) > 0 Then
        lstComponents.Selected(lstComponents.ListIndex) = True
    Else
        lstComponents.Selected(lstComponents.ListIndex) = False
    End If
End Sub
