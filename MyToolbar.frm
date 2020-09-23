VERSION 5.00
Begin VB.Form MyToolbar 
   BorderStyle     =   0  'None
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   ControlBox      =   0   'False
   FillColor       =   &H80000010&
   Icon            =   "MyToolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   19
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDown 
      Height          =   270
      Left            =   480
      MaskColor       =   &H00FF00FF&
      Picture         =   "MyToolbar.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Move Component Down"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton btnUp 
      Height          =   270
      Left            =   0
      MaskColor       =   &H00FF00FF&
      Picture         =   "MyToolbar.frx":0256
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Move Component Up"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   375
   End
End
Attribute VB_Name = "MyToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ToolbarItems
    [Move Up] = 1
    [Move Down]
End Enum

Event ItemClicked(ByVal Index As ToolbarItems)

Private Sub btnDown_Click()
    RaiseEvent ItemClicked([Move Down])
End Sub

Private Sub btnUp_Click()
    RaiseEvent ItemClicked([Move Up])
End Sub

Private Sub Form_Paint()
    ForeColor = &H80000014
    Line (0, 0)-Step(ScaleWidth, 0)
    ForeColor = &H80000010
    'ForeColor = &H80000015
    Line (0, ScaleHeight - 1)-Step(ScaleWidth, 0)
End Sub

Private Sub Form_Resize()
    Dim dx As Long, y As Long
    y = (CLng(ScaleHeight) - CLng(btnUp.Height)) \ 2
    dx = (CLng(ScaleWidth) - CLng(btnUp.Width) * 2) \ 3
    btnUp.Move dx, y
    btnDown.Move btnUp.Left + btnUp.Width + dx, y
End Sub

Public Sub SetStatus(ByVal ButtonsEnabled As Boolean)
    btnUp.Enabled = ButtonsEnabled
    btnDown.Enabled = ButtonsEnabled
End Sub
