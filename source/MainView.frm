VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6630
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Private Const MaxSize As Double = 45000

Public IsOk As Boolean
Public IsCancel As Boolean

Public Columns As TextBoxHandler
Public Rows As TextBoxHandler
Public Width As TextBoxHandler
Public Height As TextBoxHandler
Public HorizontalSpace As TextBoxHandler
Public VerticalSpace As TextBoxHandler

'===============================================================================

Private Sub UserForm_Initialize()
  Caption = APP_NAME
  Logo.ControlTipText = APP_URL
  Set Columns = TextBoxHandler.Create(tbColumns, TextBoxTypeLong)
  Set Rows = TextBoxHandler.Create(tbRows, TextBoxTypeLong)
  Set Width = TextBoxHandler.Create(tbWidth, TextBoxTypeDouble, 0, MaxSize)
  Set Height = TextBoxHandler.Create(tbHeight, TextBoxTypeDouble, 0, MaxSize)
  Set HorizontalSpace = _
    TextBoxHandler.Create(tbHorizontalSpace, TextBoxTypeDouble, -MaxSize, MaxSize)
  Set VerticalSpace = _
    TextBoxHandler.Create(tbVerticalSpace, TextBoxTypeDouble, -MaxSize, MaxSize)
End Sub

Private Sub UserForm_Activate()
  VisibilityControl
End Sub

Private Sub obByCount_Click()
  VisibilityControl
End Sub

Private Sub opByArea_Click()
  VisibilityControl
End Sub

Private Sub opByPage_Click()
  VisibilityControl
End Sub

Private Sub btnOk_Click()
  FormŒ 
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub Logo_Click()
  With VBA.CreateObject("WScript.Shell")
    .Run APP_URL
  End With
End Sub

'===============================================================================

Private Sub VisibilityControl()
  If obByCount Then
    tbColumns.Enabled = True
    tbRows.Enabled = True
    tbWidth.Enabled = False
    tbHeight.Enabled = False
  ElseIf opByArea Then
    tbColumns.Enabled = False
    tbRows.Enabled = False
    tbWidth.Enabled = True
    tbHeight.Enabled = True
  ElseIf opByPage Then
    tbColumns.Enabled = False
    tbRows.Enabled = False
    tbWidth.Enabled = False
    tbHeight.Enabled = False
  End If
End Sub

Private Sub FormŒ ()
  Me.Hide
  IsOk = True
End Sub

Private Sub FormCancel()
  Me.Hide
  IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    —ancel = True
    FormCancel
  End If
End Sub
