VERSION 5.00
Begin VB.UserControl SimpleBox 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5295
   ScaleWidth      =   6735
End
Attribute VB_Name = "SimpleBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ShowThisBool As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MakeVisible(VisibleControl As Boolean)
Public Event Click()

Private Sub CxBCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Frame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let Enabled(ByVal EnabledValue As Boolean)
UserControl.Enabled = EnabledValue
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
'CxBImage.Width = UserControl.Width - (CxBImage.Left * 2)
'VScrollButton.Left = UserControl.Width - VScrollButton.Width
'VScrollButton.Height = UserControl.Height
End Sub

Public Property Let ShowThis(ByVal ShowThisValue As Boolean)
ShowThisBool = ShowThisValue
RaiseEvent MakeVisible(ShowThisValue)
End Property
Public Property Get ShowThis() As Boolean
ShowThis = ShowThisBool
End Property

Public Sub ResetControl()
UserControl.BackColor = Theme_Light
End Sub

