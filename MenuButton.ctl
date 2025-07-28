VERSION 5.00
Begin VB.UserControl MenuButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipBehavior    =   0  'None
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   HitBehavior     =   2  'Use Paint
   MaskColor       =   &H00FFFFFF&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   255
   ScaleWidth      =   510
   Windowless      =   -1  'True
   Begin VB.Label MuBHotSpot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.Label MuBLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   45
   End
   Begin VB.Shape MuBEffect 
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "MenuButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, DownKey As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event Click()
Public Event SelfAlign()

Private Sub MuBHotSpot_Click()
RaiseEvent Click
End Sub

Private Sub MuBHotSpot_DblClick()
RaiseEvent Click
End Sub

Private Sub MuBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
RaiseEvent MouseMove(0, 0, X, Y)
End Sub

Private Sub UserControl_GotFocus()
MuBEffect.BorderStyle = 1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyPress(KeyCode)
DownKey = KeyCode
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13, 32
    RaiseEvent MouseMove(0, 0, 0, 0)
    RaiseEvent Click
End Select
End Sub

Private Sub UserControl_LostFocus()
MuBEffect.BorderStyle = 0
End Sub

Private Sub UserControl_Resize()
UserControl.Width = MuBLabel.Width + 240
MuBEffect.Move 0, 0, UserControl.Width, UserControl.Height
MuBHotSpot.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Public Property Let Caption(ByVal CaptionValue As String)
MuBLabel.Caption = JumpTo_Char & CaptionValue
Call UserControl_Resize
RaiseEvent SelfAlign
End Property

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = True Then
    MuBEffect.FillStyle = 0
    MuBLabel.ForeColor = Theme_Light
Else
    MuBEffect.FillStyle = 1
    MuBLabel.ForeColor = Theme_Pitch
End If
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

Public Sub ResetControl()
MuBEffect.FillColor = Theme_High
MuBEffect.BorderColor = Theme_High
MuBLabel.Font = Theme_Text
MuBEffect.BorderStyle = 0
End Sub

