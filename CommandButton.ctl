VERSION 5.00
Begin VB.UserControl CommandButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   1005
   ScaleWidth      =   1875
   Begin VB.Label CmbHotSpot 
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H002DAEA8&
      Height          =   375
      Index           =   6
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H002DAEA8&
      Height          =   375
      Index           =   5
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00808080&
      Height          =   735
      Index           =   1
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00800000&
      Height          =   855
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label CmBLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   60
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00E0E0E0&
      Height          =   495
      Index           =   4
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Index           =   2
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape CmBShape 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "CommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, EnabledControl As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Event WindowLess(TabStopValue As Boolean)
Public Event Click()

Private Sub CmbHotSpot_Click()
If EnabledControl = True Then
    RaiseEvent Click
End If
End Sub

Public Sub ClickNow()
RaiseEvent Click
End Sub

Private Sub CmbHotSpot_DblClick()
Call PressDown
If EnabledControl = True Then
    RaiseEvent Click
End If
End Sub

Private Sub CmbHotSpot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PressDown
End Sub

Private Sub CmbHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub CmbHotSpot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PressUp
End Sub

Private Sub UserControl_GotFocus()
CmBShape(4).BorderStyle = 3
CmBShape(4).BorderColor = Theme_High
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
RaiseEvent MouseMove(0, 0, X, Y)
End Sub

Private Sub UserControl_Initialize()
EnabledControl = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 32 Then
    Call PressDown
End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 32 Then
    Call PressUp
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_LostFocus()
CmBShape(4).BorderStyle = 1
CmBShape(4).BorderColor = Theme_Color
Call PressUp
End Sub

Private Sub UserControl_Resize()
CmBLabel.Move (UserControl.Width / 2) - (CmBLabel.Width / 2) - OnePix, (UserControl.Height / 2) - (CmBLabel.Height / 2) - OnePix
CmBShape(0).Move 0, 0, UserControl.Width - OnePix, UserControl.Height - OnePix
CmBShape(1).Move 0, 0, UserControl.Width - TwoPix, UserControl.Height - TwoPix
CmBShape(2).Move OnePix, OnePix, UserControl.Width - FourPix, UserControl.Height - FourPix
CmBShape(3).Move 0, 0, UserControl.Width, UserControl.Height
CmBShape(4).Move TwoPix, TwoPix, UserControl.Width - SixPix + OnePix, UserControl.Height - SixPix + OnePix
CmBShape(5).Move TwoPix, TwoPix, UserControl.Width - FourPix, UserControl.Height - FourPix
CmBShape(6).Move OnePix, OnePix, UserControl.Width - FourPix, UserControl.Height - FourPix
CmbHotSpot.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Sub PressUp()
CmBLabel.Top = (UserControl.Height / 2) - (CmBLabel.Height / 2) - OnePix
CmBShape(3).FillColor = Theme_Light
CmBShape(5).BorderColor = Theme_Invert
CmBShape(6).BorderColor = Theme_InvertLight
End Sub

Private Sub PressDown()
CmBLabel.Top = (UserControl.Height / 2) - (CmBLabel.Height / 2) - 1
CmBShape(3).FillColor = Theme_Shade
CmBShape(5).BorderColor = Theme_HighLight
CmBShape(6).BorderColor = Theme_High
End Sub

Public Sub ResetControl()
CmBShape(0).BorderColor = Theme_Pitch
CmBShape(1).BorderColor = Theme_Shadow
CmBShape(2).BorderColor = Theme_Dark
CmBShape(3).FillColor = Theme_Light
CmBShape(4).BorderColor = Theme_Color
CmBShape(5).BorderColor = Theme_Invert
CmBShape(6).BorderColor = Theme_InvertLight
CmBLabel.ForeColor = Theme_Pitch
CmBLabel.Font = Theme_Text
End Sub

'Private Sub UserControl_Show()
'CmBShape(0).BorderColor = Theme_Pitch
'CmBShape(1).BorderColor = Theme_Shadow
'CmBShape(2).BorderColor = Theme_Dark
'CmBShape(3).FillColor = Theme_Light
'CmBShape(4).BorderColor = Theme_Color
'CmBShape(5).BorderColor = Theme_Invert
'CmBShape(6).BorderColor = Theme_Invert
'CmBLabel.ForeColor = Theme_Pitch
'End Sub

Public Property Let Caption(ByVal CaptionValue As String)
CmBLabel.Caption = JumpTo_Char & CaptionValue
Call UserControl_Resize
End Property
Public Property Get Caption() As String
Caption = CmBLabel.Caption
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
'EnabledControl = EnabledValue
UserControl.Enabled = EnabledValue
'RaiseEvent WindowLess(EnabledValue)
CmBShape(1).Visible = EnabledValue
CmBShape(2).Visible = EnabledValue
CmBShape(4).Visible = EnabledValue
If EnabledValue = False Then
    CmBShape(0).BorderColor = Theme_Shadow
    'CmBShape(3).FillColor = Theme_Shade
    CmBLabel.ForeColor = Theme_Shadow
Else
    CmBShape(0).BorderColor = Theme_Pitch
    CmBLabel.ForeColor = Theme_Pitch
End If
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = True And UserControl.Enabled = True Then
    CmBShape(5).Visible = True
    CmBShape(6).Visible = True
Else
    CmBShape(5).Visible = False
    CmBShape(6).Visible = False
End If
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

