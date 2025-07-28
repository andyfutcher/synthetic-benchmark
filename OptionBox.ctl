VERSION 5.00
Begin VB.UserControl OptionBox 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
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
   MaskColor       =   &H00000000&
   ScaleHeight     =   240
   ScaleWidth      =   1815
   Windowless      =   -1  'True
   Begin VB.Label OpBHotSpot 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   285
   End
   Begin VB.Label OpBIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape OpBHighFocus 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label OpBCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "•"
      Height          =   195
      Left            =   285
      TabIndex        =   1
      Top             =   15
      Width           =   90
   End
   Begin VB.Shape OpBEffect 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Index           =   2
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape OpBEffect 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   495
      Index           =   1
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape OpBEffect 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Index           =   0
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape OpBEffect 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00F0F0F0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "OptionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, OptionVal As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()

Public Sub ResetControl()
OpBEffect(0).BorderColor = Theme_Pitch
OpBEffect(1).BorderColor = Theme_Shadow
OpBEffect(2).BorderColor = Theme_Dark
'OpBEffect(3).BorderColor = Theme_Shade
OpBEffect(3).FillColor = Theme_Light
OpBIcon.ForeColor = Theme_Icon
OpBIcon.Font = "Wingdings 2"
OpBHighFocus.BorderColor = Theme_HighLight
OpBCaption.ForeColor = Theme_Pitch
OpBCaption.Font = Theme_Font
End Sub

Public Sub Auto_Ret()
UserControl.Width = OpBHighFocus.Left + OpBHighFocus.Width
End Sub

Private Sub Update_Control()
If OptionVal = True Then
    OpBIcon.Visible = True
Else
    OpBIcon.Visible = False
End If
End Sub

Private Sub Switch_Values()
If OptionVal = True Then
    OptionVal = False
    GoTo Ed
End If
OptionVal = True
Ed: Call Update_Control
End Sub

Private Sub OpBHotSpot_Click()
Call Switch_Values
RaiseEvent Click
End Sub

Private Sub OpBHotSpot_DblClick()
Call Switch_Values
RaiseEvent Click
End Sub

Private Sub OpBHotSpot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OpBEffect(3).FillColor = Theme_Shade
End Sub

Private Sub OpBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub OpBHotSpot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
OpBEffect(3).FillColor = Theme_Light
End Sub

Private Sub UserControl_GotFocus()
OpBHighFocus.Visible = True
End Sub

Private Sub UserControl_Initialize()
OptionVal = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
OpBEffect(3).FillColor = Theme_Shade
'OpBEffect(3).BorderColor = Theme_Dark
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
OpBEffect(3).FillColor = Theme_Light
'OpBEffect(3).BorderColor = Theme_Shade
If KeyCode = 13 Or KeyCode = 32 Then Call Switch_Values
RaiseEvent Click
End Sub

Private Sub UserControl_LostFocus()
OpBHighFocus.Visible = False
End Sub

Private Sub UserControl_Resize()
Dim ResizeValue As Integer
UserControl.Height = 240
ResizeValue = (OnePix * 14)
'OpBCaption.Move (UserControl.Width / 2) - (OpBCaption.Width / 2), (UserControl.Height / 2) - (OpBCaption.Height / 2) - OnePix
OpBEffect(0).Move 0, OnePix, ResizeValue, ResizeValue
OpBEffect(1).Move OnePix, TwoPix, ResizeValue - TwoPix, ResizeValue - TwoPix
OpBEffect(2).Move TwoPix, TwoPix, ResizeValue - FourPix, ResizeValue - TwoPix
OpBEffect(3).Move 0, OnePix, ResizeValue, ResizeValue
OpBHighFocus.Width = OpBCaption.Width + SixPix
End Sub

Public Property Let Caption(ByVal CaptionValue As String)
OpBCaption.Caption = CaptionValue
Call UserControl_Resize
End Property
Public Property Get Caption() As String
Caption = OpBCaption.Caption
End Property

Public Property Let ValueDesc(ByVal OptionValue As Boolean)
OptionVal = OptionValue
Call Update_Control
End Property
Public Property Let Value(ByVal OptionValue As Boolean)
OptionVal = OptionValue
Call Update_Control
RaiseEvent Click
End Property
Public Property Get Value() As Boolean
Value = OptionVal
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
UserControl.Enabled = EnabledValue
OpBEffect(1).Visible = EnabledValue
OpBEffect(2).Visible = EnabledValue
If EnabledValue = True Then
    OpBEffect(0).BorderColor = Theme_Pitch
    OpBCaption.ForeColor = Theme_Pitch
Else
    OpBEffect(0).BorderColor = Theme_Shadow
    OpBCaption.ForeColor = Theme_Shadow
End If
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property


Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = True Then
    OpBEffect(1).BorderColor = Theme_Invert
    OpBEffect(2).BorderColor = Theme_InvertLight
Else
    OpBEffect(1).BorderColor = Theme_Shadow
    OpBEffect(2).BorderColor = Theme_Dark
End If
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

