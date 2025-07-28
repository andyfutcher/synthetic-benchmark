VERSION 5.00
Begin VB.UserControl ProgressBox 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   300
   ScaleWidth      =   4455
   Begin VB.Label PrBCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   30
      Width           =   45
   End
   Begin VB.Shape PrBShape 
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00987054&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   2295
   End
   Begin VB.Shape PrBShape 
      BorderColor     =   &H00C0C0C0&
      Height          =   270
      Index           =   1
      Left            =   15
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   4425
   End
   Begin VB.Shape PrBShape 
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   4410
   End
   Begin VB.Shape PrBShape 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   305
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "ProgressBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim MaxVal As Long, ValVal As Long
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Sub ResetControl()
PrBShape(0).BorderColor = Theme_Shadow
PrBShape(0).FillColor = Theme_Light
PrBShape(1).BorderColor = Theme_Dark
PrBShape(2).BorderColor = Theme_Color
PrBShape(3).FillColor = Theme_High
PrBCaption.ForeColor = Theme_Light
PrBCaption.Font = Theme_Text
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 300
PrBShape(0).Move 0, 0, UserControl.Width, UserControl.Height - OnePix
PrBShape(1).Move OnePix, OnePix, UserControl.Width - TwoPix, UserControl.Height - ThreePix
PrBShape(2).Move TwoPix, TwoPix, UserControl.Width - ThreePix, UserControl.Height - FourPix
PrBShape(3).Move 0, 0, OnePix, UserControl.Height
Call Update_Progress
End Sub

Public Property Let Max(ByVal MaxValue As Integer)
MaxVal = MaxValue
Call Update_Progress
End Property
Public Property Get Max() As Integer
Max = MaxVal
End Property

Public Property Let Value(ByVal ValueValue As Integer)
ValVal = ValueValue
Call Update_Progress
End Property
Public Property Get Value() As Integer
Value = ValVal
End Property

Private Sub Update_Progress()
If ValVal = 0 Then
    PrBShape(3).Visible = False
    PrBCaption.Caption = Empty_Code
Else
    PrBShape(3).Width = Int((PrBShape(0).Width / MaxVal) * ValVal)
    PrBCaption.Caption = Round((100 / MaxVal) * ValVal) & "%"
    If PrBShape(3).Width > PrBCaption.Width Then
        PrBCaption.Left = (PrBShape(3).Width / 2) - (PrBCaption.Width / 2)
        PrBShape(3).Visible = True
    Else
        PrBShape(3).Visible = False
    End If
End If
End Sub
