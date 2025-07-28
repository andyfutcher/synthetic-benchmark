VERSION 5.00
Begin VB.UserControl BenchMarkBox 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   CanGetFocus     =   0   'False
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4575
   ScaleWidth      =   6255
   Begin VB.Frame BmbFrame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   360
      Begin VB.Image BmBImage 
         Height          =   3000
         Left            =   0
         Picture         =   "BenchMarkBox.ctx":0000
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape BmbEffect 
         BackColor       =   &H00987054&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00987054&
         Height          =   1335
         Index           =   2
         Left            =   0
         Top             =   3000
         Width           =   375
      End
   End
   Begin VB.Shape BmbEffect 
      BorderColor     =   &H00404040&
      Height          =   4560
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6240
   End
   Begin VB.Shape BmbEffect 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4530
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   6210
   End
End
Attribute VB_Name = "BenchMarkBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_Resize()
On Error Resume Next
BmbEffect(0).Move 0, 0, UserControl.Width - OnePix, UserControl.Height - OnePix
BmbEffect(1).Move OnePix, OnePix, BmbEffect(0).Width - TwoPix, BmbEffect(0).Height - TwoPix
'BmBTextBox.Move EightPix, EightPix, BmbEffect(0).Width - (EightPix * 6), BmbEffect(0).Height - (EightPix * 2)
BmbFrame.Left = UserControl.Width - BmbFrame.Width - (EightPix + OnePix)
BmbFrame.Height = UserControl.Height - (EightPix * 2)
End Sub

Public Sub ResetControl()
BmbEffect(2).Height = Screen.Height
'BmBTextBox.BackColor = Theme_shade
BmbEffect(1).FillColor = Theme_Color
BmbFrame.BackColor = Theme_Color
BmbEffect(0).BorderColor = Theme_Pitch
BmbEffect(1).BorderColor = Theme_Shadow
End Sub

