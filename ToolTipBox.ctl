VERSION 5.00
Begin VB.UserControl ToolTipBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   360
   ScaleWidth      =   3615
   Begin VB.Label TiPLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   45
   End
   Begin VB.Image TiPImage 
      Height          =   240
      Left            =   3360
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape TiPShape 
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "ToolTipBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, OldTipString As String, ToolTipNo As Integer, OriginalText As String
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Sub ResetControl()
TiPShape.BorderColor = Theme_High
TiPShape.FillColor = Theme_HighLight
UserControl.BackColor = Theme_HighLight
TiPLabel.ForeColor = Theme_Light
TiPLabel.Font = Theme_Text
TiPImage.Picture = Manager.PictureLoader(2).ListImages.Item(9).Picture
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
TiPShape.Move 0, 0, UserControl.Width, UserControl.Height
TiPImage.Move UserControl.Width - TiPImage.Width - FourPix, UserControl.Height - TiPImage.Height - FourPix
End Sub

Public Property Let Caption(ByVal CaptionValue As String)
If OldTipString <> CaptionValue Then
    OldTipString = CaptionValue
    TiPLabel.Caption = WordWrapper(CaptionValue, EightPix * 28, Theme_Text, False).WrapText
    UserControl.Width = TiPLabel.Width + SixTeenPix
    UserControl.Height = TiPLabel.Height + (OnePix * 12)
End If
End Property
Public Property Get Caption() As String
Caption = TiPLabel.Caption
End Property
