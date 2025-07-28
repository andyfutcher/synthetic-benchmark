VERSION 5.00
Begin VB.UserControl StylishLabel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   300
   ScaleWidth      =   3840
   Windowless      =   -1  'True
   Begin VB.Label StLCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(caption)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   30
      Width           =   645
   End
   Begin VB.Label StLCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(caption)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   395
      TabIndex        =   0
      Top             =   30
      Width           =   780
   End
   Begin VB.Image StLImage 
      Height          =   15
      Left            =   0
      Stretch         =   -1  'True
      Top             =   280
      Width           =   3840
   End
   Begin VB.Image StIcon 
      Height          =   240
      Left            =   60
      Picture         =   "StylishLabel.ctx":0000
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "StylishLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub ResetControl()
UserControl.Height = 300
StLCaption(0).ForeColor = Theme_Pitch
StLCaption(1).ForeColor = Theme_Pitch
StLCaption(0).Font = Theme_Text
StLCaption(1).Font = Theme_Text
StLImage.Picture = Manager.PictureLoader(2).ListImages.Item(7).Picture
End Sub

Public Sub LabelProperty(CaptionValue As String, AltCaptionValue As String, PicLdrValue As Integer)
StIcon.Picture = Manager.PictureLoader(3).ListImages.Item(PicLdrValue).Picture
StLCaption(0).Caption = CaptionValue
StLCaption(1).Caption = AltCaptionValue
If StLCaption(1).Width + EightPix > StLCaption(0).Width + EightPix Then
    StLCaption(1).Left = StLCaption(0).Left + StLCaption(0).Width + SixPix
    UserControl.Width = StLCaption(1).Left + StLCaption(1).Width + (EightPix * 9)
    StLImage.Width = UserControl.Width
End If
End Sub

Public Property Get Caption() As String
Caption = StLCaption(0).Caption
End Property
Public Property Get AltCaption() As String
AltCaption = StLCaption(1).Caption
End Property

Private Sub StLCaption_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub StLImage_Click()
RaiseEvent Click
End Sub

