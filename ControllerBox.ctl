VERSION 5.00
Begin VB.UserControl ControllerBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3945
   ScaleWidth      =   5280
   Begin VB.Label ClBHotSpot 
      BackStyle       =   0  'Transparent
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.Line NormLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   375
      X2              =   5040
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line NormLine 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   375
      X2              =   5040
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label ClBCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
   Begin VB.Image ClBImage 
      Height          =   240
      Index           =   0
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   135
      Width           =   240
   End
   Begin VB.Label ClBCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   15
      Width           =   120
   End
   Begin VB.Shape ClBEffect 
      BorderColor     =   &H00404040&
      Height          =   3840
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5155
   End
   Begin VB.Shape ClBEffect 
      BorderColor     =   &H00C0C0C0&
      Height          =   3810
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   5130
   End
   Begin VB.Shape ClBEffect 
      BorderColor     =   &H00C0C0C0&
      Height          =   3810
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   5130
   End
   Begin VB.Image ClBHolder 
      Appearance      =   0  'Flat
      Height          =   3855
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   315
   End
   Begin VB.Image ClBBackGround 
      Height          =   3840
      Left            =   315
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "ControllerBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Click()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Let Caption(ByVal CaptionValue As String)
ClBCaption(0).Caption = CaptionValue
ClBCaption(1).Caption = CaptionValue
End Property
Public Property Get Caption() As String
Caption = ClBCaption(0).Caption
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
UserControl.Enabled = EnabledValue
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Sub ControllerProperty(CaptionValue As String, PicLdrValue As Integer)
ClBCaption(0).Caption = CaptionValue
ClBCaption(1).Caption = CaptionValue
If PicLdrValue = 0 Then
    ClBImage(0).Picture = ManagerSub.Icon
Else
    ClBImage(0).Picture = Manager.PictureLoader(3).ListImages.Item(PicLdrValue).Picture
End If
End Sub

Private Sub ClBHotSpot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Click
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub ClBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub ClBHotSpot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
ClBHotSpot.Move 0, 0, UserControl.Width, UserControl.Height
ClBEffect(0).Move 0, 0, UserControl.Width - OnePix, UserControl.Height - OnePix
ClBEffect(1).Move OnePix, OnePix, ClBEffect(0).Width - TwoPix, ClBEffect(0).Height - TwoPix
ClBEffect(2).Move 0, 0, UserControl.Width, UserControl.Height
ClBHolder(0).Height = ClBEffect(0).Height
ClBBackGround.Move ClBHolder(0).Width, OnePix, ClBEffect(0).Width - ClBHolder(0).Width, ClBEffect(0).Height
ClBImage(0).Left = ClBEffect(0).Width - ClBImage(0).Width - EightPix
NormLine(0).X2 = ClBEffect(0).Width - EightPix
NormLine(1).X2 = ClBEffect(0).Width - EightPix
End Sub

Public Sub ResetControl()
UserControl.BackColor = Theme_Color
ClBEffect(0).BorderColor = Theme_Pitch
ClBEffect(1).BorderColor = Theme_Dark
ClBEffect(2).BorderColor = Theme_Pitch
ClBHolder(0).Picture = Manager.PictureLoader(1).ListImages.Item(2).Picture
ClBBackGround.Picture = Manager.PictureLoader(0).ListImages.Item(1).Picture
ClBImage(0).Picture = ManagerSub.Icon

ClBCaption(0).ForeColor = Theme_Shadow
ClBCaption(1).ForeColor = Theme_Vague
ClBCaption(0).Font = Theme_Text
ClBCaption(1).Font = Theme_Text

NormLine(0).Y1 = ClBCaption(0).Top + ClBCaption(0).Height + OnePix
NormLine(1).Y1 = ClBCaption(0).Top + ClBCaption(0).Height + TwoPix
NormLine(0).Y2 = NormLine(0).Y1
NormLine(1).Y2 = NormLine(1).Y1

NormLine(0).BorderColor = Theme_Dark
NormLine(1).BorderColor = Theme_Shade
End Sub
