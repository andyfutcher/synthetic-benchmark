VERSION 5.00
Begin VB.UserControl ExplorerFrame 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ClipBehavior    =   0  'None
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
   HitBehavior     =   0  'None
   ScaleHeight     =   1695
   ScaleWidth      =   2820
   Begin VB.Label EfBHotSpot 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   0
      MouseIcon       =   "ExplorerFrame.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image EfBImage 
      Height          =   240
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   240
   End
   Begin VB.Label EfBCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   380
      TabIndex        =   0
      Top             =   75
      Width           =   60
   End
   Begin VB.Image EfBUpDown 
      Height          =   240
      Left            =   2460
      Stretch         =   -1  'True
      Top             =   60
      Width           =   240
   End
   Begin VB.Shape EfBEffect 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      Height          =   330
      Index           =   1
      Left            =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image EfbHeadImage 
      Height          =   360
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label EfBHeader 
      BackColor       =   &H000000FF&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Shape EfBEffect 
      BorderColor     =   &H00000000&
      Height          =   1695
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "ExplorerFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, ExpandHeight As Integer, Expanded As Boolean, ImageNumber As Integer, SupposedText As String
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event HeaderMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SelfAlign()
Public Event Click()

Private Sub EfBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent HeaderMove(Button, Shift, X, Y)
End Sub

Private Sub Switch_Values()
If Expanded = True Then
    Expanded = False
    GoTo Ed
End If
Expanded = True
Ed: Call Update_Control
End Sub

Private Sub EfBHotSpot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Switch_Values
DoEvents
RaiseEvent Click
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 13, 32
    Call Switch_Values
    RaiseEvent Click
End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_GotFocus()
EfBEffect(1).Visible = True
End Sub
Private Sub UserControl_LostFocus()
EfBEffect(1).Visible = False
End Sub

Private Sub UserControl_Resize()
EfBEffect(0).Move 0, 0, UserControl.Width, UserControl.Height
EfBEffect(1).Move OnePix, OnePix, UserControl.ScaleWidth - TwoPix, EfBHeader.Height - TwoPix
EfBHeader.Width = UserControl.ScaleWidth
EfbHeadImage.Width = UserControl.ScaleWidth
EfBHotSpot.Width = UserControl.ScaleWidth
EfBUpDown.Left = UserControl.ScaleWidth - EfBUpDown.Width - FourPix
End Sub

Public Sub Elipser_Check()
EfBCaption.Caption = WordElipser(SupposedText, EfBUpDown.Left - EfBCaption.Left - TwoPix, Theme_Font, True)
End Sub

Public Sub Refresh()
UserControl.Refresh
End Sub

Public Property Let PanelHeight(ByVal HeightValue As Integer)
ExpandHeight = EfBHotSpot.Height + HeightValue + EightPix
Call Update_Control
End Property
Public Property Get PanelHeight() As Integer
PanelHeight = ExpandHeight
End Property

Public Property Let Width(ByVal WidthValue As Integer)
UserControl.Width = WidthValue
End Property
Public Property Get Width() As Integer
Width = UserControl.Width
End Property
Public Property Get HotSpotHeight() As Integer
HotSpotHeight = EfBHotSpot.Height
End Property

Public Property Let ShowPanel(ByVal ShowValue As Boolean)
Expanded = ShowValue
Call Update_Control
End Property
Public Property Get ShowPanel() As Boolean
ShowPanel = Expanded
End Property

Private Sub Update_Control()
If Expanded = True Then
    UserControl.Height = ExpandHeight
    ImageNumber = 8
Else
    UserControl.Height = EfBHotSpot.Height
    ImageNumber = 9
End If
Call Update_Display
RaiseEvent SelfAlign
End Sub
Private Sub Update_Display()
If ImHighBool = True Then
    EfBCaption.ForeColor = Theme_Light
    EfBUpDown.Picture = Manager.PictureLoader(1).ListImages.Item(ImageNumber).Picture
Else
    EfBCaption.ForeColor = Theme_Color
    EfBUpDown.Picture = Manager.PictureLoader(0).ListImages.Item(ImageNumber).Picture
End If
End Sub

Public Property Let Caption(ByVal CaptionValue As String)
EfBCaption.Caption = CaptionValue
SupposedText = CaptionValue
End Property
Public Property Get Caption() As String
Caption = EfBCaption.Caption
End Property
Public Property Let FrameIcon(ByVal PicLdrValue As Integer)
EfBUpDown.Picture = Manager.PictureLoader(0).ListImages.Item(PicLdrValue).Picture
End Property
Public Property Let FrameHighIcon(ByVal PicLdrValue As Integer)
EfBUpDown.Picture = Manager.PictureLoader(1).ListImages.Item(PicLdrValue).Picture
End Property

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
Call Update_Display
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

Public Sub ResetControl()
ImageNumber = 8
ExpandHeight = EfBHotSpot.Height + EightPix
UserControl.BackColor = Theme_Shade
EfBEffect(0).BorderColor = Theme_High
EfBEffect(1).BorderColor = Theme_Color
EfBHeader.BackColor = Theme_High
EfBCaption.BackColor = Theme_High
EfBCaption.ForeColor = Theme_Color
EfBCaption.Font = Theme_Font
EfBUpDown.Picture = Manager.PictureLoader(0).ListImages.Item(ImageNumber).Picture
Call Update_Control
End Sub

Public Sub FrameProperty(CaptionValue As String, PicLdrValue As Integer)
EfBCaption.Caption = CaptionValue
SupposedText = Filter_Html(CaptionValue)
EfBImage.Picture = Manager.PictureLoader(3).ListImages.Item(PicLdrValue).Picture
EfbHeadImage.Picture = Manager.PictureLoader(2).ListImages.Item(8).Picture
End Sub
