VERSION 5.00
Begin VB.UserControl ToolButton 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
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
   HasDC           =   0   'False
   HitBehavior     =   2  'Use Paint
   MaskColor       =   &H00FFFFFF&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   1305
   ScaleWidth      =   2625
   Windowless      =   -1  'True
   Begin VB.Label TlBHotSpot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape TlBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1095
   End
   Begin VB.Line NormLine 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      Visible         =   0   'False
      X1              =   15
      X2              =   15
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line NormLine 
      BorderColor     =   &H00808080&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Label TlBCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   405
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
   Begin VB.Image TlBImage 
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape TlBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape TlBShape 
      BackColor       =   &H00F0F0F0&
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "ToolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, NewCatVal As Boolean, ToolCommandID As Integer
Dim SunkenBool As Boolean, EnabledControl As Boolean, ControlIcon As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click(CommandID As Integer)

Private Sub TlBHotSpot_Click()
If EnabledControl = True Then RaiseEvent Click(ToolCommandID)
End Sub

Private Sub TlBHotSpot_DblClick()
If EnabledControl = True Then RaiseEvent Click(ToolCommandID)
Call TlBHotSpot_MouseDown(1, 0, 0, 0)
End Sub

Private Sub TlBHotSpot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TlBShape(0).FillColor = Theme_Color
TlBShape(0).BorderColor = Theme_Pitch
TlBShape(1).BorderColor = Theme_Shadow
TlBShape(2).BorderColor = Theme_Dark
TlBCaption.ForeColor = Theme_Pitch
TlBCaption.Top = Int((UserControl.Height / 2) - (TlBCaption.Height / 2) + OnePix)
TlBImage.Top = Int((UserControl.Height / 2) - (TlBImage.Height / 2)) + OnePix - 1
End Sub
Private Sub TlBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
Call Show_ToolTip(ToolCommandID, 0)
End Sub
Private Sub TlBHotSpot_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TlBShape(0).FillColor = Theme_Shade
TlBShape(0).BorderColor = Theme_Shadow
TlBShape(1).BorderColor = Theme_Dark
TlBShape(2).BorderColor = Theme_Color
TlBCaption.ForeColor = Theme_Shadow
TlBCaption.Top = Int((UserControl.Height / 2) - (TlBCaption.Height / 2))
TlBImage.Top = (UserControl.Height / 2) - (TlBImage.Height / 2)
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
If UserControl.Enabled = True Then RaiseEvent MouseMove(0, 0, X, Y)
End Sub

Private Sub UserControl_Resize()
TlBImage.Top = (UserControl.Height / 2) - (TlBImage.Height / 2)
TlBCaption.Top = Int((UserControl.Height / 2) - (TlBCaption.Height / 2))
If NewCatVal = True Then
    TlBImage.Left = OnePix * 12
    TlBCaption.Left = OnePix * 32
    If TlBCaption.Caption = Empty_Code Then
        UserControl.Width = TlBImage.Left + TlBImage.Width + EightPix
    Else
        UserControl.Width = TlBCaption.Left + TlBCaption.Width + EightPix '+ (OnePix * 1)
    End If
    TlBShape(0).Move OnePix * 3, 0, UserControl.Width - OnePix * 3, UserControl.Height
    TlBShape(1).Move OnePix * 4, OnePix, UserControl.Width - OnePix * 5, UserControl.Height - TwoPix
    TlBShape(2).Move OnePix * 5, TwoPix, UserControl.Width - OnePix * 7, UserControl.Height - FourPix
    TlBHotSpot.Move OnePix * 3, 0, UserControl.Width, UserControl.Height
    NormLine(0).X1 = OnePix
    NormLine(0).X2 = OnePix
    NormLine(1).X1 = TwoPix
    NormLine(1).X2 = TwoPix
    NormLine(0).Y1 = TwoPix
    NormLine(1).Y1 = TwoPix
    NormLine(0).Y2 = UserControl.Height - OnePix
    NormLine(1).Y2 = UserControl.Height - OnePix
    NormLine(0).Visible = True
    NormLine(1).Visible = True
Else
    TlBImage.Left = OnePix * 8
    TlBCaption.Left = OnePix * 28
    If TlBCaption.Caption = Empty_Code Then
        UserControl.Width = TlBImage.Left + TlBImage.Width + EightPix + OnePix
    Else
        UserControl.Width = TlBCaption.Left + TlBCaption.Width + EightPix
    End If
    TlBShape(0).Move 0, 0, UserControl.Width, UserControl.Height
    TlBShape(1).Move OnePix, OnePix, UserControl.Width - TwoPix, UserControl.Height - TwoPix
    TlBShape(2).Move TwoPix, TwoPix, UserControl.Width - FourPix, UserControl.Height - FourPix
    TlBHotSpot.Move 0, 0, UserControl.Width, UserControl.Height
    NormLine(0).Visible = False
    NormLine(1).Visible = False
End If
End Sub

Public Property Get Caption() As String
Caption = TlBCaption.Caption
End Property
Public Property Get NewCategory() As Boolean
NewCategory = NewCatVal
End Property

Public Sub ToolProperty(CmdIDValue As Integer, CaptionValue As String, PicLdrValue As Integer, EnaControl As Boolean, NewCatValue As Integer)
NewCatVal = NewCatValue
TlBCaption.Caption = CaptionValue
ToolCommandID = CmdIDValue
EnabledControl = EnaControl
ControlIcon = PicLdrValue
Call Switch_Enabled_Mode
Call UserControl_Resize
End Sub

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighBool = True Then
    TlBShape(0).FillStyle = 0
    TlBShape(0).BorderStyle = 1
    TlBShape(1).BorderStyle = 1
    TlBShape(2).BorderStyle = 1
    'TlBCaption.ForeColor = Theme_Pitch
Else
    If SunkenBool = False Then
        TlBShape(0).FillStyle = 1
        TlBShape(0).BorderStyle = 0
    End If
    TlBShape(1).BorderStyle = 0
    TlBShape(2).BorderStyle = 0
    'TlBShape(0).FillColor = Theme_Shade
    'TlBShape(0).BorderColor = Theme_Dark
End If
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

Public Property Let Sunken(ByVal SunkenValue As Boolean)
SunkenBool = SunkenValue
If SunkenValue = True Then
    TlBShape(0).FillStyle = 0
    TlBShape(0).BorderStyle = 1
Else
    TlBShape(0).FillStyle = 1
    TlBShape(0).BorderStyle = 0
End If
End Property
Public Property Get Sunken() As Boolean
Sunken = SunkenBool
End Property

Public Property Get Command() As Integer
Command = ToolCommandID
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
If UserControl.Enabled = EnabledValue Then GoTo Ed
UserControl.Enabled = EnabledValue
Call Switch_Enabled_Mode
Ed: End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Private Sub Switch_Enabled_Mode()
If UserControl.Enabled = True Then
    TlBCaption.ForeColor = Theme_Shadow
    TlBImage.Picture = Manager.PictureLoader(3).ListImages.Item(ControlIcon).Picture
Else
    TlBCaption.ForeColor = Theme_Dark
    TlBImage.Picture = Manager.PictureLoader(4).ListImages.Item(ControlIcon).Picture
End If
End Sub

Public Sub ResetControl()
ImHigh = False
Call TlBHotSpot_MouseUp(0, 0, 0, 0)
NormLine(0).BorderColor = Theme_Dark
NormLine(1).BorderColor = Theme_Shade
TlBCaption.Font = Theme_Text
End Sub

