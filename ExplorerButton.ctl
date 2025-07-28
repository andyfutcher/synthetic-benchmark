VERSION 5.00
Begin VB.UserControl ExplorerButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ClipBehavior    =   0  'None
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   MaskColor       =   &H00E0E0E0&
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   300
   ScaleWidth      =   3030
   Begin VB.Label ExBHotSpot 
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
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label ExBLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   45
      Width           =   45
   End
   Begin VB.Image ExBImage 
      Height          =   240
      Left            =   60
      Stretch         =   -1  'True
      Top             =   30
      Width           =   240
   End
   Begin VB.Shape ExBShape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00987054&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "ExplorerButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, ToolCommandID As Integer
Public Event MouseMove(MenuCmdID As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SelfAlign()
Public Event Click(MenuCmdID As Integer)

Private Sub ExBHotSpot_Click()
RaiseEvent Click(ToolCommandID)
End Sub

Private Sub ExBHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(ToolCommandID, Button, Shift, X, Y)
End Sub

Private Sub UserControl_GotFocus()
ExBShape.Visible = True
End Sub

Private Sub UserControl_LostFocus()
ExBShape.Visible = False
End Sub

Private Sub Align_Text()
ExBLabel.Caption = WordWrapper(ExBLabel.Caption, UserControl.Width - ExBLabel.Left, Theme_Text, False).WrapText
End Sub

Public Sub Align_Controls()
'On Error Resume Next
Call Align_Text
Dim MostWidth As Integer, MostHeight As Integer
MostHeight = ExBImage.Top + ExBImage.Height
If ExBLabel.Top + ExBLabel.Height > MostHeight Then MostHeight = ExBLabel.Top + ExBLabel.Height + OnePix
MostWidth = ExBLabel.Left + ExBLabel.Width + (OnePix * 5)

ExBShape.Move 0, 0, MostWidth, MostHeight + TwoPix
ExBHotSpot.Move ExBShape.Left, ExBShape.Top, ExBShape.Width, ExBShape.Height
UserControl.Height = MostHeight + TwoPix
End Sub

Public Property Let FontUnderline(ByVal UnderLineValue As Boolean)
ExBLabel.FontUnderline = UnderLineValue
End Property
Public Property Get FontUnderline() As Boolean
FontUnderline = ExBLabel.FontUnderline
End Property

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = True Then
    Manager.ExplorerButton(ControlCount).FontUnderline = True
Else
    Manager.ExplorerButton(ControlCount).FontUnderline = False
End If
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

Public Property Get Command() As Integer
Command = ToolCommandID
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
UserControl.Enabled = EnabledValue
'RaiseEvent WindowLess(EnabledValue)
ExBImage.Visible = EnabledValue
If EnabledValue = True Then
    ExBLabel.ForeColor = Theme_Pitch
Else
    ExBLabel.ForeColor = Theme_Dark
End If
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Sub ResetControl()
ExBLabel.ForeColor = Theme_Pitch
ExBLabel.Font = Theme_Text
ExBShape.BorderColor = Theme_InvertLight
ExBLabel.BackColor = Theme_Shade
UserControl.BackColor = Theme_Shade
UserControl.MaskColor = Theme_Shade
ExBHotSpot.MouseIcon = Manager.PictureLoader(2).ListImages.Item(1).Picture
End Sub

Public Sub ButtonProperty(CmdIDValue As Integer, CaptionValue As String, PicLdrValue As Integer)
ToolCommandID = CmdIDValue
ExBImage.Picture = Manager.PictureLoader(3).ListImages.Item(PicLdrValue).Picture
ExBLabel.Caption = Filter_Html(CaptionValue)
Call Align_Controls
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(0, Button, Shift, X, Y)
End Sub
