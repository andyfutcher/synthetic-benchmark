VERSION 5.00
Begin VB.UserControl MenuList 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
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
   ScaleHeight     =   360
   ScaleWidth      =   2775
   Windowless      =   -1  'True
   Begin VB.Label MuLHotSpot 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Image MuLArrow 
      Height          =   135
      Left            =   2520
      Top             =   90
      Width           =   120
   End
   Begin VB.Label MuLCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Text"
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
      Left            =   460
      TabIndex        =   0
      Top             =   75
      Width           =   765
   End
   Begin VB.Image MuLImage 
      Height          =   240
      Left            =   75
      Stretch         =   -1  'True
      Top             =   60
      Width           =   240
   End
   Begin VB.Label MuLShader 
      BackColor       =   &H00C0C0C0&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "MenuList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim InUseBool As Boolean, MenuCommandId As Integer, EnabledVal As Boolean, StayOpenVal As Boolean
Dim StayOpenNow As Boolean, AlreadyFocued As Boolean, XHotSpot As Integer
Public Event KeyPress(KeyAscii As Integer)
Public Event Click(ClickType As Integer, MenuCmdID As Integer)
Public Event ForceLooseFocus()
Public Event CleanOthers()
Public Event MouseMoved(MenuCmdID As Integer)

Private Const Def_Height = 330

Private Sub MuLHotSpot_Click()
If XHotSpot < MuLShader.Width Then
    Call Process_Click(1)
Else
    Call Process_Click(0)
End If
End Sub

Private Sub MuLHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMoved(MenuCommandId)
If AlreadyFocued = True Then GoTo Ed
UserControl.SetFocus
XHotSpot = X
Ed: End Sub

Private Sub MuLShader_Click()
Call Process_Click(1)
End Sub

Private Sub MuLShader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl.SetFocus
RaiseEvent MouseMoved(MenuCommandId)
End Sub

Public Sub Dont_Stay_Open()
StayOpenNow = False
Call UserControl_LostFocus
End Sub

Private Sub UserControl_GotFocus()
AlreadyFocued = True
If EnabledVal = True Then
    MuLCaption.ForeColor = Theme_Light
    UserControl.BackColor = Theme_High
Else
    MuLCaption.ForeColor = Theme_Dark
    UserControl.BackColor = Theme_Color
End If
MuLShader.Visible = False
End Sub

Private Sub UserControl_Hide()
Call UserControl_LostFocus
End Sub

Public Sub Loose_Focus()
Call UserControl_LostFocus
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13, 32
    Call Process_Click(0)
End Select
RaiseEvent KeyPress(KeyCode)
End Sub

Private Sub UserControl_LostFocus()
If StayOpenNow = True Then GoTo Ed
UserControl.BackColor = Theme_Light
MuLShader.Visible = True
If EnabledVal = True Then MuLCaption.ForeColor = Theme_Pitch
AlreadyFocued = False
Ed: End Sub
Private Sub UserControl_Resize()
'UserControl.Height = (20 * Screen.TwipsPerPixelY)
MuLHotSpot.Width = UserControl.Width
MuLArrow.Left = UserControl.Width - MuLArrow.Width - 75
End Sub

Public Property Get MinimumLength() As Integer
Dim MinLen As Integer
MinLen = MuLCaption.Left + MuLCaption.Width + 150
MinimumLength = MinLen
End Property

Public Property Let MenuIcon(ByVal PicLdrValue As Integer)
If EnabledVal = True Then
    MuLImage.Picture = Manager.PictureLoader(3).ListImages.Item(PicLdrValue).Picture
Else
    MuLImage.Picture = Manager.PictureLoader(4).ListImages.Item(PicLdrValue).Picture
End If
End Property
Public Property Let Caption(ByVal CaptionValue As String)
MuLCaption.Caption = CaptionValue
End Property
Public Property Get Caption() As String
Caption = MuLCaption.Caption
End Property
Public Property Let StayOpen(ByVal StayOpenValue As Boolean)
StayOpenVal = StayOpenValue
End Property
Public Property Get StayOpen() As Boolean
StayOpen = StayOpenVal
End Property
Public Property Let IsParent(ByVal ParentValue As Boolean)
MuLArrow.Visible = ParentValue
End Property
Public Property Get IsParent() As Boolean
IsParent = MuLArrow.Visible
End Property
Public Property Get SuggestedWidth() As String
SuggestedWidth = MuLCaption.Left + MuLCaption.Width + (24 * Screen.TwipsPerPixelX)
End Property

Public Property Let CommandID(ByVal CmdIDValue As Integer)
MenuCommandId = CmdIDValue
End Property
Public Property Get CommandID() As Integer
CommandID = MenuCommandId
End Property

Public Property Let InUse(ByVal InUseValue As Boolean)
InUseBool = InUseValue
End Property
Public Property Get InUse() As Boolean
InUse = InUseBool
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
EnabledVal = EnabledValue
If EnabledValue = True Then
    MuLCaption.ForeColor = Theme_Pitch
    'MuLImage.Visible = True
Else
    MuLCaption.ForeColor = Theme_Dark
    'MuLImage.Visible = False
End If
End Property
Public Property Get Enabled() As Boolean
Enabled = EnabledVal
End Property

Public Sub ResetControl()
UserControl.Height = Def_Height
MuLHotSpot.Height = Def_Height
MuLArrow.Picture = Manager.PictureLoader(2).ListImages.Item(3).Picture
UserControl.BackColor = Theme_Light
MuLShader.BackColor = Theme_Shade
MuLCaption.Font = Theme_Text
End Sub

Private Sub Process_Click(ClickType As Integer)
'UserControl.BackColor = Theme_Invert
If EnabledVal = True Then
    If MuLArrow.Visible = False And StayOpenVal = False Then
        RaiseEvent ForceLooseFocus
    Else
        RaiseEvent CleanOthers
        If MuLArrow.Visible = True Then
            StayOpenNow = True
        Else
            StayOpenNow = False
        End If
    End If
    If MuLArrow.Visible = False Then
        ClickType = 0
    End If
    RaiseEvent Click(ClickType, MenuCommandId)
End If
End Sub

