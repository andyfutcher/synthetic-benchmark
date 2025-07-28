VERSION 5.00
Begin VB.UserControl WriteBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
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
   KeyPreview      =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   3630
   Begin VB.TextBox WrBText 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   210
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   3255
   End
   Begin VB.Label WrBWhiteness 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label WrbHotSpot 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   3240
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Image WrBArrow 
      Height          =   120
      Index           =   0
      Left            =   3390
      Stretch         =   -1  'True
      Top             =   75
      Width           =   135
   End
   Begin VB.Shape WrBEffect 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   3585
   End
   Begin VB.Shape WrBEffect 
      BorderColor     =   &H00808080&
      Height          =   270
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "WriteBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, ControlHasFocus As Boolean, DDParent As String, DDListArray() As String
Dim InternalSelection As Integer, SelectStartIndex As Integer, SelectLenIndex As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPressed(KeyCode As Integer, Shift As Integer)
Public Event Changed()
Public Event DropClick(DDListArray() As String, SelectedIndex As Integer)
Public Event SelfAlign()

Private Sub UserControl_Click()
WrBText.SetFocus
End Sub

Private Sub UserControl_GotFocus()
WrBText.SetFocus
'WrBEffect(1).BorderStyle = 3
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 270
WrBEffect(0).Move 0, 0, UserControl.Width
WrBEffect(1).Move OnePix, OnePix, UserControl.Width - TwoPix
WrBWhiteness.Move TwoPix, TwoPix, TwoPix, 270 - FourPix
If DDParent <> Empty_Code Then
    WrbHotSpot.Left = UserControl.Width - TwoPix - SixTeenPix
    WrBArrow(0).Left = UserControl.Width - TwoPix - EightPix - (WrBArrow(0).Width / 2)
    WrBArrow(0).Visible = True
    WrBText.Move FourPix, TwoPix, UserControl.Width - FivePix - SixTeenPix
Else
    WrBText.Move FourPix, TwoPix, UserControl.Width - SixPix
    WrBArrow(0).Visible = False
End If
End Sub

Public Sub WriteProperty(DDPValue As String, LockedValue As Boolean, BoldValue As Boolean, PasswordType As Boolean)
DDParent = DDPValue
WrBText.Locked = LockedValue
WrBText.FontBold = BoldValue
If PasswordType = True Then WrBText.PasswordChar = Trim(Bullet_Char)
Call UserControl_Resize
End Sub

Public Sub SwitchToIndex(Index As Integer)
WrBText.Text = DDListArray(0, Index)
WrBText.Tag = DDListArray(1, Index)
InternalSelection = Index
End Sub

Public Property Let ListIndex(ByVal IndexValue As Integer)
Call SwitchToIndex(IndexValue)
End Property
Public Property Get ListIndex() As Integer
ListIndex = InternalSelection
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
'EnabledControl = EnabledValue
UserControl.Enabled = EnabledValue
If EnabledValue = False Then
    WrBEffect(0).BorderColor = Theme_Dark
    WrBEffect(1).FillColor = Theme_Light
    WrBEffect(1).BorderColor = Theme_Light
    WrBText.ForeColor = Theme_Shadow
    WrBText.BackColor = Theme_Light
    WrBWhiteness.BackColor = Theme_Light
Else
    WrBEffect(0).BorderColor = Theme_Shadow
    WrBEffect(1).FillColor = Theme_Shade
    WrBEffect(1).BorderColor = Theme_Shade
    WrBText.ForeColor = Theme_Pitch
    WrBText.BackColor = Theme_Light
    WrBWhiteness.BackColor = Theme_Light
End If
WrBArrow(0).Refresh
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let DropDownParent(ByVal DDPValue As String)
DDParent = DDPValue
Call UserControl_Resize
End Property
Public Property Get DropDownParent() As String
DropDownParent = DDParent
End Property

Public Property Let SelectionStart(ByVal SStartValue As Integer)
WrBText.SelStart = SStartValue
End Property
Public Property Get SelectionStart() As Integer
SelectionStart = WrBText.SelStart
End Property

Public Property Let Text(ByVal TextValue As String)
WrBText.Text = TextValue
End Property
Public Property Get Text() As String
Text = Replace(WrBText.Text, LineFeed, Empty_Code)
End Property
Public Property Get ClickTag() As String
ClickTag = WrBText.Tag
End Property

Private Sub WrbHotSpot_Click()
If DDParent <> Empty_Code Then RaiseEvent DropClick(DDListArray(), InternalSelection)
End Sub

Private Sub WrbHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
'ImHighBool = True
'WrBEffect(1).FillColor = Theme_Color
End Sub

Private Sub WrBText_Change()
RaiseEvent Changed
End Sub

Private Sub WrBText_GotFocus()
ControlHasFocus = True
End Sub

Private Sub WrBText_KeyDown(KeyCode As Integer, Shift As Integer)
If WrBText.Locked = True Then RaiseEvent DropClick(DDListArray(), InternalSelection)
End Sub

Private Sub WrBText_KeyUp(KeyCode As Integer, Shift As Integer)
SelectStartIndex = WrBText.SelStart
SelectLenIndex = WrBText.SelLength
WrBText.Text = Replace(WrBText.Text, LineFeed, Empty_Code)
If SelectStartIndex > Len(WrBText.Text) Then
    WrBText.SelStart = Len(WrBText.Text)
Else
    WrBText.SelStart = SelectStartIndex
End If
WrBText.SelLength = SelectLenIndex
RaiseEvent KeyPressed(KeyCode, Shift)
End Sub

Private Sub WrBText_LostFocus()
ControlHasFocus = False
End Sub

Private Sub WrBText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
'WrBEffect(1).FillColor = Theme_Shade
End Sub

Public Sub DDList_Clear()
ReDim DDListArray(1, 0)
End Sub
Public Sub DDList_Add(AddValue As String, TagValue As String)
Call Add_Index_To_StringArray(DDListArray())
DDListArray(0, UBound(DDListArray, 2)) = Filter_QuickLinks(AddValue)
DDListArray(1, UBound(DDListArray, 2)) = TagValue
If UBound(DDListArray(), 2) = 0 Then Call SwitchToIndex(0)
Call UserControl_Resize
End Sub

Public Sub ResetControl()
If NowLoading = True Then Call DDList_Clear
If App_WinVersion = Windows98_Code Then
    UserControl.BackStyle = 1
    WrBEffect(0).Shape = 0
End If
UserControl.Height = 270
WrBEffect(0).BorderColor = Theme_Shadow
WrBEffect(1).BorderColor = Theme_Shade
WrBEffect(1).FillColor = Theme_Shade
UserControl.BackColor = Theme_Shade
WrBWhiteness.BackColor = Theme_Light
WrBArrow(0).Picture = Manager.PictureLoader(2).ListImages.Item(4).Picture
WrBText.BackColor = Theme_Light
WrBText.ForeColor = Theme_Pitch
WrBText.Font = Theme_Text
RaiseEvent Changed
'Call UserControl_Resize
End Sub

'Public Property Let ImHigh(ByVal ImHighValue As Boolean)
'ImHighBool = ImHighValue
'If ImHighValue = False Then
'    WrBEffect(1).FillColor = Theme_Shade
'End If
'End Property
'Public Property Get ImHigh() As Boolean
'ImHigh = ImHighBool
'End Property

'Dim IntSel As Integer
    'IntSel = InternalSelection
    'Select Case KeyCode
    'Case vbKeyDown
    '    IntSel = IntSel + 1
    '    If UBound(DDListArray()) < IntSel Then IntSel = 0
    'Case vbKeyUp
    '    IntSel = IntSel - 1
    '    If IntSel < 0 Then IntSel = UBound(DDListArray(), 2)
    'End Select
    'Call SwitchToIndex(IntSel)
    'Select Case KeyCode
    'Case vbKeyDown
    '    WrBText.SelStart = Len(DDListArray(0, IntSel))
    'Case vbKeyUp
    '    WrBText.SelStart = 0
    'End Select
