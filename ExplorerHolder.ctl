VERSION 5.00
Begin VB.UserControl ExplorerHolder 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
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
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   7365
   ScaleWidth      =   3465
   Begin VB.Label DragSlider 
      BackColor       =   &H00E0E0E0&
      Height          =   7335
      Left            =   3360
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   0
      Width           =   75
   End
   Begin VB.Image BackGround 
      Height          =   7335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5000
   End
End
Attribute VB_Name = "ExplorerHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, LastExpWidth As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SelfAlign()
Public Event Click()

Const MaxExpWidth = 5000
Const MinExpWidth = 1800

Private Sub BackGround_Click()
RaiseEvent Click
End Sub

Private Sub BackGround_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let Enabled(ByVal EnabledValue As Boolean)
UserControl.Enabled = EnabledValue
End Property
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Private Sub DragSlider_Click()
RaiseEvent Click
End Sub

Private Sub DragSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If UserDoneOnce = True Then GoTo Ed
    UserDoneOnce = True
    If UserControl.Width + X - 30 < MaxExpWidth And UserControl.Width + X - 30 > MinExpWidth Then
        UserControl.Width = UserControl.Width + X - TwoPix
    Else
        If UserControl.Width + X - 30 < MaxExpWidth Then
            UserControl.Width = MinExpWidth
        Else
            UserControl.Width = MaxExpWidth
        End If
        'DragSlider.Left = UserControl.Width - DragSlider.Width
        'RaiseEvent SelfAlign
    End If
    If UserControl.Width <> LastExpWidth Then
        DragSlider.Left = UserControl.Width - DragSlider.Width
        RaiseEvent SelfAlign
        LastExpWidth = UserControl.Width
    End If
    UserDoneOnce = False
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
Ed: End Sub

Public Sub Refresh()
UserControl.Refresh
End Sub

Public Sub Width_Resize()
'BackGround.Width = UserControl.Width - FourPix
'UserControl.Refresh
RaiseEvent SelfAlign
End Sub

Public Sub ResetControl()
DragSlider.Height = Screen_Height
BackGround.Height = Screen_Height - Desktop_Top
BackGround.Width = MaxExpWidth - FourPix
UserControl.BackColor = Theme_Shade
DragSlider.BackColor = Theme_Dark
BackGround.Picture = LoadPicture(System_Path & Theme_Path & Theme_Current & Texture_Name & One_Code & Resource_Ext)
'BackGround.Picture = Manager.PictureLoader(0).ListImages.Item(1).Picture
RaiseEvent SelfAlign
DragSlider.Left = UserControl.Width - DragSlider.Width
RaiseEvent SelfAlign
End Sub

Private Sub VScrollButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get UseableArea() As Integer
UseableArea = UserControl.Width - DragSlider.Width
End Property

