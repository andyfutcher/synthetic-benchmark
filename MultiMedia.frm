VERSION 5.00
Begin VB.Form MultiMedia 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MultiMedia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MediaWindow 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label FrameLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "MultiMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MultiMedia.MouseIcon = Manager.PictureLoader(3).ListImages.Item(1).Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoved = True
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseClicked = True
End Sub
Private Sub FrameLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoved = True
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
KeyHeld = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
KeyTapp = True
End Sub


