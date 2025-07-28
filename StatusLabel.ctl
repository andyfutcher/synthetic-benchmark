VERSION 5.00
Begin VB.UserControl StatusLabel 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
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
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   225
   ScaleWidth      =   3555
   Windowless      =   -1  'True
   Begin VB.Line NormLine 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      Visible         =   0   'False
      X1              =   15
      X2              =   15
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line NormLine 
      BorderColor     =   &H00808080&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Label StLCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   45
   End
End
Attribute VB_Name = "StatusLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim NewCatVal As Boolean, SupposedText As String
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SelfAlign()
Public Event Click()

Public Sub ResetControl()
NormLine(0).BorderColor = Theme_Dark
NormLine(1).BorderColor = Theme_Light
StLCaption.ForeColor = Theme_Pitch
StLCaption.Font = Theme_Text
End Sub

Public Property Let Caption(ByVal CaptionValue As String)
If CaptionValue = SupposedText Then GoTo Ed
SupposedText = CaptionValue
StLCaption.ToolTipText = CaptionValue
Call Elipser_Check
Ed: End Property
Public Property Get Caption() As String
Caption = StLCaption.Caption
End Property

Public Property Let NewCategory(ByVal NewCatValue As Boolean)
NewCatVal = NewCatValue
If NewCatValue = True Then
    NormLine(0).Visible = True
    NormLine(1).Visible = True
Else
    NormLine(0).Visible = False
    NormLine(1).Visible = False
End If
End Property
Public Property Get NewCategory() As Boolean
NewCategory = NewCatVal
End Property

Public Sub Elipser_Check()
StLCaption.Caption = WordElipser(SupposedText, UserControl.Width - (StLCaption.Left * 2), Theme_Text, False)
End Sub

Private Sub StLCaption_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
RaiseEvent Click
End Sub
