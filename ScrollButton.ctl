VERSION 5.00
Begin VB.UserControl VScrollButton 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00F0F0F0&
   CanGetFocus     =   0   'False
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
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
   ScaleHeight     =   1695
   ScaleWidth      =   255
   Begin VB.Frame ScbFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   255
      Begin VB.Label ScBHotSpot 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.Image ScBArrow 
         Height          =   120
         Index           =   1
         Left            =   60
         Stretch         =   -1  'True
         Top             =   75
         Width           =   135
      End
      Begin VB.Shape ScBEffect 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame ScbFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
      Begin VB.Label ScBHotSpot 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin VB.Image ScBArrow 
         Height          =   120
         Index           =   0
         Left            =   60
         Stretch         =   -1  'True
         Top             =   60
         Width           =   135
      End
      Begin VB.Shape ScBEffect 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame ScbFrame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   255
      Width           =   255
      Begin VB.Label ScBHotSpot 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.Image ScBArrow 
         Height          =   165
         Index           =   2
         Left            =   60
         Stretch         =   -1  'True
         Top             =   105
         Width           =   135
      End
      Begin VB.Shape ScBEffect 
         BackColor       =   &H00808080&
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer ScrollTimer 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "VScrollButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, GapVal As Long, PlayVal As Long, ValVal As Long, Slide_Len As Integer
Dim NewTopCoOrd As Integer, NewHeightCoOrd As Integer, VscrollCoOrd(0) As Long, LastVal As Long
Dim ScrollRate As Integer, LastXY(1) As Single
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event SliderDrag(NowValue As Long)
Const ScrollSpeed = 10

Private Sub ScBHotSpot_DblClick(Index As Integer)
Select Case Index
Case 0
    ScrollTimer.Tag = ScrollRate
    ScrollTimer.Interval = ScrollSpeed
Case 1
    ScrollTimer.Tag = (0 - ScrollRate)
    ScrollTimer.Interval = ScrollSpeed
End Select
ImHighBool = False
End Sub

Private Sub ScBHotSpot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
    ScrollTimer.Tag = ScrollRate
    ScrollTimer.Interval = ScrollSpeed
Case 1
    ScrollTimer.Tag = (0 - ScrollRate)
    ScrollTimer.Interval = ScrollSpeed
Case 2
    VscrollCoOrd(0) = Y
End Select
ImHighBool = False
End Sub

Private Sub ScBHotSpot_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If UserCount = True Then GoTo Ed
UserCount = True
If VscrollCoOrd(0) <> -1 And Button <> 0 And Index = 2 Then
    ValVal = ((ScbFrame(2).Top - ScbFrame(0).Height) - VscrollCoOrd(0) + Y)
    'ValVal = (Int((ScbFrame(2).Top - ScbFrame(0).Height) - VscrollCoOrd(0) + Y) / OnePix) * OnePix
    Call Process_CoOrdinates(0)
Else
    For ControlCount = 0 To ScBHotSpot.Count - 1
        If ControlCount = Index Then
            ScBEffect(ControlCount).FillColor = Theme_High
            ScbFrame(2).BackColor = Theme_High
        Else
            If ControlCount = 2 Then
                ScBEffect(ControlCount).FillColor = Theme_Color
            Else
                ScBEffect(ControlCount).FillColor = Theme_Dark
            End If
        End If
        Next ControlCount
    RaiseEvent MouseMove(Button, Shift, X, Y)
    ImHighBool = True
End If
UserCount = False
Ed: End Sub

Private Sub ScBHotSpot_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0, 1
    ScrollTimer.Interval = 0
Case 2
    VscrollCoOrd(0) = -1
End Select
ImHighBool = True
End Sub

Private Sub ScrollTimer_Timer()
ValVal = (Int((ScbFrame(2).Top - ScbFrame(0).Height) - (ScrollTimer.Tag)) / OnePix) * OnePix
Call Process_CoOrdinates(0)
End Sub


Private Sub UserControl_DblClick()
Call UserControl_MouseDown(1, 0, LastXY(0), LastXY(1))
End Sub

Private Sub UserControl_Initialize()
VscrollCoOrd(0) = -1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < ScbFrame(2).Top Then
    ScrollTimer.Tag = ScrollRate * 2
    ScrollTimer.Interval = ScrollSpeed
Else
    ScrollTimer.Tag = (0 - (ScrollRate * 2))
    ScrollTimer.Interval = ScrollSpeed
End If
ImHighBool = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
Call ForceLooseFocus
LastXY(0) = X
LastXY(1) = Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ScrollTimer.Interval = 0
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 255
ScbFrame(1).Top = UserControl.Height - ScbFrame(1).Height

Call Process_CoOrdinates(0)
End Sub

Public Sub Process_CoOrdinates(ProcessType As Integer)
Dim NewCoOrdSubmit As Long
If PlayVal <= 0 Or GapVal <= 0 Then GoTo Ed
Slide_Len = UserControl.Height - ScbFrame(0).Height - ScbFrame(1).Height
If Slide_Len < 0 Then GoTo Ed
'If Slide_Len * (GapVal / PlayVal) < SixTeenPix Then
'    ScbFrame(2).Height = EightPix * 2
'Else
ScbFrame(2).Height = Slide_Len * (GapVal / PlayVal)
'End If
'If Slide_Len < ScbFrame(2).Height Then
'    ScbFrame(2).Visible = False
'Else
'    ScbFrame(2).Visible = True
'End If

If ValVal < 0 Then ValVal = 0
If ValVal > Slide_Len - ScbFrame(2).Height Then ValVal = Slide_Len - ScbFrame(2).Height
If ProcessType = 0 Then If LastVal = ValVal Then GoTo Ed
LastVal = ValVal

UserControl.Refresh
If (GapVal - (UserControl.Height - Slide_Len)) > 0 Then NewCoOrdSubmit = (PlayVal / (GapVal - (UserControl.Height - Slide_Len))) * ValVal  'Int((ValVal / PlayVal) * GapVal) - OnePix
RaiseEvent SliderDrag(NewCoOrdSubmit)
Ed: If ScbFrame(2).Top <> ScbFrame(0).Height + ValVal Or ScbFrame(2).Height <> ScBEffect(2).Height Then
    ScbFrame(2).Top = ScbFrame(0).Height + ValVal
    ScBEffect(2).Height = ScbFrame(2).Height
    ScBHotSpot(2).Height = ScbFrame(2).Height
    ScBArrow(2).Top = (ScbFrame(2).Height / 2) - (ScBArrow(2).Height / 2)
End If
End Sub

Public Sub ForceLooseFocus()
ScBEffect(0).FillColor = Theme_Dark
ScBEffect(1).FillColor = Theme_Dark
ScBEffect(2).FillColor = Theme_Color
ScbFrame(2).BackColor = Theme_Color
ImHighBool = False
End Sub

Public Sub ResetControl()
Call ForceLooseFocus
ScBEffect(0).BorderColor = Theme_Shade
ScBEffect(1).BorderColor = Theme_Shade
ScBEffect(2).BorderColor = Theme_Shade
ScBArrow(0).Picture = Manager.PictureLoader(2).ListImages.Item(5).Picture
ScBArrow(1).Picture = Manager.PictureLoader(2).ListImages.Item(4).Picture
ScBArrow(2).Picture = Manager.PictureLoader(2).ListImages.Item(6).Picture
ScbFrame(2).BackColor = Theme_Color
UserControl.BackColor = Theme_Shade
End Sub

Public Sub Submit_New_Coord(NewCoord As Long)
ValVal = NewCoord / (PlayVal / (GapVal - (UserControl.Height - Slide_Len)))
Call Process_CoOrdinates(1)
End Sub

Public Property Let Gap(ByVal GapValue As Long)
GapVal = GapValue
Call Check_Rate_Value
End Property
Public Property Get Gap() As Long
Gap = GapVal
End Property
Public Property Get VeryBottom() As Boolean
If ScbFrame(2).Top + ScbFrame(2).Height + SixPix > ScbFrame(1).Top Then
    VeryBottom = True
Else
    VeryBottom = False
End If
End Property
Public Property Let Play(ByVal PlayValue As Long)
PlayVal = PlayValue
Call Check_Rate_Value
End Property
Public Property Get Play() As Long
Play = PlayVal
End Property
Public Property Let Value(ByVal ValValue As Long)
ValVal = ValValue
End Property
Public Property Get Value() As Long
Value = ValVal
End Property

Private Sub Check_Rate_Value()
If PlayVal <> 0 And GapVal <> 0 Then ScrollRate = Int((EightPix / PlayVal) * GapVal)
End Sub

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
End Property
Public Property Get ImHigh() As Boolean
ImHigh = ImHighBool
End Property

