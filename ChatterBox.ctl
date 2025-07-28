VERSION 5.00
Begin VB.UserControl ChatterBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5415
   ScaleWidth      =   4455
   Begin SynthMark_XP.VScrollButton VScrollButton 
      Height          =   5265
      Left            =   4185
      Top             =   15
      Width           =   255
      _extentx        =   450
      _extenty        =   9287
   End
   Begin VB.Shape CtBSelector 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   10  'Mask Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape CtBEffect 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00F0F0F0&
      Height          =   5265
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   4425
   End
   Begin VB.Shape CtBEffect 
      BorderColor     =   &H00808080&
      Height          =   5295
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label CtBText 
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
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   45
   End
   Begin VB.Label CtBText 
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
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "ChatterBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim UseableArea As Integer, ImHighBool As Boolean, OneLineHeight As Integer, LastWidth As Integer, SelectIndex As Integer, StickOnce As Boolean
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()

Public Sub ResetControl()
Call VScrollButton.ResetControl
CtBEffect(0).BorderColor = Theme_Color
CtBEffect(1).BorderColor = Theme_Shade
CtBText(0).Font = Theme_Font
CtBText(0).FontBold = True
CtBText(1).Font = Theme_Font
CtBText(0).ForeColor = Theme_Shadow
CtBText(1).ForeColor = Theme_Shadow
CtBText(0).Top = FourPix
CtBText(1).Top = FourPix
UserControl.FillColor = Theme_Light
CtBSelector.BorderColor = Theme_High
CtBSelector.FillColor = Theme_HighLight
OneLineHeight = WordHieght(Theme_Text, False)
CtBSelector.Height = OneLineHeight + OnePix
Call Align_Columns
End Sub

Private Sub CtBText_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub CtBText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SelectIndex = Int(Y / OneLineHeight)
Call Select_Follower
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CtBText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
Call Align_Columns
End Sub

Private Sub UserControl_GotFocus()
CtBSelector.Visible = True
End Sub

Private Sub UserControl_LostFocus()
CtBSelector.Visible = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call VScrollButton.ForceLooseFocus
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Update_UseableArea()
If CtBText(0).Height > UserControl.ScaleHeight - FourPix Then
    UseableArea = UserControl.Width - VScrollButton.Width - FivePix
    If VScrollButton.Visible = False Then
        VScrollButton.Visible = True
        StickOnce = True
    End If
Else
    UseableArea = UserControl.Width - FourPix
    VScrollButton.Visible = False
End If
End Sub

Public Sub Submit_Data_Array(DatArray() As String)
If Is_Array_Empty(DatArray(), 2) = True Then GoTo Ed
Dim ColStringTempName As String, ColStringTemp As String, ColWrapp As TEXTWRAPPER
For UserCount = 0 To UBound(DatArray(), 2)
    ColWrapp = WordWrapper(DatArray(1, UserCount), UseableArea - CtBText(0).Width - SixTeenPix, Theme_Font, False)
    ColStringTemp = ColStringTemp & ColWrapp.WrapText & Chr(13)
    ColStringTempName = ColStringTempName & DatArray(0, UserCount) & String(ColWrapp.WrapLines, Chr(13))
Next UserCount
ColStringTempName = Left(ColStringTempName, Len(ColStringTempName) - 1)
ColStringTemp = Left(ColStringTemp, Len(ColStringTemp) - 1)
CtBText(0).Caption = ColStringTempName
CtBText(1).Caption = ColStringTemp
Call Align_Columns
Ed: End Sub

Public Sub Align_Columns()
Call Update_UseableArea
CtBText(1).Left = CtBText(0).Width + SixTeenPix
If VScrollButton.Visible = True Then
    VScrollButton.Play = CtBText(0).Height
    VScrollButton.Gap = UserControl.Height - SixPix
    If VScrollButton.VeryBottom = True Then
        Call VScrollButton.Submit_New_Coord(CtBText(0).Height)
    Else
        Call VScrollButton.Process_CoOrdinates(0)
    End If
    If StickOnce = True Then
        Call VScrollButton.Submit_New_Coord(CtBText(0).Height)
        StickOnce = False
    End If
Else
    For UserCount = 0 To CtBText.Count - 1
        CtBText(UserCount).Top = UserControl.Height - CtBText(UserCount).Height - FourPix
    Next UserCount
End If
Call Select_Follower
End Sub

Private Sub UserControl_Resize()
If UserControl.Width < VScrollButton.Width Then UserControl.Width = VScrollButton.Width
On Error GoTo TooSmall1
Dim ScrollWidth As Integer
Call Update_UseableArea
If VScrollButton.Visible = True Then ScrollWidth = VScrollButton.Width

CtBEffect(0).Move 0, 0, UserControl.Width, UserControl.Height
CtBEffect(1).Move OnePix, OnePix, UserControl.Width - TwoPix, UserControl.Height - TwoPix
CtBSelector.Left = ThreePix
CtBSelector.Width = UseableArea - TwoPix
VScrollButton.Move UserControl.Width - VScrollButton.Width - ThreePix, ThreePix
VScrollButton.Height = UserControl.Height - SixPix
If NowLoading = False Then Call Align_Columns
GoTo Ed

TooSmall1: Resume Ed
Ed: End Sub

Public Sub Do_Resize()
Call UserControl_Resize
End Sub

Private Sub Select_Follower()
CtBSelector.Top = CtBText(0).Top + (SelectIndex * OneLineHeight)
If CtBSelector.Top > CtBText(0).Top + CtBText(0).Height Then
    CtBSelector.Top = CtBText(0).Top + CtBText(0).Height - CtBSelector.Height
    SelectIndex = Int(CtBSelector.Top / OneLineHeight)
End If
End Sub

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = False Then
    Call VScrollButton.ForceLooseFocus
End If
End Property
Public Property Get ImHigh() As Boolean
If VScrollButton.ImHigh = True Then
    ImHigh = True
    GoTo Ed
End If
ImHigh = ImHighBool
Ed: End Property

Private Sub UserControl_Show()
If NowLoading = False Then Call Align_Columns
End Sub

Private Sub VScrollButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImHighBool = True
End Sub

Private Sub VScrollButton_SliderDrag(NowValue As Long)
CtBText(0).Top = ThreePix - NowValue
CtBText(1).Top = ThreePix - NowValue
Call Select_Follower
End Sub
