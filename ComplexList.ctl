VERSION 5.00
Begin VB.UserControl ComplexList 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5460
   ScaleWidth      =   4455
   Begin SynthMark_XP.VScrollButton VScrollButton 
      Height          =   5270
      Left            =   4185
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9287
   End
   Begin VB.Frame CxLHeader 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4155
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHotSpot 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   6
         Top             =   15
         Width           =   45
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   30
         TabIndex        =   5
         Top             =   15
         Width           =   45
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   4
         Top             =   15
         Width           =   45
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   3
         Top             =   15
         Width           =   45
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   2
         Top             =   15
         Width           =   45
      End
      Begin VB.Label CxLColHead 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   15
         Width           =   45
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape CxLColShape 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image CxLHeadImage 
         Height          =   240
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Shape CxLEffect 
      BorderColor     =   &H00808080&
      Height          =   5295
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape CxLEffect 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFFFFF&
      Height          =   5265
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   4425
   End
   Begin VB.Shape CxLSelector 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   10  'Mask Pen
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   45
      Top             =   285
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLColText 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   45
   End
   Begin VB.Label CxLHotSpot 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   5055
      Left            =   0
      TabIndex        =   19
      Top             =   275
      Width           =   3405
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape CxLTextShape 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   1
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   2
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   3
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   4
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
   Begin VB.Line CxLColLine 
      BorderColor     =   &H00F0F0F0&
      BorderStyle     =   3  'Dot
      Index           =   5
      Visible         =   0   'False
      X1              =   30
      X2              =   0
      Y1              =   0
      Y2              =   5160
   End
End
Attribute VB_Name = "ComplexList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ColRelativ(5) As Integer, UseableArea As Integer, ImHighBool As Boolean, OneLineHeight As Integer, LastWidth As Integer
Dim SelectIndex As Integer, DataIndex() As String, KeepFocusVal As Boolean, MouseNowDown As Boolean
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DoubleClick(Selection As Integer)
Public Event HeaderClick(HeaderIndex As Integer)
Public Event Resized()
Public Event Click()

Public Sub ResetControl()
Call VScrollButton.ResetControl
CxLEffect(0).BorderColor = Theme_Color
CxLEffect(1).BorderColor = Theme_Shade
UserControl.BackColor = Theme_Vague
CxLHeader(0).BackColor = Theme_Shade
For UserCount = 0 To CxLColHead.Count - 1
    CxLColHead(UserCount).ForeColor = Theme_Pitch
    'CxLColShape(UserCount).BorderColor = Theme_Shade
    CxLColShape(UserCount).FillColor = Theme_Shade
    'CxLTextShape(UserCount).FillColor = Theme_Shade
    CxLColText(UserCount).Top = 19 * OnePix
    CxLColHead(UserCount).Font = Theme_Text
    CxLColText(UserCount).Font = Theme_Text
    OneLineHeight = WordHieght(Theme_Text, False)
Next UserCount
For UserCount = 0 To CxLColLine.Count - 1
    CxLColLine(UserCount).BorderColor = Theme_Color
Next UserCount
CxLHeadImage.Picture = Manager.PictureLoader(0).ListImages.Item(1).Picture
CxLSelector.Height = OneLineHeight + OnePix
End Sub

Private Sub CxLColHotSpot_Click(Index As Integer)
'RaiseEvent Click
RaiseEvent HeaderClick(Index)
End Sub

Private Sub CxLColHotSpot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
MouseNowDown = True
Do While MouseNowDown = True
    If Button = 1 Then
        CxLColHead(Index).Tag = CxLColHead(Index).Tag + 1
    Else
        CxLColHead(Index).Tag = CxLColHead(Index).Tag - 1
    End If
    If CxLColHead(Index).Tag < 100 Then CxLColHead(Index).Tag = 100
    If CxLColHead(Index).Tag > 1000 Then CxLColHead(Index).Tag = 1000
    DoEvents
    LastWidth = -1
    Call Align_Columns
Loop
End Sub

Private Sub CxLColHotSpot_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If CxLColShape(Index).Visible = False Then
    Call Clear_Col_Selection(Index)
    CxLColShape(Index).Visible = True
    'CxLTextShape(Index).Visible = True
    CxLColHead(Index).ForeColor = Theme_High
    Call VScrollButton.ForceLooseFocus
End If
ImHighBool = True
'RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Sub Clear_Col_Selection(Exeption As Integer)
For UserCount = 0 To CxLColLine.Count - 1
    If UserCount <> Exeption Then
        CxLColShape(UserCount).Visible = False
        'CxLTextShape(UserCount).Visible = False
        CxLColHead(UserCount).ForeColor = Theme_Pitch
    End If
Next UserCount
If Exeption <> -2 Then
    Call VScrollButton.ForceLooseFocus
End If
ImHighBool = False
End Sub

Private Sub CxLColHotSpot_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseNowDown = False
RaiseEvent Resized
End Sub

Private Sub CxLColText_Click(Index As Integer)
RaiseEvent Click
End Sub

Private Sub CxLColText_DblClick(Index As Integer)
RaiseEvent DoubleClick(SelectIndex)
End Sub

Private Sub CxLColText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SelectIndex = Int(Y / OneLineHeight)
Call Select_Follower
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CxLColText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Clear_Col_Selection(-1)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CxLHeader_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub CxLHotSpot_Click()
RaiseEvent Click
End Sub

Private Sub CxLHotSpot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Clear_Col_Selection(-1)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_EnterFocus()
Call Align_Columns
End Sub

Private Sub UserControl_GotFocus()
CxLSelector.Visible = True
CxLSelector.BorderColor = Theme_High
CxLSelector.FillColor = Theme_HighLight
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 38
    SelectIndex = SelectIndex - 1
    If SelectIndex < 0 Then SelectIndex = 0
    Call Select_Follower
Case 40
    SelectIndex = SelectIndex + 1
    If SelectIndex > Int((CxLColText(0).Height / OneLineHeight) - 1) Then SelectIndex = Int((CxLColText(0).Height / OneLineHeight) - 1)
    Call Select_Follower
End Select
End Sub

Private Sub UserControl_LostFocus()
If KeepFocusVal = True Then
    CxLSelector.BorderColor = Theme_Dark
    CxLSelector.FillColor = Theme_Color
Else
    CxLSelector.Visible = False
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub Update_UseableArea()
If CxLColText(0).Height > UserControl.ScaleHeight - CxLHeader(0).Height - FourPix Then
    UseableArea = UserControl.Width - VScrollButton.Width - FivePix
    VScrollButton.Visible = True
Else
    UseableArea = UserControl.Width - FivePix
    VScrollButton.Visible = False
End If
End Sub

Public Sub Setup_Cols_Headers(ColArray() As String)
CxLColHotSpot(0).Tag = -1
For UserCount = 0 To UBound(ColArray, 1)
    If ColArray(UserCount, 0) <> Empty_Code Then
        CxLColHead(UserCount).Caption = ColArray(UserCount, 0)
        CxLColHead(UserCount).Tag = ColArray(UserCount, 1) * 100
        CxLColHead(UserCount).Visible = True
        If UserCount > 0 Then CxLColLine(UserCount - 1).Visible = True
    Else
        CxLColHead(UserCount).Tag = 0
        CxLColHead(UserCount).Visible = False
        CxLColLine(UserCount).Visible = False
    End If
Next UserCount
Call Align_Columns
End Sub

Public Sub Submit_Data_Array(DatArray() As String, SectionNum As Integer, VisibleCols As Integer)
Call Align_Columns
Dim CurrentCol As Integer, ColStringTemp As String
If SectionNum <> -1 Then
    For CurrentCol = 0 To VisibleCols 'UBound(DatArray(), 2)
        For UserCount = 0 To UBound(DatArray(), 3)
            If DatArray(SectionNum, 0, UserCount) <> Empty_Code Then ColStringTemp = ColStringTemp & WordElipser(DatArray(SectionNum, CurrentCol, UserCount), CxLColText(CurrentCol).Tag, Theme_Text, False) & Chr(13)
        Next UserCount
        ColStringTemp = Left(ColStringTemp, Len(ColStringTemp) - 1)
        CxLColText(CurrentCol).Caption = ColStringTemp
        ColStringTemp = Empty_Code
    Next CurrentCol
Else
    For CurrentCol = 0 To VisibleCols 'UBound(DatArray(), 1)
        For UserCount = 0 To UBound(DatArray(), 2)
            If DatArray(0, UserCount) <> Empty_Code Then ColStringTemp = ColStringTemp & WordElipser(DatArray(CurrentCol, UserCount), CxLColText(CurrentCol).Tag, Theme_Text, False) & Chr(13)
        Next UserCount
        'Manager.Caption = DatArray(CurrentCol, UserCount - 1)
        ColStringTemp = Left(ColStringTemp, Len(ColStringTemp) - 1)
        CxLColText(CurrentCol).Caption = ColStringTemp
        ColStringTemp = Empty_Code
    Next CurrentCol
End If
Call Select_Follower
Ed: End Sub

Public Sub Empty_Data()
Dim CurrentCol As Integer
For CurrentCol = 0 To CxLColText.Count - 1
    CxLColText(CurrentCol).Caption = Empty_Code
Next CurrentCol
End Sub

Public Sub Align_Columns()
'On Error GoTo TooSmall
Call Update_UseableArea
If LastWidth = UseableArea Then GoTo Ed
Dim MaxRelativity As Integer, LeftRelativity As Integer, MaxShown As Integer
If LastWidth <> -1 Then CxLSelector.Width = UseableArea - OnePix
For UserCount = 0 To CxLColHead.Count - 1
    If CxLColHead(UserCount).Tag <> Empty_Code Then
        CxLColHead(UserCount).Visible = True
        CxLColHotSpot(UserCount).Visible = True
        MaxRelativity = MaxRelativity + CxLColHead(UserCount).Tag
        MaxShown = UserCount
    Else
        CxLColHead(UserCount).Visible = False
        CxLColHotSpot(UserCount).Visible = False
    End If
Next UserCount

For UserCount = 0 To MaxShown
    CxLColHead(UserCount).Width = (UseableArea / MaxRelativity) * CxLColHead(UserCount).Tag
    CxLColText(UserCount).Width = CxLColHead(UserCount).Width - TwoPix
    'CxLColText(UserCount).Tag = CxLColHead(UserCount).Width - EightPix
    CxLColHotSpot(UserCount).Width = CxLColHead(UserCount).Width + OnePix
    CxLColShape(UserCount).Width = CxLColHead(UserCount).Width + OnePix
    'CxLTextShape(UserCount).Width = CxLColHead(UserCount).Width + OnePix
    'If UserCount = 0 Then
        CxLColHead(UserCount).Left = LeftRelativity + TwoPix
        CxLColText(UserCount).Left = LeftRelativity + FivePix
        CxLColText(UserCount).Tag = CxLColText(UserCount).Width
        CxLColHotSpot(UserCount).Left = LeftRelativity
        CxLColShape(UserCount).Left = LeftRelativity
        'CxLTextShape(UserCount).Left = LeftRelativity
    'Else
    '    CxLColHead(UserCount).Left = LeftRelativity + ThreePix
    '    CxLColText(UserCount).Left = LeftRelativity + SixPix
    '    CxLColText(UserCount).Tag = CxLColText(UserCount).Width
    '    CxLColHotSpot(UserCount).Left = LeftRelativity + OnePix
    '    CxLColShape(UserCount).Left = LeftRelativity + OnePix
    '   'CxLTextShape(UserCount).Left = LeftRelativity + OnePix
    'End If
    LeftRelativity = LeftRelativity + CxLColHead(UserCount).Width + OnePix
    CxLColLine(UserCount).X1 = LeftRelativity + TwoPix
    CxLColLine(UserCount).X2 = LeftRelativity + TwoPix
Next UserCount
GoTo Ed

TooSmall: Resume Ed
Ed:
If LastWidth <> -1 Then
If VScrollButton.Visible = True Then
    VScrollButton.Play = CxLColText(0).Height
    VScrollButton.Gap = UserControl.Height - (CxLHeader(0).Top + CxLHeader(0).Height + FourPix)
    Call VScrollButton.Process_CoOrdinates(0)
Else
    For UserCount = 0 To CxLColHead.Count - 1
        CxLColText(UserCount).Top = ThreePix + CxLHeader(0).Height
    Next UserCount
End If
End If
LastWidth = UseableArea
If LastWidth <> -1 Then Call Select_Follower
End Sub

Private Sub UserControl_Resize()
If UserControl.Width < VScrollButton.Width Then UserControl.Width = VScrollButton.Width
On Error GoTo TooSmall1
Dim ScrollWidth As Integer
If VScrollButton.Visible = True Then ScrollWidth = VScrollButton.Width

CxLEffect(0).Move 0, 0, UserControl.Width, UserControl.Height
CxLEffect(1).Move OnePix, OnePix, UserControl.Width - TwoPix, UserControl.Height - TwoPix
CxLHeader(0).Move TwoPix, TwoPix, UserControl.Width - FourPix
CxLSelector.Left = ThreePix
CxLHotSpot.Move CxLHeader(0).Height + TwoPix, 0, UserControl.ScaleWidth, UserControl.Height
CxLHeadImage.Width = CxLHeader(0).Width
VScrollButton.Move UserControl.Width - VScrollButton.Width - TwoPix, CxLHeader(0).Top + CxLHeader(0).Height + OnePix
VScrollButton.Height = UserControl.Height - ThreePix - VScrollButton.Top

For UserCount = 0 To CxLColLine.Count - 1
    CxLColLine(UserCount).Y1 = TwoPix
    CxLColLine(UserCount).Y2 = UserControl.Height - TwoPix
Next UserCount
If NowLoading = False Then Call Align_Columns
RaiseEvent Resized
GoTo Ed

TooSmall1: Resume Ed
Ed: End Sub

Private Sub VScrollButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImHighBool = True
Call Clear_Col_Selection(-2)
'RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let ListIndex(ByVal IndexValue As Integer)
SelectIndex = IndexValue
Call Select_Follower
End Property
Public Property Get ListIndex() As Integer
ListIndex = SelectIndex
End Property

Public Property Let KeepFocus(ByVal FocusValue As Boolean)
KeepFocusVal = FocusValue
Call Select_Follower
CxLSelector.Visible = True
End Property
Public Property Get KeepFocus() As Boolean
KeepFocus = KeepFocusVal
End Property

Public Property Let ImHigh(ByVal ImHighValue As Boolean)
ImHighBool = ImHighValue
If ImHighValue = False Then
    Call Clear_Col_Selection(-1)
End If
End Property
Public Property Get ImHigh() As Boolean
If VScrollButton.ImHigh = True Then
    ImHigh = True
    GoTo Ed
End If
ImHigh = ImHighBool
Ed: End Property

Private Sub Select_Follower()
CxLSelector.Top = CxLColText(0).Top + (SelectIndex * OneLineHeight)
End Sub

Private Sub VScrollButton_SliderDrag(NowValue As Long)
For UserCount = 0 To CxLColHead.Count - 1
    CxLColText(UserCount).Top = (ThreePix + CxLHeader(0).Height) - NowValue
Next UserCount
Call Select_Follower
End Sub
