VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.UserControl GraphBox 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5295
   ScaleWidth      =   6255
   Begin MSChart20Lib.MSChart GrBChart 
      Height          =   5235
      Left            =   30
      OleObjectBlob   =   "GraphBox.ctx":0000
      TabIndex        =   0
      Top             =   30
      Width           =   5715
   End
   Begin VB.Shape GrBEffect 
      BorderColor     =   &H00808080&
      Height          =   5295
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Shape GrBEffect 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5265
      Index           =   1
      Left            =   15
      Top             =   15
      Width           =   5745
   End
End
Attribute VB_Name = "GraphBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ThreeDMode As Boolean, LineMode As Boolean, AllowLedgend As Boolean
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()


Public Sub ResetControl()
GrBEffect(0).BorderColor = Theme_Color
GrBEffect(1).BorderColor = Theme_Shade
GrBEffect(1).FillColor = Theme_Vague
GrBChart.Backdrop.Fill.Style = VtFillStyleBrush
Call GrBChart.Backdrop.Fill.Brush.FillColor.Set(Theme_Vague_R, Theme_Vague_G, Theme_Vague_B)
'GrBChart.Backdrop.Frame.Style = VtFrameStyleNull
'GrBChart.Backdrop.Frame.Width = 0
AllowLedgend = True
Call Select_Mode
End Sub

Public Sub EditCopyNow()
GrBChart.EditCopy
End Sub
Public Sub Submit_Data_Array(TargetArray() As Variant)
GrBChart.ChartData = TargetArray()
End Sub

Public Property Let ThreeD(ByVal ThreeDValue As Boolean)
ThreeDMode = ThreeDValue
Call Select_Mode
End Property
Public Property Get ThreeD() As Boolean
ThreeD = ThreeDMode
End Property
Public Property Let ShowLedgend(ByVal LedgendValue As Boolean)
AllowLedgend = LedgendValue
Call Select_Mode
End Property
Public Property Get ShowLedgend() As Boolean
ShowLedgend = AllowLedgend
End Property
Public Property Let LineGraph(ByVal LineValue As Boolean)
LineMode = LineValue
Call Select_Mode
End Property
Public Property Get LineGraph() As Boolean
LineGraph = LineMode
End Property

Private Sub Select_Mode()
If ThreeDMode = False And LineMode = False Then GrBChart.chartType = VtChChartType2dBar
If ThreeDMode = True And LineMode = False Then GrBChart.chartType = VtChChartType3dBar
If ThreeDMode = False And LineMode = True Then GrBChart.chartType = VtChChartType2dLine
If ThreeDMode = True And LineMode = True Then GrBChart.chartType = VtChChartType3dLine
GrBChart.ShowLegend = AllowLedgend
End Sub

Private Sub GrBChart_Click()
RaiseEvent Click
End Sub

Private Sub GrBChart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
GrBEffect(0).Move 0, 0, UserControl.Width, UserControl.Height
GrBEffect(1).Move OnePix, OnePix, UserControl.Width - TwoPix, UserControl.Height - TwoPix
GrBChart.Move TwoPix, TwoPix, UserControl.Width - FourPix, UserControl.Height - FourPix
End Sub
