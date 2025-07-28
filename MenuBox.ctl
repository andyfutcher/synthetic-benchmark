VERSION 5.00
Begin VB.UserControl MenuBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ScaleHeight     =   3300
   ScaleWidth      =   3135
   Begin VB.PictureBox MenuBoxFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   60
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   10
      Top             =   15
      Width           =   3015
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Tag             =   "0"
         Top             =   240
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Tag             =   "0"
         Top             =   0
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Tag             =   "0"
         Top             =   480
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   3
         Tag             =   "0"
         Top             =   720
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   4
         Left            =   0
         TabIndex        =   4
         Tag             =   "0"
         Top             =   960
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Tag             =   "0"
         Top             =   1200
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   6
         Left            =   0
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1440
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1680
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   8
         Left            =   0
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1920
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   9
         Left            =   0
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2160
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
   End
   Begin VB.Image FrameImage 
      Height          =   3135
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "MenuBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ImHighBool As Boolean, LastMenuNo As Integer, UserIndex As Integer, LastExeption As Integer
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click(ClickIndex As Integer, ClickType As Integer, MenuCmdID As Integer)
Public Event HideMe()

Public Sub Setup_Menu(Index As Integer, CommandID As Integer, MenuCaption As String, ImageNumber As Integer, ParentLevel As Boolean, NewCategory As Integer, ItemEnabled() As Boolean, StayOpn As Boolean)
MenuList(Index).InUse = True
If MenuCaption <> Empty_Code Then MenuList(Index).Caption = MenuCaption
MenuList(Index).StayOpen = StayOpn
MenuList(Index).IsParent = ParentLevel
MenuList(Index).CommandID = CommandID
If UBound(ItemEnabled) <> 0 Then
    MenuList(Index).Enabled = ItemEnabled(CommandID)
Else
    MenuList(Index).Enabled = True
End If
If ImageNumber = -1 Then
    Call Check_Selection_Value(Index, CommandID)
Else
    MenuList(Index).MenuIcon = ImageNumber
End If
MenuList(Index).Tag = NewCategory
MenuList(Index).Visible = True
LastMenuNo = Index
End Sub

Public Sub Align_MenuItems()
Dim MaximumWidth As Integer
MaximumWidth = 0

For MenuCount = 1 To MenuList.Count - 1
    If MenuList(MenuCount).Tag = 1 Then
        MenuList(MenuCount).Top = MenuList(MenuCount - 1).Top + MenuList(MenuCount - 1).Height + (1 * Screen.TwipsPerPixelY)
    Else
        MenuList(MenuCount).Top = MenuList(MenuCount - 1).Top + MenuList(MenuCount - 1).Height
    End If
Next MenuCount
For MenuCount = 0 To MenuList.Count - 1
    If MenuList(MenuCount).SuggestedWidth > MaximumWidth Then MaximumWidth = MenuList(MenuCount).SuggestedWidth
Next MenuCount
For MenuCount = 0 To MenuList.Count - 1
    MenuList(MenuCount).Width = MaximumWidth
Next MenuCount
MenuBoxFrame.Width = MaximumWidth
UserControl.Width = MenuBoxFrame.Left + MenuBoxFrame.Width + Screen.TwipsPerPixelY

For MenuCount = 0 To MenuList.Count - 1
    If MenuList(MenuCount).InUse = True Then
        MenuBoxFrame.Height = MenuList(MenuCount).Top + MenuList(MenuCount).Height
    End If
Next MenuCount
UserControl.Height = MenuBoxFrame.Height + (2 * Screen.TwipsPerPixelY)
FrameImage(0).Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Public Sub Clear_Entire_List()
For MenuCount = 0 To MenuList.Count - 1
    MenuList(MenuCount).InUse = False
    MenuList(MenuCount).Visible = False
    MenuList(MenuCount).Loose_Focus
    MenuList(MenuCount).Caption = Empty_Code
Next MenuCount
LastMenuNo = 0
End Sub

Public Function Give_Menu_Item_Top(Index As Integer) As Integer
Give_Menu_Item_Top = MenuList(Index).Top
End Function

Public Sub Check_Selection_Value(Index As Integer, CommandID As Integer)
If BenchSelArray(CommandID - Bench_ID_Start) = True Then
    MenuList(Index).MenuIcon = Image_Checked
Else
    MenuList(Index).MenuIcon = Image_Unchecked
End If
End Sub

Private Sub Process_KeyBoard_Int(KeyCode As Integer)
Select Case KeyCode
Case vbKeyDown, vbKeyRight
    MenuList(0).SetFocus
Case vbKeyUp, vbKeyLeft
    MenuList(LastMenuNo).SetFocus
Case 27
    RaiseEvent HideMe
End Select
End Sub

Public Sub HideFocus()
'On Error Resume Next
Call Kill_StayOpenNow(-1)
MenuBoxFrame.SetFocus
End Sub

Public Sub Kill_StayOpenNow(Exeption As Integer)
If LastExeption = Exeption Then GoTo Ed
LastExeption = Exeption
For MenuCount = 0 To MenuList.Count - 1
    If MenuCount <> Exeption Then Call MenuList(MenuCount).Dont_Stay_Open
Next MenuCount
Ed: End Sub

Public Sub ResetControl(Index As Integer)
For UserCount = 0 To MenuList.Count - 1
    Call MenuList(UserCount).ResetControl
Next UserCount
FrameImage(0).Picture = LoadPicture(System_Path & Theme_Path & Theme_Current & Texture_Name & Zero_Code & Resource_Ext)
UserControl.BackColor = Theme_High
MenuBoxFrame.BackColor = Theme_Dark
UserIndex = Index
UserControl.Tag = -2
Call Clear_Entire_List
End Sub

Private Sub MenuBoxFrame_GotFocus()
Call Kill_StayOpenNow(-1)
End Sub

Private Sub MenuBoxFrame_KeyDown(KeyCode As Integer, Shift As Integer)
Call Process_KeyBoard_Int(KeyCode)
End Sub

Private Sub MenuList_CleanOthers(Index As Integer)
Call Kill_StayOpenNow(Index)
End Sub

Private Sub MenuList_Click(Index As Integer, ClickType As Integer, MenuCmdID As Integer)
RaiseEvent Click(Index, ClickType, MenuCmdID)
End Sub

Private Sub MenuList_ForceLooseFocus(Index As Integer)
RaiseEvent HideMe
End Sub

Private Sub MenuList_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then RaiseEvent HideMe
End Sub

Private Sub MenuList_MouseMoved(Index As Integer, MenuCmdID As Integer)
Call Normalize_Controls(-1)
If MenuList(Index).Enabled = True Then Call Show_ToolTip(MenuCmdID, UserIndex + 1)
End Sub

Private Sub UserControl_GotFocus()
MenuList(0).SetFocus
End Sub
