Attribute VB_Name = "ControlEffects"
Option Explicit
Global ExplorerFrame_Count As Integer, Short() As String

Public Sub Prepare_Controls()
If Theme_Current = Empty_Code Then
    Call Search_Themes(True)
Else
    Call Search_Themes(False)
End If
If Does_File_Exist(System_Path & Theme_Path & Theme_Current & Config_Name & Zero_Code & Resource_Ext) = False Then Call Search_Themes(True)
Call Load_Application_ThemeData
Dim ConfigBuffer As String
ConfigBuffer = Load_File_Into_String(System_Path & Theme_Path & Theme_Current & Config_Name & Zero_Code & Resource_Ext)
Call Process_Info_File(ConfigBuffer)

'Done loading
For ControlCount = 0 To Manager.ExplorerButton.Count - 1
    Call Manager.ExplorerButton(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.FormBackGround.Count - 1
    Manager.FormBackGround(ControlCount).Picture = Manager.PictureLoader(0).ListImages.Item(1).Picture
Next ControlCount
For ControlCount = 0 To Manager.MenuButton.Count - 1
    Call Manager.MenuButton(ControlCount).ResetControl
Next
For ControlCount = 0 To Manager.MenuBox.Count - 1
    Call Manager.MenuBox(ControlCount).ResetControl(ControlCount)
Next
For ControlCount = 0 To Manager.ExplorerFrame.Count - 1
    Call Manager.ExplorerFrame(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ControllerBox.Count - 1
    Call Manager.ControllerBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.SimpleBox.Count - 1
    Call Manager.SimpleBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ComplexList.Count - 1
    Call Manager.ComplexList(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ChatterBox.Count - 1
    Call Manager.ChatterBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.FrameList.Count - 1
    Call Manager.FrameList(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.CommandButton.Count - 1
    Call Manager.CommandButton(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ProgressBox.Count - 1
    Call Manager.ProgressBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.FrameButton.Count - 1
    Call Manager.FrameButton(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.WriteBox.Count - 1
    Call Manager.WriteBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ToolButton.Count - 1
    Call Manager.ToolButton(ControlCount).ResetControl
    Manager.ToolButton(ControlCount).Height = 525
Next ControlCount
For ControlCount = 0 To Manager.OtherTool.Count - 1
    Call Manager.OtherTool(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.StatusLabel.Count - 1
    Call Manager.StatusLabel(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.StylishLabel.Count - 1
    Call Manager.StylishLabel(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.ScoreLabel.Count - 1
    Call Manager.ScoreLabel(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.OptionBox.Count - 1
    Manager.OptionBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.VScrollButton.Count - 1
    Manager.VScrollButton(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.FormContainer.Count - 1
    Manager.FormContainer(ControlCount).BackColor = Theme_Shade
Next ControlCount
For ControlCount = 0 To Manager.GraphBox.Count - 1
    Call Manager.GraphBox(ControlCount).ResetControl
Next ControlCount
For ControlCount = 0 To Manager.FrameLabel.Count - 1
    Manager.FrameLabel(ControlCount).ForeColor = Theme_Pitch
    If Manager.FrameLabel(ControlCount).Tag = One_Code Then
        Manager.FrameLabel(ControlCount).Font = Theme_Font
    Else
        Manager.FrameLabel(ControlCount).Font = Theme_Text
    End If
Next ControlCount
For ControlCount = 0 To Manager.LinkLabel.Count - 1
    Manager.LinkLabel(ControlCount).Font = Theme_Text
    Manager.LinkLabel(ControlCount).ForeColor = Theme_High
    Manager.LinkLabel(ControlCount).MouseIcon = Manager.PictureLoader(2).ListImages.Item(1).Picture
Next ControlCount

For ControlCount = 0 To Manager.LightLabel.Count - 1
    If Manager.LightLabel(ControlCount).Tag = One_Code Then
        Manager.LightLabel(ControlCount).Font = Theme_Font
    Else
        Manager.LightLabel(ControlCount).Font = Theme_Text
    End If
Next ControlCount
Call Manager.PlatInfoCplxList(0).ResetControl
Call Manager.ExplorerHolder.ResetControl
Call Manager.ToolTipBox.ResetControl
Call Process_Flexible_Lables

Manager.FormSlider(0).Picture = Manager.PictureLoader(2).ListImages.Item(2).Picture
Manager.FormImage.Picture = Manager.PictureLoader(0).ListImages.Item(2).Picture
Manager.FormCaption(0).ForeColor = Theme_Light
Manager.FormCaption(0).Font = Theme_Font

For ControlCount = 0 To UBound(Can_Menu())
    Can_Menu(ControlCount) = True 'True
Next ControlCount

Manager.FormIcon.Picture = ManagerSub.Icon
Manager.Icon = ManagerSub.Icon
Manager.BackColor = Theme_Shadow
Manager.WriteBoxList.Font = Theme_Text
Manager.WriteBoxList.BackColor = Theme_Light
Manager.WriteBoxList.ForeColor = Theme_Pitch

'Manager.WaterMark(0).BackColor = Manager.BackColor
Manager.FormHeader(1).BackColor = Theme_Light
Manager.NormLine(2).BorderColor = Theme_Dark
Manager.NormLine(3).BorderColor = Theme_Dark
Manager.NormLine(4).BorderColor = Theme_Light
Manager.NormLine(5).BorderColor = Theme_Dark
Manager.NormLine(6).BorderColor = Theme_Shade
Manager.NormLine(7).BorderColor = Theme_Dark
Manager.NormLine(8).BorderColor = Theme_Shade
Manager.NormLine(9).BorderColor = Theme_Dark
Manager.NormLine(10).BorderColor = Theme_Dark
Manager.NormLine(11).BorderColor = Theme_Shade

ManagerSub.GenLabel.Font = Theme_Text
ManagerSub.GenLabel.FontBold = False
ManagerSub.GenLabel.Caption = "..."
DotXThreeWidth = ManagerSub.GenLabel.Width
DefPosHolder = EightPix
Text_Height = WordHieght(Theme_Text, False)
End Sub
Public Sub Normalize_Controls(Index As Integer)
If NowLoading = True Then GoTo Ed
Call Highlight_Explorer_Button(-1)
Call Highlight_Menu_Button(-1)
Call Highlight_Explorer_Header(-1)
Call Highlight_FormControl(-1)
Call Highlight_Tool_Button(-1)
Call Highlight_Other_Tool(-1)
Call Highlight_Command_Button(-1)
Call Highlight_Frame_Button(-1)
Call Highlight_Option_Box(-1)
'Call Highlight_WriteText_Box(-1)
Call Normalize_VScroll_Buttons
'DoEvents
'If Manager.WriteBoxList.Visible = True Then Manager.WriteBox(Manager.WriteBoxList.Tag).SetFocus
If Index <> 1 Then Call Normalize_ComplexList_Box
If Index <> 1 Then Call Normalize_FrameList_Box
If Index <> 2 Then Call Normalize_PlatInfoCplxList_Box
If Index <> 3 Then Call Normalize_ChatterBox_Box
If Index <> 4 Then Call Show_ToolTip(0, 0)
Ed: End Sub
Public Sub Normalize_On_Click(Index As Integer)
If Manager.WriteBoxList.Visible = True Then Manager.WriteBoxList.Visible = False
Call Hide_MenuBox(1)
Call Hide_MenuBox(0)
End Sub

Public Sub Normalize_Menu_Lists()
If Manager.MenuBox(0).Visible = True And Manager.MenuBox(1).Visible = False Then Manager.MenuBox(0).HideFocus
If Manager.MenuBox(1).Visible = True Then Manager.MenuBox(1).HideFocus
End Sub

Public Sub Normalize_ComplexList_Box()
For ControlCount = 0 To Manager.ComplexList.Count - 1
    If Manager.ComplexList(ControlCount).Visible = True Then
        If Manager.ComplexList(ControlCount).ImHigh = True Then
            Manager.ComplexList(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Normalize_ChatterBox_Box()
For ControlCount = 0 To Manager.ChatterBox.Count - 1
    If Manager.ChatterBox(ControlCount).Visible = True Then
        If Manager.ChatterBox(ControlCount).ImHigh = True Then
            Manager.ChatterBox(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Normalize_FrameList_Box()
For ControlCount = 0 To Manager.FrameList.Count - 1
    If Manager.FrameList(ControlCount).Visible = True Then
        If Manager.FrameList(ControlCount).ImHigh = True Then
            Manager.FrameList(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Normalize_PlatInfoCplxList_Box()
If Manager.PlatInfoCplxList(0).Visible = True Then
    If Manager.PlatInfoCplxList(0).ImHigh = True Then
        Manager.PlatInfoCplxList(0).ImHigh = False
    End If
End If
End Sub
Public Sub Normalize_VScroll_Buttons()
For ControlCount = 0 To Manager.VScrollButton.Count - 1
    If Manager.VScrollButton(ControlCount).ImHigh = True Then Call Manager.VScrollButton(ControlCount).ForceLooseFocus
Next ControlCount
End Sub

Public Sub Highlight_Explorer_Button(Index As Integer)
For ControlCount = 0 To Manager.ExplorerButton.Count - 1
    If ControlCount = Index Then
        If Manager.ExplorerButton(ControlCount).ImHigh = False Then
            Manager.ExplorerButton(ControlCount).ImHigh = True
        End If
    Else
        If Manager.ExplorerButton(ControlCount).ImHigh = True Then
            Manager.ExplorerButton(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Explorer_Header(Index As Integer)
For ControlCount = 0 To Manager.ExplorerFrame.Count - 1
    If ControlCount = Index Then
        If Manager.ExplorerFrame(ControlCount).ImHigh = False Then
            Manager.ExplorerFrame(ControlCount).ImHigh = True
        End If
    Else
        If Manager.ExplorerFrame(ControlCount).ImHigh = True Then
            Manager.ExplorerFrame(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_FormControl(Index As Integer)
For ControlCount = 0 To Manager.FormControl.Count - 1
    If ControlCount = Index Then
        If Manager.FormControl(ControlCount).Tag = 0 Then
            Manager.FormControl(ControlCount).Picture = Manager.PictureLoader(1).ListImages.Item(ControlCount + 3).Picture
            Manager.FormControl(ControlCount).Tag = 1
        End If
    Else
        If Manager.FormControl(ControlCount).Tag = 1 Then
            Manager.FormControl(ControlCount).Picture = Manager.PictureLoader(0).ListImages.Item(ControlCount + 3).Picture
            Manager.FormControl(ControlCount).Tag = 0
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Menu_Button(Index As Integer)
For ControlCount = 0 To Manager.MenuButton.Count - 1
    If ControlCount = Index Then
        If Manager.MenuButton(ControlCount).ImHigh = False Then
            Manager.MenuButton(ControlCount).ImHigh = True
        End If
    Else
        If Manager.MenuBox(0).Tag <> ControlCount Then
        If Manager.MenuButton(ControlCount).ImHigh = True Then
            Manager.MenuButton(ControlCount).ImHigh = False
        End If
        End If
    End If
Next ControlCount
Ed: End Sub
Public Sub Highlight_Tool_Button(Index As Integer)
For ControlCount = 0 To Manager.ToolButton.Count - 1
    If ControlCount = Index Then
        If Manager.ToolButton(ControlCount).ImHigh = False Then
            Manager.ToolButton(ControlCount).ImHigh = True
        End If
    Else
        If Manager.ToolButton(ControlCount).ImHigh = True Then
            Manager.ToolButton(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Other_Tool(Index As Integer)
For ControlCount = 0 To Manager.OtherTool.Count - 1
    If ControlCount = Index Then
        If Manager.OtherTool(ControlCount).ImHigh = False Then
            Manager.OtherTool(ControlCount).ImHigh = True
        End If
    Else
        If Manager.OtherTool(ControlCount).ImHigh = True Then
            Manager.OtherTool(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Command_Button(Index As Integer)
For ControlCount = 0 To Manager.CommandButton.Count - 1
    If ControlCount = Index Then
        If Manager.CommandButton(ControlCount).ImHigh = False Then
            Manager.CommandButton(ControlCount).ImHigh = True
        End If
    Else
        If Manager.CommandButton(ControlCount).ImHigh = True Then
            Manager.CommandButton(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Frame_Button(Index As Integer)
For ControlCount = 0 To Manager.FrameButton.Count - 1
    If ControlCount = Index Then
        If Manager.FrameButton(ControlCount).ImHigh = False Then
            Manager.FrameButton(ControlCount).ImHigh = True
        End If
    Else
        If Manager.FrameButton(ControlCount).ImHigh = True Then
            Manager.FrameButton(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub
Public Sub Highlight_Option_Box(Index As Integer)
For ControlCount = 0 To Manager.OptionBox.Count - 1
    If ControlCount = Index Then
        If Manager.OptionBox(ControlCount).ImHigh = False Then
            Manager.OptionBox(ControlCount).ImHigh = True
        End If
    Else
        If Manager.OptionBox(ControlCount).ImHigh = True Then
            Manager.OptionBox(ControlCount).ImHigh = False
        End If
    End If
Next ControlCount
End Sub

Public Sub Prepare_Language()
Dim PlatInfoColumns(5, 1) As String, ServerListColumns(1, 1) As String, WaitingRoomColumns(1, 1) As String
Dim TierInfoColumns(5, 1) As String, URLBenchListColumns(1, 1) As String, BenchmarkTasksColumns(2, 1) As String
Dim ATMColumns(1, 1) As String, ConnStatusColumns(1, 1) As String, ComunityInfoColumns(2, 1) As String

ReDim Short(0)
ReDim TipTextArray(0)
ReDim Language(0)
TipTextIndex = 0
If Language_Current = Empty_Code Then
    Call Search_Language(True)
Else
    Call Search_Language(False)
End If
If Does_File_Exist(Application_Path & Lang_Path & Language_Current & Config_Name & Zero_Code & Resource_Ext) = False Then Call Search_Themes(True)

Dim LangBuffer() As String, LangCount As Integer, LangIndex As Integer, LangName As String, LangText As String
Call Compile_Info_File(Application_Path & Lang_Path & Language_Current & Config_Name & Zero_Code & Resource_Ext, LangBuffer())
For LangCount = 0 To UBound(LangBuffer(), 2)
    LangIndex = Filter_Sort(LangBuffer(0, LangCount))
    LangName = LCase(Replace(LangBuffer(0, LangCount), LangIndex, Empty_Code))
    LangText = LangBuffer(1, LangCount)
    Select Case LangName
    Case "lang"
        Call Add_Index_To_StringList(Language())
        Language(LangIndex) = LangText
    Case "short"
        Call Add_Index_To_StringList(Short())
        Short(LangIndex) = LangText
    Case "lighl"
        Manager.LightLabel(LangIndex).Caption = LangText
    Case "frmlb"
        Manager.FrameLabel(LangIndex).Caption = Filter_Html(LangText)
    Case "lnklb"
        Manager.LinkLabel(LangIndex).Caption = Filter_Html(LangText)
    Case "menub"
        Manager.MenuButton(LangIndex).Caption = LangText
    Case "expfr"
        Manager.ExplorerFrame(LangIndex).Tag = LangText
    Case "optbx"
        Manager.OptionBox(LangIndex).Caption = LangText
    Case "cmdbt"
        Manager.CommandButton(LangIndex).Caption = LangText
    Case "tltpa"
        ToolTipArray(LangIndex) = LangText
    Case "benaa"
        BenchDatArray(LangIndex, 0) = LangText
    Case "benab"
        BenchDatArray(LangIndex, 1) = LangText
    Case "benac"
        BenchDatArray(LangIndex, 2) = LangText
    Case "benad"
        BenchDatArray(LangIndex, 3) = LangText
    Case "prgtip"
        Call Add_Index_To_StringList(TipTextArray())
        Call Format_Text_String(LangText)
        TipTextArray(UBound(TipTextArray)) = LangText
    End Select
Next LangCount

Manager.FrameLabel(7).Caption = Short(117)
Manager.FrameLabel(55).Caption = Short(118)
Manager.FrameLabel(64).Caption = Short(119)

Call Manager.ExplorerFrame(0).FrameProperty(Short(0), 47)
Call Manager.ExplorerFrame(1).FrameProperty(Short(1), 48)
Call Manager.ExplorerFrame(2).FrameProperty(Short(2), 50)
Call Manager.ExplorerFrame(3).FrameProperty(Short(3), 52)
Call Manager.ExplorerFrame(4).FrameProperty(Short(4), 49)
Call Manager.ExplorerFrame(5).FrameProperty(Short(5), 53)
Call Manager.ExplorerFrame(6).FrameProperty(Short(6), 54)
Call Manager.ExplorerFrame(7).FrameProperty(Short(7), 55)
Call Manager.ExplorerFrame(8).FrameProperty(Short(8), 51)

PlatInfoColumns(0, 0) = Short(9)
PlatInfoColumns(1, 0) = Short(10)
PlatInfoColumns(2, 0) = Short(11)
PlatInfoColumns(3, 0) = Short(12)
PlatInfoColumns(4, 0) = Short(13)
PlatInfoColumns(5, 0) = Short(14)
PlatInfoColumns(0, 1) = 2
PlatInfoColumns(1, 1) = 3
PlatInfoColumns(2, 1) = 3
PlatInfoColumns(3, 1) = 2
PlatInfoColumns(4, 1) = 2
PlatInfoColumns(5, 1) = 4

TierInfoColumns(0, 0) = Short(15)
TierInfoColumns(1, 0) = Short(16)
TierInfoColumns(2, 0) = Short(17)
TierInfoColumns(3, 0) = Short(18)
TierInfoColumns(4, 0) = Short(19)
TierInfoColumns(5, 0) = Short(20)
TierInfoColumns(0, 1) = 3
TierInfoColumns(1, 1) = 2
TierInfoColumns(2, 1) = 2
TierInfoColumns(3, 1) = 2
TierInfoColumns(4, 1) = 2
TierInfoColumns(5, 1) = 3

ServerListColumns(0, 0) = Short(21)
ServerListColumns(1, 0) = Short(22)
ServerListColumns(0, 1) = 2
ServerListColumns(1, 1) = 1

WaitingRoomColumns(0, 0) = Short(15)
WaitingRoomColumns(1, 0) = Short(23)
WaitingRoomColumns(0, 1) = 5
WaitingRoomColumns(1, 1) = 2

URLBenchListColumns(0, 0) = Short(24)
URLBenchListColumns(1, 0) = Short(25)
URLBenchListColumns(0, 1) = 4
URLBenchListColumns(1, 1) = 1

BenchmarkTasksColumns(0, 0) = Short(26)
BenchmarkTasksColumns(1, 0) = Short(27)
BenchmarkTasksColumns(2, 0) = Short(28)
BenchmarkTasksColumns(0, 1) = 4
BenchmarkTasksColumns(1, 1) = 1
BenchmarkTasksColumns(2, 1) = 2

ComunityInfoColumns(0, 0) = Short(113)
ComunityInfoColumns(1, 0) = Short(114)
ComunityInfoColumns(2, 0) = Short(115)
ComunityInfoColumns(0, 1) = 4
ComunityInfoColumns(1, 1) = 3
ComunityInfoColumns(2, 1) = 1

ATMColumns(0, 0) = Short(29)
ATMColumns(1, 0) = Short(30)
ATMColumns(0, 1) = 5
ATMColumns(1, 1) = 1

ConnStatusColumns(0, 0) = Short(31)
ConnStatusColumns(1, 0) = Short(27)
ConnStatusColumns(0, 1) = 5
ConnStatusColumns(1, 1) = 2

Call Manager.WriteBox(0).WriteProperty("fh4", False, False, False)
Call Manager.WriteBox(3).WriteProperty(Empty_Code, False, False, True)
Call Manager.WriteBox(4).WriteProperty("cb4", True, False, False)
Call Manager.WriteBox(5).WriteProperty(Empty_Code, False, True, False)
Call Manager.WriteBox(6).WriteProperty("cb6", False, False, False)
Call Manager.WriteBox(10).WriteProperty(Empty_Code, False, False, True)
Call Manager.WriteBox(11).WriteProperty("cb2", True, False, False)
Call Manager.WriteBox(12).WriteProperty("cb2", True, False, False)
Call Manager.WriteBox(13).WriteProperty(Empty_Code, False, False, True)
Call Manager.WriteBox(14).WriteProperty("sb0", True, False, False)
Call Manager.WriteBox(16).WriteProperty("cb14", True, False, False)
Call Manager.WriteBox(17).WriteProperty("cb14", True, False, False)
Call Manager.WriteBox(18).WriteProperty("cb14", True, False, False)
Call Manager.WriteBox(21).WriteProperty("cb3", True, False, False)
Call Manager.WriteBox(22).WriteProperty(Empty_Code, False, False, False)
Call Manager.WriteBox(23).WriteProperty(Empty_Code, False, False, False)
Call Manager.WriteBox(24).WriteProperty("cb16", True, False, False)
Call Manager.WriteBox(26).WriteProperty(Empty_Code, False, False, True)

Call Manager.WriteBox(6).DDList_Clear
Call Manager.WriteBox(14).DDList_Clear
Call Manager.WriteBox(16).DDList_Clear
Call Manager.WriteBox(17).DDList_Clear
Call Manager.WriteBox(18).DDList_Clear
Call Manager.WriteBox(15).DDList_Clear

Call Manager.WriteBox(6).DDList_Add(App_UserName, App_UserName)
Call Manager.WriteBox(6).DDList_Add(App_CompName, App_CompName)
Call Manager.WriteBox(14).DDList_Add(Short(32), Zero_Code)
Call Manager.WriteBox(14).DDList_Add(Short(33), One_Code)
Call Manager.WriteBox(16).DDList_Add(Short(34), "20")
Call Manager.WriteBox(16).DDList_Add(Short(35), "50")
Call Manager.WriteBox(16).DDList_Add(Short(36), "100")
Call Manager.WriteBox(17).DDList_Add(Short(37), "9")
Call Manager.WriteBox(17).DDList_Add(Short(38), "7")
Call Manager.WriteBox(17).DDList_Add(Short(39), "4")
Call Manager.WriteBox(17).DDList_Add(Short(40), One_Code)
Call Manager.WriteBox(18).DDList_Add(Short(41), Zero_Code)
Call Manager.WriteBox(18).DDList_Add(Short(42), One_Code)
Call Manager.WriteBox(18).DDList_Add(Short(43), "2")
For ControlCount = 0 To 15
    Call Manager.WriteBox(24).DDList_Add(ControlCount + 1, ControlCount + 1)
Next ControlCount

'manager.WriteBox(14).Width =

Call Refresh_Benchmark_Configuration

Call Manager.ExplorerButton(0).ButtonProperty(10, Short(44), 3)
Call Manager.ExplorerButton(1).ButtonProperty(11, Short(45), 4)
Call Manager.ExplorerButton(2).ButtonProperty(65, Short(46), 27)
Call Manager.ExplorerButton(3).ButtonProperty(60, Short(47), 32)
Call Manager.ExplorerButton(4).ButtonProperty(54, Short(48), 31)
Call Manager.ExplorerButton(5).ButtonProperty(144, Short(49), 30)
Call Manager.ExplorerButton(6).ButtonProperty(50, Short(50), 28)
Call Manager.ExplorerButton(7).ButtonProperty(52, Short(51), 29)
Call Manager.ExplorerButton(8).ButtonProperty(101, Short(52), 14)
Call Manager.ExplorerButton(9).ButtonProperty(32, Short(53), 15)
Call Manager.ExplorerButton(10).ButtonProperty(110, Short(54), 16)
Call Manager.ExplorerButton(11).ButtonProperty(23, Short(55), 10)
Call Manager.ExplorerButton(12).ButtonProperty(62, Short(56), 33)
Call Manager.ExplorerButton(13).ButtonProperty(17, Short(57), 7)
Call Manager.ExplorerButton(14).ButtonProperty(63, Short(58), 34)
Call Manager.ExplorerButton(15).ButtonProperty(65, Short(59), 36)
Call Manager.ExplorerButton(16).ButtonProperty(64, Short(60), 35)
Call Manager.ExplorerButton(17).ButtonProperty(35, Short(61), 35)
Call Manager.ExplorerButton(18).ButtonProperty(61, Short(62), 35)
Call Manager.ExplorerButton(19).ButtonProperty(20, Short(63), 8)
Call Manager.ExplorerButton(20).ButtonProperty(21, Short(64), 9)
Call Manager.ExplorerButton(21).ButtonProperty(22, Short(65), 1)
Call Manager.ExplorerButton(22).ButtonProperty(60, Short(66), 32)
Call Manager.ExplorerButton(23).ButtonProperty(49, Short(67), 27)
Call Manager.ExplorerButton(24).ButtonProperty(71, Short(68), 38)
Call Manager.ExplorerButton(25).ButtonProperty(72, Short(69), 39)
Call Manager.ExplorerButton(26).ButtonProperty(73, Short(70), 40)
Call Manager.ExplorerButton(27).ButtonProperty(74, Short(71), 41)
Call Manager.ExplorerButton(28).ButtonProperty(75, Short(72), 57)
Call Manager.ExplorerButton(29).ButtonProperty(80, Short(73), 2)
Call Manager.ExplorerButton(30).ButtonProperty(150, Short(74), 43)
Call Manager.ExplorerButton(31).ButtonProperty(83, Short(75), 56)
Call Manager.ExplorerButton(32).ButtonProperty(85, Short(76), 44)

Call Manager.ControllerBox(0).ControllerProperty(Short(77), 0)
Call Manager.ControllerBox(1).ControllerProperty(Short(78), 44)
Call Manager.ControllerBox(2).ControllerProperty(Short(79), 13)
Call Manager.ControllerBox(3).ControllerProperty(Short(80), 14)
Call Manager.ControllerBox(4).ControllerProperty(Short(81), 15)
Call Manager.ControllerBox(5).ControllerProperty(Short(82), 7)
Call Manager.ControllerBox(6).ControllerProperty(Short(83), 10)
Call Manager.ControllerBox(7).ControllerProperty(Short(84), 0)
Call Manager.ControllerBox(8).ControllerProperty(Short(85), 0)
Call Manager.ControllerBox(9).ControllerProperty(Short(86), 0)
Call Manager.ControllerBox(10).ControllerProperty(Short(87), 32)
Call Manager.ControllerBox(11).ControllerProperty(Short(88), 57)
Call Manager.ControllerBox(12).ControllerProperty(Short(89), 24)
Call Manager.ControllerBox(14).ControllerProperty(Short(90), 27)
Call Manager.ControllerBox(15).ControllerProperty(Short(91), 16)
Call Manager.ControllerBox(16).ControllerProperty(Short(92), 34)
Call Manager.ControllerBox(17).ControllerProperty(Short(93), 18)
Call Manager.ControllerBox(18).ControllerProperty(Short(94), 18)
Call Manager.ControllerBox(19).ControllerProperty(Short(95), 71)
Call Manager.ControllerBox(20).ControllerProperty(Short(110), 71)
If Date_Left = 0 Then
    Call Manager.ControllerBox(21).ControllerProperty(Short(111), 62)
Else
    Call Manager.ControllerBox(21).ControllerProperty(Date_Left & Space_Code & Short(120), 62)
End If
Call Manager.ControllerBox(22).ControllerProperty(Short(112), 38)
Call Manager.ControllerBox(23).ControllerProperty(Short(116), 66)

Call Manager.OptionBox(14).Auto_Ret

Manager.FrameButton(0).Caption = Short(96)
Manager.FrameButton(1).Caption = Short(97)

Call Manager.ToolButton(0).ToolProperty(10, Short(98), 3, True, False)
Call Manager.ToolButton(1).ToolProperty(11, Empty_Code, 4, True, False)
Call Manager.ToolButton(2).ToolProperty(14, Short(99), 5, True, False)
Call Manager.ToolButton(3).ToolProperty(17, Empty_Code, 7, True, True)
Call Manager.ToolButton(4).ToolProperty(32, Empty_Code, 15, True, False)
Call Manager.ToolButton(5).ToolProperty(71, Short(100), 38, True, True)
Call Manager.ToolButton(6).ToolProperty(72, Short(101), 39, True, False)
Call Manager.ToolButton(7).ToolProperty(73, Short(102), 40, True, False)
Call Manager.ToolButton(8).ToolProperty(74, Short(103), 41, True, False)
Call Manager.ToolButton(9).ToolProperty(75, Short(104), 57, True, False)
Call Manager.ToolButton(10).ToolProperty(83, Empty_Code, 56, True, True)
Call Manager.ToolButton(11).ToolProperty(80, Short(105), 42, True, False)

Call Manager.OtherTool(0).ToolProperty(9, Short(106), 58, True, False)
Call Manager.OtherTool(1).ToolProperty(49, Empty_Code, 27, True, True)
Call Manager.OtherTool(2).ToolProperty(50, Short(107), 28, True, False)
Call Manager.OtherTool(3).ToolProperty(52, Empty_Code, 29, True, False)
Call Manager.OtherTool(4).ToolProperty(53, Empty_Code, 13, True, True)
Call Manager.OtherTool(5).ToolProperty(23, Empty_Code, 10, True, False)
Call Manager.OtherTool(6).ToolProperty(20, Empty_Code, 8, True, True)
Call Manager.OtherTool(7).ToolProperty(21, Empty_Code, 9, True, False)

Manager.FormControl(0).ToolTipText = Short(121)
Manager.FormControl(1).ToolTipText = Short(122)
Manager.FormControl(2).ToolTipText = Short(123)
Manager.FormControl(3).ToolTipText = Short(124)
Manager.FormControl(4).ToolTipText = Short(125)

Call Align_ToolBar_Buttons

BenchDatArray(0, 4) = 18
BenchDatArray(1, 4) = 18
BenchDatArray(2, 4) = 19
BenchDatArray(3, 4) = 19
BenchDatArray(4, 4) = 20
BenchDatArray(5, 4) = 20
BenchDatArray(6, 4) = 20
BenchDatArray(7, 4) = 21
BenchDatArray(8, 4) = 21
BenchDatArray(9, 4) = 21
BenchDatArray(10, 4) = 21
BenchDatArray(11, 4) = 22
BenchDatArray(12, 4) = 22
BenchDatArray(13, 4) = 23
BenchDatArray(14, 4) = 23
BenchDatArray(15, 4) = 23
BenchDatArray(16, 4) = 24
BenchDatArray(17, 4) = 25
BenchDatArray(18, 4) = 25
BenchDatArray(19, 4) = 25
BenchDatArray(20, 4) = 26
BenchDatArray(21, 4) = 26
BenchDatArray(22, 4) = 26
BenchDatArray(23, 4) = 26

Call Change_StatusBar_Text(0, Empty_Code)

For ControlCount = 0 To Manager.StylishLabel.Count - 1
    Call Manager.StylishLabel(ControlCount).LabelProperty(BenchDatArray(ControlCount, 0), Short(108) & Space_Code & LCase(BenchDatArray(ControlCount, 2)) & Space_Code & Short(109) & Space_Code & LCase(BenchDatArray(ControlCount, 1)), Val(BenchDatArray(ControlCount, 4)))
    Call Manager.ScoreLabel(ControlCount).LabelProperty(BenchDatArray(ControlCount, 0), Short(108) & Space_Code & LCase(BenchDatArray(ControlCount, 2)) & Space_Code & Short(109) & Space_Code & LCase(BenchDatArray(ControlCount, 1)), Val(BenchDatArray(ControlCount, 4)))
Next ControlCount
For ControlCount = 0 To Manager.ComplexList.Count - 1
    Call Manager.ComplexList(ControlCount).Setup_Cols_Headers(TierInfoColumns())
Next ControlCount
For ControlCount = 0 To Manager.PlatInfoCplxList.Count - 1
    Call Manager.PlatInfoCplxList(ControlCount).Setup_Cols_Headers(PlatInfoColumns())
Next ControlCount
Call Manager.FrameList(0).Setup_Cols_Headers(ServerListColumns())
Call Manager.FrameList(1).Setup_Cols_Headers(WaitingRoomColumns())
Call Manager.FrameList(2).Setup_Cols_Headers(URLBenchListColumns())
Call Manager.FrameList(3).Setup_Cols_Headers(BenchmarkTasksColumns())
Call Manager.FrameList(4).Setup_Cols_Headers(ATMColumns())
Call Manager.FrameList(5).Setup_Cols_Headers(ConnStatusColumns())
Call Manager.FrameList(6).Setup_Cols_Headers(ComunityInfoColumns())

Manager.FrameList(1).KeepFocus = True

Manager.FrameImage(0).Picture = Manager.PictureLoader(3).ListImages.Item(16).Picture
Manager.FrameImage(1).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & One_Code & Resource_Ext)
Manager.FrameImage(3).Picture = Manager.PictureLoader(3).ListImages.Item(14).Picture
Manager.FrameImage(4).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "4" & Resource_Ext)
Manager.FrameImage(5).Picture = Manager.PictureLoader(3).ListImages.Item(26).Picture
Manager.FrameImage(6).Picture = Manager.PictureLoader(3).ListImages.Item(24).Picture
Manager.FrameImage(7).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "5" & Resource_Ext)
Manager.FrameImage(8).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "6" & Resource_Ext)
Manager.FrameImage(9).Picture = Manager.PictureLoader(3).ListImages.Item(31).Picture
Manager.FrameImage(10).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "7" & Resource_Ext)
Manager.FrameImage(11).Picture = Manager.PictureLoader(3).ListImages.Item(31).Picture
Manager.FrameImage(12).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "9" & Resource_Ext)
Manager.FrameImage(13).Picture = Manager.PictureLoader(3).ListImages.Item(31).Picture
If FullVersion = False Then
    Manager.FrameImage(14).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "10" & Resource_Ext)
    Manager.FrameImage(15).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "11" & Resource_Ext)
End If
Manager.FrameImage(17).Picture = Manager.PictureLoader(3).ListImages.Item(31).Picture
Manager.FrameImage(18).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "12" & Resource_Ext)
Manager.FrameImage(19).Picture = Manager.PictureLoader(3).ListImages.Item(16).Picture

ManagerSub.Caption = "SynthMark� XP Benchmarking Suite - Copyright� 1999-2AndyFutchericro� Software"
'SupposedCaption = "AndyFutcher� SynthMark� XP"

ReDim Short(0)
End Sub
