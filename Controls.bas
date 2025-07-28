Attribute VB_Name = "Controls"
Option Explicit

Public Sub Form_Control_Click(Index As Integer)
Select Case Index
Case 0
    If UBound(ThemeFindList()) <> 0 Then
        For SettingsCount = 0 To UBound(ThemeFindList())
            If SettingsCount = UBound(ThemeFindList()) Then
                Theme_Current = Add_The_Slash(ThemeFindList(0))
                Call Prepare_Controls
                GoTo Ed
            End If
            If Theme_Current = Add_The_Slash(ThemeFindList(SettingsCount)) Then
                Theme_Current = Add_The_Slash(ThemeFindList(SettingsCount + 1))
                Call Prepare_Controls
                GoTo Ed
            End If
        Next SettingsCount
    End If
Case 1
    Call Open_Help_System(0)
Case 2
    Manager.Hide
    ManagerSub.WindowState = 1
    Call Manager_WindowState_Check
Case 3
    If Manager.WindowState = 0 Then
        Manager.WindowState = 2
    Else
        Manager.WindowState = 0
    End If
    Call Manager.Form_Resize_Now
    Call Center_ControllerBoxes(-1)
Case 4
    Call Save_Settings_Data(0)
    Call Prompt_Quit
End Select
Ed: End Sub

Public Sub Check_All_Menus()
Select Case Manager.WindowState
Case 0
    Can_Menu(5) = True
    Can_Menu(4) = True
    Can_Menu(3) = False
Case 1
    Can_Menu(5) = False
    Can_Menu(4) = False
    Can_Menu(3) = True
Case 2
    Can_Menu(5) = True
    Can_Menu(4) = False
    Can_Menu(3) = True
End Select

Can_Menu(24) = Can_Menu(1) 'beep

If (Net_Single = False And Net_ServerType = False And Net_ClientType = False) Or Net_Dedicated = True Then
    Call Is_Something_Open(False)
Else
    Call Is_Something_Open(True)
End If
If Net_ServerType = False And Net_ClientType = False Then
    Can_Menu(54) = True
    Can_Menu(36) = False
Else
    Can_Menu(54) = False
    Can_Menu(36) = True
End If
If Net_ClientType = True Then
    Can_Menu(47) = True
Else
    Can_Menu(47) = False
End If
If Are_There_Benchmarks = True And Net_Dedicated = False Then
    Can_Menu(72) = True
    Can_Menu(73) = True
Else
    Can_Menu(72) = False
    Can_Menu(73) = False
End If
Can_Menu(50) = False
Can_Menu(52) = False
If SectionSelect <> 5 Then
For ControlCount = 0 To UBound(BenchSelArray)
    If BenchSelArray(ControlCount) = True Then
        Can_Menu(50) = True
        If ConnectedUsers(1, 0) = App_Full Or ConnectedUsers(1, 0) = App_Timed Then Can_Menu(52) = True
    End If
Next ControlCount
End If
If SectionSelect = 4 Then
    Can_Menu(26) = True
Else
    Can_Menu(26) = False
End If
If LoadPath <> Empty_Code Then
    Can_Menu(13) = True
Else
    Can_Menu(13) = False
End If
Manager.CommandButton(37).Enabled = Can_Menu(36)
If ConnectedUsers(1, 0) = App_Trial And (Net_ServerType = True Or Net_ClientType = True) Then
    Can_Menu(43) = False
    Can_Menu(44) = False
    Can_Menu(45) = False
    Can_Menu(46) = False
    Can_Menu(47) = False
    Can_Menu(52) = False
    Manager.OptionBox(1).Enabled = False
    Manager.OptionBox(2).Enabled = False
    Manager.OptionBox(3).Enabled = False
    For SettingsCount = 7 To 20
        BenchSelArray(SettingsCount) = False
    Next SettingsCount
End If
If ConnectedUsers(1, 0) = App_Trial Then
    Manager.WriteBox(3).Enabled = False
    Manager.WriteBox(3).Text = Empty_Code
    Manager.OptionBox(16).ValueDesc = False
    Manager.OptionBox(16).Enabled = False
    Can_Menu(18) = False
    Can_Menu(16) = False
    Can_Menu(62) = False
Else
    Manager.WriteBox(3).Enabled = True
    Manager.OptionBox(16).Enabled = True
End If

'Make True for all.
For ControlCount = 0 To Manager.ToolButton.Count - 1
    Manager.ToolButton(ControlCount).Enabled = Can_Menu(Manager.ToolButton(ControlCount).Command)
Next ControlCount
For ControlCount = 0 To Manager.OtherTool.Count - 1
    Manager.OtherTool(ControlCount).Enabled = Can_Menu(Manager.OtherTool(ControlCount).Command)
Next ControlCount
For ControlCount = 0 To Manager.ExplorerButton.Count - 1
    Manager.ExplorerButton(ControlCount).Enabled = Can_Menu(Manager.ExplorerButton(ControlCount).Command)
Next ControlCount
Call Update_Address_Bar
End Sub

Private Sub Is_Something_Open(YesOrNo As Boolean)
Can_Menu(12) = YesOrNo
Can_Menu(13) = YesOrNo
Can_Menu(14) = YesOrNo
Can_Menu(15) = YesOrNo
Can_Menu(16) = YesOrNo
Can_Menu(17) = YesOrNo
Can_Menu(18) = YesOrNo
Can_Menu(20) = YesOrNo
Can_Menu(21) = YesOrNo
Can_Menu(22) = YesOrNo
Can_Menu(27) = YesOrNo
Can_Menu(40) = YesOrNo
Can_Menu(41) = YesOrNo
Can_Menu(42) = YesOrNo
Can_Menu(43) = YesOrNo
Can_Menu(44) = YesOrNo
Can_Menu(45) = YesOrNo
Can_Menu(46) = YesOrNo
'Can_Menu(47) = YesOrNo
Can_Menu(48) = YesOrNo
Can_Menu(49) = YesOrNo
Can_Menu(50) = YesOrNo
Can_Menu(51) = YesOrNo
Can_Menu(52) = YesOrNo
Can_Menu(53) = YesOrNo
'Can_Menu(54) = YesOrNo
'Can_Menu(61) = YesOrNo
Can_Menu(62) = YesOrNo
'Can_Menu(63) = YesOrNo
Can_Menu(71) = YesOrNo
Can_Menu(72) = YesOrNo
Can_Menu(73) = YesOrNo
Can_Menu(74) = YesOrNo
'Can_Menu(75) = YesOrNo
Can_Menu(76) = Switch_Boolean(YesOrNo)
Can_Menu(144) = YesOrNo

'Can_Menu(150) = YesOrNo
'Can_Menu(151) = YesOrNo
'Can_Menu(152) = YesOrNo
'Can_Menu(153) = YesOrNo
End Sub

Public Sub Process_CommandButton_Click(Index As Integer)
Select Case Index
Case 0
    Call Hide_Controller_Box(1)
Case 1
    TipTextIndex = TipTextIndex + 1
    Call Prepare_Tip_Display
Case 2
    Call Unload_Application
Case 3
    Call Hide_Controller_Box(0)
Case 4, 5
    Call Hide_Controller_Box(2)
Case 6
    Call Process_WriteList_Click(11)
    Call Process_WriteList_Click(12)
Case 7
    Call Reset_Defaults(3)
Case 8
    Call Hide_Controller_Box(3)
    Call Multi_New_Session
    If Net_Dedicated = False Then
        Call Switch_Sections_To(0)
    Else
        Call Process_Menu_Command_Ids(-1, 0, 30)
    End If
Case 9
    Call Hide_Controller_Box(3)
Case 10
    Call Multi_Find_Session
Case 11
    Call Hide_Controller_Box(4)
    Call Multi_Join_Session(0)
Case 12
    Call Hide_Controller_Box(4)
    Call Show_Controller_Box(7)
Case 13
    Call Hide_Controller_Box(4)
Case 14
    Call Reset_Defaults(5)
Case 15
    Call Print_RTF_Box
    Call Hide_Controller_Box(5)
Case 16
    Call Hide_Controller_Box(5)
Case 17
    If Manager.WriteBox(6).Text <> Empty_Code Then
        Call Update_System_Information
    End If
    If Net_ServerType = True Or Net_Single = True Then Call Net_UpDataUsers
    If Net_ClientType = True Then
        Call Net_Send_Data(NetCode_Comment & ConnectedUsers(2, 0) & PlatFormUsers(5, 0), -1, -1)
        Call Net_Send_Data(NetCode_Name & ConnectedUsers(2, 0) & ConnectedUsers(0, 0), -1, -1)
    End If
    Call Hide_Controller_Box(6)
Case 18
    Call Hide_Controller_Box(6)
Case 19
    Call Reset_Defaults(7)
Case 20
    Call Hide_Controller_Box(7)
    Call Multi_Join_Session(-1)
Case 21
    Call Hide_Controller_Box(7)
Case 22
    Manager.WriteBox(10).Tag = 1
    'Call Hide_Controller_Box(8)
Case 23
    Manager.WriteBox(10).Tag = 2
    'Call Hide_Controller_Box(8)
Case 24
    Call Start_BenchMarks(LastBenType)
Case 25
    If BenchDoneOnce = False Then
        Call Hide_Controller_Box(9)
    Else
        StopBenchmarks = True
    End If
Case 26
    Call Manager.WriteBox_KeyPressed(15, 0, 0)
    Call Show_Controller_Box(11)
Case 27
    Call Remove_Index_From_StringArray(URLBenchList(), Manager.FrameList(2).ListIndex)
    If URLBenchList(0, 0) <> Empty_Code Then
        Call Manager.FrameList(2).Submit_Data_Array(URLBenchList, -1, 1)
    Else
        Call Manager.FrameList(2).Empty_Data
    End If
    Call Check_Target_Url_Buttons
    Call Net_URLBench_Equalize(-1)
Case 28
    For NetCount = 0 To UBound(URLBenchList(), 2)
        URLBenchList(1, NetCount) = Language(199)
    Next NetCount
    Call Begin_URL_Check_List
Case 29
    BenchPassword = LCase(Manager.WriteBox(13).Text)
    Call Hide_Controller_Box(18)
Case 30
    Call Hide_Controller_Box(18)
Case 31
    If Manager.CommandButton(Index).Tag = One_Code Then
        Manager.ControllerBox(9).Height = 5880
        Manager.CommandButton(Index).Tag = Zero_Code
        Manager.CommandButton(Index).Caption = "<< " & Language(226)
        Manager.NormLine(7).Visible = True
        Manager.NormLine(8).Visible = True
    Else
        Manager.ControllerBox(9).Height = 2295
        Manager.CommandButton(Index).Tag = One_Code
        Manager.CommandButton(Index).Caption = Language(226) & " >>"
        Manager.NormLine(7).Visible = False
        Manager.NormLine(8).Visible = False
    End If
    Manager.Refresh
Case 32
    Call ATM_Switch_Processes
Case 33
    Call ATM_End_Process
Case 34
    Call Hide_Controller_Box(10)
    Manager.Timer(1).Interval = 0
Case 35
    Dim TempURLString As String
    TempURLString = Manager.WriteBox(15).Text
    Manager.WriteBox(15).Text = Empty_Code
    Call Hide_Controller_Box(11)
    Call Add_URLList(TempURLString)
    Call Net_URLBench_Equalize(-1)
    Call Begin_URL_Check_List
Case 36
    Call Hide_Controller_Box(11)
Case 37
    Call Process_Menu_Command_Ids(-1, 0, 36)
Case 38
    Call Hide_Controller_Box(12)
Case 39
    Call Hide_Controller_Box(13)
    Manager.CommandButton(52).Tag = 0
Case 40
    Call Save_Settings_Data(1)
    Call Hide_Controller_Box(14)
Case 41
    Call Load_Settings_Data(1)
    Call Hide_Controller_Box(14)
Case 42
    Call Save_Settings_Data(1)
Case 43
    Call Switch_UpdateWindow_Index(5)
    'Call Choose_Manager_Functionality(False, 15, 1)
    Call Check_Online_Status(0)
    Call Do_Software_Update_Now(True)
Case 44
    Call Reset_Data_Limits(True)
    Call Save_Settings_Data(1)
    Call Hide_Controller_Box(16)
Case 45
    Call Load_Settings_Data(1)
    Call Hide_Controller_Box(16)
Case 46
    Call Process_CommandButton_Click(43)
Case 47
    Call Begin_Mail_CPU
    Call Hide_Controller_Box(17)
Case 48
    Call Hide_Controller_Box(17)
Case 49
    BenchPassword = Empty_Code
    Call Hide_Controller_Box(18)
Case 50
    Call Hide_Controller_Box(19)
Case 51
    Call Process_Menu_Command_Ids(-1, 0, 14)
    DoEvents
    Call Unload_Application
Case 52
    Call Hide_Controller_Box(13)
    Manager.CommandButton(52).Tag = 1
Case 53
    Call Select_Common_Benchmarks(3)
    Call Hide_Controller_Box(20)
    DoBurnin = True
    Call Show_Controller_Box(9)
    Do While StopBenchmarks = False
        Call Start_BenchMarks(False)
    Loop
    DoBurnin = False
    If Net_Single = False And Net_ServerType = False And Net_ClientType = False Then Call Hide_Controller_Box(9)
Case 54
    Call Hide_Controller_Box(20)
Case 55
    'Call Hide_Controller_Box(21)
    Call Browse_To_Site(-2, Empty_Code)
    'Call Switch_Sections_To(4)
Case 56
    Call Hide_Controller_Box(21)
    Manager.Timer(5).Interval = 250
    Call Show_Controller_Box(22)
Case 57
    Call Hide_Controller_Box(21)
Case 58
    Call Hide_Controller_Box(19)
    Call Show_Controller_Box(21)
Case 59
    Dim NewCDKeyString As String
    For SettingsCount = 27 To 30
        NewCDKeyString = NewCDKeyString & Manager.WriteBox(SettingsCount).Text
    Next SettingsCount
    ConnectedUsers(3, 0) = NewCDKeyString
    Call Add_Version_Status
    If ConnectedUsers(1, 0) = App_Full Then
        If Net_ServerType = True Or Net_Single = True Then Call Net_UpDataUsers
        If Net_ClientType = True Then
            Call Net_Send_Data(NetCode_Version & ConnectedUsers(2, 0) & ConnectedUsers(1, 0), -1, -1)
            Call Net_Send_Data(NetCode_CDKey & ConnectedUsers(2, 0) & ConnectedUsers(3, 0), -1, -1)
            'Call Net_Send_Data(NetCode_Name & ConnectedUsers(2, 0) & ConnectedUsers(0, 0), -1, -1)
            Call Net_Send_Data(NetCode_CheckImIn & ConnectedUsers(2, 0), -1, -1)
        End If
        Call Hide_Controller_Box(22)
    End If
Case 60
    Call Hide_Controller_Box(22)
Case 61
    If NetSearchCom(5, Manager.FrameList(6).ListIndex) <> Empty_Code Then
        Call Switch_Sections_To(4)
        Call Browse_To_Site(-3, NetSearchCom(5, Manager.FrameList(6).ListIndex))
    End If
Case 62
    Call Download_WWW_Photo
Case 63
    Call Multi_Join_Session(1)
Case 64
    Call Hide_Controller_Box(23)
End Select
End Sub

Public Sub Process_OptionBox_Click(Index As Integer)
Select Case Index
Case 1
    For ControlCount = 7 To 10
        BenchSelArray(ControlCount) = Manager.OptionBox(Index).Value
    Next ControlCount
    Call Check_Option_Enabled_States
Case 2
    For ControlCount = 11 To 12
        BenchSelArray(ControlCount) = Manager.OptionBox(Index).Value
    Next ControlCount
    Call Check_Option_Enabled_States
Case 3
    BenchSelArray(16) = Manager.OptionBox(Index).Value
    Call Check_Option_Enabled_States
Case 4, 13
    Manager.WriteBox(21).Enabled = Manager.OptionBox(4).Value
    Manager.LightLabel(10).Tag = Manager.OptionBox(4).Value
    If Manager.OptionBox(Index).Value = True Then Call Check_Online_Status(0)
Case 12
    Manager.OptionBox(17).ValueDesc = Manager.OptionBox(12).Value
Case 14
    Manager.OptionBox(15).ValueDesc = Manager.OptionBox(14).Value
Case 15
    Manager.OptionBox(14).ValueDesc = Manager.OptionBox(15).Value
Case 16
    Call ATM_Process_System_Tasks
Case 17
    Manager.OptionBox(12).ValueDesc = Manager.OptionBox(17).Value
Case 20, 21, 22
    Call Check_BurnIn_CommandButton
End Select
Call Process_Flexible_Lables
End Sub

Public Sub Check_BurnIn_CommandButton()
If Manager.OptionBox(20).Value = True Or Manager.OptionBox(21).Value = True Or Manager.OptionBox(22).Value = True Then
    Manager.CommandButton(53).Enabled = True
Else
    Manager.CommandButton(53).Enabled = False
End If
End Sub

Public Sub Check_NetCanConnect_Button()
If NetSearchList(2, 0) <> Empty_Code Then
    Manager.CommandButton(11).Enabled = True
Else
    Manager.CommandButton(11).Enabled = False
End If
If NetSearchCom(3, 0) <> Empty_Code Then
    Manager.CommandButton(61).Enabled = True
    Manager.CommandButton(62).Enabled = True
    Manager.CommandButton(63).Enabled = True
Else
    Manager.CommandButton(61).Enabled = False
    Manager.CommandButton(62).Enabled = False
    Manager.CommandButton(63).Enabled = False
End If
End Sub

Public Sub Check_Option_Enabled_States()
Manager.WriteBox(11).Enabled = Manager.OptionBox(1).Value
Manager.WriteBox(12).Enabled = Manager.OptionBox(2).Value
Manager.CommandButton(26).Enabled = Manager.OptionBox(3).Value
If Manager.OptionBox(3).Enabled = False Then Manager.CommandButton(26).Enabled = True
Manager.CommandButton(27).Enabled = Manager.OptionBox(3).Value
Manager.CommandButton(28).Enabled = Manager.OptionBox(3).Value
Call Check_Target_Url_Buttons
Call Check_All_Menus
End Sub

Public Sub Check_Password_Similarity()
If Trim(Manager.WriteBox(13).Text) <> Empty_Code And LCase(Manager.WriteBox(13).Text) = LCase(Manager.WriteBox(26).Text) Then
    Manager.CommandButton(29).Enabled = True
Else
    Manager.CommandButton(29).Enabled = False
End If
If BenchPassword <> Empty_Code Then
    Manager.CommandButton(49).Enabled = True
Else
    Manager.CommandButton(49).Enabled = False
End If
End Sub

Public Sub Select_Common_Benchmarks(TypeIndex As Integer)
If TypeIndex <> -1 Then Call Select_No_Benchmarks
Select Case TypeIndex
Case 1
    For VisualCount = 0 To 6
        BenchSelArray(VisualCount) = True
    Next VisualCount
    For VisualCount = 13 To 15
        BenchSelArray(VisualCount) = True
    Next VisualCount
Case 2
    For VisualCount = 0 To UBound(BenchSelArray())
        If Can_Menu(120 + VisualCount) = True Then BenchSelArray(VisualCount) = True
    Next VisualCount
    If Can_Menu(47) = False Then
        BenchSelArray(17) = False
        BenchSelArray(18) = False
        BenchSelArray(19) = False
    End If
Case 3
    If Manager.OptionBox(20).Value = True Then
        For VisualCount = 0 To 6
            BenchSelArray(VisualCount) = True
        Next VisualCount
    End If
    If Manager.OptionBox(21).Value = True Then
        For VisualCount = 13 To 15
            BenchSelArray(VisualCount) = True
        Next VisualCount
    End If
    If Manager.OptionBox(22).Value = True Then
        For VisualCount = 7 To 10
            BenchSelArray(VisualCount) = True
        Next VisualCount
    End If
End Select
If ConnectedUsers(1, 0) = App_Trial Then
    For VisualCount = 7 To 20
        BenchSelArray(VisualCount) = False
    Next VisualCount
End If
End Sub

Public Sub Select_No_Benchmarks()
For VisualCount = 0 To UBound(BenchSelArray())
    BenchSelArray(VisualCount) = False
Next VisualCount
End Sub

Public Sub Process_WriteList_Click(Index As Integer)
Select Case Index
Case 0
    Select Case SectionSelect
    Case 0, 3
        Call Switch_Sections_To(Val(Manager.WriteBox(Index).ClickTag))
    Case 1
        Manager.ComplexList(Val(Manager.WriteBox(Index).ClickTag)).SetFocus
    Case 2
        Manager.GraphBox(Val(Manager.WriteBox(Index).ClickTag)).SetFocus
    Case 4, 5
        Call Process_Menu_Command_Ids(0, 0, Manager.WriteBox(Index).ClickTag)
    End Select
Case 4
    If Manager.WriteBox(4).ClickTag = Net_Code Then Call Check_Online_Status(0)
Case 11
    Manager.LightLabel(0).Caption = Language(14) & Space_Code & Drv_Free_Space(Manager.WriteBox(Index).ClickTag) & Space_Code & Language(15)
Case 12
    If Drv_Total_Size(Manager.WriteBox(Index).ClickTag) <> Empty_Code Then
        Manager.LightLabel(7).Caption = Language(16) & Space_Code & Drv_Total_Size(Manager.WriteBox(Index).ClickTag) & Space_Code & Language(17)
        Manager.LightLabel(7).Tag = Empty_Code
    Else
        Manager.LightLabel(7).Caption = Language(13)
        Manager.LightLabel(7).Tag = Bad_Code
    End If
    Call Process_Flexible_Lables
Case 18
    Call AutoDetect_Internet_Connection
    If Manager.WriteBox(Index).ClickTag < 2 Then
        Manager.WriteBox(19).Enabled = False
        Manager.WriteBox(20).Enabled = False
        Manager.LightLabel(8).Tag = False
        Manager.LightLabel(9).Tag = False
    Else
        Manager.WriteBox(19).Enabled = True
        Manager.WriteBox(20).Enabled = True
        Manager.LightLabel(8).Tag = True
        Manager.LightLabel(9).Tag = True
    End If
    Call Process_Flexible_Lables
End Select
Ed: End Sub

Public Sub Process_System_Defualts()
For ControlCount = 0 To Manager.WriteBox.Count - 1
    If Manager.WriteBox(ControlCount).Text = Empty_Code Then Manager.WriteBox(ControlCount).Text = Manager.WriteBox(ControlCount).Tag
Next ControlCount
End Sub

Public Sub Manager_Display_Menu(Index As Integer, FormNumber As Integer)
Select Case FormNumber
Case 1
    Call Manager_Menu_Loader(Index)
    Call Manager.MenuBox(0).Align_MenuItems
    Call Hide_MenuBox(1)
    Manager.MenuBox(0).Tag = Index
    If Index = -1 Then
        Manager.MenuBox(0).Move Manager.FormIcon.Left, Manager.FormIcon.Top + Manager.FormIcon.Height
    Else
        Manager.MenuBox(0).Move Manager.MenuButton(Index).Left, Manager.FormHeader(2).Top + Manager.MenuButton(Index).Top + Manager.MenuButton(Index).Height - Screen.TwipsPerPixelY
    End If
    Manager.MenuBox(0).Visible = True
    Call Manager.MenuBox(0).ZOrder(0)
    Call Manager.MenuBox(0).SetFocus
Case 2
    Call Manager.MenuBox(1).Align_MenuItems
    Manager.MenuBox(1).Tag = Index
    Manager.MenuBox(1).Move Manager.MenuBox(0).Left + Manager.MenuBox(0).Width, Manager.MenuBox(0).Top + Manager.MenuBox(0).Give_Menu_Item_Top(Index) '- Screen.TwipsPerPixelY
    'Manager.MenuBox(1).Move Manager.MenuBox(0).Left + Manager.MenuBox(0).Width, Manager.FormHeader(2).Top + Manager.MenuButton(0).Top + Manager.MenuButton(0).Height + Manager.MenuBox(0).Give_Menu_Item_Top(Index) - Screen.TwipsPerPixelY
    Manager.MenuBox(1).Visible = True
    Call Manager.MenuBox(1).ZOrder(0)
    Call Manager.MenuBox(1).SetFocus
Case 3
    Call Manager.MenuBox(1).Align_MenuItems
    Manager.MenuBox(1).Tag = Index
    Manager.MenuBox(1).Move Manager.OtherTool(Index).Left, Manager.FormHeader(4).Top + Manager.OtherTool(Index).Top
    Manager.MenuBox(1).Visible = True
    Call Manager.MenuBox(1).ZOrder(0)
    Call Manager.MenuBox(1).SetFocus
Case 4
    Call Manager.MenuBox(1).Align_MenuItems
    Manager.MenuBox(1).Tag = Index
    Manager.MenuBox(1).Move Manager.ExplorerFrame(Manager.ExplorerButton(Index).Tag).Width, Manager.ExplorerHolder.Top + Manager.ExplorerFrame(Manager.ExplorerButton(Index).Tag).Top + Manager.ExplorerButton(Index).Top
    Manager.MenuBox(1).Visible = True
    Call Manager.MenuBox(1).ZOrder(0)
    Call Manager.MenuBox(1).SetFocus
Case 5
    Call Manager.MenuBox(1).Align_MenuItems
    Manager.MenuBox(1).Tag = Index
    RetVal = APIControls.GetCursorPos(MouseLoc)
    Manager.MenuBox(1).Move (MouseLoc.X * Screen.TwipsPerPixelX) - (Manager.Left + FormWall) - Manager.MenuBox(1).Width, (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall) - Manager.MenuBox(1).Height
    
    Manager.MenuBox(1).Visible = True
    Call Manager.MenuBox(1).ZOrder(0)
    Call Manager.MenuBox(1).SetFocus
End Select
Call Normalize_Controls(4)
End Sub

Public Sub Align_Explorer_Frames(Index As Integer)
'On Error Resume Next
'DoEvents
Dim FrameHeights(), FrameTops() As Integer, TopPosHolder As Integer, ExpLastWidth As Integer, ExpLastScroll As Boolean
ReDim FrameHeights(ExplorerFrame_Count)
ReDim FrameTops(ExplorerFrame_Count)

TopPosHolder = EightPix
For ExplorerCount = 0 To Manager.ExplorerFrame.Count - 1
    If Manager.ExplorerFrame(ExplorerCount).Visible = True Then
        TopPosHolder = TopPosHolder + Manager.ExplorerFrame(ExplorerCount).Height + EightPix
        LastExplorerEnd = TopPosHolder
    End If
Next ExplorerCount
Manager.VScrollButton(0).Gap = Manager.ExplorerHolder.Height
Manager.VScrollButton(0).Play = LastExplorerEnd

If LastExplorerEnd > Manager.ExplorerHolder.Height Then
    If Manager.VScrollButton(0).Visible = False Then
        Manager.VScrollButton(0).Visible = True
        Manager.VScrollButton(0).Value = 0
    End If
Else
    If Manager.VScrollButton(0).Visible = True Then
        Manager.VScrollButton(0).Visible = False
        DefPosHolder = EightPix
    End If
End If
Call Manager.VScrollButton(0).Process_CoOrdinates(0)

If Index = 0 And ExpLastWidth = Manager.ExplorerHolder.Width And ExpLastScroll = Manager.VScrollButton(0).Visible And NowLoading = False Then GoTo Ed
ExpLastWidth = Manager.ExplorerHolder.Width
ExpLastScroll = Manager.VScrollButton(0).Visible

For ExplorerCount = 0 To Manager.ExplorerFrame.Count - 1
    If Manager.VScrollButton(0).Visible = True Then
        Manager.ExplorerFrame(ExplorerCount).Width = Manager.ExplorerHolder.UseableArea - Manager.VScrollButton(0).Width - SixTeenPix
    Else
        Manager.ExplorerFrame(ExplorerCount).Width = Manager.ExplorerHolder.UseableArea - SixTeenPix
    End If
    Call Manager.ExplorerFrame(ExplorerCount).Elipser_Check
Next ExplorerCount
For ExplorerCount = 0 To Manager.ExplorerButton.Count - 1
    If Manager.ExplorerButton(ExplorerCount).Visible = True Then
        Manager.ExplorerButton(ExplorerCount).Width = Manager.ExplorerFrame(Manager.ExplorerButton(ExplorerCount).Tag).Width - SixTeenPix
        Call Manager.ExplorerButton(ExplorerCount).Align_Controls
    End If
Next ExplorerCount

For ExplorerCount = 0 To Manager.ExplorerButton.Count - 1
    FrameTops(Manager.ExplorerButton(ExplorerCount).Tag) = (31 * OnePix)
Next ExplorerCount
For ExplorerCount = 0 To Manager.ExplorerButton.Count - 1
    Manager.ExplorerButton(ExplorerCount).Top = FrameTops(Manager.ExplorerButton(ExplorerCount).Tag)
    FrameHeights(Manager.ExplorerButton(ExplorerCount).Tag) = FrameHeights(Manager.ExplorerButton(ExplorerCount).Tag) + Manager.ExplorerButton(ExplorerCount).Height + SixPix
    FrameTops(Manager.ExplorerButton(ExplorerCount).Tag) = FrameTops(Manager.ExplorerButton(ExplorerCount).Tag) + Manager.ExplorerButton(ExplorerCount).Height + SixPix
Next ExplorerCount
For ExplorerCount = 0 To Manager.ExplorerFrame.Count - 1
    Manager.ExplorerFrame(ExplorerCount).PanelHeight = FrameHeights(ExplorerCount)
Next ExplorerCount

Ed: TopPosHolder = DefPosHolder
NormHolder = EightPix
For ExplorerCount = 0 To Manager.ExplorerFrame.Count - 1
    If Manager.ExplorerFrame(ExplorerCount).Visible = True Then
        Manager.ExplorerFrame(ExplorerCount).Top = TopPosHolder
        Manager.ExplorerFrame(ExplorerCount).Refresh
        TopPosHolder = TopPosHolder + Manager.ExplorerFrame(ExplorerCount).Height + EightPix
        NormHolder = NormHolder + Manager.ExplorerFrame(ExplorerCount).Height + EightPix
        'LastExplorerEnd = TopPosHolder
    End If
Next ExplorerCount
End Sub

Public Sub Manager_Tab_Movements()
Dim TabCounter As Integer, LastFrameCount As Integer
LastFrameCount = -1

For ControlCount = 0 To Manager.MenuButton.Count - 1
    Manager.MenuButton(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.ExplorerButton.Count - 1
    If Manager.ExplorerButton(ControlCount).Tag <> LastFrameCount Then
        Manager.ExplorerFrame(Manager.ExplorerButton(ControlCount).Tag).TabIndex = TabCounter
        LastFrameCount = Manager.ExplorerButton(ControlCount).Tag
        TabCounter = TabCounter + 1
    End If
    Manager.ExplorerButton(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.CommandButton.Count - 1
    Manager.CommandButton(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.WriteBox.Count - 1
    Manager.WriteBox(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.FrameList.Count - 1
    Manager.FrameList(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
Manager.PlatInfoCplxList(0).TabIndex = TabCounter
TabCounter = TabCounter + 1
For ControlCount = 0 To Manager.OptionBox.Count - 1
    If Manager.OptionBox(ControlCount).Tag <> "lts" Then Manager.OptionBox(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.ComplexList.Count - 1
    Manager.ComplexList(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.PlatInfoCplxList.Count - 1
    Manager.PlatInfoCplxList(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.ChatterBox.Count - 1
    Manager.ChatterBox(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
For ControlCount = 0 To Manager.FrameButton.Count - 1
    Manager.FrameButton(ControlCount).TabIndex = TabCounter
    TabCounter = TabCounter + 1
Next ControlCount
End Sub

Public Sub Prompt_Quit()
Manager.Show 0, ManagerSub
Call Show_Controller_Box(0)
If Net_ServerType = True Or Net_ClientType = True Or Net_Single = True Then
    Manager.CommandButton(51).Visible = True
    Call SetFocus_Class(0, 51)
Else
    Manager.CommandButton(51).Visible = False
    Call SetFocus_Class(0, 2)
End If
End Sub

Public Sub SetFocus_Class(ClassIndex As Integer, AltIndex As Integer)
On Error Resume Next
Select Case ClassIndex
Case 0
    Manager.CommandButton(AltIndex).SetFocus
Case 1
    Manager.MenuButton(AltIndex).SetFocus
Case 2
    Manager.WriteBoxList.SetFocus
Case 3
    Manager.WriteBox(AltIndex).SetFocus
Case 4
    Manager.SetFocus
End Select
End Sub

Public Sub Invoke_Keyboard_Shortcuts(KeyCodes As Integer)
If LastKeyCode = 18 And KeyCodes = 115 Then Call Form_Control_Click(4)
If LastKeyCode = 18 And KeyCodes = 114 Then Call Form_Control_Click(2)
'If KeyCodes = 13 Then Beep
LastKeyCode = KeyCodes
End Sub

Public Sub Show_ToolTip(ToolTipNumber As Integer, TipLocation As Integer)
If SubDoneOnce = True Or Manager.OptionBox(14).Value = False Then GoTo Ed
SubDoneOnce = True
If ToolTipNumber = 0 Then GoTo Ed1
Dim Additional_Left As Integer
If LastTipShown <> ToolTipArray(ToolTipNumber) Then
    Manager.ToolTipBox.Caption = ToolTipArray(ToolTipNumber)
    LastTipShown = ToolTipArray(ToolTipNumber)
End If
If Manager.ToolTipBox.Caption = Empty_Code Then GoTo Ed1
RetVal = APIControls.GetCursorPos(MouseLoc)
Select Case TipLocation
Case 0
    If (MouseLoc.X * Screen.TwipsPerPixelX) - (Manager.Left + FormWall) + Manager.ToolTipBox.Width > Manager.ScaleWidth Then
        Manager.ToolTipBox.Move (MouseLoc.X * Screen.TwipsPerPixelX) - Manager.ToolTipBox.Width - (Manager.Left + FormWall), (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall) + (EightPix * 4)
    Else
        Manager.ToolTipBox.Move (MouseLoc.X * Screen.TwipsPerPixelX) - (Manager.Left + FormWall), (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall) + (EightPix * 4)
    End If
Case 1
    Manager.ToolTipBox.Move Manager.MenuBox(0).Left + Manager.MenuBox(0).Width + EightPix, (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall)
Case 2
    Manager.ToolTipBox.Move Manager.MenuBox(1).Left + Manager.MenuBox(1).Width + EightPix, (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall)
Case 3
    Manager.ToolTipBox.Move ((MouseLoc.X * Screen.TwipsPerPixelX) - Manager.Left) + (EightPix * 4), (MouseLoc.Y * Screen.TwipsPerPixelY) - (Manager.Top + FormWall)
End Select
If Manager.ToolTipBox.Visible = False Then
    Manager.ToolTipBox.Visible = True
    Manager.ToolTipBox.ZOrder 0
End If
DoEvents
SubDoneOnce = False
GoTo Ed
Ed1: Manager.ToolTipBox.Visible = False
SubDoneOnce = False
Ed: End Sub

Public Sub Open_Help_System(HelpIndex As Integer)
Dim OpenPath As String
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_Global, Reg_ATXCPUInfo, PlatFormUsers(0, 0) & HTML_Enter & PlatFormUsers(2, 0) & HTML_Enter & PlatFormUsers(1, 0) & HTML_Enter & PlatFormUsers(4, 0))
OpenPath = Application_Path & Lang_Path & Language_Current & "synthmark_xp.rlp"
Select Case HelpIndex
Case 1
    OpenPath = OpenPath & Space_Code & "-index"
Case 2
    OpenPath = OpenPath & Space_Code & "-search"
Case 3
    OpenPath = OpenPath & Space_Code & "-atx"
Case 4
    OpenPath = OpenPath & Space_Code & "-reg"
End Select
If ATXFilePath <> Empty_Code Then Shell ATXFilePath & Space_Code & OpenPath
End Sub

Public Sub Hide_MenuBox(Index As Integer)
If Index = 0 Then If Manager.MenuBox(0).Tag <> -1 Then Call SetFocus_Class(1, Manager.MenuBox(0).Tag) 'Manager.MenuButton(Manager.MenuBox(0).Tag).SetFocus
Manager.MenuBox(Index).Tag = -1
Manager.MenuBox(Index).Visible = False
End Sub
