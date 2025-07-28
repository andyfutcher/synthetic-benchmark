Attribute VB_Name = "MenuControl"
Option Explicit

Public Sub Manager_Menu_Loader(Index As Integer)
Call Manager.MenuBox(0).Clear_Entire_List
Select Case Index
Case -1
    Call Manager.MenuBox(0).Setup_Menu(0, 1, Language(69), 11, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 2, Language(70), 2, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 3, Language(71), 1, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 4, Language(72), 45, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 5, Language(73), 46, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 6, Language(74), 1, False, 1, Can_Menu(), False)
Case 0
    Call Manager.MenuBox(0).Setup_Menu(0, 10, Language(75), 3, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 11, Language(76), 4, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 12, Language(77), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 14, Language(78), 5, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 15, Language(79), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 16, Language(80), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(6, 13, Language(81), 6, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(7, 17, Language(82), 7, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(8, 18, Language(83), 1, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(9, 6, Language(84), 1, False, 1, Can_Menu(), False)
Case 1
    Call Manager.MenuBox(0).Setup_Menu(0, 20, Language(85), 8, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(1, 21, Language(86), 9, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(2, 22, Language(87), 1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(3, 23, Language(88), 10, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 24, Language(89), 11, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 25, Language(90), 1, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(6, 26, Language(91), 12, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(7, 27, Language(92), 1, False, 0, Can_Menu(), False)
    'Call Manager.MenuBox(0).Setup_Menu(8, 28, "Refresh", 1, False, 1, Can_Menu(), False)
Case 2
    Call Manager.MenuBox(0).Setup_Menu(0, 30, Language(93), 30, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 31, Language(94), 14, True, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 32, Language(95), 15, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 33, Language(96), 1, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 34, Language(97), 16, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 35, Language(98), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(6, 36, Language(99), 17, False, 1, Can_Menu(), False)
Case 3
    Call Manager.MenuBox(0).Setup_Menu(0, 40, BenchDatArray(0, 2), 18, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(1, 41, BenchDatArray(2, 2), 19, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(2, 42, BenchDatArray(4, 2), 20, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(3, 43, BenchDatArray(7, 2), 21, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(4, 44, BenchDatArray(11, 2), 22, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(5, 45, BenchDatArray(13, 2), 23, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(6, 46, BenchDatArray(16, 2), 24, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(7, 47, BenchDatArray(17, 2), 25, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(8, 48, BenchDatArray(20, 2), 26, True, 0, Can_Menu(), True)
    Call Manager.MenuBox(0).Setup_Menu(9, 49, Language(100), 27, True, 1, Can_Menu(), True)
Case 4
    Call Manager.MenuBox(0).Setup_Menu(0, 50, Language(101), 28, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 51, Language(102), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 52, Language(103), 29, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 53, Language(104), 13, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 54, Language(105), 31, False, 0, Can_Menu(), False)
Case 5
    Call Manager.MenuBox(0).Setup_Menu(0, 60, Language(106), 32, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 61, Language(107), 1, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 62, Language(108), 33, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 63, Language(109), 34, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 64, Language(110), 35, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 65, Language(111), 36, False, 1, Can_Menu(), False)
Case 6
    Call Manager.MenuBox(0).Setup_Menu(0, 70, Language(112), 37, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 71, Language(1), 38, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 72, Language(2), 39, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 73, Language(3), 40, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 74, Language(4), 41, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 75, Language(5), 57, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(6, 76, Language(6), 71, False, 1, Can_Menu(), False)
Case 7
    Call Manager.MenuBox(0).Setup_Menu(0, 80, Language(113), 2, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(1, 81, Language(114), 1, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(2, 82, Language(115), 42, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(3, 83, Language(116), 56, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(4, 84, Language(117), 43, True, 0, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(5, 85, Language(118), 44, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(0).Setup_Menu(6, 86, Language(119), 1, False, 0, Can_Menu(), False)
End Select
End Sub

Public Sub Process_Menu_Command_Ids(Index As Integer, ClickType As Integer, MenuCmdID As Integer)
Call Process_Dynamic_Command_Ids(Index, MenuCmdID)
Select Case MenuCmdID
Case 1, 24
    Call Manager.MenuBox(1).Clear_Entire_List
    For SettingsCount = 0 To UBound(ThemeFindList)
        Call Manager.MenuBox(1).Setup_Menu(SettingsCount, 170 + SettingsCount, Remove_The_Slash(ThemeFindList(SettingsCount)), 1, False, 0, Can_Menu(), False)
    Next SettingsCount
    If SettingsCount < 10 Then Call Manager.MenuBox(1).Setup_Menu(SettingsCount, 170 + SettingsCount, Language(141), 35, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 2
    Call Open_Help_System(0)
Case 3, 4
    Call Form_Control_Click(3)
Case 5
    Call Form_Control_Click(2)
Case 6
    Call Form_Control_Click(4)
Case 9
    Call Process_WriteList_Click(0)
Case 6, 19
    Call Prompt_Quit
Case 10
    Call Single_New_Session
    Call Switch_Sections_To(0)
Case 11
    LoadPath = Show_Common_Dialogue(0, Language(120), Language(121), SaveFolder)
    If LoadPath <> Empty_Code Then
        Call Single_New_Session
        Call Load_Current_Benchmark(LoadPath)
        Call Switch_Sections_To(0)
    End If
Case 12
    Call Chat_Score_Reset(0)
    Call All_Close_Session
Case 13
    Call Single_New_Session
    Call Load_Current_Benchmark(LoadPath)
    Call Switch_Sections_To(0)
Case 14
    Call Check_All_Menus
    If SavePath = Empty_Code Then SavePath = Show_Common_Dialogue(1, Language(122), Language(121), SaveFolder)
    If SavePath <> Empty_Code Then Call Save_Current_Benchmark(SavePath)
Case 15
    Call Check_All_Menus
    SavePath = Show_Common_Dialogue(1, Language(123), Language(121), SaveFolder)
    If SavePath <> Empty_Code Then Call Save_Current_Benchmark(SavePath)
Case 16
    Dim SaveRTFPath As String
    Call Check_All_Menus
    SaveRTFPath = Show_Common_Dialogue(1, Language(124), Language(125), SaveFolder)
    If SaveRTFPath <> Empty_Code Then Call Save_RTF_Box(SaveRTFPath)
Case 17
    Call Show_Controller_Box(5)
Case 18
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 90, Language(126), 59, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(1, 91, Language(127), 60, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(2, 92, Language(128), 61, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(3, 93, Language(129), 62, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 20
    If Manager.GraphBox(0).ThreeD = True Then
        Manager.GraphBox(0).ThreeD = False
    Else
        Manager.GraphBox(0).ThreeD = True
    End If
    Call Setup_Graphs_Like_First
Case 21
    If Manager.GraphBox(0).LineGraph = True Then
        Manager.GraphBox(0).LineGraph = False
    Else
        Manager.GraphBox(0).LineGraph = True
    End If
    Call Setup_Graphs_Like_First
Case 22
    If Manager.GraphBox(0).ShowLedgend = True Then
        Manager.GraphBox(0).ShowLedgend = False
    Else
        Manager.GraphBox(0).ShowLedgend = True
    End If
    Call Setup_Graphs_Like_First
Case 23
    Call Show_Controller_Box(6)
Case 25
    Call Manager.MenuBox(1).Clear_Entire_List
    For SettingsCount = 0 To UBound(LangFindList)
        Call Manager.MenuBox(1).Setup_Menu(SettingsCount, 180 + SettingsCount, Remove_The_Slash(LangFindList(SettingsCount)), 1, False, 0, Can_Menu(), False)
    Next SettingsCount
    If SettingsCount < 10 Then Call Manager.MenuBox(1).Setup_Menu(SettingsCount, 180 + SettingsCount, Language(141), 35, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 26
    If SectionSelect = 4 Then
        On Error Resume Next
        Call Manager.WeBrowser.ExecWB(OLECMDID_ENABLE_INTERACTION, OLECMDEXECOPT_DODEFAULT)
        Call Manager.WeBrowser.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT)
        On Error GoTo 0
    Else
        Call Show_Controller_Box(20)
    End If
Case 27
    Manager.WeBrowser.Refresh
Case 30
    Call Show_Controller_Box(12)
    Call Update_Connection_List
Case 31
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 100, Language(130), 63, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(1, 101, Language(131), 64, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(2, 102, Language(132), 65, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 32
    Call Check_NetCanConnect_Button
    If NetSearchList(0, 0) = Empty_Code Then
        ReDim ArbitoryList(1, 0)
        ArbitoryList(0, 0) = Language(133)
        Call Manager.FrameList(0).Submit_Data_Array(ArbitoryList(), -1, 1)
    End If
    Call Show_Controller_Box(4)
Case 33
    Call Show_Controller_Box(7)
Case 34
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 110, Language(134), 66, False, 0, Can_Menu(), False)
    For SettingsCount = 0 To UBound(RecentComList(), 2)
        If RecentComList(2, SettingsCount) <> Empty_Code Then
        If SettingsCount = 0 Then
            Call Manager.MenuBox(1).Setup_Menu(SettingsCount + 1, 111 + SettingsCount, RecentComList(0, SettingsCount), 1, False, 1, Can_Menu(), False)
        Else
            Call Manager.MenuBox(1).Setup_Menu(SettingsCount + 1, 111 + SettingsCount, RecentComList(0, SettingsCount), 1, False, 0, Can_Menu(), False)
        End If
        End If
    Next SettingsCount
    Call Manager_Display_Menu(Index, 2)
    
    'Call Manager.MenuBox(1).Setup_Menu(1, 111, "(recent list)", 1, False, 1, Can_Menu(), True)
    'Call Manager.MenuBox(1).Setup_Menu(2, 112, "Dedicated Server", 1, False, 0, Can_Menu(), False)
    'Call Manager_Display_Menu(Index, 2)
Case 35
    Call Open_Help_System(4)
Case 36
    Call All_Close_Session
Case 40
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 120, BenchDatArray(0, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 121, BenchDatArray(1, 0), -1, False, 0, Can_Menu(), True)
    Call Manager_Display_Menu(Index, 2)
Case 41
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 122, BenchDatArray(2, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 123, BenchDatArray(3, 0), -1, False, 0, Can_Menu(), True)
    Call Manager_Display_Menu(Index, 2)
Case 42
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 124, BenchDatArray(4, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 125, BenchDatArray(5, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 126, BenchDatArray(6, 0), -1, False, 0, Can_Menu(), True)
    Call Manager_Display_Menu(Index, 2)
Case 43
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 127, BenchDatArray(7, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 128, BenchDatArray(8, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 129, BenchDatArray(9, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(3, 130, BenchDatArray(10, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(4, 144, Language(135), 13, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 44
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 131, BenchDatArray(11, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 132, BenchDatArray(12, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 144, Language(135), 13, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 45
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 133, BenchDatArray(13, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 134, BenchDatArray(14, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 135, BenchDatArray(15, 0), -1, False, 0, Can_Menu(), True)
    Call Manager_Display_Menu(Index, 2)
Case 46
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 136, BenchDatArray(16, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 144, Language(135), 13, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 47
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 137, BenchDatArray(17, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 138, BenchDatArray(18, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 139, BenchDatArray(19, 0), -1, False, 0, Can_Menu(), True)
    'Call Manager.MenuBox(1).Setup_Menu(3, 147, Language(135), 13, False, 1, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 48
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 140, BenchDatArray(20, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(1, 141, BenchDatArray(21, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(2, 142, BenchDatArray(22, 0), -1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(3, 143, BenchDatArray(23, 0), -1, False, 0, Can_Menu(), True)
    Call Manager_Display_Menu(Index, 2)
Case 49, 200
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 160, Language(136), 69, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(1, 161, Language(137), 70, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(2, 162, Language(138), 1, False, 1, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(3, 163, Language(139), 1, False, 0, Can_Menu(), True)
    Call Manager.MenuBox(1).Setup_Menu(4, 164, Language(140), 1, False, 0, Can_Menu(), True)
    'Call Manager.MenuBox(1).Setup_Menu(3, 163, "Configure...", 0, False, 0, Can_Menu(), True)
    Select Case ClickType
    Case 0
        Call Manager_Display_Menu(Index, 2)
    Case 1
        Call Manager_Display_Menu(Index, 4)
    Case 2
        Call Manager_Display_Menu(Index, 3)
    End Select
Case 50, 51
    If MenuCmdID = 51 Then Call Select_Common_Benchmarks(1)
    Call Show_Controller_Box(9)
    Call Start_BenchMarks(False)
Case 52
    Call Show_Controller_Box(9)
    Call Start_BenchMarks(True)
Case 53, 144
    Call Show_Controller_Box(2)
    Call Check_Option_Enabled_States
Case 54
    Call Show_Controller_Box(20)
Case 60
    Call Show_Controller_Box(10)
    Call ATM_Process_System_Tasks
    Manager.Timer(1).Interval = 500
Case 61
    Call Switch_Sections_To(4)
    Call Browse_To_Site(4, Empty_Code)
Case 62
    Call Check_Password_Similarity
    Call Show_Controller_Box(18)
Case 63
    Call Show_Controller_Box(16)
Case 64
    Call Process_CommandButton_Click(43)
Case 65
    Call Show_Controller_Box(14)
Case 70
    Call Center_ControllerBoxes(-1)
Case 71
    Call Switch_Sections_To(0)
Case 72
    Call Switch_Sections_To(1)
Case 73
    Call Switch_Sections_To(2)
Case 74
    Call Switch_Sections_To(3)
Case 75
    Call Switch_Sections_To(4)
Case 76
    Call Switch_Sections_To(5)
Case 80
    Call Open_Help_System(0)
Case 81
    Call Open_Help_System(1)
Case 82
    Call Open_Help_System(2)
Case 83
    Call Open_Help_System(3)
Case 84
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 150, Language(141), 62, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(1, 151, Language(142), 61, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(2, 152, Language(143), 35, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(3, 153, Language(144), 60, False, 0, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 2)
Case 85
    Call Prepare_Tip_Display
    Call Show_Controller_Box(1)
Case 86
    Call Show_Controller_Box(19)
Case 90
    Call Save_RTF_Box(Add_The_Slash(GetSpecialFolder(GetDesktop)) & Language(250) & Richtext_Ext)
    Call Show_Msg_Window(Language(251), Language(252), 1)
Case 91
    Call Save_RTF_Box(Application_Path & Language(250) & Richtext_Ext)
    Call Begin_Mail_Recpient
    Call Show_Msg_Window(Language(251), Language(252), 1)
Case 92
    Call Save_RTF_Box(Add_The_Slash(GetSpecialFolder(GetDocuments)) & Language(250) & Richtext_Ext)
    Call Show_Msg_Window(Language(251), Language(252), 1)
Case 93
    Call Save_Current_Benchmark(Application_Path & Language(250) & Project_Ext)
    Call Begin_Mail_AndyFutcherSubmit
    Call Show_Msg_Window(Language(251), Language(252), 1)
Case 100, 101, 102
    If Manager.WriteBox(2).Text = Empty_Code Then Call Reset_Defaults(3)
    Call Show_Controller_Box(3)
    If MenuCmdID = 100 Then Manager.OptionBox(4).Value = True
    If MenuCmdID = 101 Then Manager.OptionBox(4).Value = False
    If MenuCmdID = 102 Then Manager.OptionBox(5).Value = True
Case 110
    Call Control_Enable_Group(1, False)
    Call Check_NetCanConnect_Button
    Call Show_Controller_Box(23)
    Call List_WWW_Hosts(1)
Case 111, 112, 113, 114, 115, 116, 117, 118, 119
    Net_ServAddress = RecentComList(1, MenuCmdID - 111)
    Net_PortNumber = RecentComList(2, MenuCmdID - 111)
    Net_InterNet = RecentComList(3, MenuCmdID - 111)
    Call Multi_Join_Session(-2)
Case 127, 128, 129, 130
    If BenchSelArray(7) = True Or BenchSelArray(8) = True Or BenchSelArray(9) = True Or BenchSelArray(10) = True Then
        Manager.OptionBox(1).ValueDesc = True
    Else
        Manager.OptionBox(1).ValueDesc = False
    End If
    Call Check_Option_Enabled_States
Case 131, 132
    If BenchSelArray(11) = True Or BenchSelArray(12) = True Then
        Manager.OptionBox(2).ValueDesc = True
    Else
        Manager.OptionBox(2).ValueDesc = False
    End If
    Call Check_Option_Enabled_States
Case 136
    If BenchSelArray(16) = True Then
        Manager.OptionBox(3).ValueDesc = True
    Else
        Manager.OptionBox(3).ValueDesc = False
    End If
    Call Check_Option_Enabled_States
Case 150
    Call Browse_To_Site(1, Empty_Code)
    Call Switch_Sections_To(4)
Case 151
    Call Browse_To_Site(2, Empty_Code)
    Call Switch_Sections_To(4)
Case 152
    Call Browse_To_Site(3, Empty_Code)
    Call Switch_Sections_To(4)
Case 153
    Call Begin_Mail_TechSupport
Case 160
    PathString = Show_Common_Dialogue(0, Language(145), Language(146), Application_Path)
    If PathString <> Empty_Code Then Call Load_Current_Preset(PathString)
Case 161
    PathString = Show_Common_Dialogue(1, Language(147), Language(146), Application_Path)
    If PathString <> Empty_Code Then Call Save_Current_Preset(PathString)
Case 162
    Call Select_Common_Benchmarks(1)
Case 163
    Call Select_Common_Benchmarks(0)
Case 164
    Call Select_Common_Benchmarks(2)
Case 170, 171, 172, 173, 174, 175, 176, 177, 178, 179
    If MenuCmdID <= (UBound(ThemeFindList) + 170) Then
        Theme_Current = Add_The_Slash(ThemeFindList(MenuCmdID - 170))
        Call Prepare_Controls
    Else
        Call Browse_To_Site(3, Empty_Code)
        Call Switch_Sections_To(4)
    End If
Case 180, 181, 182, 183, 184, 185, 186, 187, 188, 189
    If MenuCmdID <= (UBound(LangFindList) + 180) Then
        Language_Current = Add_The_Slash(LangFindList(MenuCmdID - 180))
        Call Prepare_Language
    Else
        Call Browse_To_Site(3, Empty_Code)
        Call Switch_Sections_To(4)
    End If
Case 190
    Manager.WriteBox(5).Text = "/" & Language(264)
    Call Manager.FrameButton(0).ClickNow
Case 191
    Manager.WriteBox(5).Text = "/" & Language(265)
    Call Manager.FrameButton(0).ClickNow
Case 192
    Manager.WriteBox(5).Text = "/" & Language(270) & Space_Code
    Manager.WriteBox(5).SelectionStart = Len(Manager.WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
Case 193
    Manager.WriteBox(5).Text = "/" & Language(271) & Space_Code
    Manager.WriteBox(5).SelectionStart = Len(Manager.WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
Case 194
    Manager.WriteBox(5).Text = "/" & Language(282) & Space_Code
    Manager.WriteBox(5).SelectionStart = Len(Manager.WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
Case 195
    Manager.WriteBox(5).Text = "/" & Language(283) & Space_Code
    Manager.WriteBox(5).SelectionStart = Len(Manager.WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
End Select
Call Check_All_Menus
End Sub

Public Sub Process_Dynamic_Command_Ids(Index As Integer, MenuCmdID As Integer)
If MenuCmdID >= 120 And MenuCmdID <= 143 Then
    If BenchSelArray(MenuCmdID - Bench_ID_Start) = True Then
        BenchSelArray(MenuCmdID - Bench_ID_Start) = False
    Else
        BenchSelArray(MenuCmdID - Bench_ID_Start) = True
    End If
    Call Manager.MenuBox(1).Check_Selection_Value(Index, MenuCmdID)
End If
End Sub
