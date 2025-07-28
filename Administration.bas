Attribute VB_Name = "Administration"
Option Explicit

Public Sub All_Close_Session()
Net_ServerType = False
Net_ClientType = False
Call Unregister_WWW_Host
Call Disconnect_Me_Client
Net_Single = False
Call Reset_Data_Arrays(-1)
Call Switch_Sections_To(5)
Call Check_All_Menus
End Sub

Public Sub Disconnect_Me_Client()
If Net_ServerType = True Or Net_ClientType = True Then
    Net_Single = True
End If
Net_ServerType = False
Net_ClientType = False
Call Change_Benchmark_State(0)
Call Hide_Controller_Box(9)
Call Net_Disconnect(-1)
StatusMenus(1) = Language(39)
Call Update_StatusBars
End Sub

Public Sub Single_New_Session()
Call Update_System_Information
Net_ServerType = False
Net_ClientType = False
Net_Single = True
Net_InterNet = False
Call Unregister_WWW_Host
Call Reset_Data_Arrays(-1)
Call Net_UpDataUsers
Call Net_Disconnect(-1)
Call Chat_Score_Reset(0)
End Sub

Public Sub Multi_New_Session()
Call Update_System_Information
'MsgBox Manager.WriteBox(21).ClickTag
Call Get_Network_Defaults
Call Unregister_WWW_Host
If Manager.OptionBox(4).Value = True And Manager.WriteBox(21).ClickTag <> LAN_Code Then
    If Def_WWWAddress <> Empty_Code Then
        Call Register_Host_WWW
    Else
        Call Show_Msg_Window(Language(165), Language(166), 0)
        GoTo Ed
    End If
End If
Net_ServerType = True
Net_ClientType = False
Net_Single = False
Net_FindType = False
Call Reset_Data_Arrays(-1)
Call Hide_Controller_Box(20)
ReDim Net_Buffer(2, 0)

Call Begin_Host_System
StatusMenus(0) = Language(167)

Call Update_StatusBars
Call Net_UpDataUsers
Call Chat_Score_Reset(0)
Ed: End Sub

Public Sub Multi_Join_Session(SourceIPFrom As Integer)
Call Update_System_Information
Net_ServerType = False
Net_ClientType = True
Net_Single = False
Net_FindType = False
Call Unregister_WWW_Host
Call Hide_Controller_Box(20)
Call Reset_Data_Arrays(-1)
ReDim Net_Buffer(2, 0)

Call Chat_Score_Reset(0)
Call Begin_Join_System(SourceIPFrom)
End Sub

Public Sub Multi_Find_Session()
Call Get_Network_Defaults
If Manager.WriteBox(4).ClickTag = LAN_Code Then
    If Def_IPAddress = Empty_Code Then
        Call Show_Msg_Window(Language(249), Language(156) & " 100003", 0)
        GoTo Ed
    End If
    If Net_Single = True Or Net_ServerType = True Or Net_ClientType = True Then If Show_Question_Window(Language(168), Language(166), 0) = False Then GoTo Ed
    Call Unregister_WWW_Host
    Call All_Close_Session
    Call Control_Enable_Group(0, False)
    Net_ServerType = False
    Net_ClientType = True
    Net_Single = False
    Net_FindType = True
    ReDim Net_Buffer(2, 0)
    Call Begin_Find_System
Else
    Call Control_Enable_Group(0, False)
    Call List_WWW_Hosts(0)
End If
Ed: End Sub

Public Sub Control_Enable_Group(EnaIndex As Integer, EnaValue As Boolean)
Select Case EnaIndex
Case 0
    Manager.WriteBox(4).Enabled = EnaValue
    Manager.CommandButton(10).Enabled = EnaValue
    Manager.CommandButton(11).Enabled = EnaValue
    Manager.CommandButton(12).Enabled = EnaValue
    Manager.FrameLabel(8).Visible = EnaValue
    Manager.FrameList(0).Visible = EnaValue
    Call Check_NetCanConnect_Button
Case 1
    Manager.FrameList(6).Visible = EnaValue
    Manager.CommandButton(61).Enabled = EnaValue
    Manager.CommandButton(62).Enabled = EnaValue
End Select
End Sub

Public Sub Begin_Join_System(SourceIPFrom As Integer)
StatusMenus(1) = Language(169)
Call Update_StatusBars
Select Case SourceIPFrom
Case -1
    Net_ServAddress = Manager.WriteBox(8).Text
    Net_PortNumber = Int(Val(Manager.WriteBox(9).Text))
    Net_InterNet = Manager.OptionBox(13).Value
    Call Add_To_RecentList(Net_ServAddress & Language(11))
Case 0
    Net_ServAddress = NetSearchList(2, Manager.FrameList(0).ListIndex)
    Net_PortNumber = NetSearchList(3, Manager.FrameList(0).ListIndex)
    If Manager.WriteBox(4).ClickTag = LAN_Code Then
        Net_InterNet = False
    Else
        Net_InterNet = True
    End If
    Call Add_To_RecentList(NetSearchList(0, Manager.FrameList(0).ListIndex))
Case 1
    Net_ServAddress = NetSearchCom(3, Manager.FrameList(6).ListIndex)
    Net_PortNumber = NetSearchCom(2, Manager.FrameList(6).ListIndex)
    Net_InterNet = True
    Call Add_To_RecentList(NetSearchCom(0, Manager.FrameList(6).ListIndex))
End Select
Net_Password = Empty_Code

Call Net_Connect
End Sub

Public Sub Begin_Find_System()
StatusMenus(1) = Language(170)
Call Update_StatusBars
Net_ClientType = True
Net_DoTwice = False
Net_SearchType = True
Net_InterNet = False

Call Net_Find_Local_Servers
End Sub

Public Sub Begin_Host_System()
Net_ServerName = Replace(Manager.WriteBox(1).Text, Net_WWWScan, Empty_Code)
Net_PortNumber = Int(Val(Manager.WriteBox(2).Text))
Net_Password = LCase(Trim(Manager.WriteBox(3).Text))
Net_InterNet = Manager.OptionBox(4).Value
Net_Dedicated = Manager.OptionBox(5).Value
Net_Public = Manager.OptionBox(6).Value
Net_LangFilter = Manager.OptionBox(7).Value

Call Net_Disconnect(-1)
Call NetSub_Next_Listener
End Sub

Public Sub Refresh_Benchmark_Configuration()
Dim RealDrivePath As String
For ControlCount = 0 To ManagerSub.GenDrive.ListCount - 1
    RealDrivePath = Left(ManagerSub.GenDrive.List(ControlCount), 2)
    Select Case Drv_Type_Code(RealDrivePath)
    Case 2
        Call Manager.WriteBox(11).DDList_Add(RealDrivePath & " (" & LCase(Drv_Type(RealDrivePath)) & Space_Code & Language(225) & ")", RealDrivePath & BackSlash_Code)
    Case 4
        Call Manager.WriteBox(12).DDList_Add(RealDrivePath & " (" & LCase(Drv_Type(RealDrivePath)) & Space_Code & Language(225) & ")", RealDrivePath & BackSlash_Code)
    End Select
Next ControlCount
Call Process_WriteList_Click(11)
Call Process_WriteList_Click(12)
End Sub

Public Sub Chat_Score_Reset(ResetType As Integer)
Select Case ResetType
Case 0
    Call Reset_Data_Limits(False)
    ReDim ChatLineInfo(1, 0)
    Call Chat_AddSay(Space_Code, App_Title & Space_Code & App_Ver & Space_Code & ConnectedUsers(1, 0) & Space_Code & Language(253))
    Call Chat_AddSay(Space_Code, "Copyright � 1999-2003 Andy Futcher")
    Call Chat_AddSay(Space_Code, Space_Code)
Case 1
    ChatLineInfo(1, 0) = App_Title & Space_Code & App_Ver & Space_Code & ConnectedUsers(1, 0) & Space_Code & Language(253)
    Call Manager.ChatterBox(0).Submit_Data_Array(ChatLineInfo())
End Select
End Sub

Public Sub Chat_AddSay(WhoString As String, SayString As String)
Call ChatSub_AddIndex
ChatLineInfo(0, UBound(ChatLineInfo, 2)) = WhoString
ChatLineInfo(1, UBound(ChatLineInfo, 2)) = FilterBadLang(SayString)
Call Manager.ChatterBox(0).Submit_Data_Array(ChatLineInfo())
End Sub

Public Sub Chat_AddArray(TargetArray() As String)
For NetCount = 0 To UBound(TargetArray(), 2)
    If TargetArray(0, NetCount) <> Space_Code Then
        Call ChatSub_AddIndex
        ChatLineInfo(0, UBound(ChatLineInfo, 2)) = TargetArray(0, NetCount)
        ChatLineInfo(1, UBound(ChatLineInfo, 2)) = FilterBadLang(TargetArray(1, NetCount))
    End If
Next NetCount
Call Manager.ChatterBox(0).Submit_Data_Array(ChatLineInfo())
End Sub

Public Sub ChatSub_AddIndex()
If UBound(ChatLineInfo(), 2) > Val(Manager.WriteBox(22).Text) Then
    For NetCount = 3 To UBound(ChatLineInfo(), 2) - 1
        ChatLineInfo(0, NetCount) = ChatLineInfo(0, NetCount + 1)
        ChatLineInfo(1, NetCount) = ChatLineInfo(1, NetCount + 1)
    Next NetCount
Else
    Call Add_Index_To_StringArray(ChatLineInfo())
End If
End Sub
