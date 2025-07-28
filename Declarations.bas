Attribute VB_Name = "Declarations"
Option Explicit

Global ToolTipArray(200) As String, BenchSelArray(23) As Boolean, BenchDatArray(23, 4) As String
Global Language() As String, TipTextArray() As String, TipTextIndex As Integer, ATXFilePath As String

Global BenchResults() As String, PlatFormUsers() As String, ChatLineInfo() As String
Global ATMProcessInfo() As String, ConnectedList() As String, BenchTaskList() As String
Global URLBenchList() As String, ConnStatusList() As String, NetSearchList() As String
Global ArbitoryList() As String, BenResultList() As String, ConnectedUsers() As String
Global PlatformList() As String, ThemeFindList() As String, LangFindList() As String
Global NetSearchCom() As String, RecentComList(3, 7) As String, BadLangList() As String

Global Net_ServerType As Boolean, Net_InterNet As Boolean, Net_Dedicated As Boolean
Global Net_PortNumber As Integer, Net_ServerName As String, Net_Password As String
Global Net_Public As Boolean, Net_LangFilter As Boolean, Net_ServAddress As String
Global Net_ClientType As Boolean, Net_WebAccess(1) As Integer, Net_SearchType As Boolean
Global Net_WaitState() As Boolean, Net_Buffer() As String, Net_DoTwice As Boolean
Global Net_WWWId As Long, Net_CantBench As Boolean, Net_StillBusy As Boolean
Global Net_PingNow As Boolean, Net_Send As Boolean, Net_Single As Boolean
Global Net_KickList() As String, Net_VoteList() As Integer, Net_FindType As Boolean

Global Def_FriendlyName As String, Def_IPAddress As String, Def_WWWAddress As String
Global Web_OurSite As String, OurSite_Counter As Integer, AlreadyOnline As String
Global EndUpdatePath As String, Date_ERT As String, Date_Left As Integer

Global SectionSelect As Integer, TasksBusy As Boolean, KeyTapp As Boolean, IsFullVer As Boolean
Global WebBusy As Boolean, StopBenchmarks As Boolean, KeyHeld As Boolean
Global MouseMoved As Boolean, MouseClicked As Boolean, DoBurnin As Boolean

Global Can_Menu(200) As Boolean, StatusMenus(2) As String, TimerState(5) As Integer
Global App_NewVer(2) As String, App_UserName As String, App_CompName As String

Global Desktop_Top As Integer, Desktop_Left As Integer, Desktop_Width As Integer, Desktop_Height As Integer
Global SupposedCaption As String, DotXThreeWidth As Integer, App_Ver As String, LastBenType As Boolean
Global LastKeyCode As Integer, LastExplorerEnd As Integer, DefPosHolder As Integer, NormHolder As Integer
Global MouseOverMoved As String, HtmlData As String, GenBrowserErr(1) As String, PathString As String
Global BenchPassword As String, SavePath As String, SaveFolder As String, LoadPath As String
Global Language_Current As String, LastTipShown As String, FullVersion As Boolean, DuelOutput As String
Global BannedList As String, CanAcceptCode As String, Text_Height As Long

Global Theme_Current As String, Theme_Text As String, Theme_Font As String, Theme_Icon As Long
Global Theme_Invert As Long, Theme_InvertLight As Long, Theme_High As Long, Theme_HighLight As Long
Global Theme_Light As Long, Theme_Shade As Long, Theme_Vague As Long, Theme_Color As Long
Global Theme_Shadow As Long, Theme_Dark As Long, Theme_Pitch As Long, Theme_Vague_R As Integer
Global Theme_Vague_G As Integer, Theme_Vague_B As Integer

Global DoneItOnce As Boolean, UserDoneOnce As Boolean, SubDoneOnce As Boolean, TimedItOnce As Boolean
Global BenchDoneOnce As Boolean, UrlDoneOnce As Boolean, Timed1Once As Boolean

Global ControlCount As Integer, VisualCount As Integer, ExplorerCount As Integer, UserCount As Integer
Global MenuCount As Integer, SettingsCount As Integer, ServerCount As Integer, NetCount As Integer
Global CheckCount As Integer, WebCount As Integer

Global Const Theme_Path = "AndyFutcher\engine001\"
Global Const Sample_Path = "data\samples\"
Global Const Update_Path = "data\updates\"
Global Const Data_Path = "data\"
Global Const Lang_Path = "data\languages\"
Global Const ResourceA_Name = "resourcea"
Global Const ResourceB_Name = "resourceb"
Global Const Resource_Name = "resource"
Global Const Config_Name = "config"
Global Const Sample_Name = "sample"
Global Const Texture_Name = "texture"
Global Const Graphic_Name = "graphic"
Global Const Winsock_Name = "wnsckrg"
Global Const Winsock2_Name = "wnscklg"
Global Const Bench_Name = "sxpdata"
Global Const Update_Name = "update.exe"
Global Const Bench_Ext = ".sxptemp"
Global Const CPUFile_Name = "sxpcpudb"
Global Const BWFFile_Name = "sxpbadlg"
Global Const Photo_Name = "WWWPhoto.dat"
Global Const Resource_Ext = ".dat"
Global Const Richtext_Ext = ".rtf"
Global Const Project_Ext = ".sxp"
Global Const Logging_Ext = ".log"

Global Const INet_TimeoutSmall = 25
Global Const Bench_ID_Start = 120
Global Const Image_Checked = 67
Global Const Image_Unchecked = 68
Global Const Allowed_Size = 1440
Global Const NetCountAscii = 97
Global Const Net_DefPort = 27960
Global Const MaxSave_Long = 2097152

Global Const Chat_BadLang = "�"
Global Const Dagger_Char = "�"
Global Const Bullet_Char = "*"
Global Const UnReg_Char = "*"
Global Const JumpTo_Char = "&"
Global Const LAN_Code = "lan"
Global Const HTML_Enter = "</b>"

Global Const DoubleAsterix_Code = "**"
Global Const Bad_Code = "�"
Global Const Net_Code = "�"
Global Const Empty_Code = ""
Global Const Colon_Code = ":"
Global Const Space_Code = " "
Global Const Apost_Code = "'"
Global Const FullStop_Code = "."
Global Const BackSlash_Code = "\"
Global Const One_Code = "1"
Global Const Zero_Code = "0"

'Global Const InvCommas_Code = """"
Global Const App_Title = "AndyFutcher� SynthMark� XP"

'Global Const Web_OurSite = "http://www.AndyFutcher.com/"
Global Const Web_IniFile = "synthxp.ini"
Global Const Web_Http = "http://"
Global Const Net_WWWScan = ";"
Global Const Net_WWWList = "db/list.php?country="
Global Const Net_WWWListAll = "db/list.php"
Global Const Net_WWWCommunity = "synthcom.ini"
Global Const Net_WWWName = "db/add.php?name="
Global Const Net_WWWRem = "db/remove.php?id="
Global Const Net_WWWIp = "&ip="
Global Const Net_WWWPort = "&port="
Global Const Net_WWWProtected = "&protected="
Global Const Net_WWWRegion = "&country="
Global Const Net_DefCant = "cannot find server"

Global Const Net_UpdateURL = "update/synthxp.upd"
Global Const Net_CPUUpdate = "update/sxpcpudb.upd"
Global Const Net_BWFUpdate = "update/sxpbadlg.upd"

Global Const WindowsXP_Code = 51
Global Const Windows2K_Code = 50
Global Const WindowsME_Code = 49
Global Const Windows98_Code = 41
Global Const Windows95_Code = 40
Global Const WindowsNT_Code = 39

Global Const Reg_Global = "SOFTWARE\Andy Futcher"
Global Const Reg_DefAddress = "SOFTWARE\Andy Futcher\SynthMark XP"
Global Const Reg_CPUAddress = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Global Const Reg_CPUVendor = "VendorIdentifier"
Global Const Reg_CPUName = "ProcessorNameString"
Global Const Reg_LastUpdateDay = "LastUpdateDay"
Global Const Reg_VersionNumber = "VersionNumber"
Global Const Reg_ExePath = "ExePath"
Global Const Reg_FirstTimeRun = "FirstTimeRun"
Global Const Reg_WindowsState = "WindowsState"
Global Const Reg_ExplorerShow = "ExplorerShow"
Global Const Reg_SelectedPort = "SelectedPort"
Global Const Reg_ServerName = "ServerName"
Global Const Reg_ListPublicly = "ListPublicly"
Global Const Reg_ChatCatchup = "ChatCatchup"
Global Const Reg_ShowTips = "ShowTips"
Global Const Reg_NetWorkSpeed = "NetWorkSpeed"
Global Const Reg_Compression = "Compression"
Global Const Reg_ConnectionType = "ConnectionType"
Global Const Reg_ProxyAddress = "ProxyAddress"
Global Const Reg_ProxyPort = "ProxyPort"
Global Const Reg_BadWords = "BadWordFilter"
Global Const Reg_AlwaysConnect = "AlwaysConnect"
Global Const Reg_CheckUpdates = "CheckForUpdates"
Global Const Reg_MaxChatLines = "MaxChatLines"
Global Const Reg_MaxScoreLines = "MaxScoreLines"
Global Const Reg_MaxGraphUsers = "MaxGraphUsers"
Global Const Reg_UserName = "UserName"
Global Const Reg_UserComment = "UserComment"
Global Const Reg_QuickHelp = "QuickHelp"
Global Const Reg_ThreeDGraphs = "ThreeDeeGraphs"
Global Const Reg_LineGraphs = "LineGraphs"
Global Const Reg_ShowLegend = "ShowLegend"
Global Const Reg_LastSavePath = "LastSavePath"
Global Const Reg_PrintChat = "PrintChat"
Global Const Reg_PrintScores = "PrintScores"
Global Const Reg_PrintGraphs = "PrintGraphs"
Global Const Reg_PrintHWInfo = "PrintHWInfo"
Global Const Reg_CurrentTheme = "CurrentTheme"
Global Const Reg_CurrentLang = "CurrentLang"
Global Const Reg_SentCPUInfo = "SentCPUInfo"
Global Const Reg_OurSiteCounter = "OurSiteCounter"
Global Const Reg_CloseSuccess = "CloseSuccess"
Global Const Reg_CDKey = "CDKey"
Global Const Reg_StartDate = "StartDate"
Global Const Reg_GUID = "GUID"
Global Const Reg_ATXPath = "ATXPath"
Global Const Reg_ATXCPUInfo = "ATXCPUInfo"

Global Const Save_Version = "verz"
Global Const Save_Password = "pass"
Global Const Save_Chat = "chat"
Global Const Save_Scores = "sc"

Global Const App_Trial = "Trial"
Global Const App_Timed = "Timed"
Global Const App_Full = "Full"

Global Const Cntrl_OptionBox = "optionbox"
Global Const Cntrl_WriteBox = "writebox"
Global Const Cntrl_BenchSel = "benchsel"

Public Sub Save_Settings_Data(TypeIndex As Integer)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_VersionNumber, App_Ver)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ExePath, Application_Path)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_WindowsState, Manager.WindowState)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_SelectedPort, Val(Manager.WriteBox(2).Text))
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ServerName, Manager.WriteBox(1).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ListPublicly, Manager.OptionBox(6).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ChatCatchup, Manager.OptionBox(7).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ShowTips, Manager.OptionBox(0).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_NetWorkSpeed, Manager.WriteBox(16).ListIndex)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_Compression, Manager.WriteBox(17).ListIndex)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ConnectionType, Manager.WriteBox(18).ListIndex)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ProxyAddress, Manager.WriteBox(19).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ProxyPort, Manager.WriteBox(20).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_BadWords, Manager.OptionBox(17).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_AlwaysConnect, Manager.OptionBox(18).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CheckUpdates, Manager.OptionBox(19).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxChatLines, Manager.WriteBox(22).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxScoreLines, Manager.WriteBox(23).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxGraphUsers, Manager.WriteBox(24).ListIndex)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_UserName, Manager.WriteBox(6).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_UserComment, Manager.WriteBox(7).Text)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_QuickHelp, Manager.OptionBox(14).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ThreeDGraphs, Manager.GraphBox(0).ThreeD)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LineGraphs, Manager.GraphBox(0).LineGraph)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ShowLegend, Manager.GraphBox(0).ShowLedgend)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LastSavePath, SaveFolder)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintChat, Manager.OptionBox(8).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintScores, Manager.OptionBox(9).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintGraphs, Manager.OptionBox(10).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintHWInfo, Manager.OptionBox(11).Value)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CurrentTheme, Theme_Current)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_Global, Reg_CurrentTheme, Theme_Current)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CurrentLang, Language_Current)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_GUID, ConnectedUsers(2, 0))
If Generate_Input(ConnectedUsers(3, 0)) = True And CanAcceptCode = True Then
    Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CDKey, ConnectedUsers(3, 0))
Else
    Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CDKey, Empty_Code)
End If
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_StartDate, Date_ERT)

For ControlCount = 0 To Manager.ExplorerFrame.Count - 1
    Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ExplorerShow & ControlCount, Manager.ExplorerFrame(ControlCount).ShowPanel)
Next ControlCount
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_FirstTimeRun, Zero_Code)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_OurSiteCounter, OurSite_Counter)
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CloseSuccess, Manager.OptionBox(23).Value)

'Call Convert_Array_To_String(RecentComList(), SaveString)
'Call Convert_String_ByteArray(SaveString, FileArray())
'Call CompressData(FileArray(), CompressLevel)
'Call Kill_File(FilePath)
'Call Write_Array_Into_File(FilePath, FileArray())
End Sub

Public Sub Load_Settings_Data(TypeIndex As Integer)
Dim FirstTimeString As String
FirstTimeString = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_FirstTimeRun)
If FirstTimeString = One_Code Or FirstTimeString = Empty_Code Then
    Call Reset_All_Defaults
    GoTo Ed
End If

For ControlCount = 0 To Manager.ExplorerFrame.Count - 1
    Manager.ExplorerFrame(ControlCount).ShowPanel = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ExplorerShow & ControlCount)
Next ControlCount
Manager.WindowState = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_WindowsState)
Manager.WriteBox(2).Text = Val(QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_SelectedPort))
Manager.WriteBox(9).Text = Manager.WriteBox(2).Text
Manager.WriteBox(1).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ServerName)
Manager.OptionBox(6).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ListPublicly)
Manager.OptionBox(7).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ChatCatchup)
Manager.OptionBox(0).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ShowTips)
Manager.WriteBox(16).ListIndex = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_NetWorkSpeed)
Manager.WriteBox(17).ListIndex = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_Compression)
Manager.WriteBox(18).ListIndex = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ConnectionType)
Manager.WriteBox(19).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ProxyAddress)
Manager.WriteBox(20).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ProxyPort)
Manager.OptionBox(17).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_BadWords)
Manager.OptionBox(18).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_AlwaysConnect)
Manager.OptionBox(19).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CheckUpdates)
Manager.WriteBox(22).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxChatLines)
Manager.WriteBox(23).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxScoreLines)
Manager.WriteBox(24).ListIndex = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_MaxGraphUsers)
Manager.WriteBox(6).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_UserName)
Manager.WriteBox(7).Text = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_UserComment)
Manager.OptionBox(14).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_QuickHelp)
Manager.GraphBox(0).ThreeD = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ThreeDGraphs)
Manager.GraphBox(0).LineGraph = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LineGraphs)
Manager.GraphBox(0).ShowLedgend = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_ShowLegend)
SaveFolder = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LastSavePath)
Manager.OptionBox(8).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintChat)
Manager.OptionBox(9).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintScores)
Manager.OptionBox(10).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintGraphs)
Manager.OptionBox(11).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_PrintHWInfo)
ConnectedUsers(2, 0) = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_GUID)
OurSite_Counter = Val(QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_OurSiteCounter))
On Error Resume Next
Manager.OptionBox(23).Value = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CloseSuccess)

On Error GoTo 0

Ed: Call Update_System_Information
ATXFilePath = QueryValue(HKEY_LOCAL_MACHINE, Reg_Global, Reg_ATXPath)
End Sub

Public Sub Check_OS_CPUDescription()
If InStr(1, CPUBitDesc, Language(8)) <> 0 Then
    If QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_SentCPUInfo) = One_Code Then GoTo Ed
    Manager.FrameLabel(39).Caption = Trim(QueryValue(HKEY_LOCAL_MACHINE, Reg_CPUAddress, Reg_CPUVendor)) & Space_Code & Trim(QueryValue(HKEY_LOCAL_MACHINE, Reg_CPUAddress, Reg_CPUName))
    Call Show_Controller_Box(17)
End If
Ed: End Sub

Public Sub Reset_All_Defaults()
For SettingsCount = 0 To Manager.ControllerBox.Count - 1
    Call Reset_Defaults(SettingsCount)
Next SettingsCount
End Sub

Public Sub Reset_Defaults(Index As Integer)
Select Case Index
Case 0
    Manager.ExplorerFrame(0).ShowPanel = True
    Manager.ExplorerFrame(1).ShowPanel = True
    Manager.ExplorerFrame(2).ShowPanel = True
    Manager.ExplorerFrame(3).ShowPanel = True
    Manager.ExplorerFrame(5).ShowPanel = True
    Manager.ExplorerFrame(6).ShowPanel = True
    Manager.ExplorerFrame(8).ShowPanel = True
Case 1
    Manager.OptionBox(0).Value = True
Case 3
    Manager.WriteBox(3).Text = Empty_Code
    Manager.WriteBox(2).Text = Net_DefPort
    Manager.WriteBox(1).Text = Normalize(App_UserName) & Language(236)
    Manager.OptionBox(6).Value = True
    Manager.OptionBox(7).Value = True
Case 5
    Manager.OptionBox(8).Value = True
    Manager.OptionBox(9).Value = True
    Manager.OptionBox(10).Value = True
    Manager.OptionBox(11).Value = True
Case 7
    Manager.WriteBox(9).Text = Net_DefPort
Case 9
    Manager.OptionBox(23).Value = True
Case 14
    Manager.OptionBox(17).Value = True
    Manager.OptionBox(19).Value = True
Case 16
    Manager.WriteBox(22).Text = "100"
    Manager.WriteBox(23).Text = "50"
    Manager.WriteBox(24).ListIndex = 9
End Select
End Sub

Public Sub Reset_Data_Arrays(RstType As Integer)
Select Case RstType
Case 0
    ReDim ConnectedUsers(3, 255)
    ReDim Net_KickList(255)
    ReDim BenchResults(23, 5, 99)
    ReDim PlatFormUsers(5, 255)
    ReDim ChatLineInfo(1, 0)
    ReDim Net_WaitState(ManagerSub.Winsock.Count - 1)
    ReDim Net_VoteList(255)
    ReDim NetSearchList(3, 0)
    ReDim NetSearchCom(5, 0)
    ReDim BadLangList(0)
    BannedList = Empty_Code
End Select
ReDim ConnectedList(2, 0)
ReDim PlatformList(5, 0)
ReDim URLBenchList(1, 0)
ReDim ArbitoryList(1, 0)
ReDim BenResultList(5)
Net_CantBench = False
BenchPassword = Empty_Code
SavePath = Empty_Code
End Sub

Public Sub Reset_Data_Limits(Prezerv As Boolean)
If Prezerv = True Then
    ReDim Preserve BenchResults(23, 5, Val(Manager.WriteBox(23).Text) - 1)
Else
    ReDim BenchResults(23, 5, Val(Manager.WriteBox(23).Text) - 1)
End If
End Sub

Public Sub Load_Current_Benchmark(FilePath As String)
Call Choose_Manager_Functionality(False, -1, 1)
Call Change_StatusBar_Text(3, FilePath)

Dim FileString As String, FileArray() As Byte, SavePacket As String, SaveCommand As String
Dim FileVersion As Integer, ProcessArray() As String, TargetSection As Integer
Dim ColumnCount As Integer, IndexCount As Integer, CheckPass As String

Call Load_File_Into_Array(FilePath, FileArray())
Call DeCompressData(FileArray(), MaxSave_Long)
Call Convert_ByteArray_To_String(FileArray(), FileString)

Do While InStr(1, FileString, Net_PakTerminator) <> 0
    SavePacket = Left(FileString, InStr(1, FileString, Net_PakTerminator) - 1)
    FileString = Right(FileString, Len(FileString) - InStr(1, FileString, Net_PakTerminator))
    SaveCommand = Left(SavePacket, 4)
    SavePacket = Right(SavePacket, Len(SavePacket) - 4)
    Select Case SaveCommand
    Case Save_Version
        FileVersion = Filter_Sort(SavePacket)
    Case Save_Password
        'Call Switch_Sections_To(5)
        Do While LCase(SavePacket) <> LCase(CheckPass)
            CheckPass = Get_Password
            If CheckPass = Bad_Code Then GoTo Ed
            'If LCase(SavePacket) <> LCase(CheckPass) Then
            '    Call Show_Msg_Window("The password you entered was incorrect? Please try again?", "Invalid Password", 0)
            'End If
        Loop
        BenchPassword = SavePacket
    Case Save_Chat
        Call Convert_String_To_Array(SavePacket, ProcessArray())
        Call Chat_AddArray(ProcessArray())
    End Select
    If Left(SaveCommand, 2) = Save_Scores Then
        TargetSection = Val(Right(SaveCommand, 2))
        Call Convert_String_To_Array(SavePacket, ProcessArray())
        'Call Chat_AddArray(ProcessArray())
        Call Flow_Triple_Array(BenchResults(), TargetSection)
        For IndexCount = 0 To UBound(ProcessArray(), 2)
            For ColumnCount = 0 To UBound(BenchResults, 2)
                BenchResults(TargetSection, ColumnCount, IndexCount) = ProcessArray(ColumnCount, IndexCount)
            Next ColumnCount
        Next IndexCount
        Call Manager.ComplexList(TargetSection).Submit_Data_Array(BenchResults(), TargetSection, 5)
        Call Graph_Update(TargetSection)
    End If
Loop
SaveFolder = Give_Path_Name_Only(FilePath)
SavePath = FilePath
Call Change_StatusBar_Text(2, Empty_Code)
Call Choose_Manager_Functionality(True, -1, 0)
Call Align_Selected_Complex_Controls
If FileVersion > Filter_Sort(App_Ver) Then
    Call Show_Msg_Window(Language(237), Language(238), 1)
End If
Ed: End Sub

Public Sub Save_Current_Benchmark(FilePath As String)
Call Choose_Manager_Functionality(False, -1, 0)
Call Change_StatusBar_Text(1, FilePath)

Dim FileString As String, FileArray() As Byte, ProcessString As String, CompressLevel As Long
CompressLevel = Manager.WriteBox(17).ClickTag

FileString = Save_Version & Filter_Sort(App_Ver) & Net_PakTerminator
If BenchPassword <> Empty_Code Then FileString = FileString & Save_Password & LCase(BenchPassword) & Net_PakTerminator
Call Convert_Array_To_String(ChatLineInfo(), ProcessString)
FileString = FileString & Save_Chat & ProcessString & Net_PakTerminator
For SettingsCount = 0 To UBound(BenchResults(), 1)
    If BenchResults(SettingsCount, 0, 0) <> Empty_Code Then
        Call Convert_Triple_Array_To_String(BenchResults(), ProcessString, SettingsCount)
        FileString = FileString & Save_Scores & Format(SettingsCount, "00") & ProcessString & Net_PakTerminator
    End If
Next SettingsCount

Call Convert_String_ByteArray(FileString, FileArray())
Call CompressData(FileArray(), CompressLevel)
Call Kill_File(FilePath)
Call Write_Array_Into_File(FilePath, FileArray())

SaveFolder = Give_Path_Name_Only(FilePath)
SavePath = FilePath
Call Change_StatusBar_Text(2, Empty_Code)
Call Choose_Manager_Functionality(True, -1, 0)
End Sub

Public Sub Save_Current_Preset(FilePath As String)
Dim SaveString As String
For SettingsCount = 0 To UBound(BenchSelArray, 1)
    If BenchSelArray(SettingsCount) = True Then
        SaveString = SaveString + One_Code
    Else
        SaveString = SaveString + Zero_Code
    End If
Next SettingsCount
If Write_String_Into_File(FilePath, SaveString) = True Then
    Call Show_Msg_Window(Language(228), Language(229), 1)
Else
    Call Show_Msg_Window(Language(230), Language(231), 1)
End If
End Sub
Public Sub Load_Current_Preset(FilePath As String)
Dim SaveString As String
SaveString = Load_File_Into_String(FilePath)
SaveString = Left(SaveString, Len(SaveString) - 2)
If Len(SaveString) <> (UBound(BenchSelArray) + 1) Then
    Call Show_Msg_Window(Language(232), Language(233), 1)
    GoTo Ed
End If

For SettingsCount = 1 To UBound(BenchSelArray, 1) + 1
    If Mid(SaveString, SettingsCount, 1) = Zero_Code Then
        BenchSelArray(SettingsCount - 1) = False
    Else
        If SettingsCount >= 17 And SettingsCount <= 19 Then
            If Can_Menu(47) = True Then BenchSelArray(SettingsCount - 1) = True
        Else
            BenchSelArray(SettingsCount - 1) = True
        End If
    End If
Next SettingsCount
Call Show_Msg_Window(Language(234), Language(235), 1)
Ed: End Sub

Public Sub Check_Current_GUID()
If Len(ConnectedUsers(2, 0)) <> 10 Then ConnectedUsers(2, 0) = Generate_GUID
Manager.FrameLabel(54).Caption = Language(253) & Space_Code & App_Ver & Space_Code & ConnectedUsers(1, 0)
If ConnectedUsers(1, 0) = App_Full Then
    Manager.CommandButton(58).Visible = False
    Manager.FrameLabel(50).Caption = App_UserName
    Manager.FrameLabel(51).Caption = "GUID: " & ConnectedUsers(2, 0)
    Manager.FrameLabel(52).Caption = "REG: " & ConnectedUsers(3, 0)
    SupposedCaption = App_Title
Else
    Manager.CommandButton(58).Visible = True
    Manager.FrameLabel(50).Caption = Empty_Code
    Manager.FrameLabel(51).Caption = Empty_Code
    Manager.FrameLabel(52).Caption = Empty_Code
    If Date_Left = 0 Then
        SupposedCaption = App_Title & Space_Code & Language(258)
    Else
        SupposedCaption = App_Title & Space_Code & "(" & Manager.ControllerBox(21).Caption & ")"
    End If
End If
End Sub
