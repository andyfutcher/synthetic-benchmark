Attribute VB_Name = "Internet"
Option Explicit

Public Sub Check_Online_Status(TypeIndex As Integer)
Call Choose_Manager_Functionality(False, -1, 1)
Call Switch_UpdateWindow_Index(5)
Call Get_Network_Defaults
Call Manager.WriteBox(4).DDList_Clear
Call Manager.WriteBox(4).DDList_Add(Language(37), LAN_Code)
Call Manager.WriteBox(21).DDList_Clear
If TypeIndex = 1 And Manager.OptionBox(18).Value = False Then GoTo Ed1
WebBusy = True
StatusMenus(1) = Language(40)
Call Update_StatusBars
Dim Target_Array() As Byte, ReadString As String
ManagerSub.Internet.RequestTimeout = INet_TimeoutSmall
Call Cycle_Our_Site
Call Download_Internet_File(Web_OurSite & Web_IniFile, Target_Array())
Call Convert_ByteArray_To_String(Target_Array(), ReadString)
'Clipboard.SetText LCase(ReadString)
If InStr(1, LCase(ReadString), "html") = 0 And ReadString <> Empty_Code Then
    'Call Manager.WriteBox(21).DDList_Add("All", Empty_Code)
    Call Process_Info_File(ReadString)
    Net_WebAccess(1) = 2
    StatusMenus(1) = Language(38)
    Call Begin_URL_Check_List
    Manager.WeBrowser.Offline = False
    ManagerSub.GenBrowser.Offline = False
    'Manager.OptionBox(13).ValueDesc = True
Else
Ed1: Net_WebAccess(1) = 1
    StatusMenus(1) = Language(39)
    'Manager.OptionBox(4).Enabled = False
    Call Manager.WriteBox(21).DDList_Add(Language(41), LAN_Code)
    Call Manager.WriteBox(4).DDList_Add(Language(42), Net_Code)
End If
WebBusy = False
Ed: Call Update_StatusBars
Call Update_Connection_List
Call Choose_Manager_Functionality(True, -1, 1)
Call Hide_Controller_Box(15)
End Sub

Public Sub Download_Internet_File(WebURL As String, Target_Array() As Byte)
On Error GoTo Cant_Download
Target_Array() = ManagerSub.Internet.OpenURL(WebURL, icByteArray)
GoTo Ed
Cant_Download: ReDim Target_Array(0)
Resume Ed
Ed: End Sub

Public Sub Add_URLList(AddString As String)
If AddString = Empty_Code Then
    Call Manager.FrameList(2).Empty_Data
    GoTo Ed
End If
Call Add_Index_To_StringArray(URLBenchList())
URLBenchList(0, UBound(URLBenchList, 2)) = Web_Http & Replace(AddString, Web_Http, Empty_Code)
If InStr(1, AlreadyOnline, LCase(AddString)) = 0 Then
    URLBenchList(1, UBound(URLBenchList, 2)) = Language(45)
Else
    URLBenchList(1, UBound(URLBenchList, 2)) = Language(44)
End If
Call Manager.FrameList(2).Submit_Data_Array(URLBenchList, -1, 1)
Call Check_Target_Url_Buttons
Ed: End Sub

Public Sub Begin_URL_Check_List()
WebBusy = True
If UrlDoneOnce = True Then GoTo Ed1
UrlDoneOnce = True
Call Check_Target_Url_Buttons
If URLBenchList(0, 0) = Empty_Code Then GoTo Ed
ManagerSub.GenBrowser.Offline = False
If URLBenchList(0, 0) <> Empty_Code Then Net_WebAccess(0) = 1
Do
    For WebCount = 0 To UBound(URLBenchList(), 2)
        If URLBenchList(1, WebCount) = Language(45) And URLBenchList(0, WebCount) <> Empty_Code Then
            Call ManagerSub.GenBrowser.Navigate2(URLBenchList(0, WebCount), 2 Or 4 Or 8)
            Call WWW_Wait
            Call Return_WebBrowser_Contents(0, HtmlData)
            If Check_HTML_Integrity(HtmlData) = False Then
                URLBenchList(1, WebCount) = Language(43)
            Else
                URLBenchList(1, WebCount) = Language(44)
                AlreadyOnline = AlreadyOnline & LCase(URLBenchList(0, WebCount))
                Net_WebAccess(0) = 2
            End If
        End If
        If URLBenchList(0, 0) <> Empty_Code Then Call Manager.FrameList(2).Submit_Data_Array(URLBenchList, -1, 1)
    Next WebCount
Loop Until WebCount > UBound(URLBenchList(), 2)
Ed: UrlDoneOnce = False
WebBusy = False
Ed1: Call Check_Target_Url_Buttons
End Sub

Public Sub Return_WebBrowser_Contents(WindowIndex As Integer, ReturnString As String)
Dim HtmlDoc As Object, ObjOleCount As Integer
ReturnString = Empty_Code
Select Case WindowIndex
Case 0
    Set HtmlDoc = ManagerSub.GenBrowser.Document
    For ObjOleCount = 1 To HtmlDoc.All.Length - 1
        ReturnString = ReturnString & HtmlDoc.All.Item(ObjOleCount).innertext & Chr(13)
    Next ObjOleCount
End Select
End Sub

Public Sub Update_Connection_List()
ReDim Preserve ConnStatusList(1, 1)
ConnStatusList(0, 0) = Language(46)
ConnStatusList(0, 1) = Language(47)

For NetCount = 0 To 1
    Select Case Net_WebAccess(NetCount)
    Case 0
        ConnStatusList(1, NetCount) = Language(0)
    Case 1
        ConnStatusList(1, NetCount) = Language(43)
    Case 2
        ConnStatusList(1, NetCount) = Language(44)
    End Select
Next NetCount

For NetCount = 0 To ManagerSub.Winsock.Count - 1
    'If ConnStatusList(0, UBound(ConnStatusList, 2)) <> "" Then ReDim Preserve ConnStatusList(UBound(ConnStatusList, 1), UBound(ConnStatusList, 2) + 1)
    Select Case ManagerSub.Winsock(NetCount).State
    Case sckConnected
        Call Add_Index_To_StringArray(ConnStatusList())
        ConnStatusList(0, UBound(ConnStatusList, 2)) = Language(48) & Space_Code & NetCount
        ConnStatusList(1, UBound(ConnStatusList, 2)) = Language(49)
    Case sckListening
        Call Add_Index_To_StringArray(ConnStatusList())
        ConnStatusList(0, UBound(ConnStatusList, 2)) = Language(48) & Space_Code & NetCount
        ConnStatusList(1, UBound(ConnStatusList, 2)) = Language(50)
    End Select
Next NetCount
Call Manager.FrameList(5).Submit_Data_Array(ConnStatusList(), -1, 1)
End Sub

Public Sub Register_Host_WWW()
'http://AndyFutcher.com/db/add.php?name=&ip=&port=&protected=&country=
Call Switch_UpdateWindow_Index(4)
Dim SrvSubString As String, TempWWWIdString As String
If Manager.OptionBox(6).Value = False Then GoTo Ed
Call Cycle_Our_Site
SrvSubString = Web_OurSite & Net_WWWName & Manager.WriteBox(1).Text
SrvSubString = SrvSubString & Net_WWWIp & Def_WWWAddress
SrvSubString = SrvSubString & Net_WWWPort & Filter_Sort(Manager.WriteBox(2).Text)
If Trim(Manager.WriteBox(3).Text) = Empty_Code Then
    SrvSubString = SrvSubString & Net_WWWProtected & Zero_Code
Else
    SrvSubString = SrvSubString & Net_WWWProtected & One_Code
End If
SrvSubString = SrvSubString & Net_WWWRegion & LCase(Manager.WriteBox(21).ClickTag)
Call Show_Controller_Box(15)
'Call ManagerSub.GenBrowser.Navigate2("_blank")
Call ManagerSub.GenBrowser.Navigate2(SrvSubString, 2 Or 4 Or 8)
Call WWW_Wait
Call Return_WebBrowser_Contents(0, HtmlData)
If Check_HTML_Integrity(HtmlData) = True Then
    Net_WWWId = Val(HtmlData)
    TempWWWIdString = Net_WWWId
    TempWWWIdString = Left(TempWWWIdString, Int(Len(TempWWWIdString) / 2))
    Net_WWWId = Val(TempWWWIdString)
End If
'MsgBox Net_WWWId
If Net_WWWId = 0 Then Call Show_Msg_Window(Language(51), Language(156) & " 100002", 0)
Call Hide_Controller_Box(15)
Ed: End Sub

Public Sub Unregister_WWW_Host()
If Net_WWWId <> 0 Then
    Dim SrvSubString As String
    Call Cycle_Our_Site
    SrvSubString = Web_OurSite & Net_WWWRem & Net_WWWId
    Call Switch_UpdateWindow_Index(4)
    Call ManagerSub.GenBrowser.Navigate2(SrvSubString, 2 Or 4 Or 8)
    Call WWW_Wait
    Net_WWWId = 0
    Call Hide_Controller_Box(15)
End If
End Sub

Public Sub List_WWW_Hosts(ListType As Integer)
Dim SearchURL As String, PacketMax As Integer, WWWArray() As Byte
Call Choose_Manager_Functionality(False, -1, 1)
Call Switch_UpdateWindow_Index(4)

Call Cycle_Our_Site
If ListType = 0 Then
    ReDim NetSearchList(3, 0)
    If Manager.WriteBox(4).ClickTag <> Empty_Code Then
        SearchURL = Web_OurSite & Net_WWWList & Manager.WriteBox(4).ClickTag
    Else
        SearchURL = Web_OurSite & Net_WWWListAll
    End If
    PacketMax = 4
    Call ManagerSub.GenBrowser.Navigate2(SearchURL, 2 Or 4 Or 8)
    Call WWW_Wait
    Call Return_WebBrowser_Contents(0, HtmlData)
Else
    ReDim NetSearchCom(5, 0)
    SearchURL = Web_OurSite & Net_WWWCommunity
    PacketMax = 5
    Call Download_Internet_File(SearchURL, WWWArray())
    Call Convert_ByteArray_To_String(WWWArray(), HtmlData)
    HtmlData = Replace(HtmlData, Chr(10), Empty_Code)
End If

If Check_HTML_Integrity(HtmlData) = True Then
    Dim WholePacket As String, PacketCoOrd(1) As Integer, PacketCount As Integer, PacketString As String
    Do While Len(HtmlData) <> 0
        If ListType = 0 Then
            PacketCoOrd(0) = 2
        Else
            PacketCoOrd(0) = 1
        End If
        PacketCoOrd(1) = InStr(PacketCoOrd(0), HtmlData, Chr(13)) - PacketCoOrd(0)
        WholePacket = Mid(HtmlData, PacketCoOrd(0), PacketCoOrd(1))
        HtmlData = Right(HtmlData, Len(HtmlData) - InStr(1, HtmlData, Chr(13)))
        
        Call Add_Index_To_StringArray(NetSearchList())
        NetSearchList(1, UBound(NetSearchList, 2)) = Language(52)
        If InStr(1, WholePacket, Net_WWWScan) <> 0 Then
        For PacketCount = 0 To PacketMax
            PacketString = Left(WholePacket, InStr(1, WholePacket, Net_WWWScan) - 1)
            Select Case PacketCount
            Case 0
                If ListType = 0 Then
                    NetSearchList(0, UBound(NetSearchList, 2)) = PacketString
                Else
                    NetSearchCom(0, UBound(NetSearchCom, 2)) = PacketString
                End If
            Case 1
                If ListType = 0 Then
                    For NetCount = 0 To UBound(NetSearchList, 2) - 1
                        If NetSearchList(2, NetCount) = PacketString Then NetSearchList(0, UBound(NetSearchList, 2)) = Empty_Code
                    Next NetCount
                    NetSearchList(2, UBound(NetSearchList, 2)) = PacketString
                Else
                    NetSearchCom(3, UBound(NetSearchCom, 2)) = PacketString
                End If
            Case 2
                If ListType = 0 Then
                    NetSearchList(3, UBound(NetSearchList, 2)) = PacketString
                    NetSearchList(1, UBound(NetSearchList, 2)) = NetSearchList(1, UBound(NetSearchList, 2)) & " (" & PacketString & ")"
                Else
                    NetSearchCom(2, UBound(NetSearchCom, 2)) = PacketString
                End If
            Case 3
                If ListType = 0 Then
                    If PacketString = 1 And NetSearchList(0, UBound(NetSearchList, 2)) <> Empty_Code Then NetSearchList(0, UBound(NetSearchList, 2)) = NetSearchList(0, UBound(NetSearchList, 2)) & UnReg_Char
                Else
                    NetSearchCom(1, UBound(NetSearchCom, 2)) = PacketString
                End If
            Case 4
                If ListType <> 0 Then
                    NetSearchCom(4, UBound(NetSearchCom, 2)) = PacketString
                End If
            Case 5
                If ListType <> 0 Then
                    NetSearchCom(5, UBound(NetSearchCom, 2)) = PacketString
                End If
            End Select
            WholePacket = Right(WholePacket, Len(WholePacket) - InStr(1, WholePacket, Net_WWWScan))
        Next PacketCount
        End If
    Loop
End If

'NetSearchCom

If ListType = 0 Then
    If NetSearchList(0, 0) = Empty_Code Then
        ArbitoryList(0, 0) = Language(53)
        Call Manager.FrameList(0).Submit_Data_Array(ArbitoryList(), -1, 1)
        StatusMenus(1) = Language(54)
    Else
        Call Manager.FrameList(0).Submit_Data_Array(NetSearchList(), -1, 1)
        StatusMenus(1) = Language(55) & Space_Code & UBound(NetSearchList, 2) & Space_Code & Language(56)
    End If
Else
    If NetSearchCom(0, 0) = Empty_Code Then
        ArbitoryList(0, 0) = Language(53)
        Call Manager.FrameList(6).Submit_Data_Array(ArbitoryList(), -1, 1)
        StatusMenus(1) = Language(54)
    Else
        Call Manager.FrameList(6).Submit_Data_Array(NetSearchCom(), -1, 2)
        StatusMenus(1) = Language(55) & Space_Code & UBound(NetSearchCom, 2) & Space_Code & Language(56)
    End If
End If

Call Update_StatusBars
Call Control_Enable_Group(0, True)
Call Control_Enable_Group(1, True)
Call Check_NetCanConnect_Button
Call Choose_Manager_Functionality(True, -1, 1)
Call Hide_Controller_Box(15)
End Sub

Public Sub Add_To_RecentList(AddComName As String)
For ControlCount = 0 To UBound(RecentComList(), 2) - 1
    RecentComList(0, ControlCount) = RecentComList(0, ControlCount + 1)
    RecentComList(1, ControlCount) = RecentComList(1, ControlCount + 1)
    RecentComList(2, ControlCount) = RecentComList(2, ControlCount + 1)
    RecentComList(3, ControlCount) = RecentComList(3, ControlCount + 1)
Next ControlCount

RecentComList(0, 0) = AddComName
RecentComList(1, 0) = Net_ServAddress
RecentComList(2, 0) = Net_PortNumber
RecentComList(3, 0) = Net_InterNet
End Sub

Public Sub Download_WWW_Photo()
If NetSearchCom(4, Manager.FrameList(6).ListIndex) <> Empty_Code Then
    Call Choose_Manager_Functionality(False, -1, 1)
    Call Switch_UpdateWindow_Index(6)
    Dim Target_Array() As Byte, ReadString As String
    Call Download_Internet_File(NetSearchCom(4, Manager.FrameList(6).ListIndex), Target_Array())
    Call Convert_ByteArray_To_String(Target_Array(), ReadString)
    If Check_HTML_Integrity(ReadString) = True Then
        Call Kill_File(Application_Path & Update_Path & Photo_Name)
        Call Write_Array_Into_File(Application_Path & Update_Path & Photo_Name, Target_Array())
        Manager.FrameImage(16).Picture = LoadPicture(Application_Path & Update_Path & Photo_Name)
    End If
    Call Choose_Manager_Functionality(True, -1, 1)
    Call Hide_Controller_Box(15)
End If
End Sub

Public Sub WWW_Wait()
Dim WWWCounter As Integer, OldTimeSec As Integer
OldTimeSec = Second(Time$)
Do While WWWCounter < 5 And ManagerSub.GenBrowser.Busy = False
    If OldTimeSec <> Second(Time$) Then
        OldTimeSec = Second(Time$)
        WWWCounter = WWWCounter + 1
    End If
    DoEvents
Loop
Do While ManagerSub.GenBrowser.Busy = True
    DoEvents
Loop
End Sub

Public Sub Browse_To_Site(SiteIndex As Integer, OrURL As String)
Dim NewUrl As String
Select Case SiteIndex
Case -3
    NewUrl = OrURL
Case -2
    NewUrl = "http://www.AndyFutcher.com/purchase001.html"
    Call Manager.WeBrowser.Navigate2(NewUrl, 1)
    GoTo Ed
Case -1
    NewUrl = Web_OurSite & "index.html"
Case 0
    NewUrl = Application_Path & Lang_Path & Language_Current & "index.html"
Case 1
    NewUrl = Web_OurSite & "support.html"
Case 2
    NewUrl = Web_OurSite & "sxpfaq.html"
Case 3
    NewUrl = Web_OurSite & "download.html"
Case 4
    NewUrl = Web_OurSite & "synthboards.html"
End Select
Call Manager.WeBrowser.Navigate2(NewUrl)
Ed: End Sub

Public Sub Switch_UpdateWindow_Index(Index As Integer)
Select Case Index
Case 0
    Manager.FrameLabel(32).Caption = Language(57)
Case 1
    Manager.FrameLabel(32).Caption = Language(58)
Case 2
    Manager.FrameLabel(32).Caption = Language(59)
Case 3
    Manager.FrameLabel(32).Caption = Language(60)
Case 4
    Manager.FrameLabel(32).Caption = Language(61)
Case 5
    Manager.FrameLabel(32).Caption = Language(62)
Case 6
    Manager.FrameLabel(32).Caption = Language(275)
End Select
Call Show_Controller_Box(15)
End Sub

Public Sub Do_Software_Update_Now(UpdUrgent As Boolean)
Dim Target_Array() As Byte
If UpdUrgent = False Then If QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LastUpdateDay) = Day(Date$) Then GoTo Ed
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_LastUpdateDay, Val(Day(Date$)))
Call Switch_UpdateWindow_Index(0)
Call Choose_Manager_Functionality(False, 15, 1)

If Filter_Sort(App_NewVer(0)) > Filter_Sort(App_Ver) Then
    UpdUrgent = False
    If Show_Question_Window(Language(63) & Space_Code & App_Ver & Space_Code & Language(64) & Space_Code & App_NewVer(0) & Language(65), Language(66), 2) = False Then GoTo Cont1
    Call Switch_UpdateWindow_Index(1)
    Call Download_Internet_File(Web_OurSite & Net_UpdateURL, Target_Array())
    Call Kill_File(Application_Path & Update_Path & Update_Name)
    Call Write_Array_Into_File(Application_Path & Update_Path & Update_Name, Target_Array())
    Call Save_Settings_Data(0)
    Call Shell(Application_Path & Update_Path & Update_Name, vbNormalFocus)
    EndUpdatePath = Application_Path & Update_Path & Update_Name
    Call Show_Msg_Window(Language(285), Language(286), 1)
End If

Cont1: If FileLen(Application_Path & CPUFile_Name & Resource_Ext) < Val(App_NewVer(1)) Then
    UpdUrgent = False
    Call Switch_UpdateWindow_Index(2)
    Call Download_Internet_File(Web_OurSite & Net_CPUUpdate, Target_Array())
    If Check_Download_Integrity(Target_Array) = True Then
        Call Kill_File(Application_Path & CPUFile_Name & Resource_Ext)
        Call Write_Array_Into_File(Application_Path & CPUFile_Name & Resource_Ext, Target_Array())
    End If
End If

If FileLen(Application_Path & BWFFile_Name & Resource_Ext) < Val(App_NewVer(2)) Then
    UpdUrgent = False
    Call Switch_UpdateWindow_Index(3)
    Call Download_Internet_File(Web_OurSite & Net_BWFUpdate, Target_Array())
    If Check_Download_Integrity(Target_Array) = True Then
        Call Kill_File(Application_Path & BWFFile_Name & Resource_Ext)
        Call Write_Array_Into_File(Application_Path & BWFFile_Name & Resource_Ext, Target_Array())
    End If
End If
Ed: Call Hide_Controller_Box(15)
Call Choose_Manager_Functionality(True, -1, 1)
If UpdUrgent = True Then Call Show_Msg_Window(Language(67), Language(68), 1)
End Sub

Public Sub AutoDetect_Internet_Connection()
ManagerSub.Internet.AccessType = Manager.WriteBox(18).ClickTag
ManagerSub.Internet.RequestTimeout = 60
If Manager.WriteBox(18).ClickTag = 2 And Trim(Manager.WriteBox(19).Text) = Empty_Code Then
    Dim ProxyRead As String
    ProxyRead = QueryValue(HKEY_CURRENT_USER, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\INTERNET SETTINGS\", "ProxyServer")
    If InStr(1, ProxyRead, Colon_Code) <> 0 Then
        Manager.WriteBox(19).Text = Left(ProxyRead, InStr(1, ProxyRead, Colon_Code) - 1)
        Manager.WriteBox(20).Text = Right(ProxyRead, Len(ProxyRead) - InStr(1, ProxyRead, Colon_Code))
    End If
End If
If Manager.WriteBox(19).Text <> Empty_Code Then ManagerSub.Internet.Proxy = Manager.WriteBox(19).Text & Colon_Code & Val(Manager.WriteBox(20).Text)
End Sub

Public Sub Cycle_Our_Site()
OurSite_Counter = OurSite_Counter + 1
If OurSite_Counter = 4 Then OurSite_Counter = 0
Select Case OurSite_Counter
Case 0
    Web_OurSite = "http://www.AndyFutcher.com/"
Case 1
    Web_OurSite = "http://www.andyfutcher.com/"
Case 2
    Web_OurSite = "http://AndyFutcher.com/"
Case 3
    Web_OurSite = "http://andyfutcher.com/"
End Select
End Sub
