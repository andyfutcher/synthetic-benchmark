Attribute VB_Name = "Networking"
Option Explicit
Dim Net_Command As String, Net_Data As String, Net_Target As String, Net_Packet As String, QueryIPStarter As String, QueryIPCounter As Integer
Dim Net_BusyPacket As String, Net_GetPacket As String, Net_TempData As String, Net_VoteCommand As String

'Net Detection IP
Const Maximum_Ips = 5
Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type
Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(Maximum_Ips) As IPINFO  'array of IP address entries
End Type
Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Global Const NetCode_Name = "name"
Global Const NetCode_Rename = "renm"
Global Const NetCode_OS = "opsy"
Global Const NetCode_CPUType = "cput"
Global Const NetCode_CPUSpeed = "cpus"
Global Const NetCode_Memory = "memi"
Global Const NetCode_Comment = "come"
Global Const NetCode_Version = "vers"
Global Const NetCode_GUID = "guid"
Global Const NetCode_CDKey = "cdky"
Global Const NetCode_AMIIn = "amin"
Global Const NetCode_CheckImIn = "ckin"
Global Const NetCode_YoureIn = "urin"
Global Const NetCode_GivePassword = "gpss"
Global Const NetCode_SayAll = "syal"
Global Const NetCode_SayBlank = "sybl"
Global Const NetCode_SayTo = "syto"
Global Const NetCode_SayTo2 = "syt2"
Global Const NetCode_ChatCatch = "catc"
Global Const NetCode_UpConnList = "upul"
Global Const NetCode_UpPlatList = "ptul"
Global Const NetCode_WhoAreYou = "whou"
Global Const NetCode_MeName = "menm"
Global Const NetCode_MessageBox = "msgb"
Global Const NetCode_Ping = "ping"
Global Const NetCode_Ring = "ring"
Global Const NetCode_Busy = "busy"
Global Const NetCode_ImDone = "imdn"
Global Const NetCode_NoNew = "nonw"
Global Const NetCode_CanNew = "canw"
Global Const NetCode_Bench = "bmu"
Global Const NetCode_UBad = "ubad"
Global Const NetCode_UBanned = "uban"
Global Const NetCode_UKicked = "ukck"
Global Const NetCode_URLClear = "urlc"
Global Const NetCode_URLAdd = "urla"
Global Const NetCode_URLCheck = "urlk"
Global Const NetCode_CantVote = "ctvt"

Public Sub Net_Send_Data(DataString As String, Exception As Integer, OnlyTarget As Integer)
'If OnlyTarget = -2 Then GoTo Ed
If Net_ServerType = True And Net_ClientType = False Then
    If OnlyTarget <> "-1" Then
        If ManagerSub.Winsock(OnlyTarget).State = sckConnected And OnlyTarget <> Exception Then
            If Net_WaitState(OnlyTarget) = True And Left(DataString, 4) <> NetCode_Ring Then
                Call NetSub_AddToBuffer(DataString, Exception, OnlyTarget)
            Else
                Call Make_NetPacket_Compat(DataString)
                Call ManagerSub.Winsock(OnlyTarget).SendData(DataString & Net_PakTerminator)
            End If
        End If
    Else
        For ServerCount = 1 To ManagerSub.Winsock.Count - 1
            If ManagerSub.Winsock(ServerCount).State = sckConnected And ServerCount <> Exception Then
                If Net_WaitState(ServerCount) = True Then
                    Call NetSub_AddToBuffer(DataString, Exception, ServerCount)
                Else
                    Call Make_NetPacket_Compat(DataString)
                    Call ManagerSub.Winsock(ServerCount).SendData(DataString & Net_PakTerminator)
                End If
            End If
        Next ServerCount
    End If
End If
If Net_ServerType = False And Net_ClientType = True Then
    If Net_WaitState(0) = True Or ManagerSub.Winsock(0).State <> sckConnected Then
        Call NetSub_AddToBuffer(DataString, Exception, 0)
        GoTo Ed
    End If
    Call Make_NetPacket_Compat(DataString)
    Call ManagerSub.Winsock(0).SendData(DataString & Net_PakTerminator)
End If
Ed: End Sub

Public Sub Net_EmptyBack_Buffer(Index As Integer)
If Net_WaitState(Index) = False And Net_Buffer(0, 0) <> Empty_Code Then
    For NetCount = 0 To UBound(Net_Buffer(), 2) '- 1
        If Val(Net_Buffer(2, 0)) = Index And NetCount <= UBound(Net_Buffer(), 2) Then
            Call Net_Send_Data(Net_Buffer(0, NetCount), Val(Net_Buffer(1, NetCount)), Val(Net_Buffer(2, NetCount)))
            If UBound(Net_Buffer, 2) > 0 Then
                If NetCount <> UBound(Net_Buffer, 2) Then  '- 1
                    For ServerCount = NetCount To UBound(Net_Buffer, 2) - 1
                        Net_Buffer(0, ServerCount) = Net_Buffer(0, ServerCount + 1)
                        Net_Buffer(1, ServerCount) = Net_Buffer(1, ServerCount + 1)
                        Net_Buffer(2, ServerCount) = Net_Buffer(2, ServerCount + 1)
                    Next ServerCount
                End If
                ReDim Preserve Net_Buffer(UBound(Net_Buffer, 1), UBound(Net_Buffer, 2) - 1)
                NetCount = NetCount - 1
            Else
                Net_Buffer(0, 0) = Empty_Code
                Net_Buffer(1, 0) = Empty_Code
                Net_Buffer(2, 0) = Empty_Code
            End If
        End If
    Next NetCount
End If
End Sub

Public Sub NetSub_AddToBuffer(DataString As String, Exception As Integer, OnlyTarget As Integer)
Call Add_Index_To_StringArray(Net_Buffer())
Net_Buffer(0, UBound(Net_Buffer, 2)) = DataString
Net_Buffer(1, UBound(Net_Buffer, 2)) = Exception
Net_Buffer(2, UBound(Net_Buffer, 2)) = OnlyTarget
End Sub

Public Sub NetSub_Check_Client_Status()
Dim NewBusyStatus As Boolean
NewBusyStatus = False
For NetCount = 0 To UBound(Net_WaitState())
    If Net_WaitState(NetCount) = True Then NewBusyStatus = True
Next NetCount
Net_StillBusy = NewBusyStatus
End Sub

Public Sub NetSub_Next_Listener()
For ServerCount = 1 To ManagerSub.Winsock.Count - 1
    If ManagerSub.Winsock(ServerCount).State = sckListening Then GoTo PortAdded
Next ServerCount
For ServerCount = 1 To ManagerSub.Winsock.Count - 1
    If ManagerSub.Winsock(ServerCount).State = sckClosed Then
        ManagerSub.Winsock(ServerCount).Protocol = sckTCPProtocol
        ManagerSub.Winsock(ServerCount).LocalPort = Net_PortNumber
        ManagerSub.Winsock(ServerCount).Listen
        StatusMenus(1) = Language(148)
        GoTo PortAdded
    End If
Next ServerCount
StatusMenus(0) = Language(149)
PortAdded: Call Update_StatusBars
Call Update_Connection_List
End Sub

Public Sub Net_ConnectRequest(Index As Integer, RequestID As Long)
If ManagerSub.Winsock(Index).State <> sckClosed Then ManagerSub.Winsock(Index).Close
ManagerSub.Winsock(Index).Accept RequestID
End Sub

Public Sub Net_DataArrival(Index As Integer, TotalBytes As Long)
Call ManagerSub.Winsock(Index).GetData(Net_GetPacket, vbString, TotalBytes)
Net_BusyPacket = Net_BusyPacket & Net_GetPacket

Do While InStr(1, Net_BusyPacket, Net_PakTerminator) <> 0
    Net_Packet = Left(Net_BusyPacket, InStr(1, Net_BusyPacket, Net_PakTerminator) - 1)
    Call UnMake_NetPacket_Compat(Net_Packet)
    'Manager.Caption = Len(Net_Packet)
    Net_BusyPacket = Right(Net_BusyPacket, Len(Net_BusyPacket) - InStr(1, Net_BusyPacket, Net_PakTerminator))
    Net_Command = Left(Net_Packet, 4)
    Net_Target = Mid(Net_Packet, 5, 10)
    Net_Data = Right(Net_Packet, Len(Net_Packet) - 14)
    Call Do_Net_Command(Index)
Loop
End Sub

Public Sub Do_Net_Command(Index As Integer)
Select Case Net_Command
Case NetCode_Name
    ConnectedUsers(0, Index) = Net_Data
    If Net_WhosDuplicate(ConnectedUsers(0, Index)) = True Then
        Call Net_Send_Data(NetCode_Rename & ConnectedUsers(2, 0), -1, Index)
        GoTo Ed
    Else
        PlatFormUsers(0, Index) = Net_Data
    End If
    If Net_ServerType = True Then Call Net_UpDataUsers
Case NetCode_Rename
    ConnectedUsers(0, 0) = ConnectedUsers(0, 0) & Trim(Str(Int(Rnd * 9) + 1))
    Call Net_Send_Data(NetCode_Name & ConnectedUsers(2, 0) & ConnectedUsers(0, 0), -1, 0)
Case NetCode_OS
    PlatFormUsers(1, Index) = Net_Data
Case NetCode_CPUType
    PlatFormUsers(2, Index) = Net_Data
Case NetCode_CPUSpeed
    PlatFormUsers(3, Index) = Net_Data
Case NetCode_Memory
    PlatFormUsers(4, Index) = Net_Data
Case NetCode_Comment
    PlatFormUsers(5, Index) = Net_Data
Case NetCode_Version
    ConnectedUsers(1, Index) = Net_Data
    'If Val(Net_Data) > 99 Then
Case NetCode_GUID
    ConnectedUsers(2, Index) = Net_Data
Case NetCode_CDKey
    ConnectedUsers(3, Index) = Net_Data
Case NetCode_AMIIn, NetCode_CheckImIn
    Dim DuplicateKey As Integer
    If LCase(Trim(Net_Data)) <> Net_Password And Net_Command = NetCode_AMIIn Then  'beta Or ConnectedUsers(1, Index) < 100
        Call Net_Send_Data(NetCode_GivePassword & ConnectedUsers(2, 0), -1, Index)
        Call NetSub_Next_Listener
    Else
        If (ConnectedUsers(1, Index) = App_Full Or ConnectedUsers(1, Index) = App_Timed Or ConnectedUsers(1, Index) = App_Trial) And ConnectedUsers(1, Index) = App_Full Then
            If Generate_Input(ConnectedUsers(3, Index)) = False Then
                Call Net_Send_Data(NetCode_UBad & ConnectedUsers(2, 0), -1, Index)
                Call Net_Sure_Kicked(Index)
                GoTo Ed
            End If
            DuplicateKey = NetSub_Check_KeyDuplicates(Index)
            If DuplicateKey <> -1 And Net_InterNet = True Then
                Call Net_Send_Data(NetCode_UBad & ConnectedUsers(2, 0), -1, Index)
                Call Net_Sure_Kicked(Index)
                If DuplicateKey <> 0 Then
                    Call Net_Send_Data(NetCode_UBad & ConnectedUsers(2, 0), -1, DuplicateKey)
                    Call Net_Sure_Kicked(DuplicateKey)
                End If
                GoTo Ed
            End If
        End If
        
        If InStr(1, BannedList, ConnectedUsers(2, Index)) <> 0 Then
            Call Net_Send_Data(NetCode_UBanned & ConnectedUsers(2, 0), -1, Index)
            Call Net_Sure_Kicked(Index)
            GoTo Ed
        End If
        Call Net_Send_Data(NetCode_YoureIn & ConnectedUsers(2, 0), -1, Index)
        Call NetSub_Next_Listener
        Call Net_UpDataUsers
        If Net_Command = NetCode_AMIIn Then
            Call Net_ChatCatchUp(Index)
            Call Net_URLBench_Equalize(Index)
        End If
    End If
Case NetCode_YoureIn
    StatusMenus(1) = Language(150)
    Call Update_StatusBars
    Call Update_Connection_List
    Call Switch_Sections_To(0)
Case NetCode_GivePassword
    Net_TempData = Get_Password
    If Net_TempData <> Bad_Code Then
        Call Net_Send_Data(NetCode_AMIIn & ConnectedUsers(2, 0) & Net_TempData, -1, -1)
    Else
        Call All_Close_Session
    End If
Case NetCode_UBad, NetCode_UBanned
    Call All_Close_Session
    If Net_Command = NetCode_UBad Then
        Call Show_Msg_Window(Language(259), Language(260), 0)
    Else
        Call Show_Msg_Window(Language(273), Language(260), 0)
    End If
    StatusMenus(1) = Language(260)
    Call Update_StatusBars
Case NetCode_UKicked
    Call Disconnect_Me_Client
    Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & Language(268) & Space_Code & DoubleAsterix_Code)
Case NetCode_CantVote
    Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & Language(261) & Space_Code & DoubleAsterix_Code)
Case NetCode_SayAll
    If Net_Is_This_A_Command(Net_Data, Index) = False Then
        Call Chat_AddSay(Net_WhoIsThat(Net_Target), Net_Data)
        If Net_ServerType = True Then
            Call Net_Send_Data(NetCode_SayAll & Net_Target & Net_Data, Index, -1)
        End If
    End If
Case NetCode_SayTo
    If Net_Is_This_A_Command(Net_Data, Index) = False Then
        If ConnectedUsers(2, 0) = Net_Target Then
            Call Chat_AddSay(ConnectedUsers(0, Index) & Dagger_Char, Net_Data) 'Net_WhoIsThat(Net_Target)
        Else
            Call Net_Send_Data(NetCode_SayTo2 & ConnectedUsers(2, Index) & Net_Data, -1, Net_WhosNameIsThat(Net_Target))
        End If
    End If
Case NetCode_SayTo2
    Call Chat_AddSay(Net_WhoIsThat(Net_Target) & Dagger_Char, Net_Data)
Case NetCode_SayBlank
    Call Chat_AddSay(Space_Code, Net_Data)
Case NetCode_ChatCatch
    Call Convert_String_To_Array(Net_Data, ArbitoryList())
    Call Chat_AddArray(ArbitoryList())
Case NetCode_UpConnList
    Call Convert_String_To_Array(Net_Data, ConnectedList())
    Call Manager.FrameList(1).Submit_Data_Array(ConnectedList(), -1, 1)
Case NetCode_UpPlatList
    Call Convert_String_To_Array(Net_Data, PlatformList())
    Call Manager.PlatInfoCplxList(0).Submit_Data_Array(PlatformList(), -1, 5)
Case NetCode_WhoAreYou
    If Net_ServerType = True Then
        Call Net_Send_Data(NetCode_MeName & ConnectedUsers(2, 0) & Net_ServerName, -1, Index)
        DoEvents
    End If
    Call Net_Disconnect(Index)
    If Net_ServerType = True Then Call NetSub_Next_Listener
Case NetCode_MeName
    Call Add_Index_To_StringArray(NetSearchList())
    NetSearchList(0, UBound(NetSearchList, 2)) = Net_Data
    NetSearchList(1, UBound(NetSearchList, 2)) = Language(151)
    NetSearchList(2, UBound(NetSearchList, 2)) = QueryIPStarter & QueryIPCounter
    NetSearchList(3, UBound(NetSearchList, 2)) = Int(Val(Manager.WriteBox(9).Text))
    StatusMenus(1) = Language(152) & Space_Code & UBound(NetSearchList, 2) + 1 & Space_Code & Language(153)
    Call Update_StatusBars
Case NetCode_URLClear
    WebBusy = True
    ReDim URLBenchList(1, 0)
    Call Add_URLList(Empty_Code)
Case NetCode_URLAdd
    Call Add_URLList(Net_Data)
Case NetCode_URLCheck
    Call Begin_URL_Check_List
Case NetCode_MessageBox
    MsgBox Net_Data, vbInformation, Language(154) & Space_Code & ConnectedList(0, Index)
Case NetCode_Ping
    Call Net_Send_Data(NetCode_Ring & ConnectedUsers(2, 0), -1, Index)
Case NetCode_Ring
    Net_PingNow = False
Case NetCode_Busy
    Net_WaitState(Index) = True
    Call NetSub_Check_Client_Status
Case NetCode_ImDone
    Net_WaitState(Index) = False
    Call Net_EmptyBack_Buffer(Index)
    Call NetSub_Check_Client_Status
Case NetCode_NoNew
    Net_CantBench = True
Case NetCode_CanNew
    Net_CantBench = False
End Select
If Left(Net_Command, 3) = NetCode_Bench Then
    QueryIPCounter = Asc(Mid(Net_Command, 4, 1)) - NetCountAscii
    If Net_ServerType = True Then
        Call Net_Send_Data(NetCode_Bench & Chr(QueryIPCounter + NetCountAscii) & Net_Target & Net_Data, Index, -1)
    End If
    Call Convert_String_To_List(Net_Data, BenResultList())
    Call Flow_Triple_Array(BenchResults(), QueryIPCounter)
    
    If Net_ServerType = True Then
        If InStr(1, BenResultList(2), Language(204)) <> 0 And ConnectedUsers(1, Index) <> App_Full Then
            Call Net_Send_Data(NetCode_UBad & ConnectedUsers(2, 0), -1, Index)
            Call Net_Sure_Kicked(Index)
        End If
        If QueryIPCounter > 6 And QueryIPCounter < 19 Then
            If ConnectedUsers(1, Index) <> App_Full Then
                Call Net_Send_Data(NetCode_UBad & ConnectedUsers(2, 0), -1, Index)
                Call Net_Sure_Kicked(Index)
            End If
        End If
    End If
    
    For NetCount = 0 To UBound(BenchResults, 2)
        BenchResults(QueryIPCounter, NetCount, 0) = BenResultList(NetCount)
    Next NetCount
    Call Manager.ComplexList(QueryIPCounter).Submit_Data_Array(BenchResults(), QueryIPCounter, 5)
    Call Graph_Update(QueryIPCounter)
    Call Check_All_Menus
    Call Align_Selected_Complex_Controls
End If
Ed: End Sub

Public Sub Net_ChatCatchUp(Index As Integer)
If Manager.OptionBox(7).Value = False Then GoTo Ed
Dim TempChatString As String
Call Convert_Limted_Array_To_String(ChatLineInfo(), TempChatString)
Call Net_Send_Data(NetCode_ChatCatch & ConnectedUsers(2, 0) & TempChatString, -1, Index)
Ed: End Sub

Public Sub Net_URLBench_Equalize(Index As Integer)
Call Net_Send_Data(NetCode_URLClear & ConnectedUsers(2, 0), -1, Index)
For NetCount = 0 To UBound(URLBenchList(), 2)
    Call Net_Send_Data(NetCode_URLAdd & ConnectedUsers(2, 0) & URLBenchList(0, NetCount), -1, Index)
Next NetCount
Call Net_Send_Data(NetCode_URLCheck & ConnectedUsers(2, 0), -1, Index)
End Sub

Public Sub Net_ConnClose(Index As Integer)
'If Net_SearchType = False Then
    If Net_ServerType = True Then
        If ConnectedUsers(0, Index) <> Empty_Code Then Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & ConnectedUsers(0, Index) & Space_Code & Language(158) & Space_Code & DoubleAsterix_Code)
        Call NetSub_Next_Listener
    End If
    If Net_ClientType = True Then Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & Language(157) & Space_Code & DoubleAsterix_Code)
    Call Net_Disconnect(Index)
    Call Net_Is_Kicked(Index)
'End If
End Sub

Public Sub Net_Connect()
ManagerSub.Winsock(0).Tag = One_Code
ManagerSub.Winsock(0).RemoteHost = Net_ServAddress
ManagerSub.Winsock(0).RemotePort = Net_PortNumber
On Error GoTo CloseConn
ManagerSub.Winsock(0).Connect
If Net_FindType = True Then
    Do While ManagerSub.Winsock(0).State <> sckConnecting
        DoEvents
    Loop
End If
GoTo Ed
CloseConn: Call NetSub_Find_Ender
Call NetSub_KillConnection(0)
Call Show_Msg_Window(Language(155), Language(156) & " 10055", 0)
Ed: End Sub

Public Sub NetSub_Connect()
If Net_SearchType = True Then
    Call Net_Send_Data(NetCode_WhoAreYou & ConnectedUsers(2, 0), -1, -1)
Else
    StatusMenus(1) = Language(239)
    Call Update_StatusBars
    Call Net_Send_Data(NetCode_GUID & ConnectedUsers(2, 0) & ConnectedUsers(2, 0), -1, 0)
    Call Net_Send_Data(NetCode_Name & ConnectedUsers(2, 0) & ConnectedUsers(0, 0), -1, 0)
    Call Net_Send_Data(NetCode_Version & ConnectedUsers(2, 0) & ConnectedUsers(1, 0), -1, 0)
    Call Net_Send_Data(NetCode_CDKey & ConnectedUsers(2, 0) & ConnectedUsers(3, 0), -1, 0)
    Call Net_Send_Data(NetCode_OS & ConnectedUsers(2, 0) & PlatFormUsers(1, 0), -1, 0)
    Call Net_Send_Data(NetCode_CPUType & ConnectedUsers(2, 0) & PlatFormUsers(2, 0), -1, 0)
    Call Net_Send_Data(NetCode_CPUSpeed & ConnectedUsers(2, 0) & PlatFormUsers(3, 0), -1, 0)
    Call Net_Send_Data(NetCode_Memory & ConnectedUsers(2, 0) & PlatFormUsers(4, 0), -1, 0)
    Call Net_Send_Data(NetCode_Comment & ConnectedUsers(2, 0) & PlatFormUsers(5, 0), -1, 0)
    Call Net_Send_Data(NetCode_AMIIn & ConnectedUsers(2, 0), -1, 0)
End If
End Sub

Public Sub Net_UpDataUsers()
Dim UNameNetSend As String

ReDim ConnectedList(UBound(ConnectedList, 1), 0)
For NetCount = 0 To UBound(ConnectedUsers, 2)
    If ConnectedUsers(0, NetCount) <> Empty_Code Then
        Call Add_Index_To_StringArray(ConnectedList())
        ConnectedList(0, UBound(ConnectedList, 2)) = ConnectedUsers(0, NetCount)
        ConnectedList(1, UBound(ConnectedList, 2)) = ConnectedUsers(1, NetCount)
        ConnectedList(2, UBound(ConnectedList, 2)) = ConnectedUsers(2, NetCount)
    End If
Next NetCount
Call Convert_Array_To_String(ConnectedList, UNameNetSend)
Call Net_Send_Data(NetCode_UpConnList & ConnectedUsers(2, 0) & UNameNetSend, -1, -1)

ReDim PlatformList(UBound(PlatformList, 1), 0)
For NetCount = 0 To UBound(PlatFormUsers, 2)
    If PlatFormUsers(0, NetCount) <> Empty_Code Then
        Call Add_Index_To_StringArray(PlatformList())
        PlatformList(0, UBound(PlatformList, 2)) = PlatFormUsers(0, NetCount)
        PlatformList(1, UBound(PlatformList, 2)) = PlatFormUsers(1, NetCount)
        PlatformList(2, UBound(PlatformList, 2)) = PlatFormUsers(2, NetCount)
        PlatformList(3, UBound(PlatformList, 2)) = PlatFormUsers(3, NetCount)
        PlatformList(4, UBound(PlatformList, 2)) = PlatFormUsers(4, NetCount)
        PlatformList(5, UBound(PlatformList, 2)) = PlatFormUsers(5, NetCount)
    End If
Next NetCount
Call Convert_Array_To_String(PlatformList, UNameNetSend)
Call Net_Send_Data(NetCode_UpPlatList & ConnectedUsers(2, 0) & UNameNetSend, -1, -1)

Call Manager.FrameList(1).Submit_Data_Array(ConnectedList(), -1, 1)
Call Manager.PlatInfoCplxList(0).Submit_Data_Array(PlatformList(), -1, 5)
End Sub

Public Sub Net_Disconnect(Index As Integer)
If Index <> -1 Then
    Call NetSub_KillConnection(Index)
Else
    For ServerCount = 0 To ManagerSub.Winsock.Count - 1
        Call NetSub_KillConnection(ServerCount)
    Next ServerCount
End If
Call Update_Connection_List
Call Net_UpDataUsers
End Sub

Public Sub NetSub_KillConnection(Index As Integer)
ManagerSub.Winsock(Index).Close
If Index <> 0 Then
    Net_WaitState(Index) = False
    ConnectedUsers(0, Index) = Empty_Code
    ConnectedUsers(1, Index) = Empty_Code
    ConnectedUsers(2, Index) = Empty_Code
    ConnectedUsers(3, Index) = Empty_Code
    PlatFormUsers(0, Index) = Empty_Code
    PlatFormUsers(1, Index) = Empty_Code
    PlatFormUsers(2, Index) = Empty_Code
    PlatFormUsers(3, Index) = Empty_Code
    PlatFormUsers(4, Index) = Empty_Code
    PlatFormUsers(5, Index) = Empty_Code
End If
End Sub

Public Sub Net_Find_Local_Servers()
Call Get_Network_Defaults
QueryIPCounter = 0
QueryIPStarter = Def_IPAddress
ReDim NetSearchList(3, 0)
Call Choose_Manager_Functionality(False, -1, 1)
Do While Right(QueryIPStarter, 1) <> FullStop_Code
    QueryIPStarter = Left(QueryIPStarter, Len(QueryIPStarter) - 1)
Loop
Call Net_Find_Try_Next_IP
End Sub

Public Sub Net_Find_Try_Next_IP()
If Net_DoTwice = True Then
    Net_DoTwice = False
Else
    If ManagerSub.Winsock(0).State = sckConnected Then
        Net_DoTwice = True
        GoTo Ed
    End If
End If
Manager.Timer(2).Enabled = False
Manager.Timer(2).Interval = 0
Call NetSub_KillConnection(0)
QueryIPCounter = QueryIPCounter + 1
If QueryIPCounter = 255 Then
    Call NetSub_Find_Ender
    GoTo Ed
End If
Net_ServAddress = QueryIPStarter & QueryIPCounter
Net_PortNumber = Int(Val(Manager.WriteBox(9).Text))
Do While ManagerSub.Winsock(0).State <> sckClosed
    DoEvents
Loop
Call Net_Connect
If Net_SearchType = True Then
    Manager.Timer(2).Interval = Manager.WriteBox(16).ClickTag
    Manager.Timer(2).Enabled = True
End If
Ed: End Sub

Public Sub NetSub_Find_Ender()
Manager.Timer(2).Enabled = False
Manager.Timer(2).Interval = 0
Net_ClientType = False
Net_SearchType = False
If NetSearchList(0, 0) = Empty_Code Then
    ArbitoryList(0, 0) = Language(159)
    Call Manager.FrameList(0).Submit_Data_Array(ArbitoryList(), -1, 1)
Else
    Call Manager.FrameList(0).Submit_Data_Array(NetSearchList(), -1, 1)
End If
Call Check_NetCanConnect_Button
Call Control_Enable_Group(0, True)
Call Choose_Manager_Functionality(True, -1, 1)
StatusMenus(1) = Language(160)
Call Update_StatusBars
End Sub

Public Function NetSub_Check_KeyDuplicates(ClientIndex As Integer) As Integer
For NetCount = 0 To UBound(ConnectedUsers(), 2)
    If ConnectedUsers(1, NetCount) = App_Full And NetCount <> ClientIndex Then
    If ConnectedUsers(3, ClientIndex) = ConnectedUsers(3, NetCount) Then
        NetSub_Check_KeyDuplicates = NetCount
        GoTo Ed
    End If
    End If
Next NetCount
NetSub_Check_KeyDuplicates = -1
Ed: End Function

Public Sub Get_Network_Defaults()
Def_FriendlyName = ManagerSub.Winsock(0).LocalHostName
'Def_IPAddress = ManagerSub.Winsock(0).LocalIP
'Def_WWWAddress =

Dim IPRet As Long, IPTableBytes() As Byte, IPListing As MIB_IPADDRTABLE, IPCounter As Integer
Dim IPSubNet As String, IPTemp As String

On Error GoTo CantDetect
GetIpAddrTable ByVal 0&, IPRet, True
If IPRet <= 0 Then
    Def_IPAddress = ManagerSub.Winsock(0).LocalIP
    Def_WWWAddress = Empty_Code
    GoTo Ed
End If

ReDim IPTableBytes(0 To IPRet - 1) As Byte
GetIpAddrTable IPTableBytes(0), IPRet, False
  
CopyMemory IPListing.dEntrys, IPTableBytes(0), 4
For IPCounter = 0 To IPListing.dEntrys - 1
    CopyMemory IPListing.mIPInfo(IPCounter), IPTableBytes(4 + (IPCounter * Len(IPListing.mIPInfo(0)))), Len(IPListing.mIPInfo(IPCounter))
    IPSubNet = ConvertAddressToString(IPListing.mIPInfo(IPCounter).dwMask)
    IPTemp = ConvertAddressToString(IPListing.mIPInfo(IPCounter).dwAddr)
    If Val(Filter_Sort(IPSubNet)) > 255000 Then
        If IPSubNet = "255.255.255.255" Then
            Def_WWWAddress = IPTemp
        Else
            Def_IPAddress = IPTemp
        End If
    End If
Next
GoTo Ed

CantDetect: Def_IPAddress = ManagerSub.Winsock(0).LocalIP
Def_WWWAddress = Empty_Code
Resume Ed
Ed: End Sub

Public Sub Net_Sure_Kicked(KickIndex As Integer)
Net_KickList(KickIndex) = ConnectedUsers(2, KickIndex)
Manager.Timer(3).Interval = 5000
End Sub
Public Sub Net_Check_Kicked()
Manager.Timer(3).Interval = 0
For CheckCount = 0 To UBound(ConnectedUsers(), 2)
    If Net_KickList(CheckCount) <> Empty_Code Then
        If Net_KickList(CheckCount) = ConnectedUsers(2, CheckCount) Then
            Call Net_Disconnect(CheckCount)
        End If
        Net_KickList(CheckCount) = Empty_Code
    End If
Next CheckCount
End Sub

Public Sub Net_Is_Kicked(KickIndex As Integer)
Net_KickList(KickIndex) = Empty_Code
Manager.Timer(3).Interval = 0
End Sub

Private Function ConvertAddressToString(longAddr As Long) As String
Dim myByte(3) As Byte, Cnt As Long
CopyMemory myByte(0), longAddr, 4
For Cnt = 0 To 3
    ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + FullStop_Code
Next Cnt
ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Function Net_Is_This_A_Command(CommandString As String, ClientIndex As Integer) As Boolean
Net_Is_This_A_Command = False
If Left(Trim(CommandString), 1) <> "/" Then
    GoTo Ed
Else
    If Net_ClientType = True Or Net_Single = True Then
        Net_Is_This_A_Command = True
        GoTo Ed
    End If
End If

Dim RawCmdString As String, RawCommand As String, RawTarget As Integer, RawData As String
If InStr(1, CommandString, Space_Code) <> 0 Then
    RawCmdString = CommandString
    RawCommand = LCase(Mid(RawCmdString, 2, InStr(1, RawCmdString, Space_Code) - 2))
    RawData = Mid(RawCmdString, InStr(1, RawCmdString, Space_Code) + 1, Len(RawCmdString) - InStr(1, RawCmdString, Space_Code))
    RawTarget = Net_WhosNameIsThat(Trim(Right(RawCmdString, Len(RawCmdString) - Len(RawCommand) - 1)))
Else
    RawCommand = LCase(Right(CommandString, Len(CommandString) - 1))
    RawTarget = -2
End If

Net_Is_This_A_Command = True
Select Case RawCommand
Case LCase(Language(282))
    If ConnectedUsers(1, ClientIndex) = App_Full Or ConnectedUsers(1, ClientIndex) = App_Timed Then
        If ClientIndex = 0 Then
            Manager.WriteBox(15).Text = RawData
            Call Process_CommandButton_Click(35)
        Else
            Call Net_Start_Vote(CommandString, RawCommand, RawTarget, ClientIndex, RawData)
        End If
    Else
        Call NetSub_Send_Confirmation(ClientIndex, 1)
    End If
Case LCase(Language(283))
    If ConnectedUsers(1, ClientIndex) = App_Full Or ConnectedUsers(1, ClientIndex) = App_Timed Then
        For NetCount = 0 To UBound(URLBenchList, 2)
            If InStr(1, LCase(URLBenchList(0, NetCount)), LCase(RawData)) <> 0 Then
                If ClientIndex = 0 Then
                    Manager.FrameList(2).ListIndex = NetCount
                    Call Process_CommandButton_Click(27)
                Else
                    Call Net_Start_Vote(CommandString, RawCommand, RawTarget, ClientIndex, RawData)
                End If
            End If
        Next NetCount
    Else
        Call NetSub_Send_Confirmation(ClientIndex, 1)
    End If
Case LCase(Language(264))
    If Net_VoteCommand <> Empty_Code Then
        Net_VoteList(ClientIndex) = 1
        Call NetSub_Send_Confirmation(ClientIndex, 0)
    End If
Case LCase(Language(265))
    If Net_VoteCommand <> Empty_Code Then
        Net_VoteList(ClientIndex) = 2
        Call NetSub_Send_Confirmation(ClientIndex, 0)
    End If
Case Language(270), Language(271)
    If RawTarget = -1 Then GoTo Ed
    If RawTarget = 0 Then
        Call Net_Send_Data(NetCode_SayBlank & ConnectedUsers(2, 0) & DoubleAsterix_Code & Space_Code & Language(269) & Space_Code & DoubleAsterix_Code, -1, ClientIndex)
        GoTo Ed
    End If
    If ClientIndex = 0 Then
        If RawCommand = Language(271) Then Call Net_Add_Banned_List(ConnectedUsers(2, RawTarget))
        Call Net_Send_Data(NetCode_UKicked & ConnectedUsers(2, 0), -1, RawTarget)
        Call Net_Sure_Kicked(RawTarget)
    Else
        Call Net_Start_Vote(CommandString, RawCommand, RawTarget, ClientIndex, RawData)
    End If
Case Else
    If Net_ServerType = False Then Net_Is_This_A_Command = False
End Select
Ed: End Function

Public Sub Net_Start_Vote(CommandString As String, CommandName As String, TargetIndex As Integer, ClientIndex As Integer, RawDataStr As String)
If Net_VoteCommand <> Empty_Code Then
    Call Net_Send_Data(NetCode_CantVote & ConnectedUsers(2, 0), -1, ClientIndex)
Else
    For NetCount = 0 To UBound(Net_VoteList)
        Net_VoteList(NetCount) = 0
    Next NetCount
    Net_VoteList(ClientIndex) = 1
    Net_VoteCommand = CommandString
    If RawDataStr = Empty_Code Then
        DuelOutput = DoubleAsterix_Code & Space_Code & ConnectedUsers(0, ClientIndex) & Space_Code & Language(262) & Space_Code & Normalize(CommandName) & Space_Code & ConnectedList(0, TargetIndex) & Space_Code & DoubleAsterix_Code
    Else
        DuelOutput = DoubleAsterix_Code & Space_Code & ConnectedUsers(0, ClientIndex) & Space_Code & Language(262) & Space_Code & Normalize(CommandName) & Space_Code & RawDataStr & Space_Code & DoubleAsterix_Code
    End If
    Call Chat_AddSay(Space_Code, DuelOutput)
    Call Net_Send_Data(NetCode_SayBlank & ConnectedUsers(2, 0) & DuelOutput, -1, -1)
    Manager.Timer(4).Interval = 15000
End If
End Sub

Public Sub Net_End_Vote()
If Net_VoteCommand = Empty_Code Then
    Manager.Timer(4).Interval = 0
    GoTo Ed
End If
Dim YesScore As Integer, NoScore As Integer, ForfietScore As Integer
Manager.Timer(4).Interval = 0
For NetCount = 0 To UBound(Net_VoteList)
    If ConnectedUsers(2, NetCount) <> Empty_Code Then
        Select Case Net_VoteList(NetCount)
        Case 0
            ForfietScore = ForfietScore + 1
        Case 1
            YesScore = YesScore + 1
        Case 2
            NoScore = NoScore + 1
        End Select
    End If
Next NetCount

DuelOutput = DoubleAsterix_Code & Space_Code & Language(263) & Space_Code & YesScore & Space_Code & Language(264) & Space_Code & NoScore & Space_Code & Language(265) & Space_Code & ForfietScore & Space_Code & Language(266) & Space_Code & DoubleAsterix_Code
Call Chat_AddSay(Space_Code, DuelOutput)
Call Net_Send_Data(NetCode_SayBlank & ConnectedUsers(2, 0) & DuelOutput, -1, -1)

If YesScore >= NoScore Then Call Net_Is_This_A_Command(Net_VoteCommand, 0)
Net_VoteCommand = Empty_Code
Ed: End Sub

Public Sub NetSub_Send_Confirmation(Index As Integer, IndexType As Integer)
Select Case IndexType
Case 0
    DuelOutput = DoubleAsterix_Code & Space_Code & Language(267) & Space_Code & DoubleAsterix_Code
Case 1
    DuelOutput = DoubleAsterix_Code & Space_Code & Language(284) & Space_Code & DoubleAsterix_Code
End Select

If Index = 0 Then
    Call Chat_AddSay(Space_Code, DuelOutput)
Else
    Call Net_Send_Data(NetCode_SayBlank & ConnectedUsers(2, 0) & DuelOutput, -1, Index)
End If
End Sub

Public Sub Net_Add_Banned_List(GUIDString As String)
BannedList = BannedList & GUIDString
End Sub

Public Function Net_WhoIsThat(GUIDString As String) As String
Net_WhoIsThat = Language(0)
For NetCount = 0 To UBound(ConnectedList, 2)
    If ConnectedList(2, NetCount) = GUIDString Then
        Net_WhoIsThat = ConnectedList(0, NetCount)
    End If
Next NetCount
End Function

Public Function Net_WhosDuplicate(NameString As String) As Boolean
Dim NameCount As Integer
Net_WhosDuplicate = False
For NetCount = 0 To UBound(ConnectedUsers, 2)
    If LCase(ConnectedUsers(0, NetCount)) = LCase(NameString) Then
        NameCount = NameCount + 1
    End If
Next NetCount
If NameCount > 1 Then Net_WhosDuplicate = True
End Function

'Public Function Net_WhosGUIDIsThat(GuidString As String) As Integer
'Net_WhosGUIDIsThat = -2
'For NetCount = 0 To UBound(ConnectedList, 2)
'    If ConnectedList(2, NetCount) = GuidString Then
'        Net_WhosGUIDIsThat = NetCount
'    End If
'Next NetCount
'End Function
Public Function Net_WhosNameIsThat(GUIDString As String) As Integer
Net_WhosNameIsThat = -1
For NetCount = 0 To UBound(ConnectedUsers(), 2)
    If LCase(ConnectedUsers(0, NetCount)) = LCase(GUIDString) Then
        Net_WhosNameIsThat = NetCount
    End If
Next NetCount
End Function


Public Sub Net_ErrorDisplay(Index As Integer, ErrorString As String, ErrorNumber As Integer)
If Net_SearchType = True Then GoTo Ed
Call Choose_Manager_Functionality(True, -1, 1)
Call Show_Msg_Window(ErrorString, Language(156) & Space_Code & Str(ErrorNumber), 0)
StatusMenus(1) = Language(156) & Space_Code & ErrorNumber
Call Update_StatusBars
Call All_Close_Session 'Call Net_Disconnect(Index)
Ed: End Sub
