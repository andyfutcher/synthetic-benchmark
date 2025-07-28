Attribute VB_Name = "BenchMarks"
Option Explicit
Dim BenchTime As Single, SecondCounter As Integer, BenchCount As Integer, BenRealtime As Boolean, BenFileCounter As Long

Dim BenchScore() As Single, BenCount As Integer, BenCounter As Integer, BenMax As Single, BenMin As Single, BenLongCount As Long, BenPath As String
Dim BenchString As String, BenchInt() As Integer, BenchLong As Long, BenchSingle() As Single, BenchByte() As Byte, BenchList() As String
Dim BenchWidth As Integer, BenchHeight As Integer, PrevWinState As Integer, BenchCurrent As Integer, BenchError As Boolean

Dim TgtMegaBytes As Long, TgtMBCount As Integer
Const IntMax = 32767
Const OneMB = 1048576
Const OneKB = 1024
Const HalfLong = 1073741822
Const MaxByte = 255
Const MediaBackGnd = &H0&

Private Sub BenchMark_Delay(TimeCount As Integer)
Call BenchMark_Ready
SecondCounter = TimeCount
Do Until SecondCounter = 0
    Call Change_BenchDisplay(-1, Str(SecondCounter))
    BenchTime = Timer + 1
    Do While BenchTime >= Timer And StopBenchmarks = False
        DoEvents
    Loop
    SecondCounter = SecondCounter - 1
Loop
End Sub
Private Sub BenchMark_Ready()
If Int(Timer) >= 86380 Then
    Call Change_BenchDisplay(-2, Language(176))
    Do While Int(Timer) <> 0 And StopBenchmarks = False
        DoEvents
    Loop
End If
End Sub

Private Sub Do_Timer_Work(EnaTimers As Boolean)
If EnaTimers = True Then
    For BenCount = 0 To Manager.Timer.Count - 1
        Manager.Timer(BenCount).Interval = TimerState(BenCount)
    Next BenCount
Else
    For BenCount = 0 To Manager.Timer.Count - 1
        TimerState(BenCount) = Manager.Timer(BenCount).Interval
        Manager.Timer(BenCount).Interval = 0
    Next BenCount
End If
End Sub

Private Function Prepare_Benchmark_Go() As Boolean
Prepare_Benchmark_Go = True
Select Case BenchCount
Case 0
    ReDim BenchByte(0)
Case 1
    BenchString = Empty_Code
Case 2
    ReDim BenchSingle(0) 'BenchSingle = 0
Case 3
    ReDim BenchInt(0)
Case 4, 6
    TaskStyle = 0
    Call ATM_Get_System_Processes(ATMProcessInfo())
    Call ATM_Repaint_All_Tasks
Case 7
    Call Setup_HDD_Benchmarks
    BenPath = Add_The_Slash(Left(Manager.WriteBox(11).Text, 2))
    ReDim BenchByte(TgtMegaBytes)
Case 8, 9
    If BenCounter <> 0 Then GoTo Ed
    Call Update_Benchmark_List(0)
    Call Change_BenchDisplay(-2, Language(177))
    BenPath = Add_The_Slash(Left(Manager.WriteBox(11).Text, 2))
    Call Setup_HDD_Benchmarks
    ReDim BenchByte(TgtMegaBytes)
    For BenLongCount = 0 To 4
        Call Write_Array_Into_File(BenPath & Bench_Name & BenLongCount & Bench_Ext, BenchByte())
    Next BenLongCount
    ReDim BenchByte(0)
    Call Update_Benchmark_List(1)
Case 10
    If BenCounter <> 0 Then GoTo Ed
    ReDim BenchList(0)
    Call Update_Benchmark_List(0)
    Call Change_BenchDisplay(-2, Language(178))
    BenPath = Add_The_Slash(Left(Manager.WriteBox(11).Text, 2))
    Call Build_File_Database(BenchList(), BenPath)
    Call Update_Benchmark_List(1)
Case 11, 12
    If BenCounter <> 0 Then GoTo Ed
    ReDim BenchList(0)
    Call Update_Benchmark_List(0)
    Call Change_BenchDisplay(-2, Language(179))
    BenPath = Add_The_Slash(Left(Manager.WriteBox(12).Text, 2))
    If BenPath = Empty_Code Then
        Prepare_Benchmark_Go = False
        Call Show_Msg_Window(Language(180), Language(181), 0)
    Else
        Call Build_File_Database(BenchList(), BenPath)
        If BenchList(0, 0) = Empty_Code Then
            Prepare_Benchmark_Go = False
            Call Show_Msg_Window(Language(180), Language(181), 0)
        End If
    End If
    Call Update_Benchmark_List(1)
Case 13
    'ManagerSub.MMControl.hWndDisplay = MultiMedia.hwnd
    ManagerSub.MMControl.DeviceType = "Waveaudio"
    ManagerSub.MMControl.TimeFormat = 0
    ManagerSub.MMControl.FileName = Application_Path & Sample_Path & Sample_Name & Zero_Code & Resource_Ext
    ManagerSub.MMControl.Wait = False
    ManagerSub.MMControl.Command = "open"
    If ManagerSub.MMControl.CanPlay = False Then
        Prepare_Benchmark_Go = False
        Call Show_Msg_Window(Language(182), Language(181), 0)
    End If
Case 14
    If BenCounter <> 0 Then GoTo Ed
    Call Show_MM_Window(True)
Case 15
    If BenCounter <> 0 Then GoTo Ed
    MultiMedia.MediaWindow.Visible = True
    MultiMedia.MediaWindow.Move (Screen.Width / 2) - ((OnePix * 640) / 2), (Screen.Height / 2) - ((OnePix * 480) / 2), (OnePix * 640), (OnePix * 480)
    MultiMedia.BackColor = MediaBackGnd
    ManagerSub.MMControl.DeviceType = "AVIVideo"
    ManagerSub.MMControl.TimeFormat = 3
    ManagerSub.MMControl.hWndDisplay = MultiMedia.MediaWindow.hWnd '
    ManagerSub.MMControl.FileName = Application_Path & Sample_Path & Sample_Name & One_Code & Resource_Ext
    ManagerSub.MMControl.Wait = False
    ManagerSub.MMControl.Command = "open"
    
    If ManagerSub.MMControl.CanPlay = False Or ManagerSub.MMControl.CanStep = False Or Does_File_Exist(ManagerSub.MMControl.FileName) = False Then '
        Prepare_Benchmark_Go = False
        Call Show_Msg_Window(Language(183), Language(181), 0)
    Else
        Call Show_MM_Window(True)
    End If
Case 16
    If BenCounter <> 0 Then GoTo Ed
    Call Change_BenchDisplay(-2, Language(184))
    Do While WebBusy = True And StopBenchmarks = False
        DoEvents
    Loop
    If Net_WebAccess(0) <> 2 Then
        Call Check_Online_Status(0)
        If Net_WebAccess(0) <> 2 Then
            Prepare_Benchmark_Go = False
            Call Show_Msg_Window(Language(185), Language(181), 0)
            GoTo Ed
        End If
    End If
    If URLBenchList(0, 0) = Empty_Code Then
        Prepare_Benchmark_Go = False
        Call Show_Msg_Window(Language(186), Language(181), 0)
    End If
Case 17, 18, 19
    If BenCounter <> 0 Then GoTo Ed
    If Net_ClientType = False Then
        Prepare_Benchmark_Go = False
        'Call Show_Msg_Window("A server or offline platform cannot perform this benchmark type.", "Skipped Benchmark", 0)
    Else
        BenchString = Empty_Code
        Call Change_BenchDisplay(-2, Language(187))
        If BenchCount = 18 Then BenchString = String(OneKB, Chr(Int(MaxByte * Rnd)))
        If BenchCount = 19 Then
            ReDim BenchByte(OneKB)
            For BenLongCount = 0 To OneKB
                BenchByte(BenLongCount) = Int(57 * Rnd) + 65
            Next BenLongCount
            Call CompressData(BenchByte, 9)
            Call Convert_ByteArray_To_String(BenchByte(), BenchString)
        End If
    End If
Case 20, 21, 22, 23
    If BenCounter <> 0 Then GoTo Ed
    Select Case BenchCount
    Case 20
        Call Change_BenchDisplay(-2, Language(188))
        MultiMedia.FrameLabel.Caption = Language(192)
    Case 21
        Call Change_BenchDisplay(-2, Language(189))
        MultiMedia.FrameLabel.Caption = Language(193)
    Case 22
        Call Change_BenchDisplay(-2, Language(190))
        MultiMedia.FrameLabel.Caption = Language(194)
    Case 23
        Call Change_BenchDisplay(-2, Language(191))
        MultiMedia.FrameLabel.Caption = Language(195)
    End Select
    MultiMedia.FrameLabel.Move (Screen.Width / 2) - (MultiMedia.FrameLabel.Width / 2), (Screen.Height / 2) - (MultiMedia.FrameLabel.Height / 2)
    MultiMedia.FrameLabel.Visible = True
    BenchTime = Timer + 4
    Do While BenchTime >= Timer And StopBenchmarks = False
        DoEvents
    Loop
    Call Show_MM_Window(True)
End Select
Ed: DoEvents
End Function

Public Sub Start_BenchMarks(IsRealTime As Boolean)
LastBenType = IsRealTime
If BenchDoneOnce = True Then
    Call Show_Msg_Window(Language(197), Language(196), 1)
    GoTo Ed
End If
If Net_CantBench = True Then
    Call Show_Msg_Window(Language(198), Language(196), 1)
    GoTo Ed
End If
BenchDoneOnce = True
StopBenchmarks = False
Call Select_Common_Benchmarks(-1)
If DoBurnin = False Then
    Call Choose_Manager_Functionality(False, 9, 0)
Else
    Call Choose_Manager_Functionality(False, 9, 1)
End If
Call Setup_HDD_Benchmarks
BenRealtime = IsRealTime
If BenFileCounter > HalfLong Then BenFileCounter = 0
ReDim BenchTaskList(2, 0)
Call Do_Timer_Work(False)
For BenchCount = 0 To 23
    If BenchSelArray(BenchCount) = True Then
        Call Add_Index_To_StringArray(BenchTaskList())
        BenchTaskList(0, UBound(BenchTaskList, 2)) = BenchDatArray(BenchCount, 0)
        BenchTaskList(1, UBound(BenchTaskList, 2)) = Language(199)
        BenchTaskList(2, UBound(BenchTaskList, 2)) = BenchDatArray(BenchCount, 2)
    End If
Next BenchCount
Call Manager.FrameList(3).Submit_Data_Array(BenchTaskList(), -1, 2) 'If BenchTaskList(0, 0) <> Empty_Code Then
Call Change_Benchmark_State(1)
If DoBurnin = False Then Call BenchMark_Delay(3)

Manager.ProgressBox(0).Max = 5 * (UBound(BenchTaskList(), 2) + 1)
BenchCount = 0
BenchCurrent = -1
BenchError = False
For BenchCount = 0 To 23
If BenchSelArray(BenchCount) = True Then
    BenchCurrent = BenchCurrent + 1
    Call Update_Benchmark_List(1)
    Call Reset_BenchMark_Memory
    For BenCounter = 0 To 4
        If Prepare_Benchmark_Go = False Then
            Call Update_Benchmark_List(3)
            BenchError = True
            GoTo DoNextBen
        End If
        If StopBenchmarks = True Then
            Call Update_Benchmark_List(3)
            GoTo Ed1
        End If
        Manager.ProgressBox(0).Value = Manager.ProgressBox(0).Value + 1
        DoEvents
        Call BenchMark_Ready
        Call Change_BenchDisplay(BenchCount, Str(BenCounter + 1))
        
        Select Case BenchCount
        Case 0
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                ReDim Preserve BenchByte(BenchScore(BenCounter))
                BenchByte(BenchScore(BenCounter)) = BenchScore(BenCounter) Mod 255
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents
            Loop
        Case 1
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                BenchString = BenchString & Chr(Int(255 * Rnd))
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents
            Loop
        Case 2
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                ReDim Preserve BenchSingle(BenchScore(BenCounter))
                BenchSingle(BenchScore(BenCounter)) = Exp(Atn(Sqr(Sqr(Cos(Sin(Tan(Sqr((BenchScore(BenCounter) * 2) / 2) ^ 2) ^ 2) ^ 2)))))
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents
            Loop
        Case 3
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                ReDim Preserve BenchInt(BenchScore(BenCounter))
                BenchInt(BenchScore(BenCounter)) = BenchInt(BenchScore(BenCounter)) + (BenchScore(BenCounter) Mod IntMax)
                BenchInt(BenchScore(BenCounter)) = BenchInt(BenchScore(BenCounter)) - (BenchScore(BenCounter) Mod IntMax)
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents
            Loop
        Case 4
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                Call ATM_Get_System_Processes(ATMProcessInfo())
                BenchScore(BenCounter) = BenchScore(BenCounter) + UBound(ATMProcessInfo(), 2)
                If BenRealtime = False Then DoEvents
            Loop
        Case 5
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                If BenCounter > 2 Then BenchString = BenchString & Space_Code
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1 'UBound(BenString(), 2)
                If BenRealtime = False Then DoEvents
            Loop
        Case 6
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                Call ATM_Get_System_Processes(ATMProcessInfo())
                Call ATM_Repaint_All_Tasks
                BenchScore(BenCounter) = BenchScore(BenCounter) + UBound(ATMProcessInfo(), 2)
                If BenRealtime = False Then DoEvents
            Loop
        Case 7
            BenchTime = Timer
            Call Write_Array_Into_File(BenPath & Bench_Name & BenCounter & Bench_Ext, BenchByte())
            BenchScore(BenCounter) = (1 / ((Timer - BenchTime) / TgtMBCount)) ', "0.00")
        Case 8, 9
            BenchTime = Timer
            If Load_File_Into_Array(BenPath & Bench_Name & BenCounter & Bench_Ext, BenchByte()) = False Then
                'Beep
            End If
            BenchScore(BenCounter) = (1 / ((Timer - BenchTime) / TgtMBCount)) ', "0.00")
        Case 10, 12
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                'BenchLong = FileLen(BenchList(0, BenchScore(BenCounter) Mod UBound(BenchList, 2)) & BenchList(1, BenchScore(BenCounter) Mod UBound(BenchList, 2)))
                BenFileCounter = BenFileCounter + 1
                BenchString = BenchList(0, BenFileCounter Mod UBound(BenchList, 2)) & BenchList(1, BenFileCounter Mod UBound(BenchList, 2))
                If Does_File_Exist(BenchString) = True Then
                    Open BenchString For Binary Access Read As #1
                    BenchString = Input(1, #1)
                    Close #1
                End If
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
        Case 11
            BenchLong = 0
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                BenFileCounter = BenFileCounter + 1
                BenchString = BenchList(0, BenFileCounter Mod UBound(BenchList, 2)) & BenchList(1, BenFileCounter Mod UBound(BenchList, 2))
                If Load_File_Into_Array(BenchString, BenchByte()) = True Then
                    BenchLong = BenchLong + UBound(BenchByte)
                End If
            Loop
            BenchTime = BenchTime - 1
            BenchScore(BenCounter) = (((BenchLong / 1024) / 1024) / (Timer - BenchTime)) '(1 / ((Timer - BenchTime) / TgtMBCount))
        Case 13
            ManagerSub.MMControl.Command = "prev"
            ManagerSub.MMControl.Command = "play"
            DoEvents
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                ManagerSub.MMControl.From = BenchScore(BenCounter) Mod ManagerSub.MMControl.Length
                ManagerSub.MMControl.Command = "play"
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents '
            Loop
        Case 14
            BenchHeight = MultiMedia.ScaleHeight
            BenchWidth = MultiMedia.ScaleWidth
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                MultiMedia.Line (Int(BenchWidth * Rnd), Int(BenchHeight * Rnd))-(Int(BenchWidth * Rnd), Int(BenchHeight * Rnd)), RGB(Int(MaxByte * Rnd), Int(MaxByte * Rnd), Int(MaxByte * Rnd)), B
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents '
            Loop
        Case 15
            ManagerSub.MMControl.Command = "prev"
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                ManagerSub.MMControl.Command = "step"
                If (BenchScore(BenCounter) Mod ManagerSub.MMControl.Length) = 0 Then ManagerSub.MMControl.Command = "prev"
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
                If BenRealtime = False Then DoEvents '
            Loop
        Case 16
            BenchLong = 0
            ManagerSub.GenBrowser.Offline = False
            BenchTime = Timer
            For NetCount = 0 To UBound(URLBenchList(), 2)
                If URLBenchList(1, NetCount) = Language(44) Then
                Call ManagerSub.GenBrowser.Navigate2(URLBenchList(0, NetCount), 2 Or 4 Or 8)
                Do While ManagerSub.GenBrowser.Busy = True
                    DoEvents
                Loop
                Call Return_WebBrowser_Contents(0, HtmlData)
                If Check_HTML_Integrity(HtmlData) = True Then
                    BenchLong = BenchLong + 1
                End If
                End If
            Next NetCount
            If BenchLong = 0 Then GoTo DoNextBen
            BenchScore(BenCounter) = (60 / ((Timer - BenchTime) / BenchLong))
        Case 17
            BenchLong = 0
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                Net_PingNow = True
                Call Net_Send_Data(NetCode_Ping & ConnectedUsers(2, 0), -1, 0)
                Do While Net_PingNow = True
                    DoEvents
                Loop
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
            'BenchScore(BenCounter) = (1 / BenchLong)
        Case 18, 19
            BenchLong = 0
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                Net_PingNow = True
                Call Net_Send_Data(NetCode_Ping & ConnectedUsers(2, 0) & BenchString, -1, 0)
                Do While Net_PingNow = True
                    DoEvents
                Loop
                BenchLong = BenchLong + 1
            Loop
            BenchTime = BenchTime - 1
            BenchScore(BenCounter) = (BenchLong / (Timer - BenchTime))
        Case 20
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                MouseMoved = False
                Do While MouseMoved = False
                    If BenchTime < Timer Then Exit Do
                    DoEvents
                Loop
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
        Case 21
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                MouseClicked = False
                Do While MouseClicked = False
                    If BenchTime < Timer Then Exit Do
                    DoEvents
                Loop
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
        Case 22
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                KeyHeld = False
                Do While KeyHeld = False
                    If BenchTime < Timer Then Exit Do
                    DoEvents
                Loop
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
        Case 23
            BenchTime = Timer + 1
            Do While BenchTime >= Timer
                KeyTapp = False
                Do While KeyTapp = False
                    If BenchTime < Timer Then Exit Do
                    DoEvents
                Loop
                BenchScore(BenCounter) = BenchScore(BenCounter) + 1
            Loop
        End Select
        'Call Choose_Manager_Functionality(True, -1, 0)
    Next BenCounter
    If BenchCount = 13 Or BenchCount = 15 Then
        ManagerSub.MMControl.Command = "stop"
        ManagerSub.MMControl.Command = "close"
    End If
    Call Update_Benchmark_List(2)
    
    If DoBurnin = False Then
    Call Do_Benchmark_Scoring
    If Net_ServerType = True Or Net_ClientType = True Then
        Call Convert_List_To_String(BenResultList(), BenchString)
        Call Net_Send_Data(NetCode_Bench & Chr(BenchCount + NetCountAscii) & ConnectedUsers(2, 0) & BenchString, -1, -1)
    End If
    
    Call Flow_Triple_Array(BenchResults(), BenchCount)
    For BenCounter = 0 To UBound(BenchResults, 2)
        BenchResults(BenchCount, BenCounter, 0) = BenResultList(BenCounter)
    Next BenCounter
    Call Manager.ComplexList(BenchCount).Submit_Data_Array(BenchResults(), BenchCount, 5)
    Call Graph_Update(BenchCount)
    Else
        If BenchCount = 14 Or BenchCount = 15 Or BenchCount >= 20 Then Call Show_MM_Window(False)
    End If
End If
DoNextBen: 'Call Choose_Manager_Functionality(True, -1, 0)
Next BenchCount
Ed1: Call Do_Timer_Work(True)
If BenchError = True Then
    Call Change_BenchDisplay(-3, Language(200))
Else
    Call Change_BenchDisplay(-3, Language(201))
    If Manager.OptionBox(23).Value = True And DoBurnin = False Then Call Hide_Controller_Box(9)
End If
'MsgBox BenchScore(0) & " " & BenchScore(1) & " " & BenchScore(2) & " " & BenchScore(3) & " " & BenchScore(4)
Ed: Manager.ProgressBox(0).Value = 0

Call Change_Benchmark_State(2)
Call Align_Selected_Complex_Controls
Call Choose_Manager_Functionality(True, -1, 0)
Call Check_All_Menus
BenchDoneOnce = False
End Sub

Private Sub Do_Benchmark_Scoring()
GlobalMemoryStatus MemoryInfo
BenResultList(0) = ConnectedUsers(0, 0)
BenResultList(3) = MemoryInfo.dwMemoryLoad & "%" & Space_Code & Language(202) & Space_Code & Round((MemoryInfo.dwTotalPhys / 1024) / 1024) & Language(20)
TaskStyle = WS_VISIBLE
Call ATM_Get_System_Processes(ATMProcessInfo())
BenResultList(4) = UBound(ATMProcessInfo(), 2)
TaskStyle = 0
Call ATM_Get_System_Processes(ATMProcessInfo())
BenResultList(4) = BenResultList(4) & " (" & UBound(ATMProcessInfo(), 2) & Space_Code & Language(203) & ")"
BenMax = 0
BenMin = BenchScore(0)
If BenchCount = 5 Then
    For BenCount = 0 To 2
        If BenchScore(BenCount) > BenMax Then BenMax = BenchScore(BenCount)
        If BenchScore(BenCount) < BenMin Then BenMin = BenchScore(BenCount)
    Next BenCount
Else
    For BenCount = 0 To 4
        If BenchScore(BenCount) > BenMax Then BenMax = BenchScore(BenCount)
        If BenchScore(BenCount) < BenMin Then BenMin = BenchScore(BenCount)
    Next BenCount
End If
BenResultList(2) = Format(((100 / BenMax) * BenMin), "0.00") & "%"
BenResultList(5) = Empty_Code
If BenchCount >= 7 And BenchCount <= 12 Then
    BenResultList(5) = UCase(Left(BenPath, 2))
    If Drv_Label(BenPath) <> Empty_Code Then BenResultList(5) = BenResultList(5) & " (" & Drv_Label(BenPath) & ")"
End If

Select Case BenchCount
Case 0, 1, 2, 3, 4, 6, 10, 12, 13, 14, 15, 17, 20, 21, 22, 23, 24
    BenResultList(1) = 0
    For BenCount = 1 To 4
        BenResultList(1) = BenResultList(1) + BenchScore(BenCount)
    Next BenCount
    BenResultList(1) = Int(BenResultList(1) / 4) & Space_Code & BenchDatArray(BenchCount, 3)
    If BenchCount <> 10 And BenchCount <> 12 Then Call Add_Realtime_Indicator
    If BenchCount = 14 Or BenchCount = 15 Or BenchCount >= 20 Then Call Show_MM_Window(False)
Case 5
    BenMax = BenchScore(1) + BenchScore(2)
    BenMin = BenchScore(3) + BenchScore(4)
    BenResultList(1) = Format(BenMax / BenMin, "0.00") & Space_Code & BenchDatArray(BenchCount, 3)
    Call Add_Realtime_Indicator
Case 7, 8, 9, 11, 16, 18, 19
    BenResultList(1) = 0
    For BenCount = 1 To 4
        BenResultList(1) = BenResultList(1) + BenchScore(BenCount)
    Next BenCount
    BenResultList(1) = Format((BenResultList(1) / 4), "0.00") & Space_Code & BenchDatArray(BenchCount, 3)
    If BenchCount <= 9 Then Call Kill_File(BenPath & "*" & Bench_Ext)
End Select
End Sub

Private Sub Reset_BenchMark_Memory()
BenchString = Empty_Code
BenchLong = 0
BenchHeight = 0
BenchWidth = 0
ReDim BenchInt(0)
ReDim BenchSingle(0)
ReDim BenchByte(0)
ReDim BenchList(0)
ReDim BenchScore(4)
BenCounter = 0
End Sub

Private Sub Add_Realtime_Indicator()
If BenRealtime = True Then
    BenResultList(2) = BenResultList(2) & " (" & Language(204) & ")"
Else
    BenResultList(2) = BenResultList(2) & " (" & Language(205) & ")"
End If
End Sub

Private Sub Setup_HDD_Benchmarks()
GlobalMemoryStatus MemoryInfo

TgtMegaBytes = Round((MemoryInfo.dwTotalPhys / 1024) / 1024) + 1
If BenchCount = 9 Then
    TgtMBCount = (TgtMegaBytes / 20)
    TgtMegaBytes = OneMB * TgtMBCount
Else
    TgtMBCount = (TgtMegaBytes / 10)
    TgtMegaBytes = OneMB * TgtMBCount
End If
End Sub

Public Sub Change_Benchmark_State(StateIndex As Integer)
Select Case StateIndex
Case 0
    StatusMenus(2) = Language(206)
    Manager.CommandButton(24).Enabled = False
    Manager.CommandButton(25).Caption = Language(9)
Case 1
    StatusMenus(2) = Language(207)
    Manager.CommandButton(24).Enabled = False
    Manager.CommandButton(25).Caption = Language(10)
    If Net_ServerType = True Then
        Call Net_Send_Data(NetCode_NoNew & ConnectedUsers(2, 0), -1, -1)
        Call Change_BenchDisplay(-4, Empty_Code)
        Do While Net_StillBusy = True And StopBenchmarks = False
            DoEvents
        Loop
    End If
    If Net_ClientType = True Or Net_ServerType = True Then Call Net_Send_Data(NetCode_Busy & ConnectedUsers(2, 0), -1, -1)
Case 2
    StatusMenus(2) = Language(208)
    Manager.CommandButton(25).Caption = Language(9)
    If Net_ClientType = True Or Net_ServerType = True Then Call Net_Send_Data(NetCode_ImDone & ConnectedUsers(2, 0), -1, -1)
    If Net_ServerType = True Then Call Net_Send_Data(NetCode_CanNew & ConnectedUsers(2, 0), -1, -1)
    Manager.CommandButton(24).Enabled = True
End Select
Call Update_StatusBars
End Sub

Public Sub Change_BenchDisplay(DisType As Integer, AltText As String)
Select Case DisType
Case -4
    Manager.FrameLabel(16).Caption = Language(209)
    Manager.FrameLabel(17).Caption = Language(210)
Case -3
    If StopBenchmarks = False Then
        Manager.FrameLabel(16).Caption = Language(211)
    Else
        Manager.FrameLabel(16).Caption = Language(227)
    End If
    Manager.FrameLabel(17).Caption = AltText
Case -2
    Manager.FrameLabel(16).Caption = Language(212)
    Manager.FrameLabel(17).Caption = AltText
Case -1
    Manager.FrameLabel(16).Caption = Language(213)
    Manager.FrameLabel(17).Caption = Language(214) & AltText & Space_Code & Language(215)
Case Is >= 0
    Manager.FrameLabel(16).Caption = Language(216) & Space_Code & BenchDatArray(DisType, 0) & Space_Code & Language(217)
    If AltText > 1 Then
        Manager.FrameLabel(17).Caption = Language(218)
    Else
        Manager.FrameLabel(17).Caption = Language(219)
    End If
End Select
DoEvents
End Sub

Public Sub Update_Benchmark_List(TypeIndex As Integer)
Select Case TypeIndex
Case 0
    BenchTaskList(1, BenchCurrent) = Language(220)
Case 1
    BenchTaskList(1, BenchCurrent) = Language(221)
Case 2
    BenchTaskList(1, BenchCurrent) = Language(222)
Case 3
    BenchTaskList(1, BenchCurrent) = Language(223)
End Select
Call Manager.FrameList(3).Submit_Data_Array(BenchTaskList(), -1, 2)
End Sub

Private Sub Show_MM_Window(EnaShow As Boolean)
If EnaShow = True Then
    PrevWinState = ManagerSub.WindowState
    MultiMedia.Show
    Call Form_Control_Click(2)
Else
    ManagerSub.WindowState = PrevWinState
    ManagerSub.Refresh
    MultiMedia.Hide
    Unload MultiMedia
End If
End Sub

'Private Sub Generate_Random_Buffer()
'For BenLongCount = 0 To TgtMegaBytes
'    BenchByte(BenLongCount) = Int(254 * Rnd) + 1
'Next BenLongCount
'End Sub

