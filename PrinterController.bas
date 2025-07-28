Attribute VB_Name = "PrinterController"
Option Explicit
Const WM_PASTE = &H302
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Print_RTF_Box()
Call Rebuild_Print_Job(False)
On Error GoTo CantPrint
ManagerSub.CommonDialog.CancelError = True
ManagerSub.CommonDialog.Flags = cdlPDReturnDC + cdlPDNoPageNums + cdlPDDisablePrintToFile + cdlPDAllPages
ManagerSub.CommonDialog.ShowPrinter
Printer.Print Empty_Code
Call ManagerSub.RichTextBox.SelPrint(ManagerSub.CommonDialog.hDC)
GoTo Ed

CantPrint: Resume Ed
Ed: ManagerSub.RichTextBox.Text = Empty_Code
End Sub

Public Sub Save_RTF_Box(SaveFilePath As String)
On Error GoTo CantSave
Call Rebuild_Print_Job(True)
Call ManagerSub.RichTextBox.SaveFile(SaveFilePath, rtfRTF)
GoTo Ed

CantSave: Call Show_Msg_Window(Language(161), Language(156) & Space_Code & "102032", 0)
Resume Ed
Ed: ManagerSub.RichTextBox.Text = Empty_Code
End Sub

Public Sub Rebuild_Print_Job(DoallSections As Boolean)
Dim TextBuilder As String, PrintCount1 As Integer, PrintCount2 As Integer
If DoallSections = False Then If Manager.OptionBox(8).Value = False Then GoTo NextSec1
TextBuilder = LineFeed
For PrintCount1 = 0 To UBound(ChatLineInfo(), 2)
    If ChatLineInfo(0, PrintCount1) = Space_Code Then
        TextBuilder = TextBuilder & ChatLineInfo(1, PrintCount1) & Chr(13)
    Else
        TextBuilder = TextBuilder & ChatLineInfo(0, PrintCount1) & Colon_Code & Space_Code & ChatLineInfo(1, PrintCount1) & Chr(13)
    End If
Next PrintCount1

NextSec1: If DoallSections = False Then If Manager.OptionBox(9).Value = False Then GoTo NextSec2
For PrintCount1 = 0 To UBound(BenchResults(), 1)
    If BenchResults(PrintCount1, 0, 0) <> Empty_Code Then
    TextBuilder = TextBuilder & Chr(13) & BenchDatArray(PrintCount1, 0) & Space_Code & Language(242) & Space_Code & LCase(BenchDatArray(PrintCount1, 2)) & Space_Code & Language(243) & Space_Code & LCase(BenchDatArray(PrintCount1, 1)) & Chr(13)
    For PrintCount2 = 0 To UBound(BenchResults(), 3)
        If BenchResults(PrintCount1, 0, PrintCount2) <> Empty_Code Then
            TextBuilder = TextBuilder & BenchResults(PrintCount1, 0, PrintCount2) & Colon_Code & TabSpaces & BenchResults(PrintCount1, 1, PrintCount2) & TabSpaces & BenchResults(PrintCount1, 2, PrintCount2) & TabSpaces & BenchResults(PrintCount1, 3, PrintCount2) & TabSpaces & BenchResults(PrintCount1, 4, PrintCount2) & TabSpaces & BenchResults(PrintCount1, 5, PrintCount2) & Chr(13)
        End If
    Next PrintCount2
    End If
Next PrintCount1

NextSec2: If DoallSections = False Then If Manager.OptionBox(11).Value = False Then GoTo NextSec3
TextBuilder = TextBuilder & LineFeed
For PrintCount1 = 0 To UBound(PlatFormUsers(), 2)
    If PlatFormUsers(0, PrintCount1) <> Empty_Code Then
        TextBuilder = TextBuilder & PlatFormUsers(0, PrintCount1) & Colon_Code & TabSpaces & PlatFormUsers(1, PrintCount1) & TabSpaces & PlatFormUsers(2, PrintCount1) & TabSpaces & PlatFormUsers(3, PrintCount1) & TabSpaces & PlatFormUsers(4, PrintCount1) & TabSpaces & PlatFormUsers(5, PrintCount1) & Chr(13)
    End If
Next PrintCount1

NextSec3: TextBuilder = Replace(TextBuilder, "&&", "&")
ManagerSub.RichTextBox.TextRTF = TextBuilder
If DoallSections = False Then If Manager.OptionBox(10).Value = False Then GoTo Ed
For PrintCount1 = 0 To UBound(BenchResults(), 1)
    If BenchResults(PrintCount1, 0, 0) <> Empty_Code Then
        Call Add_Chart_to_RichtextBox(PrintCount1)
    End If
Next PrintCount1

Ed: 'Call ManagerSub.RichTextBox.SaveFile("c:\temp.rtf", rtfRTF)
End Sub

Public Sub Check_Whats_OnTop()
If Manager.ControllerBox(22).Visible = True Then
    TaskStyle = WS_VISIBLE
    Call ATM_Get_System_Processes(ATMProcessInfo())
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "gen") <> 0 Then GoTo CantAccept
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "key") <> 0 Then GoTo CantAccept
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "maker") <> 0 Then GoTo CantAccept
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "patch") <> 0 Then GoTo CantAccept
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "command") <> 0 Then GoTo CantAccept
    If InStr(1, LCase(ATMProcessInfo(0, 0)), "serial") <> 0 Then GoTo CantAccept
Else
    Manager.Timer(5).Interval = 0
    CanAcceptCode = True
End If
GoTo Ed
CantAccept: CanAcceptCode = False
If InStr(1, LCase(ATMProcessInfo(0, 0)), "pad") <> 0 Then CanAcceptCode = True
If InStr(1, LCase(ATMProcessInfo(0, 0)), "word") <> 0 Then CanAcceptCode = True
Ed: End Sub

Public Sub Add_Chart_to_RichtextBox(ChartIndex As Integer)
Dim HasImage As Boolean
HasImage = Clipboard.GetFormat(vbCFBitmap)
If HasImage Then ManagerSub.Clipbuffer.Picture = Clipboard.GetData

Call Manager.GraphBox(ChartIndex).EditCopyNow
ManagerSub.ChartBuffer.Picture = Clipboard.GetData(vbCFMetafile)
Clipboard.Clear
Clipboard.SetData ManagerSub.ChartBuffer.Picture
SendMessage ManagerSub.RichTextBox.hWnd, WM_PASTE, 0, 0

If HasImage Then Clipboard.SetData ManagerSub.Clipbuffer.Picture
Clipboard.Clear
End Sub

Public Sub Add_Version_Status()
If Generate_Input(ConnectedUsers(3, 0)) = False Then GoTo NotFullVersion
If CanAcceptCode = False Then GoTo NotFullVersion
Call Under_Analyse
FullVersion = True
ConnectedUsers(1, 0) = App_Full
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CDKey, ConnectedUsers(3, 0))
Call Show_Msg_Window(Language(254), Language(256), 1)
GoTo Ed

NotFullVersion: FullVersion = False
ConnectedUsers(1, 0) = App_Trial
ConnectedUsers(3, 0) = Empty_Code
Call Show_Msg_Window(Language(255), Language(256), 0)
Ed: Call Change_StatusBar_Text(0, Empty_Code)
Call Check_Current_GUID
Call Chat_Score_Reset(1)
Call Manager.Form_Caption_Refresh
End Sub

Public Sub Check_TimeTrial_Status()
If ConnectedUsers(1, 0) = App_Full Then GoTo Ed1
If ConnectedUsers(1, 0) = App_Trial Then
    Dim Date_Day As Integer, Date_Month As Integer, Date_Year As Integer, Date_Temp As String
    Dim Curr_Day As Integer, Curr_Month As Integer, Curr_Year As Integer
    Curr_Day = Day(Date)
    Curr_Month = Month(Date)
    Curr_Year = Year(Date)
    If Date_ERT = Empty_Code And Over_Excess = Empty_Code Then
        If Month(Date) = 12 Then
            Date_ERT = Curr_Day & Space_Code & Curr_Month & Space_Code & Curr_Year + 1
        Else
            Date_ERT = Curr_Day & Space_Code & Curr_Month + 1 & Space_Code & Curr_Year
        End If
        Call Under_Excess
        Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_StartDate, Date_ERT)
    End If
    If Date_ERT <> Over_Excess Then
        Date_Left = 0
        GoTo Ed
    End If
    Date_Temp = Date_ERT
    
    Date_Day = Left(Date_Temp, InStr(1, Date_Temp, Space_Code) - 1)
    Date_Temp = Right(Date_Temp, Len(Date_Temp) - InStr(1, Date_Temp, Space_Code))
    Date_Month = Left(Date_Temp, InStr(1, Date_Temp, Space_Code) - 1)
    Date_Temp = Right(Date_Temp, Len(Date_Temp) - InStr(1, Date_Temp, Space_Code))
    Date_Year = Date_Temp
    
    
    
    If Curr_Year <= Date_Year Then
        If Curr_Month <= Date_Month - 2 Then
            Date_Left = 0
            GoTo Ed
        End If
        If Curr_Month = Date_Month - 1 And Curr_Day < Date_Day Then
            Date_Left = 0
            GoTo Ed
        End If
        If Curr_Month = Date_Month - 1 And Curr_Day = Date_Day Then
            Date_Left = 30
            GoTo Ed
        End If
        If Curr_Month <= Date_Month And Curr_Day <= Date_Day Then
            Date_Left = Date_Day - Curr_Day
            GoTo Ed
        End If
        If Curr_Month < Date_Month And Curr_Day > Date_Day Then
            Date_Left = 30 - Curr_Day + Date_Day
            GoTo Ed
        End If
    End If
End If

OutofTime: Date_Left = 0
Ed: If Date_Left = 0 Then
    ConnectedUsers(1, 0) = App_Trial
Else
    ConnectedUsers(1, 0) = App_Timed
End If
Ed1: End Sub

Public Sub Check_Version_Status()
ConnectedUsers(3, 0) = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CDKey)
Date_ERT = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_StartDate)

'Call Over_Analyse
If ConnectedUsers(3, 0) = Empty_Code Then
    FullVersion = False
    ConnectedUsers(1, 0) = App_Trial
Else
    If Generate_Input(ConnectedUsers(3, 0)) = False Then GoTo NotFullVersion
    If Over_Analyse = False Then GoTo NotFullVersion
    FullVersion = True
    ConnectedUsers(1, 0) = App_Full
End If
GoTo Ed

NotFullVersion: FullVersion = False
ConnectedUsers(1, 0) = App_Trial
Ed: End Sub
Public Function Generate_Input(CdKeyString As String) As Boolean
Dim CurrentValue As Integer, RawString As String, OldCount As Integer

Dim LastValue As Integer, CurrentFifth As Integer, LastFifth As Integer
Dim OddCount As Integer, EvenCount As Integer, TotalValue As Integer

RawString = Replace(CdKeyString, "-", "")
If Len(RawString) <> 20 Then GoTo NotIt
For OldCount = 0 To 19
    LastValue = CurrentValue
    CurrentValue = Get_KeyVal(Mid(RawString, OldCount + 1, 1))
    
    If CurrentValue = LastValue Then GoTo NotIt
    If CurrentValue = LastValue + 1 Then GoTo NotIt
    If CurrentValue = LastValue - 1 Then GoTo NotIt
    
    If (OldCount + 1) / 5 = Int((OldCount + 1) / 5) Then
        If LastFifth = CurrentFifth Then GoTo NotIt
        LastFifth = CurrentFifth
        CurrentFifth = 0
    End If
    CurrentFifth = CurrentFifth + CurrentValue
    
    If (OldCount + 1) / 2 = Int((OldCount + 1) / 2) Then
        EvenCount = EvenCount + CurrentValue
        TotalValue = TotalValue + CurrentValue
    Else
        OddCount = OddCount + CurrentValue
        TotalValue = TotalValue - CurrentValue
    End If
    If EvenCount = OddCount Then GoTo NotIt
    
Next OldCount
If TotalValue <> 27 Then GoTo NotIt
If EvenCount - OddCount < 0 And EvenCount - OddCount > -3 Then GoTo NotIt
If (OddCount + 1) / 2 = Int((OddCount + 1) / 2) Then GoTo NotIt
If (EvenCount + 1) / 2 <> Int((EvenCount + 1) / 2) Then GoTo NotIt
Generate_Input = True

GoTo Ed
NotIt: Generate_Input = False
Ed: End Function
Private Function Get_KeyVal(KeyValue As String) As Integer
If Asc(KeyValue) < 64 Then
    Get_KeyVal = Val(KeyValue)
Else
    Get_KeyVal = (Asc(KeyValue) - 55)
End If
End Function

