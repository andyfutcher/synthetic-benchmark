Attribute VB_Name = "GenFunctions"
Option Explicit
Global Screen_Width As Long, Screen_Height As Long, MouseClick As Boolean
Global ThreePix As Integer, TwoPix As Integer, OnePix As Integer, EightPix As Integer
Global FourPix As Integer, FivePix As Integer, SixPix As Integer, SixTeenPix As Integer
Global FormWall As Long, FormHead As Long

Global Application_Path As String, Windows_Path As String, System_Path As String
Global LineFeed As String, TabSpaces As String

Global Net_PakTerminator As String, Net_PakSpiltter As String, Net_PakTransmitter As String
Dim ConvString As String, ConvArray() As String, ConvCount As Integer, ConvWidth As Integer

Type TEXTWRAPPER
    WrapText As String
    WrapHeight As Integer
    WrapLines As Integer
End Type
Type PAKSPILTER
    PakPart1 As String
    PakPart2 As String
End Type

Public Sub Get_Common_Vars()
OnePix = Screen.TwipsPerPixelX * 1
TwoPix = Screen.TwipsPerPixelX * 2
ThreePix = Screen.TwipsPerPixelX * 3
FourPix = Screen.TwipsPerPixelX * 4
FivePix = Screen.TwipsPerPixelX * 5
SixPix = Screen.TwipsPerPixelX * 6
EightPix = Screen.TwipsPerPixelX * 8
SixTeenPix = Screen.TwipsPerPixelX * 16
Screen_Width = Screen.Width + (8 * Screen.TwipsPerPixelX)
Screen_Height = Screen.Height + (8 * Screen.TwipsPerPixelY)
Application_Path = Add_The_Slash(LCase(App.Path))
Windows_Path = Add_The_Slash(LCase(WindowsDirectory))
System_Path = Add_The_Slash(LCase(SystemDirectory))

FormWall = ((Manager.Width - Manager.ScaleWidth) / 2)
FormHead = (Manager.Height - Manager.ScaleHeight) - (FormWall / 2)
App_Ver = App.Major & FullStop_Code & App.Minor & FullStop_Code & App.Revision

App_UserName = Normalize(GetUserName)
App_CompName = Normalize(GetComputerName)

SaveFolder = Application_Path

Net_PakTerminator = Chr(31)
Net_PakTransmitter = Chr(30) & Chr(30) & Chr(30) & Chr(30)
Net_PakSpiltter = Chr(30)
LineFeed = Chr(13) & Chr(10)
TabSpaces = String(8, Space_Code)
End Sub

Public Sub Make_NetPacket_Compat(RawString As String)
RawString = Replace(RawString, Net_PakTerminator, Net_PakTransmitter)
End Sub
Public Sub UnMake_NetPacket_Compat(RawString As String)
RawString = Replace(RawString, Net_PakTransmitter, Net_PakTerminator)
End Sub

Public Function Is_Array_Empty(ArrayString() As String, NumCols As Integer) As Boolean
On Error GoTo IsEmpty
Select Case NumCols
Case 1
    If ArrayString(0) <> Empty_Code Then GoTo NotEmpty
Case 2
    If ArrayString(0, 0) <> Empty_Code Then GoTo NotEmpty
Case 3
    If ArrayString(0, 0, 0) <> Empty_Code Then GoTo NotEmpty
End Select
Is_Array_Empty = True
GoTo Ed

IsEmpty: Is_Array_Empty = True
Resume Ed
NotEmpty: Is_Array_Empty = False
Ed: End Function

Public Function Add_The_Slash(RawString As String) As String
If Right(RawString, 1) <> BackSlash_Code Then RawString = RawString & BackSlash_Code
Add_The_Slash = RawString
End Function
Public Function Remove_The_Slash(RawString As String) As String
If Right(RawString, 1) = BackSlash_Code Then RawString = Left(RawString, Len(RawString) - 1)
Remove_The_Slash = RawString
End Function

Public Function Does_File_Exist(File_Path As String) As Boolean
On Error GoTo DoesNot
Open File_Path For Binary Access Read As #99
If LOF(99) <> 0 Then
    Does_File_Exist = True
    GoTo Ed
End If
GoTo Ed
DoesNot: Does_File_Exist = False
Resume Ed
Ed: Close #99
End Function

Public Sub Unload_Application()
Call Unregister_WWW_Host
ManagerSub.Hide
Manager.Hide
DoEvents
Call Hide_Tray_Icon
Unload ManagerSub
Unload Manager
DoEvents
If EndUpdatePath <> Empty_Code Then Shell EndUpdatePath, vbNormalFocus
End
End Sub

Public Function WordWrapper(WrapString As String, MaxTwipLen As Integer, FontValue As String, BoldValue As Boolean) As TEXTWRAPPER
Dim WrapTemp As String, NewWrap As String, NextWord As String, NewStrLen As Long, NewWrapLine As String, FirstLine As String, WrapperLineHeight As Long, WrapperLineCount As Long
If WrapString = Empty_Code Then GoTo Ed
ManagerSub.GenLabel.Font = FontValue
ManagerSub.GenLabel.FontBold = BoldValue
WrapTemp = Replace(WrapString, Chr(13), Space_Code)
WrapTemp = Replace(WrapTemp & Space_Code, "  ", Space_Code)
WrapperLineHeight = ManagerSub.GenLabel.Height
WrapperLineCount = 1

Do
    NextWord = Left(WrapTemp, InStr(1, WrapTemp, Space_Code, vbTextCompare))
    ManagerSub.GenLabel.Caption = NextWord
    If Trim(ManagerSub.GenLabel.Caption) = Empty_Code Then ManagerSub.GenLabel.Width = 0
    NewStrLen = ManagerSub.GenLabel.Width
    If NewStrLen > MaxTwipLen Then
        NewWrapLine = NewWrapLine & NewWrap & Chr(13) & NextWord
        NewWrap = Empty_Code
        WrapperLineHeight = WrapperLineHeight + ManagerSub.GenLabel.Height
        WrapperLineCount = WrapperLineCount + 1
    Else
        If FirstLine <> Empty_Code Then
            ManagerSub.GenLabel.Caption = FirstLine & NewWrap
            'FirstLine = Empty_Code
        Else
            ManagerSub.GenLabel.Caption = NewWrap
        End If
        If Trim(ManagerSub.GenLabel.Caption) = Empty_Code Then ManagerSub.GenLabel.Width = 0
        If ManagerSub.GenLabel.Width + NewStrLen > MaxTwipLen Then
            NewWrapLine = NewWrapLine & Trim(NewWrap) & Chr(13) & NextWord
            NewWrap = Empty_Code
            WrapperLineHeight = WrapperLineHeight + ManagerSub.GenLabel.Height
            WrapperLineCount = WrapperLineCount + 1
            FirstLine = NextWord
        Else
            NewWrap = NewWrap & NextWord
        End If
    End If
    If Len(WrapTemp) = 0 Then GoTo Ed
    WrapTemp = Replace(WrapTemp, NextWord, Empty_Code, , 1)
    'DoEvents
Loop
Ed: If NewWrap <> Empty_Code Then
    NewWrapLine = NewWrapLine & Trim(NewWrap)
End If
If NewWrapLine <> Empty_Code Then
    NewWrapLine = Replace(NewWrapLine, Space_Code & Chr(13), Chr(13))
    Do While Left(NewWrapLine, 1) = Chr(13)
        NewWrapLine = Right(NewWrapLine, Len(NewWrapLine) - 1)
        WrapperLineCount = WrapperLineCount - 1
    Loop
    Do While Right(NewWrapLine, 1) = Chr(13)
        NewWrapLine = Left(NewWrapLine, Len(NewWrapLine) - 1)
        WrapperLineCount = WrapperLineCount - 1
    Loop
End If
WordWrapper.WrapText = Trim(NewWrapLine)
WordWrapper.WrapHeight = WrapperLineHeight
WordWrapper.WrapLines = WrapperLineCount
End Function

Public Function WordElipser(ElipString As String, MaxTwipLen As Integer, FontValue As String, BoldValue As Boolean) As String
If MaxTwipLen = 0 Then
    WordElipser = Empty_Code
    GoTo Ed
Else
    If InStr(1, ElipString, Space_Code, vbTextCompare) = 0 Then
        WordElipser = ElipString
        GoTo Ed
    End If
End If
ManagerSub.GenLabel.Font = FontValue
ManagerSub.GenLabel.FontBold = BoldValue
ManagerSub.GenLabel.Caption = ElipString
If ManagerSub.GenLabel.Width < MaxTwipLen Then
    WordElipser = ElipString
    GoTo Ed
End If
'ManagerSub.GenLabel.Caption = ElipString & "..."
Do While (ManagerSub.GenLabel.Width + DotXThreeWidth) > MaxTwipLen
    ManagerSub.GenLabel.Caption = Left(ManagerSub.GenLabel.Caption, Len(ManagerSub.GenLabel.Caption) - 1)
    If ManagerSub.GenLabel.Caption = Empty_Code Then GoTo Ed
Loop
WordElipser = ManagerSub.GenLabel.Caption & "..."
Ed: End Function

Public Function WordHieght(FontValue As String, BoldValue As Boolean) As Integer
ManagerSub.GenLabel.Font = FontValue
ManagerSub.GenLabel.FontBold = BoldValue
WordHieght = ManagerSub.GenLabel.Height
End Function

Public Sub Sort_Advanced_Array(TargetArray() As String, SectionNum As Integer, SortIndex As Integer)
Dim CurrentCol As Integer, CurrentRep As Integer, CurrentLine As String, SortMode As Integer, SortCount As Long
SortCount = 0
SortMode = -1
If SectionNum <> -1 Then
    On Error Resume Next
    SortCount = Len(Filter_Sort(TargetArray(SectionNum, SortIndex, 0)))
    On Error GoTo 0
    If Len(TargetArray(SectionNum, SortIndex, 0)) >= (SortCount * 3) Then SortMode = 1
    For CurrentRep = 0 To UBound(TargetArray(), 3)
    For CurrentCol = 0 To (UBound(TargetArray(), 3) - 1)
        If TargetArray(SectionNum, SortIndex, CurrentCol + 1) <> Empty_Code Then
        If SortMode = 1 Then
            If StrComp(TargetArray(SectionNum, SortIndex, CurrentCol), TargetArray(SectionNum, SortIndex, CurrentCol + 1), vbTextCompare) = SortMode Then
                For UserCount = 0 To UBound(TargetArray(), 2)
                    CurrentLine = TargetArray(SectionNum, UserCount, CurrentCol)
                    TargetArray(SectionNum, UserCount, CurrentCol) = TargetArray(SectionNum, UserCount, CurrentCol + 1)
                    TargetArray(SectionNum, UserCount, CurrentCol + 1) = CurrentLine
                Next UserCount
            End If
        Else
            'If Filter_Sort(TargetArray(SectionNum, SortIndex, CurrentCol)) <= Filter_Sort(TargetArray(SectionNum, SortIndex, CurrentCol + 1)) = True Then
            If Val(TargetArray(SectionNum, SortIndex, CurrentCol)) <= Val(TargetArray(SectionNum, SortIndex, CurrentCol + 1)) = True Then
                For UserCount = 0 To UBound(TargetArray(), 2)
                    CurrentLine = TargetArray(SectionNum, UserCount, CurrentCol)
                    TargetArray(SectionNum, UserCount, CurrentCol) = TargetArray(SectionNum, UserCount, CurrentCol + 1)
                    TargetArray(SectionNum, UserCount, CurrentCol + 1) = CurrentLine
                Next UserCount
            End If
        End If
        End If
    Next CurrentCol
    Next CurrentRep
Else
    On Error Resume Next
    SortCount = Len(Filter_Sort(TargetArray(SortIndex, 0)))
    On Error GoTo 0
    If Len(TargetArray(SortIndex, 0)) >= (SortCount * 3) Then SortMode = 1
    For CurrentRep = 0 To UBound(TargetArray(), 2)
    For CurrentCol = 0 To (UBound(TargetArray(), 2) - 1)
        If TargetArray(SortIndex, CurrentCol + 1) <> Empty_Code Then
        If SortMode = 1 Then
            If StrComp(TargetArray(SortIndex, CurrentCol), TargetArray(SortIndex, CurrentCol + 1), vbTextCompare) = SortMode Then
                For UserCount = 0 To UBound(TargetArray(), 1)
                    CurrentLine = TargetArray(UserCount, CurrentCol)
                    TargetArray(UserCount, CurrentCol) = TargetArray(UserCount, CurrentCol + 1)
                    TargetArray(UserCount, CurrentCol + 1) = CurrentLine
                Next UserCount
            End If
        Else
            'If Filter_Sort(TargetArray(SortIndex, CurrentCol)) <= Filter_Sort(TargetArray(SortIndex, CurrentCol + 1)) = True Then
            If Val(TargetArray(SortIndex, CurrentCol)) <= Val(TargetArray(SortIndex, CurrentCol + 1)) = True Then
                For UserCount = 0 To UBound(TargetArray(), 1)
                    CurrentLine = TargetArray(UserCount, CurrentCol)
                    TargetArray(UserCount, CurrentCol) = TargetArray(UserCount, CurrentCol + 1)
                    TargetArray(UserCount, CurrentCol + 1) = CurrentLine
                Next UserCount
            End If
        End If
        End If
    Next CurrentCol
    Next CurrentRep
End If
End Sub

Public Function Filter_QuickLinks(RawString As String) As String
Dim RawStr As String
RawStr = Replace(RawString, "&&", "&")
Filter_QuickLinks = RawStr
End Function

Public Function Filter_Sort(RawString As String) As String
Dim RawStr As String, RawCount As Integer
RawStr = RawString
For RawCount = 32 To 47
    RawStr = Replace(RawStr, Chr(RawCount), Empty_Code)
Next RawCount
For RawCount = 58 To 255
    RawStr = Replace(RawStr, Chr(RawCount), Empty_Code)
Next RawCount
Filter_Sort = RawStr
End Function

Public Function Filter_Html(RawString As String) As String
Dim RawStr As String
RawStr = Replace(RawString, HTML_Enter, Space_Code)
Filter_Html = RawStr
End Function

Public Function Drv_Free_Space(DrivePath As String) As String
On Error GoTo DiskNotReady
Dim FS_System, FS_Drive
Set FS_System = CreateObject("Scripting.FileSystemObject")
Set FS_Drive = FS_System.GetDrive(FS_System.GetDriveName(DrivePath))
Drv_Free_Space = FormatNumber(Int(FS_Drive.FreeSpace / 1024), 0)
GoTo Ed
DiskNotReady: Drv_Free_Space = Empty_Code
Ed: End Function

Public Function Drv_Total_Size(DrivePath As String) As String
On Error GoTo DiskNotReady
Dim FS_System, FS_Drive
Set FS_System = CreateObject("Scripting.FileSystemObject")
Set FS_Drive = FS_System.GetDrive(FS_System.GetDriveName(DrivePath))
Drv_Total_Size = FormatNumber(Int(FS_Drive.TotalSize / 1024), 0)
GoTo Ed
DiskNotReady: Drv_Total_Size = Empty_Code
Ed: End Function

Public Function Drv_Type(DrivePath As String) As String
Dim FS_System, FS_Drive
Set FS_System = CreateObject("Scripting.FileSystemObject")
Set FS_Drive = FS_System.GetDrive(FS_System.GetDriveName(DrivePath))
Select Case FS_Drive.DriveType
    Case 0: Drv_Type = Language(0)
    Case 1: Drv_Type = Language(244)
    Case 2: Drv_Type = Language(245)
    Case 3: Drv_Type = Language(246)
    Case 4: Drv_Type = Language(247)
    Case 5: Drv_Type = Language(248)
End Select
End Function

Public Function Drv_Type_Code(DrivePath As String) As Integer
Dim FS_System, FS_Drive
Set FS_System = CreateObject("Scripting.FileSystemObject")
Set FS_Drive = FS_System.GetDrive(FS_System.GetDriveName(DrivePath))
Drv_Type_Code = FS_Drive.DriveType
End Function

Public Function Drv_Label(DrivePath As String) As String
On Error GoTo GotNoName
Drv_Label = Dir(DrivePath, vbVolume)
GoTo Ed
GotNoName: Drv_Label = Empty_Code
Resume Ed
Ed: End Function

Public Function Normalize(OriginalName As String) As String
Dim Trailing_Space As Boolean, Trailing_Dbl_Space As Boolean, Temp_Count As Long, TempStr As String, OldName As String
OldName = Filter_Html(OriginalName)
If Trim(OldName) = Empty_Code Then
    Normalize = Empty_Code
    GoTo Ed
End If
Trailing_Space = True
Temp_Count = 1
OldName = LCase(OldName)
Do While Temp_Count <> Len(OldName)
    If Trailing_Space = True Or Trailing_Dbl_Space = True Then
        TempStr = UCase(Mid(OldName, Temp_Count, 1))
        OldName = Mid(OldName, 1, Temp_Count - 1) & TempStr & Mid(OldName, Temp_Count + 1, Len(OldName) - Temp_Count + 1)
    End If
    Trailing_Space = False
    If Mid(OldName, Temp_Count, 1) = Space_Code Then Trailing_Space = True
   'If Mid(OldName, Temp_Count, 1) = FrontSlash_Code Then Trailing_Space = True
   'If Mid(OldName, Temp_Count, 1) = BackSlash_Code Then Trailing_Space = True
   'If Mid(OldName, Temp_Count, 1) = RightParenth_Code Then Trailing_Space = True
    If Mid(OldName, Temp_Count, 1) = "-" Then Trailing_Space = True
    If Mid(OldName, Temp_Count, 1) = ". " Then Trailing_Dbl_Space = True
    If Mid(OldName, Temp_Count, 1) = Chr(13) Then Trailing_Space = True
    Temp_Count = Temp_Count + 1
Loop
Normalize = Trim(OldName)
Ed: End Function

Public Sub Convert_List_To_String(ArrayData() As String, StringData As String)
ConvString = Empty_Code
For ConvWidth = 0 To UBound(ArrayData, 1)
    ConvString = ConvString & ArrayData(ConvWidth) & Chr(0)
Next ConvWidth
StringData = ConvString
End Sub
Public Sub Convert_String_To_List(StringData As String, ArrayData() As String)
ReDim ArrayData(0)
Do While Len(StringData) <> 0
    Call Add_Index_To_StringList(ArrayData())
    ArrayData(UBound(ArrayData, 1)) = Left(StringData, InStr(1, StringData, Chr(0)) - 1)
    StringData = Right(StringData, Len(StringData) - InStr(1, StringData, Chr(0)))
Loop
End Sub

Public Sub Convert_Limted_Array_To_String(ArrayData() As String, StringData As String)
Dim LimitIndex As Integer
LimitIndex = UBound(ArrayData, 2)
If LimitIndex > 49 Then LimitIndex = 49
ConvString = UBound(ArrayData, 1) & Chr(0)
For ConvCount = 0 To LimitIndex
    For ConvWidth = 0 To UBound(ArrayData, 1)
        ConvString = ConvString & ArrayData(ConvWidth, ConvCount) & Chr(0)
    Next ConvWidth
Next ConvCount
StringData = ConvString
End Sub
Public Sub Convert_Array_To_String(ArrayData() As String, StringData As String)
ConvString = UBound(ArrayData, 1) & Chr(0)
For ConvCount = 0 To UBound(ArrayData, 2)
    For ConvWidth = 0 To UBound(ArrayData, 1)
        ConvString = ConvString & ArrayData(ConvWidth, ConvCount) & Chr(0)
    Next ConvWidth
Next ConvCount
StringData = ConvString
End Sub
Public Sub Convert_Triple_Array_To_String(ArrayData() As String, StringData As String, SectionIndex As Integer)
ConvString = UBound(ArrayData, 2) & Chr(0)
For ConvCount = 0 To UBound(ArrayData, 3)
    For ConvWidth = 0 To UBound(ArrayData, 2)
        ConvString = ConvString & ArrayData(SectionIndex, ConvWidth, ConvCount) & Chr(0)
    Next ConvWidth
Next ConvCount
StringData = ConvString
End Sub
Public Sub Convert_String_To_Array(StringData As String, ArrayData() As String)
ReDim ArrayData(Left(StringData, 1), 0)
StringData = Right(StringData, Len(StringData) - InStr(1, StringData, Chr(0)))
Do While Len(StringData) <> 0
    Call Add_Index_To_StringArray(ArrayData())
    For ConvWidth = 0 To UBound(ArrayData, 1)
        ArrayData(ConvWidth, UBound(ArrayData, 2)) = Left(StringData, InStr(1, StringData, Chr(0)) - 1)
        StringData = Right(StringData, Len(StringData) - InStr(1, StringData, Chr(0)))
    Next ConvWidth
Loop
End Sub

Public Function Give_Path_Name_Only(RawString As String) As String
If RawString = Empty_Code Then GoTo Ed
Dim RawString1 As String
RawString1 = RawString
While Right(RawString1, 1) <> BackSlash_Code
    RawString1 = Left(RawString1, Len(RawString1) - 1)
Wend
Give_Path_Name_Only = RawString1
Ed: End Function
Public Function Give_Last_Name_Only(RawString As String) As String
Dim RawString1 As String
RawString1 = RawString
RawString1 = Remove_The_Slash(RawString1)
'If Right(RawString1, 1) = BackSlash_Code Then RawString1 = Left(RawString1, Len(RawString1) - 1)
While InStr(1, RawString1, BackSlash_Code) <> 0
    RawString1 = Right(RawString1, Len(RawString1) - InStr(1, RawString1, BackSlash_Code))
Wend
Give_Last_Name_Only = RawString1
End Function

Public Function Generate_GUID() As String
Dim TempGUID As String, GUID_Count As Integer
Randomize
For GUID_Count = 0 To 9
    If (Int(Rnd * 2) + 1) = 1 Then
        TempGUID = TempGUID & Chr(Int(Rnd * 25) + 65)
    Else
        TempGUID = TempGUID & (Int(Rnd * 9) + 1)
    End If
    DoEvents
Next GUID_Count
Generate_GUID = UCase(TempGUID)
End Function

Public Function Is_Array_NonZero(TargetArray() As Byte) As Boolean
Dim Array_NZ As Long
On Error GoTo Non_Array
Array_NZ = UBound(TargetArray)
Is_Array_NonZero = True
GoTo Ed
Non_Array: Is_Array_NonZero = False
Resume Ed
Ed: End Function

Public Sub Convert_String_ByteArray(SourceString As String, TargetArray() As Byte)
Dim SourceCount As Long
ReDim TargetArray(0 To Len(SourceString) - 1)
For SourceCount = 0 To Len(SourceString) - 1
    TargetArray(SourceCount) = Asc(Mid(SourceString, SourceCount + 1, 1))
Next SourceCount
End Sub

Public Sub Convert_ByteArray_To_String(SourceArray() As Byte, TargetString As String)
Dim SourceCount As Long
For SourceCount = 0 To UBound(SourceArray)
    TargetString = TargetString & Chr(SourceArray(SourceCount))
Next SourceCount
End Sub

Public Sub Format_Text_String(New_Text As String)
New_Text = Replace(New_Text, HTML_Enter, Chr(13))
End Sub

Public Function Load_File_Into_Array(FileName As String, Target_Array() As Byte) As Boolean
On Error GoTo FileErr
Open FileName For Binary Access Read As #1
    ReDim Target_Array(0 To LOF(1) - 1)
    Get #1, , Target_Array
Close #1
Load_File_Into_Array = True
GoTo Ed
FileErr: Load_File_Into_Array = False
    Close #1
    Resume Ed
Ed: End Function

Public Function Load_File_Into_String(FileName As String) As String
Dim Readbuffer As String
If FileLen(FileName) = 0 Then
    Load_File_Into_String = Empty_Code
    GoTo Ed
End If
Open FileName For Binary Access Read As #1
    Readbuffer = Input(LOF(1), #1)
Close #1
Load_File_Into_String = Readbuffer
Ed: End Function

Public Function Switch_Boolean(SwitchBool As Boolean) As Boolean
If SwitchBool = True Then
    Switch_Boolean = False
Else
    Switch_Boolean = True
End If
End Function

Public Function Write_Array_Into_File(FileName As String, Source_Array() As Byte) As Boolean
On Error GoTo CouldNot
Open FileName For Binary Access Write As #1
    Put #1, 1, Source_Array
Close #1
Write_Array_Into_File = True
GoTo Ed
CouldNot: Write_Array_Into_File = False
    Close #1
    Resume Ed
Ed: DoEvents
End Function

Public Function Write_String_Into_File(FileName As String, SourceData As String) As Boolean
On Error GoTo CouldNot
Open FileName For Output As #1
    Print #1, SourceData
Close #1
Write_String_Into_File = True
GoTo Ed
CouldNot: Write_String_Into_File = False
    Close #1
    Resume Ed
Ed: DoEvents
End Function

Public Sub Add_Index_To_StringArray(TargetArray() As String)
If TargetArray(0, UBound(TargetArray, 2)) <> Empty_Code Then ReDim Preserve TargetArray(UBound(TargetArray, 1), UBound(TargetArray, 2) + 1)
End Sub
Public Sub Add_Index_To_StringList(TargetArray() As String)
If TargetArray(0) <> Empty_Code Then ReDim Preserve TargetArray(UBound(TargetArray, 1) + 1)
End Sub

Public Sub Remove_Index_From_StringArray(TargetArray() As String, IndexNum As Long)
Dim FlowCount1 As Integer, FlowCount2 As Integer
For FlowCount1 = IndexNum To UBound(TargetArray(), 2) - 1
    For FlowCount2 = 0 To UBound(TargetArray(), 1)
        TargetArray(FlowCount2, FlowCount1) = TargetArray(FlowCount2, FlowCount1 + 1)
    Next FlowCount2
Next FlowCount1
If UBound(TargetArray(), 2) = 0 Then
    For FlowCount2 = 0 To UBound(TargetArray(), 1)
        TargetArray(FlowCount2, 0) = Empty_Code
    Next FlowCount2
Else
    ReDim Preserve TargetArray(UBound(TargetArray(), 1), UBound(TargetArray(), 2) - 1)
End If
End Sub

Public Function FilterBadLang(OldChat As String) As String
If Manager.OptionBox(17).Value = False Then GoTo Ed
For SettingsCount = 0 To UBound(BadLangList())
    Do While InStr(1, LCase(OldChat), BadLangList(SettingsCount)) <> 0
        OldChat = Mid(OldChat, 1, InStr(1, LCase(OldChat), BadLangList(SettingsCount))) & String(Len(BadLangList(SettingsCount)) - 1, Chat_BadLang) & Mid(OldChat, InStr(1, LCase(OldChat), BadLangList(SettingsCount)) + Len(BadLangList(SettingsCount)), Len(OldChat) - InStr(1, LCase(OldChat), BadLangList(SettingsCount)) + Len(BadLangList(SettingsCount)))
    Loop
Next SettingsCount
Ed: FilterBadLang = OldChat
End Function

Public Sub Flow_Triple_Array(TargetArray() As String, SectionNum As Integer)
Dim FlowCount1 As Integer, FlowCount2 As Integer
'If SectionNum > UBound(TargetArray(), 1) Then GoTo Ed
FlowCount1 = UBound(TargetArray, 3)
Do While FlowCount1 <> 0
    For FlowCount2 = 0 To UBound(TargetArray, 2)
        TargetArray(SectionNum, FlowCount2, FlowCount1) = TargetArray(SectionNum, FlowCount2, FlowCount1 - 1)
    Next FlowCount2
    FlowCount1 = FlowCount1 - 1
Loop
'Ed:
End Sub

Public Function Paket_Spiltter(PakData As String) As PAKSPILTER
Paket_Spiltter.PakPart1 = Left(PakData, InStr(1, PakData, Net_PakSpiltter) - 1)
Paket_Spiltter.PakPart2 = Right(PakData, Len(Net_PakSpiltter) - InStr(1, PakData, Net_PakSpiltter))
End Function

Public Function Kill_File(File_Path As String) As Boolean
On Error GoTo CantKill
Kill File_Path
Kill_File = True
GoTo Ed
CantKill: Resume Ed
Ed: End Function

Public Function Check_Download_Integrity(TargetArray() As Byte) As Boolean
Dim HTMLString As String
Call Convert_ByteArray_To_String(TargetArray(), HTMLString)
If InStr(1, LCase(HTMLString), Net_DefCant) = 0 And InStr(1, LCase(HTMLString), Language(18)) = 0 And Len(HTMLString) > 1 Then
    Check_Download_Integrity = True
Else
    Check_Download_Integrity = False
End If
End Function

Public Function Check_HTML_Integrity(HTMLString As String) As Boolean
If InStr(1, LCase(HTMLString), Net_DefCant) = 0 And InStr(1, LCase(HTMLString), Language(18)) = 0 And Len(HTMLString) > 1 Then
    Check_HTML_Integrity = True
Else
    Check_HTML_Integrity = False
End If
End Function

Public Sub Build_File_Database(FilesArray() As String, FolderPath As String)
Dim Folder() As String, DirName As String, DirPath As String, DirLoop As Long, FileCount As Long
ReDim Folder(0)
ReDim FilesArray(1, 0)
Folder(0) = Add_The_Slash(FolderPath)
DirPath = Folder(0)
On Error Resume Next
Loop1: DirName = Dir(DirPath, vbReadOnly Or vbArchive Or vbNormal Or vbArchive Or vbSystem Or vbHidden Or vbDirectory)
While DirName <> Empty_Code
    If DirName <> FullStop_Code And DirName <> ".." Then
        If (GetAttr(DirPath & DirName) And vbDirectory) = vbDirectory Then
            ReDim Preserve Folder(UBound(Folder) + 1)
            Folder(UBound(Folder)) = DirPath & DirName
        Else
            ReDim Preserve FilesArray(1, FileCount + 1)
            FilesArray(0, FileCount) = LCase(DirPath)
            FilesArray(1, FileCount) = LCase(DirName)
            'FilesArray(2, FileCount) = FileLen(DirPath & DirName)
            FileCount = FileCount + 1
        End If
    End If
    DirName = Dir
Wend
Next_FOL: DirLoop = DirLoop + 1
DirPath = Add_The_Slash(Folder(DirLoop))
DoEvents
DirName = Empty_Code
If DirLoop <> UBound(Folder) + 1 Then GoTo Loop1
Ed: End Sub

Public Function Over_Analyse() As Boolean
Over_Analyse = False
Dim FileString As String, FileArray() As Byte

If Does_File_Exist(Windows_Path & Winsock_Name & Logging_Ext) = False Then GoTo Ed

Call Load_File_Into_Array(Windows_Path & Winsock_Name & Logging_Ext, FileArray())
Call DeCompressData(FileArray(), MaxSave_Long)
Call Convert_ByteArray_To_String(FileArray(), FileString)
If FileString <> ConnectedUsers(3, 0) Then GoTo Ed
Over_Analyse = True
Ed: End Function

Public Function Over_Excess() As String
Over_Excess = Empty_Code
Dim FileString As String, FileArray() As Byte

If Does_File_Exist(Windows_Path & Winsock2_Name & Logging_Ext) = False Then GoTo Ed

Call Load_File_Into_Array(Windows_Path & Winsock2_Name & Logging_Ext, FileArray())
Call DeCompressData(FileArray(), MaxSave_Long)
Call Convert_ByteArray_To_String(FileArray(), FileString)
Over_Excess = FileString
Ed: End Function

Public Sub Under_Analyse()
Dim FileString As String, FileArray() As Byte

FileString = ConnectedUsers(3, 0)
Call Convert_String_ByteArray(FileString, FileArray())
Call CompressData(FileArray(), 9)
Call Kill_File(Windows_Path & Winsock_Name & Logging_Ext)
Call Write_Array_Into_File(Windows_Path & Winsock_Name & Logging_Ext, FileArray())
End Sub

Public Sub Under_Excess()
Dim FileString As String, FileArray() As Byte

FileString = Date_ERT
Call Convert_String_ByteArray(FileString, FileArray())
Call CompressData(FileArray(), 9)
Call Kill_File(Windows_Path & Winsock_Name & Logging_Ext)
Call Write_Array_Into_File(Windows_Path & Winsock2_Name & Logging_Ext, FileArray())
End Sub


'Public Sub Force_Make_Folder(FolderName As String)
'On Error Resume Next
'Dim FolderStr As String, FolderCnt As Integer
'FolderName = Add_The_Slash(FolderName)
'For FolderCnt = 1 To Len(FolderName)
'    FolderStr = Left(FolderName, FolderCnt)
'    If Right(FolderStr, 1) = BackSlash_Code Then
'        MkDir Left(FolderStr, Len(FolderStr) - 1)
'    End If
'Next X
'End Sub
