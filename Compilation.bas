Attribute VB_Name = "Compilation"
Option Explicit

Public Function Compile_Info_File(FilePath As String, TargetArray() As String) As Boolean
Dim RawData As String, RawLine As String, CtrlCode As String, CtrlText As String
If Does_File_Exist(FilePath) = True Then
    ReDim TargetArray(1, 0)
    RawData = Load_File_Into_String(FilePath)
    Compile_Info_File = True
Else
    Compile_Info_File = False
    GoTo Ed
End If
While Len(RawData) <> 0
    If InStr(1, RawData, Chr(13)) = 0 Then GoTo Ed
    RawLine = Left(RawData, InStr(1, RawData, Chr(13)) - 1)
    RawData = Right(RawData, Len(RawData) - InStr(1, RawData, Chr(13)) - 1)
    If Left(RawLine, 1) <> Apost_Code And InStr(1, RawLine, Chr(13)) = 0 And Len(Trim(RawLine)) <> 0 And Len(Trim(RawLine)) <> 1 Then
        CtrlCode = Trim(Left(RawLine, InStr(1, RawLine, "=") - 1))
        CtrlText = Trim(Right(RawLine, Len(RawLine) - InStr(1, RawLine, "=")))
        Call Format_Text_String(CtrlText)
        Call Add_Index_To_StringArray(TargetArray())
        TargetArray(0, UBound(TargetArray, 2)) = CtrlCode
        TargetArray(1, UBound(TargetArray, 2)) = CtrlText
    End If
Wend
Ed: End Function

Public Sub Process_Info_File(RawData As String)
Dim RawLine As String, CtrlCode As String, CtrlText As String
While Len(RawData) <> 0
    If InStr(1, RawData, Chr(13)) = 0 Then GoTo Ed
    RawLine = Left(RawData, InStr(1, RawData, Chr(13)) - 1)
    RawData = Right(RawData, Len(RawData) - InStr(1, RawData, Chr(13)) - 1)
    If Left(RawLine, 1) <> Apost_Code And InStr(1, RawLine, Chr(13)) = 0 And Len(Trim(RawLine)) <> 0 And Len(Trim(RawLine)) <> 1 Then
        CtrlCode = Trim(Left(RawLine, InStr(1, RawLine, "=") - 1))
        CtrlText = Trim(Right(RawLine, Len(RawLine) - InStr(1, RawLine, "=")))
        Call Format_Text_String(CtrlText)
        Call Process_Inserter(LCase(CtrlCode), CtrlText)
    End If
Wend
Ed: End Sub
Public Sub Process_Inserter(Target As String, Target_Txt As String)
Select Case Target
Case "bwf"
    Call Add_Index_To_StringList(BadLangList())
    BadLangList(UBound(BadLangList())) = LCase(Target_Txt)
Case "new_bwf"
    App_NewVer(2) = Target_Txt
Case "new_cpu"
    App_NewVer(1) = Target_Txt
Case "new_ver"
    App_NewVer(0) = Target_Txt
    Call Manager.WriteBox(4).DDList_Add(Language(224), Empty_Code)
    Call Manager.WriteBox(4).SwitchToIndex(1)
Case "add_www"
    Dim RegionSection(1) As String
    RegionSection(0) = Left(Target_Txt, InStr(1, Target_Txt, ";") - 1)
    RegionSection(1) = Right(Target_Txt, Len(Target_Txt) - InStr(1, Target_Txt, ";"))
    Call Manager.WriteBox(4).DDList_Add(RegionSection(0), RegionSection(1))
    Call Manager.WriteBox(21).DDList_Add(RegionSection(0), RegionSection(1))
Case "theme_text"
    Theme_Text = Target_Txt
Case "theme_font"
    Theme_Font = Target_Txt
Case "theme_icon"
    Theme_Icon = Target_Txt
Case "theme_invert"
    Theme_Invert = Target_Txt
Case "theme_invertlight"
    Theme_InvertLight = Target_Txt
Case "theme_high"
    Theme_High = Target_Txt
Case "theme_highlight"
    Theme_HighLight = Target_Txt
Case "theme_light"
    Theme_Light = Target_Txt
Case "theme_shade"
    Theme_Shade = Target_Txt
Case "theme_vague"
    Theme_Vague = Target_Txt
Case "theme_color"
    Theme_Color = Target_Txt
Case "theme_shadow"
    Theme_Shadow = Target_Txt
Case "theme_dark"
    Theme_Dark = Target_Txt
Case "theme_pitch"
    Theme_Pitch = Target_Txt
Case "theme_vague_r"
    Theme_Vague_R = Target_Txt
Case "theme_vague_g"
    Theme_Vague_G = Target_Txt
Case "theme_vague_b"
    Theme_Vague_B = Target_Txt
End Select
Ed: End Sub


