Attribute VB_Name = "Graphical"
Option Explicit
Dim GraphCount1 As Integer, GraphCount2 As Integer, GraphCount3 As Integer
Dim GraphList() As Variant, GraphArray() As Variant

Public Sub Load_Application_ThemeData()
For ControlCount = 0 To 4
    Manager.PictureLoader(ControlCount).ListImages.Clear
Next ControlCount
For ControlCount = 0 To 99
    If Does_File_Exist(System_Path & Theme_Path & Add_The_Slash(Theme_Current) & ResourceA_Name & ControlCount & Resource_Ext) = False Then Exit For
    Manager.PictureLoader(0).ListImages.Add , , LoadPicture(System_Path & Theme_Path & Theme_Current & ResourceA_Name & ControlCount & Resource_Ext)
    'DoEvents
Next ControlCount
For ControlCount = 0 To 99
    If Does_File_Exist(System_Path & Theme_Path & Add_The_Slash(Theme_Current) & ResourceB_Name & ControlCount & Resource_Ext) = False Then Exit For
    Manager.PictureLoader(1).ListImages.Add , , LoadPicture(System_Path & Theme_Path & Theme_Current & ResourceB_Name & ControlCount & Resource_Ext)
    'DoEvents
Next ControlCount
For ControlCount = 0 To 99
    If Does_File_Exist(System_Path & Theme_Path & Add_The_Slash(Theme_Current) & Resource_Name & ControlCount & Resource_Ext) = False Then Exit For
    Manager.PictureLoader(2).ListImages.Add , , LoadPicture(System_Path & Theme_Path & Theme_Current & Resource_Name & ControlCount & Resource_Ext)
    'DoEvents
Next ControlCount
For ControlCount = 0 To 99
    If Does_File_Exist(Application_Path & Data_Path & Resource_Name & ControlCount & Resource_Ext) = False Then Exit For
    Manager.PictureLoader(3).ListImages.Add , , LoadPicture(Application_Path & Data_Path & Resource_Name & ControlCount & Resource_Ext, , vbLPColor)
    'DoEvents
Next ControlCount
For ControlCount = 0 To 99
    If Does_File_Exist(Application_Path & Data_Path & Resource_Name & ControlCount & Resource_Ext) = False Then Exit For
    Manager.PictureLoader(4).ListImages.Add , , LoadPicture(Application_Path & Data_Path & Resource_Name & ControlCount & Resource_Ext, , vbLPVGAColor)
    'DoEvents
Next ControlCount
End Sub

Public Sub Search_Themes(SetDefs As Boolean)
Dim SkinFileList() As String, SkinCount As Long
ReDim ThemeFindList(0)
Call Build_File_Database(SkinFileList(), System_Path & Theme_Path)
For SkinCount = 0 To UBound(SkinFileList(), 2)
    If InStr(1, LCase(SkinFileList(1, SkinCount)), Config_Name) Then
        Call Add_Index_To_StringList(ThemeFindList())
        ThemeFindList(UBound(ThemeFindList)) = Normalize(Give_Last_Name_Only(SkinFileList(0, SkinCount)))
    End If
Next SkinCount
If SetDefs = True Then Theme_Current = Add_The_Slash(ThemeFindList(0))
End Sub
Public Sub Search_Language(SetDefs As Boolean)
Dim LangFileList() As String, LangCount As Long
ReDim LangFindList(0)
Call Build_File_Database(LangFileList(), Application_Path & Lang_Path)
For LangCount = 0 To UBound(LangFileList(), 2)
    If InStr(1, LCase(LangFileList(1, LangCount)), Config_Name) Then
        Call Add_Index_To_StringList(LangFindList())
        LangFindList(UBound(LangFindList)) = Normalize(Give_Last_Name_Only(LangFileList(0, LangCount)))
    End If
Next LangCount
If SetDefs = True Then Language_Current = Add_The_Slash(LangFindList(0))
End Sub

Public Sub Graph_Update(BenchIndex As Integer)
Dim MaxTierCount As Integer, MaxPlatCount As Integer

' Getting Diamention Info
ReDim GraphList(1, Val(Manager.WriteBox(24).Text))
For GraphCount1 = 0 To UBound(BenchResults(), 3)
    For GraphCount2 = 0 To UBound(GraphList(), 2)
        If GraphList(0, GraphCount2) = Empty_Code Then
            If ConnectedUsers(0, 0) = BenchResults(BenchIndex, 0, GraphCount1) Then
                GraphList(0, GraphCount2) = GraphList(0, 0)
                GraphList(1, GraphCount2) = GraphList(1, 0)
                GraphList(0, 0) = BenchResults(BenchIndex, 0, GraphCount1)
                GraphList(1, 0) = 1
            Else
                GraphList(0, GraphCount2) = BenchResults(BenchIndex, 0, GraphCount1)
                GraphList(1, GraphCount2) = 1
            End If
            GoTo NextLevel1
        Else
            If GraphList(0, GraphCount2) = BenchResults(BenchIndex, 0, GraphCount1) Then
                GraphList(1, GraphCount2) = Val(GraphList(1, GraphCount2)) + 1
                GoTo NextLevel1
            End If
        End If
    Next GraphCount2
NextLevel1: Next GraphCount1

' Processing Diamention Info
For GraphCount1 = 0 To UBound(GraphList(), 2)
    If GraphList(0, GraphCount1) <> Empty_Code Then
        If MaxTierCount < GraphList(1, GraphCount1) Then MaxTierCount = GraphList(1, GraphCount1)
        MaxPlatCount = MaxPlatCount + 1
    End If
Next GraphCount1

'MsgBox MaxTierCount & " - " & MaxPlatCount
' Process Headers and Footers
ReDim GraphArray(MaxTierCount, MaxPlatCount)
For GraphCount1 = 1 To MaxTierCount
    GraphArray(GraphCount1, 0) = Language(240) & Space_Code & GraphCount1
Next GraphCount1
For GraphCount1 = 1 To MaxPlatCount
    GraphArray(0, GraphCount1) = GraphList(0, GraphCount1 - 1)
Next GraphCount1

' Insert Scores
'GraphArray(1, 1) = "50"
For GraphCount1 = 0 To UBound(BenchResults(), 3)
    If BenchResults(BenchIndex, 0, GraphCount1) <> Empty_Code Then
    For GraphCount2 = 1 To MaxPlatCount
        If GraphArray(0, GraphCount2) = BenchResults(BenchIndex, 0, GraphCount1) Then
            For GraphCount3 = 1 To MaxTierCount
                If GraphArray(GraphCount3, GraphCount2) = Empty_Code Then
                    GraphArray(GraphCount3, GraphCount2) = Val(BenchResults(BenchIndex, 1, GraphCount1))
                    GoTo NextLevel2
                End If
            Next GraphCount3
        End If
    Next GraphCount2
    End If
NextLevel2: Next GraphCount1

' Submit Data Array
Call Manager.GraphBox(BenchIndex).Submit_Data_Array(GraphArray())
End Sub

Public Sub Switch_Sections_To(Section_Number As Integer)
If SectionSelect = Section_Number Then GoTo Ed
SectionSelect = Section_Number
'If Section_Number = 3 And Net_ClientType = False Then
'    Call Update_System_Information
'    Call Net_UpDataUsers
'End If
Call Align_Selected_Complex_Controls
'Manager.SimpleBox(1).Top = Desktop_Top
'Manager.SimpleBox(2).Top = Desktop_Top

For VisualCount = 0 To Manager.SimpleBox.Count - 1
    Manager.SimpleBox(VisualCount).ShowThis = False
Next VisualCount
For VisualCount = 5 To 9
    Manager.ToolButton(VisualCount).Sunken = False
Next VisualCount

Select Case Section_Number
Case 5
    Manager.BackColor = Theme_Shadow
    GoTo Ed
Case 0, 1, 2, 3, 4
    Manager.BackColor = Theme_Light
End Select
Manager.ToolButton(Section_Number + 5).Sunken = True
Manager.SimpleBox(Section_Number).ShowThis = True
'If Section_Number = 4 Then Manager.Webrowser.ClientToWindow 0, 0
Ed: Call Check_All_Menus
Call Check_All_Explorer_Panels
Call Align_Explorer_Frames(0)
Call Update_Address_Bar
End Sub

Public Sub Prepare_Tip_Display()
If TipTextIndex > UBound(TipTextArray) Then TipTextIndex = 0
Manager.FrameLabel(0).Caption = "TIP: " & (TipTextIndex + 1) & " of " & (UBound(TipTextArray) + 1)
Manager.FrameLabel(1).Caption = TipTextArray(TipTextIndex)
Ed: End Sub

Public Sub Update_System_Information()
If NowLoading = False Then If CPUBitMHz = 0 Then CPUBitMHz = cpunormspeed()
GlobalMemoryStatus MemoryInfo
ConnectedUsers(0, 0) = Manager.WriteBox(6).Text
PlatFormUsers(0, 0) = Manager.WriteBox(6).Text
PlatFormUsers(1, 0) = App_WinDesc & Space_Code & App_WinBuild
PlatFormUsers(2, 0) = CPUBitDesc
If CPUBitMHz = 0 Then
    PlatFormUsers(3, 0) = Language(0)
Else
    PlatFormUsers(3, 0) = CPUBitMHz & Language(19)
End If
PlatFormUsers(4, 0) = Round((MemoryInfo.dwTotalPhys / 1024) / 1024) & Language(20) & " (" & Round((MemoryInfo.dwTotalPageFile / 1024) / 1024) & Language(20) & ")"
PlatFormUsers(5, 0) = Manager.WriteBox(7).Text
End Sub

Public Function Get_Password() As String
Manager.WriteBox(10).Text = Empty_Code
Manager.WriteBox(10).Tag = 0
Call Show_Controller_Box(8)
Call SetFocus_Class(3, 10)
Call Choose_Manager_Functionality(False, 8, 1)
Do While Manager.WriteBox(10).Tag = 0
    DoEvents
Loop
Call Hide_Controller_Box(8)
Call Choose_Manager_Functionality(True, -1, 0)
If Manager.WriteBox(10).Tag = 2 Then
    Get_Password = Bad_Code
Else
    Get_Password = Manager.WriteBox(10).Text
End If
End Function

Public Sub Choose_Manager_Functionality(EnaValue As Boolean, ExceptController As Integer, AllowHeader As Integer)
For VisualCount = AllowHeader To Manager.FormHeader.Count - 1
    Manager.FormHeader(VisualCount).Enabled = EnaValue
Next VisualCount
For VisualCount = 0 To Manager.SimpleBox.Count - 1
    Manager.SimpleBox(VisualCount).Enabled = EnaValue
Next VisualCount
For VisualCount = AllowHeader To Manager.ControllerBox.Count - 1
    If VisualCount <> ExceptController And VisualCount <> 13 Then
        Manager.ControllerBox(VisualCount).Enabled = EnaValue
    Else
        Manager.ControllerBox(VisualCount).Enabled = True
    End If
Next VisualCount
Manager.ExplorerHolder.Enabled = EnaValue
Manager.WriteBoxList.Enabled = EnaValue
End Sub

Public Sub Update_Address_Bar()
Manager.WriteBox(0).DDList_Clear
Select Case SectionSelect
Case 0, 3
    For VisualCount = 0 To 4
        If Can_Menu(71 + VisualCount) = True Then Call Manager.WriteBox(0).DDList_Add(Language(7) & Space_Code & Language(VisualCount + 1), Str(VisualCount))
    Next VisualCount
Case 1
    For VisualCount = 0 To Manager.ComplexList.Count - 1
        If Manager.ComplexList(VisualCount).Visible = True Then Call Manager.WriteBox(0).DDList_Add(Language(7) & Space_Code & BenchDatArray(VisualCount, 2) & Space_Code & BenchDatArray(VisualCount, 0), Str(VisualCount))
    Next VisualCount
Case 2
    For VisualCount = 0 To Manager.GraphBox.Count - 1
        If Manager.GraphBox(VisualCount).Visible = True Then Call Manager.WriteBox(0).DDList_Add(Language(7) & Space_Code & BenchDatArray(VisualCount, 2) & Space_Code & BenchDatArray(VisualCount, 0), Str(VisualCount))
    Next VisualCount
Case 4
    Call Manager.WriteBox(0).DDList_Add(Language(21), 150)
    Call Manager.WriteBox(0).DDList_Add(Language(22), 151)
    Call Manager.WriteBox(0).DDList_Add(Language(23), 152)
    Call Manager.WriteBox(0).DDList_Add(Language(24), 153)
Case 5
    Call Manager.WriteBox(0).DDList_Add(Language(25), 10)
    Call Manager.WriteBox(0).DDList_Add(Language(26), 11)
    Call Manager.WriteBox(0).DDList_Add(Language(27), 101)
    Call Manager.WriteBox(0).DDList_Add(Language(28), 32)
    Call Manager.WriteBox(0).DDList_Add(Language(29), 23)
End Select
End Sub

Public Sub ReProcess_Desktop_CoOrd()
Desktop_Left = Manager.ExplorerHolder.Width
If Manager.VScrollButton(1).Visible = True Then
    Desktop_Width = Manager.ScaleWidth - Manager.ExplorerHolder.Width - Manager.VScrollButton(1).Width
Else
    Desktop_Width = Manager.ScaleWidth - Manager.ExplorerHolder.Width
End If
Desktop_Top = Manager.FormHeader(4).Top + Manager.FormHeader(4).Height
Desktop_Height = Manager.ScaleHeight - Desktop_Top - Manager.FormHeader(1).Height
End Sub

Public Sub Align_Selected_Complex_Controls()
Dim VisCmplxBxs As Integer, VisCmplxTop As Long
Select Case SectionSelect
Case 0, 3, 4
    Manager.VScrollButton(1).Visible = False
Case 1, 2
    Manager.VScrollButton(1).Visible = True
End Select
Call ReProcess_Desktop_CoOrd
For VisualCount = 0 To Manager.SimpleBox.Count - 1
    Select Case VisualCount
    Case 0, 3, 4
        Manager.SimpleBox(VisualCount).Move Desktop_Left, Desktop_Top, Desktop_Width, Desktop_Height
    Case 1, 2
        Manager.SimpleBox(VisualCount).Left = Desktop_Left
        Manager.SimpleBox(VisualCount).Width = Desktop_Width
    End Select
Next VisualCount
VisCmplxBxs = SectionSelect

Select Case VisCmplxBxs
Case 0
    'Manager.VScrollButton(1).Visible = False
    Manager.FormContainer(0).Move 0, Desktop_Height - Manager.FormContainer(0).Height, Desktop_Width
    'Manager.FormContainer(1).Move 0, 0, Desktop_Width - 2760, Desktop_Height - Manager.FormContainer(0).Height - OnePix
    Manager.ChatterBox(0).Move 0 - OnePix, 0 - OnePix, Desktop_Width - 2760 + TwoPix, Desktop_Height - Manager.FormContainer(0).Height + TwoPix
    Call Manager.ChatterBox(0).Do_Resize
    Call Manager.ChatterBox(0).Submit_Data_Array(ChatLineInfo())
    Manager.FrameList(1).Move Desktop_Width - 2775 + OnePix, Manager.WriteBox(14).Height, 2775, Desktop_Height - Manager.FormContainer(0).Height - Manager.WriteBox(14).Height + ThreePix
    Manager.WriteBox(14).Move Manager.FrameList(1).Left, 0, Manager.FrameList(1).Width - OnePix
    Manager.FrameButton(1).Left = Desktop_Width - Manager.FrameButton(1).Width - EightPix
    Manager.FrameButton(0).Left = Manager.FrameButton(1).Left - Manager.FrameButton(0).Width - EightPix
    Manager.WriteBox(5).Width = Manager.FrameButton(0).Left - Manager.WriteBox(5).Left - EightPix
Case 1
    For VisualCount = 0 To Manager.ComplexList.Count - 1
        If BenchResults(VisualCount, 0, 0) <> Empty_Code Then
            Manager.ComplexList(VisualCount).Visible = True
            Manager.ScoreLabel(VisualCount).Visible = True
            Manager.ScoreLabel(VisualCount).Top = VisCmplxTop
            If Manager.SimpleBox(VisCmplxBxs).Width - (EightPix * 4) > SixTeenPix Then Manager.ComplexList(VisualCount).Move SixTeenPix, VisCmplxTop + (EightPix * 4), Manager.SimpleBox(VisCmplxBxs).Width - (EightPix * 4), Desktop_Height - (EightPix * 5)
            Call Manager.ComplexList(VisualCount).Align_Columns
            VisCmplxTop = VisCmplxTop + Desktop_Height
        Else
            Manager.ComplexList(VisualCount).Visible = False
            Manager.ScoreLabel(VisualCount).Visible = False
        End If
    Next VisualCount
    Manager.SimpleBox(VisCmplxBxs).Height = VisCmplxTop
    Manager.VScrollButton(1).Play = VisCmplxTop
    Manager.VScrollButton(1).Gap = Desktop_Height
    Call Manager.VScrollButton(1).Process_CoOrdinates(1)
Case 2
    For VisualCount = 0 To Manager.GraphBox.Count - 1
        If BenchResults(VisualCount, 0, 0) <> Empty_Code Then
            Manager.GraphBox(VisualCount).Visible = True
            Manager.StylishLabel(VisualCount).Visible = True
            Manager.StylishLabel(VisualCount).Top = VisCmplxTop
            Manager.GraphBox(VisualCount).Move SixTeenPix, VisCmplxTop + (EightPix * 4), Manager.SimpleBox(VisCmplxBxs).Width - (EightPix * 4), Desktop_Height - (EightPix * 5)
            VisCmplxTop = VisCmplxTop + Desktop_Height
        Else
            Manager.GraphBox(VisualCount).Visible = False
            Manager.StylishLabel(VisualCount).Visible = False
        End If
    Next VisualCount
    Manager.SimpleBox(VisCmplxBxs).Height = VisCmplxTop
    Manager.VScrollButton(1).Play = VisCmplxTop
    Manager.VScrollButton(1).Gap = Desktop_Height
    Call Manager.VScrollButton(1).Process_CoOrdinates(1)
Case 3
    Manager.PlatInfoCplxList(0).Move 0 - OnePix, 0 - OnePix, Desktop_Width + TwoPix, Desktop_Height + TwoPix
    Call Manager.PlatInfoCplxList(0).Align_Columns
Case 4
    'For VisualCount = 0 To Manager.WeBrowser.Count - 1
        'Manager.ScoreLabel(VisualCount).Top = VisCmplxTop
        Manager.WeBrowser.Move 0 - TwoPix, 0 - TwoPix, Desktop_Width + FourPix, Desktop_Height + FourPix
        'Manager.WeBrowser.Move 0, 0, Desktop_Width, Desktop_Height
    'Next VisualCount
Case 5
    Manager.VScrollButton(1).Visible = False
End Select

Ed: End Sub

Public Sub Make_Controller_Ontop(Index As Integer)
If Index = -1 Then
    Manager.ExplorerHolder.ZOrder 0
    Manager.FormHeader(4).ZOrder 0
    Manager.FormHeader(3).ZOrder 0
    Manager.FormHeader(2).ZOrder 0
    For ControlCount = 0 To Manager.ControllerBox.Count - 1
        Manager.ControllerBox(ControlCount).ZOrder 0
    Next ControlCount
Else
    Manager.ControllerBox(Index).ZOrder 0
End If
Manager.FormHeader(1).ZOrder 0
Manager.FormHeader(0).ZOrder 0
End Sub

Public Function Are_There_Benchmarks() As Boolean
Are_There_Benchmarks = False
For VisualCount = 0 To UBound(BenchSelArray())
    If BenchResults(VisualCount, 0, 0) <> Empty_Code Then Are_There_Benchmarks = True
Next VisualCount
End Function

Public Sub Check_Target_Url_Buttons()
If Net_ClientType = True Then
    Manager.CommandButton(26).Enabled = False
    Manager.CommandButton(27).Enabled = False
    Manager.CommandButton(28).Enabled = False
    GoTo Ed
End If
If URLBenchList(0, 0) = Empty_Code Then
    Manager.CommandButton(27).Enabled = False
    Manager.CommandButton(28).Enabled = False
Else
    Manager.CommandButton(27).Enabled = True
    Manager.CommandButton(28).Enabled = True
End If
Ed: If UrlDoneOnce = True Then
    Manager.CommandButton(27).Enabled = False
    Manager.CommandButton(28).Enabled = False
End If
End Sub

Public Sub Align_Status_Labels()
Manager.StatusLabel(2).Width = (EightPix * 20)
Manager.StatusLabel(2).Left = Manager.ScaleWidth - Manager.StatusLabel(2).Width - (EightPix * 3)
Manager.StatusLabel(1).Width = (EightPix * 12)
Manager.StatusLabel(1).Left = Manager.StatusLabel(2).Left - Manager.StatusLabel(1).Width
Manager.StatusLabel(0).Width = Manager.StatusLabel(1).Left
For VisualCount = 0 To Manager.StatusLabel.Count - 1
    Call Manager.StatusLabel(VisualCount).Elipser_Check
Next VisualCount
End Sub

Public Sub Manager_WindowState_Check()
If Manager.WindowState = 0 Then
    Manager.ToolButton(0).Left = 0
    Manager.FormSlider(0).Visible = True
Else
    Manager.ToolButton(0).Left = OnePix
    Manager.FormSlider(0).Visible = False
End If
End Sub

Public Sub Process_Flexible_Lables()
For ControlCount = 0 To Manager.LightLabel.Count - 1
    Select Case Manager.LightLabel(ControlCount).Tag
    Case Bad_Code
        Manager.LightLabel(ControlCount).ForeColor = Theme_Invert
    Case False
        Manager.LightLabel(ControlCount).ForeColor = Theme_Shadow
    Case Else
        Manager.LightLabel(ControlCount).ForeColor = Theme_Pitch
    End Select
Next ControlCount
End Sub

Public Sub Show_Controller_Box(BoxIndex As Integer)
Call Center_ControllerBoxes(BoxIndex)
Manager.ControllerBox(BoxIndex).ZOrder 0
Manager.ControllerBox(BoxIndex).Visible = True
Manager.MenuBox(1).Visible = False
Manager.MenuBox(0).Visible = False
Manager.FormHeader(1).ZOrder 0
Manager.FormHeader(0).ZOrder 0
Call Normalize_On_Click(0)
'Select Case BoxIndex
'Case 4
'    Call Manager.FrameList(0).Align_Columns
'End Select
End Sub

Public Sub Hide_Controller_Box(BoxIndex As Integer)
On Error Resume Next
Call Highlight_Menu_Button(-1)
'Manager.MenuButton(0).SetFocus
Manager.ControllerBox(BoxIndex).Visible = False
End Sub

Public Sub Center_ControllerBoxes(BoxIndex As Integer)
If BoxIndex = -1 Then
For VisualCount = 0 To Manager.ControllerBox.Count - 1
    If Desktop_Width < Manager.ControllerBox(VisualCount).Width Then
        Manager.ControllerBox(VisualCount).Move (Manager.ScaleWidth / 2) - (Manager.ControllerBox(VisualCount).Width / 2), (Desktop_Height / 2) - (Manager.ControllerBox(VisualCount).Height / 2) + Desktop_Top
    Else
        Manager.ControllerBox(VisualCount).Move (Desktop_Width / 2) - (Manager.ControllerBox(VisualCount).Width / 2) + Desktop_Left, (Desktop_Height / 2) - (Manager.ControllerBox(VisualCount).Height / 2) + Desktop_Top
    End If
Next VisualCount
Else
    If Desktop_Width < Manager.ControllerBox(BoxIndex).Width Then
        Manager.ControllerBox(BoxIndex).Move (Manager.ScaleWidth / 2) - (Manager.ControllerBox(BoxIndex).Width / 2), (Desktop_Height / 2) - (Manager.ControllerBox(BoxIndex).Height / 2) + Desktop_Top
    Else
        Manager.ControllerBox(BoxIndex).Move (Desktop_Width / 2) - (Manager.ControllerBox(BoxIndex).Width / 2) + Desktop_Left, (Desktop_Height / 2) - (Manager.ControllerBox(BoxIndex).Height / 2) + Desktop_Top
    End If
End If
'Manager.ControllerBox(BoxIndex).Move (Desktop_Width / 2) - (Manager.ControllerBox(BoxIndex).Width / 2) + Desktop_Left, (Desktop_Height / 2) - (Manager.ControllerBox(BoxIndex).Height / 2) + Desktop_Top
End Sub

Public Sub Align_ToolBar_Buttons()
For ControlCount = 1 To Manager.ToolButton.Count - 1
    Manager.ToolButton(ControlCount).Left = Manager.ToolButton(ControlCount - 1).Left + Manager.ToolButton(ControlCount - 1).Width
Next ControlCount
For ControlCount = 1 To Manager.OtherTool.Count - 1
    Manager.OtherTool(ControlCount).Left = Manager.OtherTool(ControlCount - 1).Left + Manager.OtherTool(ControlCount - 1).Width
Next ControlCount
End Sub

Public Sub Change_StatusBar_Text(BarIndex As Integer, AlText As String)
Select Case BarIndex
Case 0
    If ConnectedUsers(1, 0) = App_Trial Or ConnectedUsers(1, 0) = App_Timed Then
        StatusMenus(0) = Language(30) & " - " & Language(31)
    Else
        StatusMenus(0) = Language(30) & " - " & Language(257)
    End If
Case 1
    StatusMenus(0) = Language(32) & Space_Code & Give_Last_Name_Only(AlText) & Space_Code
Case 2
    StatusMenus(0) = StatusMenus(0) & Language(33)
Case 3
    StatusMenus(0) = Language(34) & Space_Code & Give_Last_Name_Only(AlText) & Space_Code
End Select
Call Update_StatusBars
End Sub

Public Sub Check_All_Explorer_Panels()
For ControlCount = 0 To Manager.ExplorerFrame.Count - 1 'DynamicExplorerExclude
    If InStr(1, Manager.ExplorerFrame(ControlCount).Tag, SectionSelect, vbTextCompare) <> 0 Then
        Manager.ExplorerFrame(ControlCount).Visible = True
    Else
        Manager.ExplorerFrame(ControlCount).Visible = False
    End If
Next ControlCount
End Sub

Public Sub Setup_Graphs_Like_First()
For VisualCount = 1 To Manager.GraphBox.Count - 1
    Manager.GraphBox(VisualCount).ThreeD = Manager.GraphBox(0).ThreeD
    Manager.GraphBox(VisualCount).LineGraph = Manager.GraphBox(0).LineGraph
    Manager.GraphBox(VisualCount).ShowLedgend = Manager.GraphBox(0).ShowLedgend
Next VisualCount
End Sub

Public Function Show_Common_Dialogue(BoxType As Integer, BoxHeader As String, BoxFilter As String, BoxPath As String) As String
ManagerSub.CommonDialog.MaxFileSize = 256
ManagerSub.CommonDialog.CancelError = False
ManagerSub.CommonDialog.DialogTitle = BoxHeader
ManagerSub.CommonDialog.InitDir = BoxPath
ManagerSub.CommonDialog.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
ManagerSub.CommonDialog.Filter = BoxFilter
ManagerSub.CommonDialog.FilterIndex = 1
ManagerSub.CommonDialog.FileName = Empty_Code
Select Case BoxType
Case 0
    ManagerSub.CommonDialog.ShowOpen
Case 1
    ManagerSub.CommonDialog.ShowSave
End Select
Show_Common_Dialogue = LCase(ManagerSub.CommonDialog.FileName)
End Function

Public Function Read_Alignment_Code(RawCode As String) As LEFTTOP
Dim DDControl_Code As String, DDIndex_Code As Integer
DDControl_Code = Left(RawCode, 2)
DDIndex_Code = Right(RawCode, Len(RawCode) - 2)
Select Case DDControl_Code
Case "fh"
    Read_Alignment_Code.ALeft = Manager.FormHeader(DDIndex_Code).Left
    Read_Alignment_Code.ATop = Manager.FormHeader(DDIndex_Code).Top
Case "cb"
    Read_Alignment_Code.ALeft = Manager.ControllerBox(DDIndex_Code).Left
    Read_Alignment_Code.ATop = Manager.ControllerBox(DDIndex_Code).Top
Case "sb"
    Read_Alignment_Code.ALeft = Manager.SimpleBox(DDIndex_Code).Left
    Read_Alignment_Code.ATop = Manager.SimpleBox(DDIndex_Code).Top
End Select
End Function

Public Sub Show_Msg_Window(MsgString As String, MsgHeader As String, WinType As Integer)
Manager.CommandButton(39).Caption = Language(35)
Manager.CommandButton(52).Visible = False
Select Case WinType
Case 0
    Manager.FrameImage(2).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "2" & Resource_Ext)
Case 1
    Manager.FrameImage(2).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "3" & Resource_Ext)
End Select
Manager.FrameLabel(25).Caption = WordWrapper(MsgString, 6800, Theme_Font, False).WrapText
Manager.ControllerBox(13).Caption = Trim(MsgHeader)
Manager.ControllerBox(13).Width = Manager.FrameLabel(25).Left + Manager.FrameLabel(25).Width + SixTeenPix
Manager.ControllerBox(13).Height = Manager.FrameLabel(25).Top + Manager.FrameLabel(25).Height + (SixTeenPix * 3) + EightPix
Manager.CommandButton(39).Move (Manager.ControllerBox(13).Width / 2) - (Manager.CommandButton(39).Width / 2) + SixTeenPix, Manager.FrameLabel(25).Top + Manager.FrameLabel(25).Height + SixTeenPix
Call Show_Controller_Box(13)
Call SetFocus_Class(0, 39)
End Sub

Public Function Show_Question_Window(MsgString As String, MsgHeader As String, WinType As Integer) As Boolean
Show_Question_Window = False
Call Choose_Manager_Functionality(False, 13, 1)
Manager.CommandButton(39).Caption = Language(36)
Manager.CommandButton(52).Visible = True
Select Case WinType
Case 0
    Manager.FrameImage(2).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "2" & Resource_Ext)
Case 1
    Manager.FrameImage(2).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "3" & Resource_Ext)
Case 2
    Manager.FrameImage(2).Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & "8" & Resource_Ext)
End Select
Manager.FrameLabel(25).Caption = WordWrapper(MsgString, 6800, Theme_Font, False).WrapText
Manager.ControllerBox(13).Caption = Trim(MsgHeader)
Manager.ControllerBox(13).Width = Manager.FrameLabel(25).Left + Manager.FrameLabel(25).Width + SixTeenPix
Manager.ControllerBox(13).Height = Manager.FrameLabel(25).Top + Manager.FrameLabel(25).Height + (SixTeenPix * 3) + EightPix
Manager.CommandButton(39).Move (Manager.ControllerBox(13).Width / 2) + SixTeenPix + FourPix, Manager.FrameLabel(25).Top + Manager.FrameLabel(25).Height + SixTeenPix
Manager.CommandButton(52).Move (Manager.ControllerBox(13).Width / 2) - Manager.CommandButton(39).Width + SixTeenPix - FourPix, Manager.FrameLabel(25).Top + Manager.FrameLabel(25).Height + SixTeenPix
Manager.CommandButton(52).Tag = Empty_Code
Call Show_Controller_Box(13)
Call SetFocus_Class(0, 52)
Do While Manager.CommandButton(52).Tag = Empty_Code
    DoEvents
Loop
If Manager.CommandButton(52).Tag = 1 Then Show_Question_Window = True
Call Choose_Manager_Functionality(True, -1, 0)
End Function

Public Sub Update_StatusBars()
For VisualCount = 0 To 2
    Manager.StatusLabel(VisualCount).Caption = StatusMenus(VisualCount)
Next VisualCount
End Sub
