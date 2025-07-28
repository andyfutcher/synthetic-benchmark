VERSION 5.00
Begin VB.Form StartForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6300
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StartForm.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   4800
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer CPUDet 
      Interval        =   4000
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer 
      Index           =   0
      Left            =   120
      Top             =   120
   End
   Begin VB.Label FrameLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Index           =   5
      Left            =   6090
      TabIndex        =   5
      Top             =   360
      Width           =   60
   End
   Begin VB.Label FrameLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   195
   End
   Begin VB.Label FrameLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reg:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label FrameLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   3
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label FrameLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"StartForm.frx":000C
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   680
      Index           =   1
      Left            =   3075
      TabIndex        =   1
      Top             =   4045
      Width           =   3135
   End
   Begin VB.Label FrameLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   45
   End
   Begin VB.Image FrameImage 
      Height          =   4800
      Left            =   0
      MousePointer    =   11  'Hourglass
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6300
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
NowLoading = True
Call GetWinVersion
If App_WinVersion <> WindowsXP_Code Then GoTo Ed
Dim Comctls As INITCOMMONCONTROLSEX_TYPE
With Comctls
    .DwSize = Len(Comctls)
    .DwICC = ICC_INTERNET_CLASSES
End With
RetVal = InitCommonControlsEx(Comctls)
Ed: End Sub

Private Sub Form_Load()
FullVersion = False
CanAcceptCode = True
Call Get_Common_Vars
Call Reset_Data_Arrays(0)
Call Check_Version_Status
Call Check_TimeTrial_Status
Call Get_Network_Defaults
Theme_Current = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CurrentTheme)
Language_Current = QueryValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_CurrentLang)
StartForm.FrameImage.Picture = LoadPicture(Application_Path & Data_Path & Graphic_Name & Zero_Code & Resource_Ext)

FrameLabel(0).Caption = ConnectedUsers(1, 0) & " Version"
FrameLabel(3).Caption = Def_IPAddress & Space_Code & Normalize(Def_FriendlyName)
If Def_WWWAddress <> Empty_Code Then
    FrameLabel(3).Caption = FrameLabel(3).Caption & " and " & Def_WWWAddress
End If
FrameLabel(3).Caption = Trim(FrameLabel(3).Caption)
FrameLabel(5).Caption = "Version " & App_Ver 'Language(253)

'Call Form_ZOrder
StartForm.Move (Screen.Width / 2) - (FrameImage.Width / 2), (Screen.Height / 2) - (FrameImage.Height / 2), FrameImage.Width, FrameImage.Height
StartForm.Show
DoEvents
CreateNewKey HKEY_LOCAL_MACHINE, Reg_DefAddress
Randomize


'Manager.WeBrowser.Offline = True
'ManagerSub.GenBrowser.Offline = True
Call Prepare_Controls
Call Prepare_Language
Call GetCPUInformation
'Call Highlight_FormControl(-1)
Call Load_Settings_Data(0)
Call Check_Current_GUID
Call Setup_Graphs_Like_First
Call Process_System_Defualts
Call Update_System_Information
Call Process_WriteList_Click(18)
Call Process_OptionBox_Click(1)
Call Process_OptionBox_Click(2)
Call Process_OptionBox_Click(3)
Call Process_OptionBox_Click(4)
Call Reset_Data_Limits(False)
Call Check_Password_Similarity
Call Check_BurnIn_CommandButton
Call Change_Benchmark_State(0)
Call GetBADLangFilter
Call Cycle_Our_Site
Call Browse_To_Site(0, Empty_Code)
Timer(0).Interval = 1000

Call Check_All_Explorer_Panels
Call Check_All_Menus
Call Manager.Check_Explorer_Frames
Call Align_Explorer_Frames(0)
Call Normalize_On_Click(0)
Call Net_UpDataUsers
Call Switch_Sections_To(5)
End Sub

Private Sub Timer_Timer(Index As Integer)
Timer(0).Interval = 0

Manager.Enabled = False
ManagerSub.Show
Manager.Show 0, ManagerSub
NowLoading = False
Call Normalize_Controls(0)
StartForm.Hide
Call Check_Online_Status(1)
Call Update_Connection_List
Call Check_OS_CPUDescription

Manager.Refresh
DoEvents
If Command <> Empty_Code Then
    Call Single_New_Session
    Manager.Enabled = True
    Call Load_Current_Benchmark(Trim(Command))
    Manager.Enabled = False
    Call Switch_Sections_To(0)
End If
If FullVersion = False Then Call Show_Controller_Box(21)
If Manager.OptionBox(0).Value = True Then Call Process_Menu_Command_Ids(-1, 0, 85)
Manager.Timer(0).Interval = 200
Manager.Enabled = True
Call SetFocus_Class(1, 0)
Unload StartForm
End Sub

