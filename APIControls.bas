Attribute VB_Name = "APIControls"
Option Explicit
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' ATM Calls
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal flgs As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpSting As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long
Private Declare Function ApiUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ApiCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'RegAccess
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

'CPU Detection
Private Declare Function wincpuidext Lib "sxpcpuid.dll" () As Integer
Private Declare Function wincpuid Lib "sxpcpuid.dll" () As Long
Private Declare Function wincpufeatures Lib "sxpcpuid.dll" () As Long
'public Declare Function cpurawspeed Lib "sxpcpuid.dll" () As Long
Public Declare Function cpunormspeed Lib "sxpcpuid.dll" () As Long
Private Declare Function ProcessorCount Lib "sxpcpuid.dll" () As Long

Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_ALL_ACCESS = &H3F
Global Const HKEY_NON_VOLATILE = 0
Global Const REG_SZ As Long = 1
Global Const ERROR_NONE = 0

Public Const ICC_INTERNET_CLASSES = &H800
Public Const GetDesktop = &H10
Public Const GetDocuments = &H5
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const HwndTopmost = -1
Private Const HwndNoTopmost = -2
Private Const SwpShowWindow = &H40
Private Const HWND_TOP = 0
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_SHOWWINDOW = &H40

Private Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WM_CLOSE = &H10
Private Const WM_PAINT = &HF
Private Const WM_SYSCOLORCHANGE = &H15
Private Const GW_HWNDFIRST = 0
Private Const GWL_STYLE = (-16)
Private Const GW_HWNDNEXT = 2
Private Const SW_RESTORE = 9

Global RetVal As Long, TrayIcon As NOTIFYICONDATA, NowLoading As Boolean, MemoryInfo As MEMORYSTATUS
Global MouseLoc As POINTAPI, AlgnCode As LEFTTOP, OSInfo As OSVERSIONINFO
Global TaskStyle As Long, CPUBitType As Long, CPUBitDesc As String, CPUBitMHz As Long
Global App_WinVersion As Integer, App_WinDesc As String, App_WinBuild As String
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
Private Type ITEMIDLIST
    shellID As SHITEMID
End Type
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Type LEFTTOP
    ALeft As Integer
    ATop As Long
End Type
Public Type INITCOMMONCONTROLSEX_TYPE
    DwSize As Long
    DwICC As Long
End Type
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Dim TypeStyle As Long, BufferString As String, lValue As Long, sValue As String
Dim ATMWindow As Long, IntLen As Long, StrTitle As String, ATM_HWnd As Long

' ZOrder Prodecures
' -----------------------------------------------------------------------------
Public Sub Form_ZOrder()
RetVal = SetWindowPos(StartForm.hWnd, HwndTopmost, 0, 0, 0, 0, SwpShowWindow)
End Sub

Public Sub Show_Tray_Icon()
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hWnd = ManagerSub.hWnd
TrayIcon.uId = vbNull
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.ucallbackMessage = WM_MOUSEMOVE
TrayIcon.hIcon = ManagerSub.Icon
TrayIcon.szTip = "SynthMark XP Edition" & Chr(13) & "Andy Futcher 2003" & Chr$(0)
Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
End Sub
Public Sub Hide_Tray_Icon()
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hWnd = ManagerSub.hWnd
TrayIcon.uId = vbNull
Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
End Sub

Public Function GetSpecialFolder(whichFolder As Long) As String
Dim FolderPath As String * 256, FolderApiID As ITEMIDLIST

RetVal = SHGetSpecialFolderLocation(Manager.hWnd, whichFolder, FolderApiID)

If RetVal = 0 Then
    RetVal = SHGetPathFromIDList(ByVal FolderApiID.shellID.SHItem, ByVal FolderPath)
    If RetVal Then
        GetSpecialFolder = Left(FolderPath, InStr(FolderPath, Chr(0)) - 1)
    End If
End If
End Function


Public Function WindowsDirectory() As String
Dim WinPath As String
Dim Temp
WinPath = String(145, Chr(0))
Temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function
Public Function SystemDirectory() As String
Dim SysPath As String
Dim Temp
SysPath = String(145, Chr(0))
Temp = GetSystemDirectory(SysPath, 145)
SystemDirectory = Left(SysPath, InStr(SysPath, Chr(0)) - 1)
End Function

Public Sub ATM_Process_System_Tasks()
If Manager.OptionBox(16).Value = False Then
    TaskStyle = WS_VISIBLE
Else
    TaskStyle = 0
End If
Call ATM_Get_System_Processes(ATMProcessInfo())
Call Manager.FrameList(4).Submit_Data_Array(ATMProcessInfo(), -1, 1)
End Sub

Public Sub ATM_Get_System_Processes(TrgtArray() As String)
ReDim ATMProcessInfo(1, 0)
ATMWindow = GetWindow(Manager.hWnd, GW_HWNDFIRST)
Do While ATMWindow
    If ATMWindow <> Manager.hWnd And TaskWindow(ATMWindow) Then
        IntLen = GetWindowTextLength(ATMWindow) + 1
        StrTitle = Space$(IntLen)
        IntLen = GetWindowText(ATMWindow, StrTitle, IntLen)
        If IntLen > 0 Then
            Call Add_Index_To_StringArray(TrgtArray())
            TrgtArray(0, UBound(TrgtArray, 2)) = StripTerminator(StrTitle)
            TrgtArray(1, UBound(TrgtArray, 2)) = ATMWindow
        End If
    End If
    ATMWindow = GetWindow(ATMWindow, GW_HWNDNEXT)
Loop
End Sub

Public Sub ATM_Switch_Processes()
Dim LngWW As Long

ATM_HWnd = ATMProcessInfo(1, Manager.FrameList(4).ListIndex)
LngWW = GetWindowLong(ATM_HWnd, GWL_STYLE)
If LngWW And WS_MINIMIZE Then
    RetVal = ShowWindow(ATM_HWnd, SW_RESTORE)
End If
RetVal = SetWindowPos(ATM_HWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
End Sub

Public Sub ATM_End_Process()
ATM_HWnd = ATMProcessInfo(1, Manager.FrameList(4).ListIndex)
SendMessage ATM_HWnd, WM_CLOSE, 0, 0
End Sub

Public Sub ATM_Repaint_All_Tasks()
For VisualCount = 0 To UBound(ATMProcessInfo, 2)
    ATM_HWnd = ATMProcessInfo(1, VisualCount)
    SendMessage ATM_HWnd, WM_SYSCOLORCHANGE, 0, 0
    SendMessage ATM_HWnd, WM_PAINT, 0, 0
Next VisualCount
End Sub

Private Function TaskWindow(ATMWindow As Long) As Long
TypeStyle = GetWindowLong(ATMWindow, GWL_STYLE)
If (TypeStyle And TaskStyle) = TaskStyle Then TaskWindow = True
End Function

Private Function StripTerminator(ByVal strString As String) As String
Dim IntZeroPos As Integer
IntZeroPos = InStr(strString, Chr$(0))
If IntZeroPos > 0 Then
    StripTerminator = Left$(strString, IntZeroPos - 1)
Else
    StripTerminator = strString
End If
End Function

Public Sub Begin_Mail_TechSupport()
Call Update_System_Information
Call Form_Control_Click(2)
ManagerSub.MAPISession.SignOn 'sign on
If ManagerSub.MAPISession.SessionID <> 0 Then 'signed on
    With ManagerSub.MAPIMessages
        .SessionID = ManagerSub.MAPISession.SessionID
        .Compose
        .RecipAddress = "support@andyfutcher.com"
        .MsgSubject = Empty_Code
        .MsgNoteText = LineFeed & LineFeed & PlatFormUsers(0, 0) & LineFeed & PlatFormUsers(1, 0) & LineFeed & PlatFormUsers(2, 0) & Space_Code & PlatFormUsers(3, 0) & LineFeed & PlatFormUsers(4, 0)
        .Send True
    End With
End If
ManagerSub.MAPISession.SignOff
End Sub

Public Sub Begin_Mail_AndyFutcherSubmit()
Call Form_Control_Click(2)
ManagerSub.MAPISession.SignOn
If ManagerSub.MAPISession.SessionID <> 0 Then
    With ManagerSub.MAPIMessages
        .SessionID = ManagerSub.MAPISession.SessionID
        .Compose
        .AttachmentName = Language(250) & Project_Ext
        .AttachmentPathName = Application_Path & Language(250) & Project_Ext
        .RecipAddress = "sxpsubmit@andyfutcher.com"
        .MsgSubject = Language(250)
        .MsgNoteText = LineFeed & LineFeed & PlatFormUsers(0, 0) & LineFeed & PlatFormUsers(1, 0) & LineFeed & PlatFormUsers(2, 0) & Space_Code & PlatFormUsers(3, 0) & LineFeed & PlatFormUsers(4, 0)
        .Send True
    End With
End If
ManagerSub.MAPISession.SignOff
End Sub

Public Sub Begin_Mail_Recpient()
Call Form_Control_Click(2)
ManagerSub.MAPISession.SignOn
If ManagerSub.MAPISession.SessionID <> 0 Then
    With ManagerSub.MAPIMessages
        .SessionID = ManagerSub.MAPISession.SessionID
        .Compose
        .AttachmentName = Language(250) & Richtext_Ext
        .AttachmentPathName = Application_Path & Language(250) & Richtext_Ext
        '.RecipAddress = "techsupport@andyfutcher.com"
        .MsgSubject = Language(250)
        .MsgNoteText = LineFeed & LineFeed & PlatFormUsers(0, 0) & LineFeed & PlatFormUsers(1, 0) & LineFeed & PlatFormUsers(2, 0) & Space_Code & PlatFormUsers(3, 0) & LineFeed & PlatFormUsers(4, 0)
        .Send True
    End With
End If
ManagerSub.MAPISession.SignOff
End Sub
Public Sub Begin_Mail_CPU()
ManagerSub.MAPISession.SignOn 'sign on
If ManagerSub.MAPISession.SessionID <> 0 Then 'signed on
    With ManagerSub.MAPIMessages
        .SessionID = ManagerSub.MAPISession.SessionID
        .Compose
        '.AttachmentName = Empty_Code
        '.AttachmentPathName = Empty_Code
        .RecipAddress = "hwsupport@andyfutcher.com"
        .MsgSubject = "CPU Info Update"
        .MsgNoteText = Manager.FrameLabel(39).Caption & LineFeed & Manager.WriteBox(25).Text & LineFeed & CPUBitType
        .Send False
    End With
End If
ManagerSub.MAPISession.SignOff
Call SetKeyValue(HKEY_LOCAL_MACHINE, Reg_DefAddress, Reg_SentCPUInfo, One_Code)

Call Show_Msg_Window(Language(171), Language(172), 1)
End Sub

Public Function GetComputerName() As String
BufferString = Space$(50)
If ApiCompName(BufferString, 50) > 0 Then
    GetComputerName = BufferString
    GetComputerName = RTrim(GetComputerName)
    GetComputerName = StripTerminator(GetComputerName)
Else
    GetComputerName = Empty_Code
End If
End Function
Public Function GetUserName() As String
BufferString = Space$(50)
If ApiUserName(BufferString, 50) > 0 Then
    GetUserName = BufferString
    GetUserName = RTrim(GetUserName)
    GetUserName = StripTerminator(GetUserName)
Else
    GetUserName = Empty_Code
End If
End Function

Public Sub GetCPUInformation()
Dim iType As Long, iFamily As Long, iModel As Long, Description As String
Dim FileString As String, CPUTypes() As String
    
iType = GetCPUType()
iFamily = wincpuid()
iModel = GetCPUModel()
iType = iType * 256
iFamily = iFamily * 16
CPUBitType = iType + iFamily + iModel

Call Compile_Info_File(Application_Path & CPUFile_Name & Resource_Ext, CPUTypes())
For VisualCount = 0 To UBound(CPUTypes, 2)
    If Val(CPUTypes(0, VisualCount)) = CPUBitType Then
        CPUBitDesc = CPUTypes(1, VisualCount)
    End If
Next VisualCount
If CPUBitDesc = Empty_Code Then CPUBitDesc = Language(8)
If CPUHasMMX = True Then CPUBitDesc = CPUBitDesc & Space_Code & Language(173)
Select Case ProcessorCount
Case 2
    CPUBitDesc = CPUBitDesc & Space_Code & Language(174)
Case Is >= 3
    CPUBitDesc = CPUBitDesc & Space_Code & Language(175)
End Select
End Sub

Public Sub GetBADLangFilter()
Dim ConfigBuffer As String
ConfigBuffer = Load_File_Into_String(Application_Path & BWFFile_Name & Resource_Ext)
Call Process_Info_File(ConfigBuffer)
End Sub

Private Function CPUHasMMX() As Boolean
Dim BitField As Long

BitField = wincpufeatures()
If BitField Then 'MMX CPUS should support CPUID Instructions
    CPUHasMMX = GetBit(BitField, 23)
End If
End Function
Public Function GetCPUType() As Long
Dim BitField As Integer, Bit1 As Boolean, Bit2 As Boolean, CPUType As Long
BitField = wincpuidext()
Bit1 = GetBit(BitField, 13)
Bit2 = GetBit(BitField, 12)
If Bit1 Then
    If Bit2 Then
        CPUType = 3 '11 - Reserved
    Else
        CPUType = 2 '10 - Dual CPU
    End If
Else
    If Bit2 Then
        CPUType = 1 '01 - OverDrive
        Else
        CPUType = 0 '00 - Standard OEM CPU
    End If
End If
GetCPUType = CPUType
End Function
Public Function GetCPUModel() As Long
Dim BitField As Integer, LowByte As Byte
BitField = wincpuidext() 'get LowByte of the 32bit return value while masking Lowest Nibble
LowByte = BitField And &HF0& 'shift High Nibble to LowNibble
If LowByte Then
    GetCPUModel = LowByte / 16 'avoid divide by 0 error
End If
End Function
Private Function GetBit(ByVal iValue As Long, ByVal iBitPos As Integer) As Boolean
Debug.Assert iBitPos >= 0 And iBitPos <= 31
Dim BitVal As Long

Select Case iBitPos
Case 0
    BitVal = &H1&
Case 1
    BitVal = &H2&
Case 2
    BitVal = &H4&
Case 3
    BitVal = &H8&
Case 4
    BitVal = &H10&
Case 5
    BitVal = &H20&
Case 6
    BitVal = &H40&
Case 7
    BitVal = &H80&
Case 8
    BitVal = &H100&
Case 9
    BitVal = &H200&
Case 10
    BitVal = &H400&
Case 11
    BitVal = &H800&
Case 12
    BitVal = &H1000&
Case 13
    BitVal = &H2000&
Case 14
    BitVal = &H4000&
Case 15
    BitVal = &H8000&
Case 16
    BitVal = &H10000
Case 17
    BitVal = &H20000
Case 18
    BitVal = &H40000
Case 19
    BitVal = &H80000
Case 20
    BitVal = &H100000
Case (21)
    BitVal = &H200000
Case (22)
    BitVal = &H400000
Case (23)
    BitVal = &H800000
Case (24)
    BitVal = &H1000000
Case (25)
    BitVal = &H2000000
Case (26)
    BitVal = &H4000000
Case (27)
    BitVal = &H8000000
Case (28)
    BitVal = &H10000000
Case (29)
    BitVal = &H20000000
Case (30)
    BitVal = &H40000000
Case (31)
    BitVal = &H80000000
End Select
GetBit = iValue And BitVal
End Function

Public Sub GetWinVersion()
On Error GoTo FixOS
    Dim ReturnStr As String, OSInformation As Integer
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    ReturnStr = GetVersionEx(OSInfo)
    If ReturnStr = 0 Then
        App_WinVersion = Windows98_Code
        '*MsgBox Err_OSVersion, vbExclamation
        GoTo Ed
    End If
    OSInformation = Val(OSInfo.dwMajorVersion) & Val(OSInfo.dwMinorVersion)
    Do While InStr(1, OSInformation, Zero_Code, vbTextCompare) <> 0
        OSInformation = Replace(OSInformation, Zero_Code, Empty_Code)
    Loop
    If Len(OSInformation) = 1 Then OSInformation = OSInformation * 10
    App_WinBuild = Trim(StripTerminator(OSInfo.szCSDVersion))
    Select Case Val(OSInformation)
    Case WindowsXP_Code
        App_WinVersion = WindowsXP_Code
        App_WinDesc = "Windows XP"
    Case Windows2K_Code
        App_WinVersion = Windows2K_Code
        App_WinDesc = "Windows 2000"
    Case WindowsME_Code
        App_WinVersion = WindowsME_Code
        App_WinDesc = "Windows Me"
    Case Windows98_Code
        App_WinVersion = Windows98_Code
        App_WinDesc = "Windows 98"
    Case Windows95_Code
        If OSInfo.dwPlatformId = 1 Then
            App_WinVersion = Windows95_Code
            App_WinDesc = "Windows 95"
        Else
            App_WinVersion = WindowsNT_Code
            App_WinDesc = "Windows NT"
        End If
    Case Else
        App_WinVersion = Val(OSInformation)
    End Select
GoTo Ed

FixOS: App_WinVersion = Windows98_Code
'MsgBox "Unsupported OS detected, please contact support with this error number: " & Val(OSInformation)
Resume Ed

Ed: End Sub

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant)
Dim hKey As Long
RetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, HKEY_ALL_ACCESS, hKey)
RetVal = SetValueEx(hKey, sValueName, vValueSetting)
RegCloseKey (hKey)
End Function
Public Function QueryValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
Dim hKey As Long, vValue As Variant
RetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, HKEY_ALL_ACCESS, hKey)
RetVal = QueryValueEx(hKey, sValueName, vValue)
QueryValue = StripTerminator(vValue)
RegCloseKey (hKey)
End Function
Public Function CreateNewKey(lPredefinedKey As Long, sNewKeyName As String)
Dim hNewKey As Long
RetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, HKEY_NON_VOLATILE, HKEY_ALL_ACCESS, 0&, hNewKey, RetVal)
RegCloseKey (hNewKey)
End Function

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, vValue As Variant) As Long
sValue = vValue
SetValueEx = RegSetValueExString(hKey, sValueName, 0&, REG_SZ, sValue, Len(sValue))
End Function
Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
Dim IpChDat As Long, LongCh As Long

On Error GoTo QueryValueExError
LongCh = RegQueryValueExNULL(lhKey, szValueName, 0&, REG_SZ, 0&, IpChDat)
If LongCh <> ERROR_NONE Then GoTo QueryValueExExit
sValue = String(IpChDat, 0)
LongCh = RegQueryValueExString(lhKey, szValueName, 0&, REG_SZ, sValue, IpChDat)
If LongCh = ERROR_NONE Then
    vValue = Left$(sValue, IpChDat)
Else
    vValue = Empty
End If

QueryValueExExit: QueryValueEx = LongCh
Exit Function
QueryValueExError: Resume QueryValueExExit
Ed: End Function
