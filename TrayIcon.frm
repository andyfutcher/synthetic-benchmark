VERSION 5.00
Begin VB.Form TrayIconForm 
   BackColor       =   &H00987054&
   BorderStyle     =   0  'None
   ClientHeight    =   3975
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5775
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
   Icon            =   "TrayIcon.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5775
End
Attribute VB_Name = "TrayIconForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
On Error Resume Next
'If NowLoading = False Then If TrayIconForm.Visible = True Then TrayIconForm.WindowState = 1
'Manager.Show 0, TrayIconForm
If Manager.Visible = True Then
    Manager.SetFocus
Else
    TrayIconForm.WindowState = 1
End If
End Sub

Private Sub Form_Load()
TrayIconForm.Move Screen.Width, Screen.Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Manager.Show 0, TrayIconForm
'Call Show_Tray_Icon
End Sub

Private Sub Form_Paint()
Manager.Show 0, TrayIconForm
'Call Show_Tray_Icon
End Sub
