VERSION 5.00
Begin VB.UserControl ToolTipper 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   810
   ScaleWidth      =   3195
   Begin VB.Shape TtREffect 
      BorderColor     =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label TtRCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   " tooltipper "
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   930
   End
End
Attribute VB_Name = "ToolTipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
