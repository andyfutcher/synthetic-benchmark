VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Manager 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F8F8F8&
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9885
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "Manager.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer 
      Index           =   5
      Left            =   4080
      Top             =   7560
   End
   Begin VB.Timer Timer 
      Index           =   4
      Left            =   3960
      Top             =   7560
   End
   Begin VB.Timer Timer 
      Index           =   3
      Left            =   3840
      Top             =   7560
   End
   Begin SynthMark_XP.VScrollButton VScrollButton 
      Align           =   4  'Align Right
      Height          =   7080
      Index           =   1
      Left            =   9630
      Top             =   1530
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   12488
   End
   Begin VB.Timer Timer 
      Index           =   2
      Left            =   3720
      Top             =   7560
   End
   Begin VB.Timer Timer 
      Index           =   1
      Left            =   3600
      Top             =   7560
   End
   Begin VB.Timer Timer 
      Index           =   0
      Left            =   3480
      Top             =   7560
   End
   Begin VB.PictureBox FormHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   2
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   9885
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   315
      Width           =   9885
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   14
         Left            =   8880
         TabIndex        =   193
         Tag             =   "lts"
         Top             =   30
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   423
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   7
         Left            =   0
         TabIndex        =   200
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   6
         Left            =   0
         TabIndex        =   201
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   5
         Left            =   0
         TabIndex        =   202
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   4
         Left            =   0
         TabIndex        =   203
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   204
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   205
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   206
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin SynthMark_XP.MenuButton MenuButton 
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   207
         Top             =   0
         Width           =   300
         _ExtentX        =   503
         _ExtentY        =   503
      End
      Begin VB.Image FormBackGround 
         Height          =   300
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9900
      End
   End
   Begin VB.PictureBox FormHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9885
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   9885
      Begin VB.Image FormControl 
         Height          =   300
         Index           =   0
         Left            =   9600
         Tag             =   "1"
         Top             =   15
         Width           =   285
      End
      Begin VB.Image FormControl 
         Height          =   285
         Index           =   1
         Left            =   9600
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   15
         Width           =   285
      End
      Begin VB.Image FormControl 
         Height          =   285
         Index           =   4
         Left            =   9600
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   15
         Width           =   285
      End
      Begin VB.Image FormControl 
         Height          =   285
         Index           =   3
         Left            =   9600
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   15
         Width           =   285
      End
      Begin VB.Image FormControl 
         Height          =   285
         Index           =   2
         Left            =   9600
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   15
         Width           =   285
      End
      Begin VB.Label FormCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   45
         Width           =   60
      End
      Begin VB.Image FormIcon 
         Height          =   240
         Left            =   30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   240
      End
      Begin VB.Image FormImage 
         Height          =   315
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15360
      End
   End
   Begin SynthMark_XP.SimpleBox SimpleBox 
      Height          =   6255
      Index           =   3
      Left            =   9720
      Tag             =   "1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   9340
      _ExtentY        =   7011
      Begin SynthMark_XP.ComplexList PlatInfoCplxList 
         Height          =   6255
         Index           =   0
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   5535
         _ExtentX        =   10821
         _ExtentY        =   11033
      End
   End
   Begin SynthMark_XP.SimpleBox SimpleBox 
      Height          =   6060
      Index           =   2
      Left            =   9720
      Tag             =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   10689
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   0
         Left            =   120
         TabIndex        =   169
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   1
         Left            =   120
         TabIndex        =   170
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   2
         Left            =   120
         TabIndex        =   171
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   3
         Left            =   120
         TabIndex        =   172
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   4
         Left            =   120
         TabIndex        =   173
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   5
         Left            =   120
         TabIndex        =   174
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   6
         Left            =   120
         TabIndex        =   175
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   7
         Left            =   120
         TabIndex        =   176
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   8
         Left            =   120
         TabIndex        =   177
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   9
         Left            =   120
         TabIndex        =   178
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   10
         Left            =   120
         TabIndex        =   179
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   11
         Left            =   120
         TabIndex        =   180
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   12
         Left            =   120
         TabIndex        =   181
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   13
         Left            =   120
         TabIndex        =   182
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   14
         Left            =   120
         TabIndex        =   183
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   15
         Left            =   120
         TabIndex        =   184
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   16
         Left            =   120
         TabIndex        =   185
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   17
         Left            =   120
         TabIndex        =   186
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   18
         Left            =   120
         TabIndex        =   187
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   19
         Left            =   120
         TabIndex        =   188
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   20
         Left            =   120
         TabIndex        =   189
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   21
         Left            =   120
         TabIndex        =   190
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   22
         Left            =   120
         TabIndex        =   191
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.GraphBox GraphBox 
         Height          =   5415
         Index           =   23
         Left            =   120
         TabIndex        =   192
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9551
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   23
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   22
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   21
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   19
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   18
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   480
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   3840
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
      Begin SynthMark_XP.StylishLabel StylishLabel 
         Height          =   855
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   3413
         _ExtentY        =   1508
      End
   End
   Begin SynthMark_XP.SimpleBox SimpleBox 
      Height          =   6495
      Index           =   0
      Left            =   3360
      Tag             =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   11456
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   14
         Left            =   3480
         TabIndex        =   208
         Top             =   0
         Width           =   2715
         _ExtentX        =   4577
         _ExtentY        =   476
      End
      Begin VB.PictureBox FormContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   6255
         TabIndex        =   128
         Top             =   6000
         Width           =   6255
         Begin SynthMark_XP.CommandButton FrameButton 
            Height          =   315
            Index           =   0
            Left            =   4440
            TabIndex        =   130
            Top             =   105
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
         End
         Begin SynthMark_XP.WriteBox WriteBox 
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   129
            Top             =   120
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   476
         End
         Begin SynthMark_XP.CommandButton FrameButton 
            Height          =   315
            Index           =   1
            Left            =   5160
            TabIndex        =   131
            Top             =   105
            Width           =   990
            _ExtentX        =   1085
            _ExtentY        =   556
         End
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   5670
         Index           =   1
         Left            =   3480
         TabIndex        =   127
         Top             =   360
         Width           =   2775
         _ExtentX        =   4260
         _ExtentY        =   10610
      End
      Begin SynthMark_XP.ChatterBox ChatterBox 
         Height          =   6015
         Index           =   0
         Left            =   0
         TabIndex        =   209
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   10610
      End
   End
   Begin VB.PictureBox FormHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9885
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1155
      Width           =   9885
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   0
         Left            =   840
         TabIndex        =   40
         Top             =   60
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   476
      End
      Begin VB.Line NormLine 
         BorderColor     =   &H00C0C0C0&
         Index           =   9
         X1              =   0
         X2              =   9960
         Y1              =   360
         Y2              =   360
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   7
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   6
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   5
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   4
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   3
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   2
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   1
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin SynthMark_XP.ToolButton OtherTool 
         Height          =   330
         Index           =   0
         Left            =   4680
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
      End
      Begin VB.Label LightLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   39
         Top             =   90
         Width           =   45
      End
      Begin VB.Line NormLine 
         Index           =   4
         X1              =   0
         X2              =   9960
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line NormLine 
         Index           =   3
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image FormBackGround 
         Height          =   330
         Index           =   3
         Left            =   0
         Stretch         =   -1  'True
         Top             =   30
         Width           =   15360
      End
   End
   Begin SynthMark_XP.ExplorerHolder ExplorerHolder 
      Align           =   3  'Align Left
      Height          =   7080
      Left            =   0
      Top             =   1530
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   12488
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   1575
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3625
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   76
            Tag             =   "1"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   42
            Tag             =   "1"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Tag             =   "1"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Tag             =   "1"
         Top             =   120
         Width           =   2895
         _ExtentX        =   4895
         _ExtentY        =   3625
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   75
            Tag             =   "0"
            Top             =   1920
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Tag             =   "0"
            Top             =   1560
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Tag             =   "0"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Tag             =   "0"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Tag             =   "0"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2990
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   78
            Tag             =   "3"
            Top             =   1560
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   49
            Tag             =   "3"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   48
            Tag             =   "3"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   47
            Tag             =   "3"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3625
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   77
            Tag             =   "2"
            Top             =   1440
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   50
            Tag             =   "2"
            Top             =   1080
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   45
            Tag             =   "2"
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Tag             =   "2"
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   79
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3836
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   82
            Tag             =   "4"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   81
            Tag             =   "4"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   80
            Tag             =   "4"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   86
            Tag             =   "5"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   85
            Tag             =   "5"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   84
            Tag             =   "5"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   87
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   89
            Tag             =   "6"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   88
            Tag             =   "6"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   90
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   95
            Tag             =   "7"
            Top             =   1920
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   94
            Tag             =   "7"
            Top             =   1560
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   93
            Tag             =   "7"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   92
            Tag             =   "7"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   91
            Tag             =   "7"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.ExplorerFrame ExplorerFrame 
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   96
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   100
            Tag             =   "8"
            Top             =   1560
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   99
            Tag             =   "8"
            Top             =   1200
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   98
            Tag             =   "8"
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
         Begin SynthMark_XP.ExplorerButton ExplorerButton 
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   97
            Tag             =   "8"
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   476
         End
      End
      Begin SynthMark_XP.VScrollButton VScrollButton 
         Height          =   7140
         Index           =   0
         Left            =   3120
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   12594
      End
   End
   Begin VB.PictureBox FormHeader 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   560
      Index           =   3
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9885
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   9885
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   11
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   10
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   9
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   8
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   7
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   6
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   5
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   4
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   3
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   2
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   1
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   926
      End
      Begin SynthMark_XP.ToolButton ToolButton 
         Height          =   525
         Index           =   0
         Left            =   0
         Top             =   30
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   979
      End
      Begin VB.Line NormLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   9960
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line NormLine 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image FormBackGround 
         Height          =   560
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9900
      End
   End
   Begin VB.PictureBox FormHeader 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   1
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   9885
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8610
      Width           =   9885
      Begin SynthMark_XP.StatusLabel StatusLabel 
         Height          =   255
         Index           =   2
         Left            =   0
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin SynthMark_XP.StatusLabel StatusLabel 
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin SynthMark_XP.StatusLabel StatusLabel 
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
      End
      Begin VB.Line NormLine 
         Index           =   2
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image FormSlider 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   0
         Left            =   9600
         MousePointer    =   8  'Size NW SE
         Stretch         =   -1  'True
         Tag             =   "1"
         Top             =   30
         Width           =   240
      End
      Begin VB.Image FormBackGround 
         Height          =   240
         Index           =   2
         Left            =   0
         Stretch         =   -1  'True
         Top             =   30
         Width           =   15360
      End
   End
   Begin SynthMark_XP.SimpleBox SimpleBox 
      Height          =   6255
      Index           =   4
      Left            =   9720
      Tag             =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11033
      Begin SHDocVwCtl.WebBrowser WeBrowser 
         CausesValidation=   0   'False
         Height          =   6255
         Left            =   0
         TabIndex        =   102
         Top             =   0
         Width           =   6015
         ExtentX         =   10610
         ExtentY         =   11033
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin SynthMark_XP.SimpleBox SimpleBox 
      Height          =   6255
      Index           =   1
      Left            =   9720
      Tag             =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11033
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   3
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   4
         Left            =   240
         TabIndex        =   55
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   5
         Left            =   240
         TabIndex        =   56
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   6
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   7
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   8
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   9
         Left            =   240
         TabIndex        =   60
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   10
         Left            =   240
         TabIndex        =   61
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   11
         Left            =   240
         TabIndex        =   62
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   12
         Left            =   240
         TabIndex        =   63
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   13
         Left            =   240
         TabIndex        =   64
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   14
         Left            =   240
         TabIndex        =   65
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   15
         Left            =   240
         TabIndex        =   66
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   16
         Left            =   240
         TabIndex        =   67
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   17
         Left            =   240
         TabIndex        =   68
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   18
         Left            =   240
         TabIndex        =   69
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   19
         Left            =   240
         TabIndex        =   70
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   20
         Left            =   240
         TabIndex        =   71
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   21
         Left            =   240
         TabIndex        =   72
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   22
         Left            =   240
         TabIndex        =   73
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.ComplexList ComplexList 
         Height          =   5295
         Index           =   23
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   4470
         _ExtentX        =   8308
         _ExtentY        =   9340
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   23
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   22
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   21
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   20
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   19
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   18
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   8
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
      Begin SynthMark_XP.StylishLabel ScoreLabel 
         Height          =   495
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   873
      End
   End
   Begin VB.ListBox WriteBoxList 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   0
      TabIndex        =   123
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin SynthMark_XP.ToolTipBox ToolTipBox 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
   End
   Begin SynthMark_XP.MenuBox MenuBox 
      Height          =   1695
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Tag             =   "-1"
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
   End
   Begin SynthMark_XP.MenuBox MenuBox 
      Height          =   1695
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Tag             =   "-1"
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2990
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   3735
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   8916
      _ExtentY        =   3836
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   15
         Left            =   3000
         TabIndex        =   133
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   14
         Left            =   600
         TabIndex        =   132
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   16
         Left            =   4320
         TabIndex        =   134
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Image FrameImage 
         Height          =   360
         Index           =   4
         Left            =   600
         Stretch         =   -1  'True
         Top             =   780
         Width           =   360
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   11
         Left            =   720
         TabIndex        =   139
         Top             =   2520
         Width           =   4665
         _ExtentX        =   8493
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   10
         Left            =   720
         TabIndex        =   138
         Top             =   2160
         Width           =   4665
         _ExtentX        =   8493
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   9
         Left            =   720
         TabIndex        =   137
         Top             =   1800
         Width           =   4665
         _ExtentX        =   8493
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   8
         Left            =   720
         TabIndex        =   136
         Top             =   1440
         Width           =   4665
         _ExtentX        =   8281
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   1140
         TabIndex        =   135
         Tag             =   "1"
         Top             =   765
         Width           =   4335
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   2895
      Index           =   18
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5106
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   29
         Left            =   2280
         TabIndex        =   265
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   30
         Left            =   3600
         TabIndex        =   266
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   13
         Left            =   2280
         TabIndex        =   267
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   26
         Left            =   2280
         TabIndex        =   269
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   49
         Left            =   600
         TabIndex        =   272
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   45
         Left            =   600
         TabIndex        =   271
         Tag             =   "1"
         Top             =   720
         Width           =   3465
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   44
         Left            =   720
         TabIndex        =   270
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   720
         TabIndex        =   268
         Top             =   1230
         Width           =   1305
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4335
      Index           =   17
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6800
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   25
         Left            =   2880
         TabIndex        =   255
         Top             =   2520
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   46
         Left            =   600
         TabIndex        =   256
         Top             =   3720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   47
         Left            =   3840
         TabIndex        =   257
         Top             =   3720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   48
         Left            =   5160
         TabIndex        =   264
         Top             =   3720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   43
         Left            =   600
         TabIndex        =   263
         Tag             =   "0"
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   38
         Left            =   600
         TabIndex        =   262
         Tag             =   "0"
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   42
         Left            =   600
         TabIndex        =   261
         Tag             =   "1"
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   41
         Left            =   840
         TabIndex        =   260
         Top             =   2535
         Width           =   1590
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   840
         TabIndex        =   259
         Top             =   3000
         Width           =   1590
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   39
         Left            =   2880
         TabIndex        =   258
         Tag             =   "1"
         Top             =   3000
         Width           =   3375
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4815
      Index           =   16
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   22
         Left            =   3840
         TabIndex        =   246
         Top             =   1560
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   23
         Left            =   3840
         TabIndex        =   248
         Top             =   2040
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   24
         Left            =   3840
         TabIndex        =   250
         Top             =   2520
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   44
         Left            =   3960
         TabIndex        =   252
         Top             =   4200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   45
         Left            =   5280
         TabIndex        =   253
         Top             =   4200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Line NormLine 
         Index           =   5
         X1              =   375
         X2              =   6615
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line NormLine 
         Index           =   6
         X1              =   375
         X2              =   6615
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   9
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   37
         Left            =   1320
         TabIndex        =   254
         Top             =   3150
         Width           =   5175
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   36
         Left            =   960
         TabIndex        =   251
         Tag             =   "1"
         Top             =   2535
         Width           =   1785
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   960
         TabIndex        =   249
         Tag             =   "1"
         Top             =   2055
         Width           =   2460
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   247
         Tag             =   "1"
         Top             =   1575
         Width           =   2430
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   33
         Left            =   1140
         TabIndex        =   245
         Tag             =   "1"
         Top             =   720
         Width           =   5175
      End
      Begin VB.Image FrameImage 
         Height          =   360
         Index           =   8
         Left            =   600
         Stretch         =   -1  'True
         Top             =   735
         Width           =   360
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   1575
      Index           =   15
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8281
      _ExtentY        =   2778
      Begin VB.Image FrameImage 
         Height          =   480
         Index           =   7
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   1320
         TabIndex        =   242
         Tag             =   "1"
         Top             =   885
         Width           =   2220
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4935
      Index           =   20
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   53
         Left            =   4560
         TabIndex        =   288
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   54
         Left            =   5880
         TabIndex        =   289
         Top             =   4320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   58
         Left            =   1080
         TabIndex        =   293
         Top             =   3270
         Width           =   6015
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   13
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   240
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   22
         Left            =   1080
         TabIndex        =   292
         Top             =   2640
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   21
         Left            =   1080
         TabIndex        =   291
         Top             =   2160
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   20
         Left            =   1080
         TabIndex        =   290
         Top             =   1680
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   57
         Left            =   1320
         TabIndex        =   287
         Top             =   990
         Width           =   45
      End
      Begin VB.Image FrameImage 
         Height          =   480
         Index           =   12
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   56
         Left            =   1320
         TabIndex        =   286
         Tag             =   "1"
         Top             =   735
         Width           =   60
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4575
      Index           =   19
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8070
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   50
         Left            =   3480
         TabIndex        =   273
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   58
         Left            =   720
         TabIndex        =   298
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   54
         Left            =   600
         TabIndex        =   284
         Top             =   4035
         Width           =   2820
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   53
         Left            =   600
         TabIndex        =   281
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   52
         Left            =   600
         TabIndex        =   280
         Top             =   2400
         Width           =   1980
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   51
         Left            =   600
         TabIndex        =   279
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   50
         Left            =   600
         TabIndex        =   278
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   49
         Left            =   600
         TabIndex        =   277
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright� 1999-200AndyFutcherro� Software"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   48
         Left            =   600
         TabIndex        =   276
         Top             =   975
         Width           =   3330
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   47
         Left            =   600
         TabIndex        =   275
         Top             =   1215
         Width           =   2340
      End
      Begin VB.Image FrameImage 
         Height          =   3000
         Index           =   10
         Left            =   4320
         Stretch         =   -1  'True
         Top             =   720
         Width           =   360
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "AndyFutcher� SynthMark� XP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   46
         Left            =   600
         TabIndex        =   274
         Top             =   720
         Width           =   2655
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   3015
      Index           =   22
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10398
      _ExtentY        =   5106
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   59
         Left            =   3120
         TabIndex        =   299
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   60
         Left            =   4560
         TabIndex        =   300
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   27
         Left            =   1320
         TabIndex        =   302
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   28
         Left            =   2520
         TabIndex        =   303
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   29
         Left            =   3720
         TabIndex        =   304
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   30
         Left            =   4920
         TabIndex        =   305
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   476
      End
      Begin VB.Label LightLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "When done click here:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1305
         TabIndex        =   306
         Top             =   2460
         Width           =   1590
      End
      Begin VB.Line NormLine 
         Index           =   14
         X1              =   4650
         X2              =   4755
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line NormLine 
         Index           =   13
         X1              =   3450
         X2              =   3555
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line NormLine 
         Index           =   12
         X1              =   2250
         X2              =   2355
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Image FrameImage 
         Height          =   480
         Index           =   15
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Your registration code can be found in your emailed reciept sent to you after your purchase is complete..."
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   60
         Left            =   1320
         TabIndex        =   301
         Tag             =   "1"
         Top             =   705
         Width           =   4455
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   1815
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3201
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   3
         Left            =   3720
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   2
         Left            =   2400
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   51
         Left            =   960
         TabIndex        =   282
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   3375
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   2895
      Index           =   7
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   12938
      _ExtentY        =   5106
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   19
         Left            =   600
         TabIndex        =   148
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   8
         Left            =   2160
         TabIndex        =   149
         Top             =   1200
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   9
         Left            =   2160
         TabIndex        =   150
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   20
         Left            =   3120
         TabIndex        =   151
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   21
         Left            =   4440
         TabIndex        =   155
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   4080
         TabIndex        =   0
         Top             =   1710
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   720
         TabIndex        =   154
         Top             =   1215
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   153
         Top             =   1710
         Width           =   45
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   13
         Left            =   600
         TabIndex        =   152
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   423
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   3015
      Index           =   6
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5318
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   17
         Left            =   6240
         TabIndex        =   140
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   6
         Left            =   1920
         TabIndex        =   142
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   7
         Left            =   1920
         TabIndex        =   144
         Top             =   1680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   18
         Left            =   6240
         TabIndex        =   147
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   15
         Left            =   1920
         TabIndex        =   194
         Top             =   2520
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   12
         Left            =   1920
         TabIndex        =   146
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   720
         TabIndex        =   145
         Top             =   1695
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   720
         TabIndex        =   143
         Top             =   1215
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   600
         TabIndex        =   141
         Tag             =   "1"
         Top             =   720
         Width           =   60
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6975
      Index           =   23
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   12091
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   64
         Left            =   6240
         TabIndex        =   317
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   3975
         Index           =   6
         Left            =   720
         TabIndex        =   308
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7011
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   62
         Left            =   2520
         TabIndex        =   309
         Top             =   6360
         Width           =   1345
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   63
         Left            =   4920
         TabIndex        =   311
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   61
         Left            =   600
         TabIndex        =   310
         Top             =   6360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
      End
      Begin VB.Image FrameImage 
         Height          =   3975
         Index           =   16
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label LightLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refreshing Server List..."
         Height          =   255
         Index           =   12
         Left            =   720
         TabIndex        =   316
         Tag             =   "1"
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label FrameLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   65
         Left            =   600
         TabIndex        =   315
         Tag             =   "1"
         Top             =   720
         Width           =   60
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   64
         Left            =   1080
         TabIndex        =   314
         Top             =   5190
         Width           =   6375
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   17
         Left            =   720
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   62
         Left            =   5280
         TabIndex        =   312
         Top             =   2520
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Label FrameLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   63
         Left            =   5280
         TabIndex        =   313
         Top             =   3600
         Width           =   2055
         WordWrap        =   -1  'True
      End
      Begin VB.Image FrameImage 
         Height          =   1200
         Index           =   18
         Left            =   5265
         Stretch         =   -1  'True
         Top             =   1185
         Width           =   2100
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6375
      Index           =   14
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11245
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   40
         Left            =   3360
         TabIndex        =   222
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   41
         Left            =   4680
         TabIndex        =   223
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   42
         Left            =   6120
         TabIndex        =   224
         Tag             =   "0"
         Top             =   5760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   315
         Index           =   43
         Left            =   6000
         TabIndex        =   227
         Tag             =   "0"
         Top             =   1155
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   16
         Left            =   3720
         TabIndex        =   232
         Top             =   2880
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   17
         Left            =   3720
         TabIndex        =   233
         Top             =   3360
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   18
         Left            =   3720
         TabIndex        =   236
         Top             =   4440
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   19
         Left            =   1920
         TabIndex        =   238
         Top             =   4920
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   20
         Left            =   5760
         TabIndex        =   240
         Top             =   4920
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   476
      End
      Begin VB.Line NormLine 
         Index           =   11
         X1              =   375
         X2              =   7455
         Y1              =   5535
         Y2              =   5535
      End
      Begin VB.Line NormLine 
         Index           =   10
         X1              =   375
         X2              =   7455
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label LightLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   5475
         TabIndex        =   241
         Top             =   4935
         Width           =   60
      End
      Begin VB.Label LightLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   239
         Top             =   4935
         Width           =   60
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   960
         TabIndex        =   237
         Top             =   4455
         Width           =   60
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   6
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   960
         TabIndex        =   235
         Tag             =   "1"
         Top             =   3990
         Width           =   60
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   960
         TabIndex        =   234
         Top             =   3375
         Width           =   60
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   19
         Left            =   960
         TabIndex        =   231
         Top             =   1920
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   960
         TabIndex        =   230
         Top             =   2895
         Width           =   60
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   960
         TabIndex        =   229
         Tag             =   "1"
         Top             =   2430
         Width           =   60
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   5
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   240
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   18
         Left            =   960
         TabIndex        =   228
         Top             =   1560
         Width           =   6375
         _ExtentX        =   11456
         _ExtentY        =   423
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   3
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   226
         Tag             =   "1"
         Top             =   750
         Width           =   60
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   255
         Index           =   17
         Left            =   960
         TabIndex        =   225
         Top             =   1200
         Width           =   5055
         _ExtentX        =   11456
         _ExtentY        =   423
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   2655
      Index           =   11
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4683
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   35
         Left            =   2520
         TabIndex        =   210
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   36
         Left            =   3840
         TabIndex        =   211
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   15
         Left            =   1320
         TabIndex        =   214
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   476
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   215
         Top             =   1680
         Width           =   2955
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   213
         Tag             =   "1"
         Top             =   1230
         Width           =   480
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   600
         TabIndex        =   212
         Top             =   720
         Width           =   3135
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4575
      Index           =   21
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13996
      _ExtentY        =   9551
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   55
         Left            =   945
         TabIndex        =   294
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   56
         Left            =   600
         TabIndex        =   295
         Top             =   3960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   57
         Left            =   6135
         TabIndex        =   296
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Image FrameImage 
         Height          =   3000
         Index           =   14
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   600
         TabIndex        =   297
         Top             =   720
         Width           =   3375
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   4095
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7223
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   1
         Left            =   4320
         TabIndex        =   17
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Tag             =   "1"
         Top             =   3530
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   60
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   4935
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6135
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   106
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   2
         Left            =   2040
         TabIndex        =   109
         Top             =   2160
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   3
         Left            =   4080
         TabIndex        =   111
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   9
         Left            =   4560
         TabIndex        =   113
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   8
         Left            =   3240
         TabIndex        =   104
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   7
         Left            =   600
         TabIndex        =   103
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   21
         Left            =   2040
         TabIndex        =   243
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   476
      End
      Begin VB.Label LightLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   244
         Tag             =   "0"
         Top             =   1200
         Width           =   45
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   7
         Left            =   2040
         TabIndex        =   116
         Top             =   3480
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   4
         Left            =   600
         TabIndex        =   115
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   6
         Left            =   2040
         TabIndex        =   108
         Top             =   3060
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   5
         Left            =   2040
         TabIndex        =   105
         Top             =   2640
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   423
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Index           =   7
         Left            =   1080
         TabIndex        =   114
         Top             =   3990
         Width           =   4575
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   0
         Left            =   720
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3120
         TabIndex        =   112
         Tag             =   "0"
         Top             =   2190
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   110
         Tag             =   "0"
         Top             =   2190
         Width           =   45
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   107
         Tag             =   "0"
         Top             =   1695
         Width           =   45
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   2055
      Index           =   8
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3625
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   22
         Left            =   2640
         TabIndex        =   156
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   10
         Left            =   2040
         TabIndex        =   157
         Top             =   840
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   23
         Left            =   3960
         TabIndex        =   158
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   600
         TabIndex        =   159
         Top             =   870
         Width           =   1185
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   1815
      Index           =   13
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   9128
      _ExtentY        =   3201
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   39
         Left            =   1440
         TabIndex        =   220
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   52
         Left            =   120
         TabIndex        =   283
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Image FrameImage 
         Height          =   480
         Index           =   2
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   480
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   1320
         TabIndex        =   221
         Top             =   840
         Width           =   5085
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6975
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12303
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   4
         Left            =   3240
         TabIndex        =   121
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   3660
         Index           =   0
         Left            =   720
         TabIndex        =   117
         Top             =   2280
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4683
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   12
         Left            =   3960
         TabIndex        =   120
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   11
         Left            =   2640
         TabIndex        =   119
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   10
         Left            =   600
         TabIndex        =   118
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   13
         Left            =   5280
         TabIndex        =   126
         Top             =   6360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin VB.Label LinkLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   6315
         MousePointer    =   99  'Custom
         TabIndex        =   319
         Top             =   6000
         Width           =   60
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   19
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image FrameImage 
         Height          =   240
         Index           =   11
         Left            =   600
         Stretch         =   -1  'True
         Top             =   5040
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   55
         Left            =   960
         TabIndex        =   285
         Top             =   5070
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   8
         Left            =   960
         TabIndex        =   125
         Top             =   1125
         Width           =   5415
      End
      Begin VB.Label LightLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   124
         Tag             =   "1"
         Top             =   2400
         Width           =   5775
      End
      Begin VB.Label LightLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   600
         TabIndex        =   122
         Top             =   750
         Width           =   2400
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   5880
      Index           =   9
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10372
      Begin SynthMark_XP.ProgressBox ProgressBox 
         Height          =   300
         Index           =   0
         Left            =   600
         Top             =   1710
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   529
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   2730
         Index           =   3
         Left            =   600
         TabIndex        =   198
         Top             =   2520
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5318
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   24
         Left            =   6000
         TabIndex        =   160
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   25
         Left            =   6000
         TabIndex        =   196
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   31
         Left            =   6000
         TabIndex        =   197
         Tag             =   "0"
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   23
         Left            =   600
         TabIndex        =   318
         Tag             =   "lts"
         Top             =   5375
         Width           =   6495
         _ExtentX        =   1720
         _ExtentY        =   423
      End
      Begin VB.Line NormLine 
         Index           =   7
         X1              =   375
         X2              =   7320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line NormLine 
         Index           =   8
         X1              =   375
         X2              =   7320
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Label FrameLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   1440
         TabIndex        =   195
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Image FrameImage 
         Height          =   720
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   720
         Width           =   720
      End
      Begin VB.Label FrameLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   1440
         TabIndex        =   161
         Tag             =   "1"
         Top             =   720
         Width           =   4455
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6735
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11456
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   1515
         Index           =   2
         Left            =   840
         TabIndex        =   165
         Top             =   4080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2672
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   11
         Left            =   1560
         TabIndex        =   162
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   6
         Left            =   600
         TabIndex        =   31
         Top             =   6120
         Width           =   1215
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   5
         Left            =   5640
         TabIndex        =   28
         Top             =   6120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   4
         Left            =   4320
         TabIndex        =   27
         Top             =   6120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.WriteBox WriteBox 
         Height          =   270
         Index           =   12
         Left            =   1560
         TabIndex        =   163
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   476
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   315
         Index           =   26
         Left            =   5760
         TabIndex        =   166
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   315
         Index           =   27
         Left            =   5760
         TabIndex        =   167
         Top             =   4560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   315
         Index           =   28
         Left            =   5760
         TabIndex        =   168
         Top             =   5280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
      End
      Begin VB.Label FrameLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   61
         Left            =   840
         TabIndex        =   307
         Top             =   5640
         Width           =   60
      End
      Begin VB.Label LightLabel 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   164
         Top             =   2790
         Width           =   2655
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   36
         Top             =   3240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   2
         Left            =   720
         TabIndex        =   32
         Top             =   2280
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   30
         Top             =   1350
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
      End
      Begin VB.Label LightLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   37
         Tag             =   "bad"
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label LightLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   35
         Top             =   2790
         Width           =   60
      End
      Begin VB.Label LightLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   34
         Top             =   1830
         Width           =   60
      End
      Begin VB.Label LightLabel 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   33
         Top             =   1830
         Width           =   2655
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   600
         TabIndex        =   29
         Top             =   720
         Width           =   6255
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   6015
      Index           =   10
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   10610
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   32
         Left            =   2040
         TabIndex        =   199
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   33
         Left            =   3360
         TabIndex        =   1
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   34
         Left            =   4800
         TabIndex        =   2
         Tag             =   "0"
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   3255
         Index           =   4
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   19
         Left            =   720
         TabIndex        =   4
         Top             =   4440
         Width           =   5175
      End
      Begin SynthMark_XP.OptionBox OptionBox 
         Height          =   240
         Index           =   16
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   5290
         _ExtentX        =   8467
         _ExtentY        =   423
      End
   End
   Begin SynthMark_XP.ControllerBox ControllerBox 
      Height          =   3735
      Index           =   12
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4683
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   37
         Left            =   2520
         TabIndex        =   216
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.CommandButton CommandButton 
         Height          =   345
         Index           =   38
         Left            =   3840
         TabIndex        =   217
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
      End
      Begin SynthMark_XP.ComplexList FrameList 
         Height          =   1575
         Index           =   5
         Left            =   720
         TabIndex        =   219
         Top             =   1320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2778
      End
      Begin VB.Label FrameLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   600
         TabIndex        =   218
         Top             =   720
         Width           =   4455
      End
   End
   Begin ComctlLib.ImageList PictureLoader 
      Index           =   4
      Left            =   3960
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList PictureLoader 
      Index           =   3
      Left            =   3840
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList PictureLoader 
      Index           =   2
      Left            =   3720
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList PictureLoader 
      Index           =   1
      Left            =   3600
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList PictureLoader 
      Index           =   0
      Left            =   3480
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DragIconCoOrd(1) As Integer

Private Sub BackGround_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub ChatterBox_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub ChatterBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(3)
End Sub

Private Sub CommandButton_Click(Index As Integer)
Call Normalize_Controls(0)
Call Normalize_On_Click(0)
Call Process_CommandButton_Click(Index)
End Sub

Private Sub CommandButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Command_Button(Index)
End Sub

Private Sub ComplexList_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub ComplexList_HeaderClick(Index As Integer, HeaderIndex As Integer)
Call Normalize_On_Click(0)
If BenchResults(Index, 0, 0) <> Empty_Code Then
    Call Sort_Advanced_Array(BenchResults(), Index, HeaderIndex)
    Call Manager.ComplexList(Index).Submit_Data_Array(BenchResults(), Index, 5)
    Call Graph_Update(Index)
End If
End Sub

Private Sub ComplexList_Resized(Index As Integer)
If BenchResults(Index, 0, 0) <> Empty_Code Then
    Call Manager.ComplexList(Index).Submit_Data_Array(BenchResults(), Index, 5)  'If Manager.ComplexList(Index).Visible = True Then
    Call Graph_Update(Index)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'beep
End Sub

Private Sub FrameImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub FrameList_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FrameList_DoubleClick(Index As Integer, Selection As Integer)
Select Case Index
Case 0
    If CommandButton(11).Enabled = True Then Call Process_CommandButton_Click(11)
Case 1
    If Trim(WriteBox(5).Text) = Empty_Code Then
        WriteBox(5).Text = ConnectedList(0, FrameList(Index).ListIndex) & Space_Code
    Else
        If Right(WriteBox(5).Text, 1) = Space_Code Then
            WriteBox(5).Text = WriteBox(5).Text & ConnectedList(0, FrameList(Index).ListIndex)
        Else
            WriteBox(5).Text = WriteBox(5).Text & Space_Code & ConnectedList(0, FrameList(Index).ListIndex)
        End If
    End If
    WriteBox(5).SelectionStart = Len(WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
Case 2
    WriteBox(5).Text = WriteBox(5).Text & Replace(URLBenchList(0, FrameList(Index).ListIndex), Web_Http, Empty_Code)
    WriteBox(5).SelectionStart = Len(WriteBox(5).Text)
    Call SetFocus_Class(3, 5)
Case 4
    Call ATM_Switch_Processes
End Select
End Sub

Private Sub FrameList_HeaderClick(Index As Integer, HeaderIndex As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub LinkLabel_Click(Index As Integer)
Select Case Index
Case 0
    FrameList(0).Height = 2670 '3660
    FrameImage(11).Visible = True
    FrameLabel(55).Visible = True
    LinkLabel(0).Visible = False
End Select
End Sub

Private Sub MenuBox_Click(Index As Integer, ClickIndex As Integer, ClickType As Integer, MenuCmdID As Integer)
Call Process_Menu_Command_Ids(ClickIndex, ClickType, MenuCmdID)
End Sub

Private Sub MenuBox_HideMe(Index As Integer)
'Manager.MenuBox(0).Tag = -1
If Index > 0 Then
    MenuBox(Index).Visible = False
    If MenuBox(Index - 1).Visible = True Then MenuBox(Index - 1).HideFocus
Else
    MenuBox(Index).Visible = False
End If
End Sub

Private Sub MenuBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub PlatInfoCplxList_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub PlatInfoCplxList_HeaderClick(Index As Integer, HeaderIndex As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub PlatInfoCplxList_Resized(Index As Integer)
If PlatformList(0, 0) <> Empty_Code Then Call Manager.PlatInfoCplxList(Index).Submit_Data_Array(PlatformList(), -1, 5)  'If Manager.PlatInfoCplxList(Index).Visible = True Then
End Sub

Private Sub SimpleBox_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub SimpleBox_MakeVisible(Index As Integer, VisibleControl As Boolean)
SimpleBox(Index).Visible = VisibleControl
End Sub

Private Sub SimpleBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub ComplexList_GotFocus(Index As Integer)
If MouseOverMoved <> "2" & Index Then
    Call Manager.VScrollButton(1).Submit_New_Coord(ComplexList(Index).Top - (EightPix * 5))
    MouseOverMoved = Empty_Code
End If
End Sub

Private Sub ComplexList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOverMoved = "2" & Index
Call Normalize_Controls(1)
End Sub

Private Sub ControllerBox_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub ControllerBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = X
DragIconCoOrd(1) = Y
Call Make_Controller_Ontop(Index)
End Sub

Private Sub ControllerBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = -1
End Sub

Private Sub ExplorerButton_Click(Index As Integer, MenuCmdID As Integer)
Call Normalize_On_Click(0)
Call Process_Menu_Command_Ids(Index, 1, MenuCmdID)
End Sub

Private Sub ExplorerButton_MouseMove(Index As Integer, MenuCmdID As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Explorer_Button(Index)
Call Show_ToolTip(MenuCmdID, 3)
End Sub

Private Sub ExplorerFrame_Click(Index As Integer)
Call Normalize_On_Click(0)
Call Align_Explorer_Frames(1)
Call VScrollButton(0).Process_CoOrdinates(0)
End Sub

Private Sub ExplorerFrame_GotFocus(Index As Integer)
If MouseOverMoved <> "3" & Index Then If Manager.VScrollButton(0).Visible = True Then Call Manager.VScrollButton(0).Submit_New_Coord(ExplorerFrame(Index).Top + SixTeenPix)
End Sub

Private Sub ExplorerFrame_SelfAlign(Index As Integer)
For VisualCount = 0 To Manager.ExplorerButton.Count - 1
    If Manager.ExplorerButton(VisualCount).Tag = Index Then
        If Manager.ExplorerFrame(Index).ShowPanel = True Then
            Manager.ExplorerButton(VisualCount).Visible = True
        Else
            Manager.ExplorerButton(VisualCount).Visible = False
        End If
    End If
Next
End Sub

Private Sub ExplorerHolder_Click()
Call Normalize_On_Click(0)
End Sub

Private Sub ExplorerHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub ExplorerHolder_SelfAlign()
If SubDoneOnce = True Then GoTo Ed
SubDoneOnce = True
VScrollButton(0).Left = ExplorerHolder.UseableArea - VScrollButton(0).Width
Call ReProcess_Desktop_CoOrd
Call Align_Explorer_Frames(0)
'SimpleBox(0).Left = Desktop_Left
'SimpleBox(0).Width = Desktop_Width
Call Align_Selected_Complex_Controls
Manager.Refresh
SubDoneOnce = False
Ed: End Sub

Private Sub Form_Click()
Call Normalize_On_Click(0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call Invoke_Keyboard_Shortcuts(KeyCode)
End Sub

Private Sub FormBackGround_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FormCaption_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FormContainer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub FormControl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_FormControl(Index)
End Sub

Private Sub FormIcon_Click()
Call Normalize_On_Click(0)
Call Manager_Display_Menu(-1, 1)
End Sub

Private Sub FormSlider_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FormSlider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = X
DragIconCoOrd(1) = Y
End Sub

Private Sub FormSlider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragIconCoOrd(0) <> -1 And Button <> 0 Then
    On Error Resume Next
    Manager.Move Manager.Left, Manager.Top, Manager.Width - DragIconCoOrd(0) + X, Manager.Height - DragIconCoOrd(1) + Y
    Manager.Refresh
    'DoEvents
End If
End Sub

Private Sub FormSlider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = -1
'NowLoading = False
End Sub

Private Sub FrameButton_Click(Index As Integer)
Call Normalize_Controls(0)
Call Normalize_On_Click(0)
Select Case Index
Case 0
    If WriteBox(5).Text <> Empty_Code Then
        If Net_Single = True Then
            If Net_Is_This_A_Command(WriteBox(5).Text, 0) = False Then
                Call Chat_AddSay(ConnectedUsers(0, 0), WriteBox(5).Text)
            End If
        End If
        If Net_ClientType = True Then
            Select Case Manager.WriteBox(14).ClickTag
            Case Zero_Code
                If Net_Is_This_A_Command(WriteBox(5).Text, 0) = False Then
                    Call Chat_AddSay(ConnectedUsers(0, 0), WriteBox(5).Text)
                End If
                Call Net_Send_Data(NetCode_SayAll & ConnectedUsers(2, 0) & WriteBox(5).Text, -1, -1)
            Case One_Code
                If Net_Is_This_A_Command(WriteBox(5).Text, 0) = True Then
                    Call Net_Send_Data(NetCode_SayAll & ConnectedUsers(2, 0) & WriteBox(5).Text, -1, -1)
                Else
                    If ConnectedList(2, FrameList(1).ListIndex) = ConnectedUsers(2, 0) Then
                        Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & Language(272) & Space_Code & DoubleAsterix_Code)
                    Else
                        Call Net_Send_Data(NetCode_SayTo & ConnectedList(2, FrameList(1).ListIndex) & WriteBox(5).Text, -1, -1)
                    End If
                End If
            End Select
        End If
        If Net_ServerType = True Then
            Select Case Manager.WriteBox(14).ClickTag
            Case Zero_Code
                If Net_Is_This_A_Command(WriteBox(5).Text, 0) = False Then
                    Call Chat_AddSay(ConnectedUsers(0, 0), WriteBox(5).Text)
                    Call Net_Send_Data(NetCode_SayAll & ConnectedUsers(2, 0) & WriteBox(5).Text, -1, -1)
                End If
            Case One_Code
                If Net_Is_This_A_Command(WriteBox(5).Text, 0) = False Then
                    If FrameList(1).ListIndex = 0 Then
                        Call Chat_AddSay(Space_Code, DoubleAsterix_Code & Space_Code & Language(272) & Space_Code & DoubleAsterix_Code)
                    Else
                        Call Net_Send_Data(NetCode_SayTo2 & ConnectedUsers(2, 0) & WriteBox(5).Text, -1, Net_WhosNameIsThat(ConnectedList(2, FrameList(1).ListIndex)))
                    End If
                End If
            End Select
        End If
        WriteBox(5).Text = Empty_Code
    End If
Case 1
    Call Manager.MenuBox(1).Clear_Entire_List
    Call Manager.MenuBox(1).Setup_Menu(0, 190, Language(276), 67, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(1, 191, Language(277), 68, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(2, 192, Language(278), 34, False, 1, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(3, 193, Language(279), 33, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(4, 194, Language(280), 63, False, 0, Can_Menu(), False)
    Call Manager.MenuBox(1).Setup_Menu(5, 195, Language(281), 17, False, 0, Can_Menu(), False)
    Call Manager_Display_Menu(Index, 5)
End Select
End Sub

Private Sub FrameButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Frame_Button(Index)
End Sub

Private Sub FrameImage_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FrameLabel_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub FrameLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub FrameList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(1)
End Sub

Private Sub GraphBox_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub GraphBox_GotFocus(Index As Integer)
If MouseOverMoved <> One_Code & Index Then
    Call Manager.VScrollButton(1).Submit_New_Coord(GraphBox(Index).Top - (EightPix * 5))
    MouseOverMoved = Empty_Code
End If
End Sub

Private Sub GraphBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseOverMoved = One_Code & Index
Call Normalize_Controls(0)
End Sub

Private Sub LightLabel_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub MenuButton_Click(Index As Integer)
Call Normalize_On_Click(0)
Call Highlight_Menu_Button(Index)
Call Manager_Display_Menu(Index, 1)
End Sub

Private Sub MenuButton_KeyPress(Index As Integer, KeyAscii As Integer)
Call Invoke_Keyboard_Shortcuts(KeyAscii)
End Sub

Private Sub MenuButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Tool_Button(-1)
Call Highlight_Menu_Button(Index)
Call Show_ToolTip(0, 0)
Call Normalize_Menu_Lists
End Sub
Private Sub MenuButton_SelfAlign(Index As Integer)
For ControlCount = 1 To Manager.MenuButton.Count - 1
    Manager.MenuButton(ControlCount).Left = Manager.MenuButton(ControlCount - 1).Left + Manager.MenuButton(ControlCount - 1).Width '+ 15
Next ControlCount
End Sub

Private Sub Form_Load()
Call Manager_Tab_Movements
StatusLabel(1).NewCategory = True
StatusLabel(2).NewCategory = True

ExplorerFrame_Count = ExplorerFrame.Count - 1
FormImage.Width = Screen_Width
FormBackGround(0).Width = Screen_Width
FormBackGround(1).Width = Screen_Width
FormBackGround(2).Width = Screen_Width
FormBackGround(3).Width = Screen_Width
'FormBackGround(1).Height = Screen_Height

NormLine(0).X2 = Screen_Width
NormLine(1).X2 = Screen_Width
NormLine(2).X2 = Screen_Width
NormLine(3).X2 = Screen_Width
NormLine(4).X2 = Screen_Width
NormLine(9).X2 = Screen_Width
DragIconCoOrd(0) = -1
Call Make_Controller_Ontop(-1)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Form_Control_Click(4)
End Sub
Private Sub ScrollButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub
Private Sub ControllerBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragIconCoOrd(0) <> -1 And DoneItOnce = False Then
    DoneItOnce = True
    ControllerBox(Index).Move ControllerBox(Index).Left - DragIconCoOrd(0) + X, ControllerBox(Index).Top - DragIconCoOrd(1) + Y
    Manager.Refresh
    DoneItOnce = False
Else
    Call Normalize_Controls(0)
End If
End Sub

Private Sub OptionBox_Click(Index As Integer)
Call Normalize_Controls(0)
Call Normalize_On_Click(0)
Call Process_OptionBox_Click(Index)
End Sub

Private Sub OptionBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Option_Box(Index)
Call Show_ToolTip(0, 0)
End Sub

Private Sub OtherTool_Click(Index As Integer, CommandID As Integer)
Call Normalize_On_Click(0)
Call Process_Menu_Command_Ids(Index, 2, CommandID)
End Sub

Private Sub OtherTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Tool_Button(-1)
Call Normalize_PlatInfoCplxList_Box
Call Highlight_Other_Tool(Index)
End Sub

Private Sub PlatInfoCplxList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(2)
End Sub

'Private Sub RichTextBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call Normalize_Controls(0)
'End Sub

Private Sub ScoreLabel_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub ScoreLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub StatusLabel_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub StatusLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub StylishLabel_Click(Index As Integer)
Call Normalize_On_Click(0)
End Sub

Private Sub StylishLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub

Private Sub ToolButton_Click(Index As Integer, CommandID As Integer)
Call Normalize_On_Click(0)
Call Process_Menu_Command_Ids(Index, 0, CommandID)
End Sub

Private Sub ToolButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Other_Tool(-1)
Call Highlight_Menu_Button(-1)
Call Highlight_Tool_Button(Index)
End Sub
Private Sub ExplorerFrame_HeaderMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Highlight_Explorer_Header(Index)
Call Normalize_Menu_Lists
MouseOverMoved = "exf" & Index
End Sub
Private Sub ExplorerFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
MouseOverMoved = "3" & Index
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub
Private Sub FormBackGround_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
End Sub
Private Sub FormImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragIconCoOrd(0) <> -1 And Manager.WindowState = 0 And DoneItOnce = False Then
    DoneItOnce = True
    Manager.Move Manager.Left - DragIconCoOrd(0) + X, Manager.Top - DragIconCoOrd(1) + Y
    ManagerSub.Move Manager.Left, Manager.Top
    'Manager.Refresh
    DoEvents
    Call Highlight_Menu_Button(-1)
    DoneItOnce = False
Else
    Call Normalize_Controls(0)
End If
End Sub
Private Sub FormImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_On_Click(0)
DragIconCoOrd(0) = X
DragIconCoOrd(1) = Y
End Sub
Private Sub FormImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = -1
End Sub
Private Sub FormCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragIconCoOrd(0) <> -1 And Manager.WindowState = 0 And DoneItOnce = False Then
    DoneItOnce = True
    Manager.Move Manager.Left - DragIconCoOrd(0) + X, Manager.Top - DragIconCoOrd(1) + Y
    ManagerSub.Move Manager.Left, Manager.Top
    'Manager.Refresh
    DoEvents
    Call Highlight_Menu_Button(-1)
    DoneItOnce = False
Else
    Call Normalize_Controls(0)
End If
End Sub
Private Sub FormCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_On_Click(0)
DragIconCoOrd(0) = X
DragIconCoOrd(1) = Y
End Sub
Private Sub FormCaption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
DragIconCoOrd(0) = -1
End Sub

Private Sub ExplorerHolder_Resize(Index As Integer)
End Sub
Private Sub Form_Resize()
If DoneItOnce = True Then GoTo Ed
DoneItOnce = True
On Error Resume Next
FormHeader(1).Move 0, Manager.ScaleHeight - FormHeader(1).Height, Manager.ScaleWidth
Call ReProcess_Desktop_CoOrd

FormSlider(0).Left = Manager.ScaleWidth - FormSlider(0).Width
'FormCaption(0).Caption = WordElipser(SupposedCaption, FormControl(0).Left - FormControl(0).Width - EightPix, Theme_Font, True)
Call Form_Caption_Refresh
VScrollButton(0).Height = Desktop_Height
If ScaleWidth - Manager.OptionBox(14).Width - TwoPix > MenuButton(7).Left + MenuButton(7).Width Then OptionBox(14).Left = Manager.ScaleWidth - OptionBox(14).Width - TwoPix
Call Align_Form_Controls
Call Align_Selected_Complex_Controls
Call Align_Explorer_Frames(0)
Call Manager_WindowState_Check
Call Align_Status_Labels
ManagerSub.Move Manager.Left, Manager.Top, Manager.Width, Manager.Height
DoneItOnce = False
Ed: End Sub
Public Sub Form_Resize_Now()
Call Form_Resize
End Sub
Public Sub Form_Caption_Refresh()
FormCaption(0).Caption = WordElipser(SupposedCaption, FormControl(0).Left - FormControl(0).Width - EightPix, Theme_Font, True)
End Sub
Public Sub Align_Form_Controls()
FormControl(4).Left = Manager.ScaleWidth - FormControl(4).Width - OnePix
FormControl(3).Left = FormControl(4).Left - FormControl(3).Width - OnePix
FormControl(2).Left = FormControl(3).Left - FormControl(2).Width - OnePix
FormControl(1).Left = FormControl(2).Left - FormControl(1).Width - FourPix
FormControl(0).Left = FormControl(1).Left - FormControl(0).Width - FourPix
End Sub

Private Sub FormControl_Click(Index As Integer)
Call Normalize_On_Click(0)
Call Form_Control_Click(Index)
End Sub

Private Sub FormImage_DblClick()
Call Form_Control_Click(3)
End Sub
Private Sub FormCaption_DblClick(Index As Integer)
Call Form_Control_Click(3)
End Sub

Private Sub Timer_Timer(Index As Integer)
Select Case Index
Case 0
    RetVal = APIControls.GetCursorPos(MouseLoc)
    If MouseLoc.X < ((Manager.Left + FormWall) / Screen.TwipsPerPixelX) Then GoTo DoNorm
    If MouseLoc.X > ((Manager.Left + Manager.Width - FormWall) / Screen.TwipsPerPixelX) Then GoTo DoNorm
    If MouseLoc.Y < ((Manager.Top + FormWall) / Screen.TwipsPerPixelY) Then GoTo DoNorm
    If MouseLoc.Y > ((Manager.Top + Manager.Height - FormWall) / Screen.TwipsPerPixelY) Then GoTo DoNorm
    TimedItOnce = False
    
    If SectionSelect = 4 And Manager.MenuBox(0).Visible = False Then
    If MouseLoc.X > ((Manager.Left + SimpleBox(0).Left + FormWall) / Screen.TwipsPerPixelX) Then
        If MouseLoc.X < ((Manager.Left + SimpleBox(0).Left + SimpleBox(0).Width - FormWall) / Screen.TwipsPerPixelX) Then
            If MouseLoc.Y > ((Manager.Top + SimpleBox(0).Top + FormWall) / Screen.TwipsPerPixelY) Then
                If MouseLoc.Y < ((Manager.Top + SimpleBox(0).Top + SimpleBox(0).Height - FormWall) / Screen.TwipsPerPixelY) Then GoTo DoWeb
            End If
        End If
    End If
    Timed1Once = False
    End If
    
Case 1
    Call ATM_Process_System_Tasks
Case 2
    Call Net_Find_Try_Next_IP
Case 3
    Call Net_Check_Kicked
Case 4
    Call Net_End_Vote
Case 5
    Call Check_Whats_OnTop
End Select
GoTo Ed
DoWeb: If Timed1Once = False Then Call Normalize_Controls(1)
Timed1Once = True
GoTo Ed
DoNorm: If TimedItOnce = False Then Call Normalize_Controls(0)
TimedItOnce = True
Ed: End Sub

Private Sub ToolTipBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ToolTipBox.Visible = False
End Sub

Private Sub VScrollButton_SliderDrag(Index As Integer, NowValue As Long)
Call Normalize_On_Click(0)
If Index = 0 And VScrollButton(Index).Visible = True Then
    DefPosHolder = (EightPix) - NowValue
    Call Align_Explorer_Frames(0)
    'Manager.ExplorerHolder.Refresh
End If
If Index = 1 And VScrollButton(Index).Visible = True Then
    Manager.SimpleBox(SectionSelect).Top = Desktop_Top - NowValue
End If
'DoEvents
End Sub

Private Sub WeBrowser_GotFocus()
Call Normalize_On_Click(0)
End Sub

Private Sub WeBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
If InStr(1, LCase(URL), "reg_class") <> 0 Then
    Call Open_Help_System(4)
    WeBrowser.GoBack
End If
If InStr(1, LCase(URL), "ask_class") <> 0 Then
    Call Open_Help_System(3)
    WeBrowser.GoBack
End If
End Sub

Private Sub WriteBox_Changed(Index As Integer)
Select Case Index
Case 5
    If Trim(WriteBox(Index).Text) = Empty_Code Then
        FrameButton(0).Enabled = False
        FrameButton(1).Enabled = True
    Else
        FrameButton(0).Enabled = True
        FrameButton(1).Enabled = False
    End If
Case 8
    If Trim(WriteBox(Index).Text) = Empty_Code Then
        CommandButton(20).Enabled = False
    Else
        CommandButton(20).Enabled = True
    End If
Case 15
    If Trim(WriteBox(15).Text) <> Empty_Code Then
        CommandButton(35).Enabled = True
    Else
        CommandButton(35).Enabled = False
    End If
Case 27, 28, 29, 30
    'On Error Resume Next
    Dim WriteBox_Index As Integer
    WriteBox_Index = WriteBox(Index).SelectionStart
    WriteBox(Index).Text = UCase(WriteBox(Index).Text)
    WriteBox(Index).SelectionStart = WriteBox_Index
    
    If Len(WriteBox(Index).Text) >= 5 Then
        If Len(WriteBox(Index).Text) > 5 Then WriteBox(Index).Text = Left(WriteBox(Index).Text, 5)
        If Index = 30 Then
            Call SetFocus_Class(0, 59) 'CommandButton(4).SetFocus
            GoTo Ed
        End If
        Call SetFocus_Class(3, Index + 1) 'WriteBox(Index + 1).SetFocus
    End If
End Select
Ed: End Sub

Private Sub WriteBox_DropClick(Index As Integer, DDListArray() As String, SelectedIndex As Integer)
Dim ListHieght As Integer, ReadCoOrd As LEFTTOP
WriteBoxList.Clear
For ControlCount = 0 To UBound(DDListArray(), 2)
    WriteBoxList.AddItem DDListArray(0, ControlCount)
Next ControlCount
ReadCoOrd = Read_Alignment_Code(WriteBox(Index).DropDownParent)
WriteBoxList.Move ReadCoOrd.ALeft + WriteBox(Index).Left, ReadCoOrd.ATop + WriteBox(Index).Top + WriteBox(Index).Height, WriteBox(Index).Width

ListHieght = WordHieght(WriteBoxList.Font, WriteBoxList.FontBold) * (UBound(DDListArray(), 2) + 1) + TwoPix
If ListHieght > Manager.ScaleHeight - WriteBoxList.Top Then ListHieght = Manager.ScaleHeight - WriteBoxList.Top

WriteBoxList.Height = ListHieght
WriteBoxList.Tag = Index
WriteBoxList.Visible = True
Call SetFocus_Class(2, 0) 'WriteBoxList.SetFocus
On Error Resume Next
WriteBoxList.ListIndex = SelectedIndex
On Error GoTo 0
WriteBoxList.ZOrder 0
End Sub

Public Sub WriteBox_KeyPressed(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case 0
    If KeyCode = 13 Then
    If InStr(1, LCase(Manager.WriteBox(Index).Text), Web_Http) <> 0 Or InStr(1, LCase(Manager.WriteBox(Index).Text), "www") <> 0 Then
        Call Switch_Sections_To(4)
        Call Manager.WeBrowser.Navigate2(Manager.WriteBox(Index).Text, 2 Or 4 Or 8)
    Else
        Call Process_WriteList_Click(0)
    End If
    End If
Case 1, 3
    If KeyCode = 13 Then Call CommandButton_Click(8)
Case 5
    If KeyCode = 13 Then Call FrameButton_Click(0)
Case 6, 7
    If Trim(WriteBox(6).Text) <> Empty_Code Then
        If KeyCode = 13 Then Call CommandButton_Click(17)
        CommandButton(17).Enabled = True
    Else
        CommandButton(17).Enabled = False
    End If
Case 8
    If KeyCode = 13 Then Call CommandButton_Click(20)
Case 10
    If KeyCode = 13 Then Call CommandButton_Click(22)
Case 15
    If KeyCode = 13 And Manager.CommandButton(35).Enabled = True Then Call CommandButton_Click(35)
Case 13, 26
    If CommandButton(29).Enabled = True Then If KeyCode = 13 Then Call CommandButton_Click(29)
    Call Check_Password_Similarity
Case 22, 23, 24
    If KeyCode = 13 Then Call CommandButton_Click(44)
End Select
End Sub

Private Sub WriteBox_LostFocus(Index As Integer)
Select Case Index
Case 13
    If LCase(Trim(WriteBox(13).Text)) <> LCase(Trim(WriteBox(26).Text)) Then Call SetFocus_Class(3, 26)
Case 22
    If Val(WriteBox(Index).Text) < 50 Then
        WriteBox(Index).Text = "50"
        Call Show_Msg_Window(Language(162) & Space_Code & "50.", Language(163), 1)
        Call SetFocus_Class(3, Index) 'WriteBox(Index).SetFocus
    End If
Case 23
    If Val(WriteBox(Index).Text) < 10 Then
        WriteBox(Index).Text = "10"
        Call Show_Msg_Window(Language(162) & Space_Code & "10.", Language(164), 1)
        Call SetFocus_Class(3, Index) 'WriteBox(Index).SetFocus
    End If
End Select
End Sub

Private Sub WriteBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
'Call Highlight_WriteText_Box(Index)
End Sub

Public Sub Check_Explorer_Frames()
Call ExplorerHolder_SelfAlign
End Sub

Private Sub WriteBoxList_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13, 32
    Call WriteBox(Manager.WriteBoxList.Tag).SwitchToIndex(WriteBoxList.ListIndex)
    Call Process_WriteList_Click(Manager.WriteBoxList.Tag)
    Call SetFocus_Class(3, Manager.WriteBoxList.Tag) 'WriteBox(Manager.WriteBoxList.Tag).SetFocus
Case 27
    Call SetFocus_Class(3, Manager.WriteBoxList.Tag) 'WriteBox(Manager.WriteBoxList.Tag).SetFocus
End Select
End Sub

Private Sub WriteBoxList_LostFocus()
If Manager.WriteBox(Manager.WriteBoxList.Tag).Visible = True Then
    'Manager.WriteBox(Manager.WriteBoxList.Tag).SetFocus
    WriteBoxList.Visible = False
End If
End Sub

Private Sub WriteBoxList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Normalize_Controls(0)
WriteBoxList.ListIndex = Int(Y / Text_Height) + WriteBoxList.TopIndex
End Sub

Private Sub WriteBoxList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call WriteBox(Manager.WriteBoxList.Tag).SwitchToIndex(WriteBoxList.ListIndex)
Call Process_WriteList_Click(Manager.WriteBoxList.Tag)
Call WriteBox_KeyPressed(Manager.WriteBoxList.Tag, 32, 0)
Call SetFocus_Class(3, Manager.WriteBoxList.Tag) 'WriteBox(Manager.WriteBoxList.Tag).SetFocus
End Sub
