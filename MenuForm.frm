VERSION 5.00
Begin VB.Form MenuFormX 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
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
   Icon            =   "MenuForm.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "-1"
   Begin VB.PictureBox MenuBoxFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   60
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   15
      Width           =   3015
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   9
         Left            =   0
         TabIndex        =   10
         Tag             =   "0"
         Top             =   2160
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   8
         Left            =   0
         TabIndex        =   9
         Tag             =   "0"
         Top             =   1920
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   7
         Left            =   0
         TabIndex        =   8
         Tag             =   "0"
         Top             =   1680
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1440
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   5
         Left            =   0
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1200
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   4
         Left            =   0
         TabIndex        =   5
         Tag             =   "0"
         Top             =   960
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   4
         Tag             =   "0"
         Top             =   720
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Tag             =   "0"
         Top             =   480
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Tag             =   "0"
         Top             =   0
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
      Begin SynthMark_XP.MenuList MenuList 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Tag             =   "0"
         Top             =   240
         Width           =   3015
         _extentx        =   5318
         _extenty        =   582
      End
   End
   Begin VB.HScrollBar FocusCatcher 
      Height          =   135
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Image FrameImage 
      Height          =   3015
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "MenuFormX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormActiveated As Boolean, LastMenuNo As Integer





