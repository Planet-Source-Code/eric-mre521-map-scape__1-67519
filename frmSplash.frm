VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplash.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      MouseIcon       =   "frmSplash.frx":0CD6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer tmrUload 
         Interval        =   5500
         Left            =   1200
         Top             =   240
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   240
         MouseIcon       =   "frmSplash.frx":19A0
         MousePointer    =   99  'Custom
         Picture         =   "frmSplash.frx":266A
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Map Copyright Jagex Ltd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         MouseIcon       =   "frmSplash.frx":3334
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Eric Eveleigh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         MouseIcon       =   "frmSplash.frx":3FFE
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Do not disassemble, decompile, or steal any part of this application without permision. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmSplash.frx":4CC8
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version: 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5490
         MouseIcon       =   "frmSplash.frx":5992
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2700
         Width           =   1365
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform: Win 9x, Me, XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3135
         MouseIcon       =   "frmSplash.frx":665C
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2340
         Width           =   3720
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Mapscape"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2040
         MouseIcon       =   "frmSplash.frx":7326
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1140
         Width           =   3135
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "License: Unregistered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmSplash.frx":7FF0
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2760
         MouseIcon       =   "frmSplash.frx":8CBA
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   720
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Load frmMap
    Load frmHelp
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub tmrUload_Timer()
Unload Me
frmMap.Visible = True
End Sub
