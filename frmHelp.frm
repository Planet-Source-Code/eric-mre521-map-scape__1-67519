VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "You use keys on the keyboard to move the map instead of the mouse. The keys are:  Left is A, Down is S, Right is D, and Up is W."
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5280
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
