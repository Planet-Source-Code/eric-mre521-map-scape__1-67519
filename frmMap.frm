VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MapScape Unregistered Version:"
   ClientHeight    =   7050
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8625
   ClipControls    =   0   'False
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMap.frx":0CCA
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   Palette         =   "frmMap.frx":1994
   ScaleHeight     =   7050
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgMap 
      Height          =   53760
      Left            =   -32640
      MouseIcon       =   "frmMap.frx":1D99
      MousePointer    =   99  'Custom
      Picture         =   "frmMap.frx":2A63
      Top             =   -23400
      Width           =   59520
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu pmnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim StartX As Long, StartY As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 100 Then
imgMap.Left = imgMap.Left - 250
Else
If KeyAscii = 115 Then
imgMap.Top = imgMap.Top - 250
Else
If KeyAscii = 97 Then
imgMap.Left = imgMap.Left + 250
Else
If KeyAscii = 119 Then
imgMap.Top = imgMap.Top + 250
Else
End If
End If
End If
End If
End Sub



Private Sub Form_Load()
Me.Caption = "MapScape Unregistered V" & " " & App.Major & "." & App.Minor
End Sub

Private Sub Form_Terminate()
Unload frmMap
Unload frmSplash
Unload frmHelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMap
Unload frmSplash
Unload frmHelp
End Sub

Private Sub imgMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' Leaves the object


End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub pmnuHelp_Click()
Show frmHelp
End Sub
