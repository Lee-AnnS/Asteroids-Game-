VERSION 5.00
Begin VB.Form frmStart3 
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16530
   LinkTopic       =   "Form1"
   Picture         =   "frmStart3.frx":0000
   ScaleHeight     =   10005
   ScaleWidth      =   16530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "START GAME"
      Height          =   735
      Left            =   9240
      TabIndex        =   0
      Top             =   6480
      Width           =   1815
   End
End
Attribute VB_Name = "frmStart3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    ' Show instructions
    MsgBox "INSTRUCTIONS:" & vbCrLf & vbCrLf & _
           "* Use arrow keys to move your ship." & vbCrLf & _
           "* Avoid the falling red asteroids!" & vbCrLf & _
           "* Reach the GREEN finish line at the top to level up!" & vbCrLf & _
           "* Each level increases the number and speed of asteroids." & vbCrLf & vbCrLf & _
           "Good luck, pilot!", vbInformation, "How to Play"
    
    
    Me.Hide
    
    frmGame3.Show
    frmGame3.StartGame
End Sub

Private Sub Form_Load()

End Sub
