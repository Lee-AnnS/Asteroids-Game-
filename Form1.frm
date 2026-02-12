VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Vertical Asteroids Game"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrGameLoop 
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "Score: 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lblGameOver 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "GAME OVER! Press SPACE to restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00000000&
      Caption         =   "Use LEFT/RIGHT arrows to move spaceship. Avoid falling asteroids!"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   6840
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Game variables
Public gameRunning As Boolean
Public score As Long
Public spaceshipX As Single
Public spaceshipY As Single
Public spaceshipWidth As Single
Public spaceshipHeight As Single

' Constants
Private Const SPACESHIP_SPEED As Single = 8
Private Const SPACESHIP_WIDTH As Single = 60
Private Const SPACESHIP_HEIGHT As Single = 40
Private Const ASTEROID_SPEED As Single = 4
Private Const MAX_ASTEROIDS As Integer = 10

Private Sub Form_Load()
    ' Initialize game
    InitializeGame
End Sub

Private Sub InitializeGame()
    ' Set up the game window
    Me.WindowState = 0 ' Normal
    Me.Width = 9600 * Screen.TwipsPerPixelX
    Me.Height = 7200 * Screen.TwipsPerPixelY
    
    ' Initialize spaceship position (bottom center)
    spaceshipX = Me.ScaleWidth / 2 - SPACESHIP_WIDTH / 2
    spaceshipY = Me.ScaleHeight - SPACESHIP_HEIGHT - 200
    spaceshipWidth = SPACESHIP_WIDTH
    spaceshipHeight = SPACESHIP_HEIGHT
    
    ' Initialize game state
    gameRunning = True
    score = 0
    lblScore.Caption = "Score: " & score
    lblGameOver.Visible = False
    
    ' Initialize asteroids
    Call InitializeAsteroids
    
    ' Start game timer
    tmrGameLoop.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If gameRunning Then
        Select Case KeyCode
            Case vbKeyLeft
                ' Move spaceship left
                spaceshipX = spaceshipX - SPACESHIP_SPEED
                If spaceshipX < 0 Then spaceshipX = 0
                
            Case vbKeyRight
                ' Move spaceship right
                spaceshipX = spaceshipX + SPACESHIP_SPEED
                If spaceshipX > Me.ScaleWidth - spaceshipWidth Then
                    spaceshipX = Me.ScaleWidth - spaceshipWidth
                End If
        End Select
    Else
        ' Game over - check for restart
        If KeyCode = vbKeySpace Then
            RestartGame
        End If
    End If
End Sub

Private Sub tmrGameLoop_Timer()
    If gameRunning Then
        ' Update game logic
        Call UpdateAsteroids
        Call CheckCollisions
        Call UpdateScore
        
        ' Redraw the game
        Call DrawGame
    End If
End Sub

Private Sub DrawGame()
    ' Clear the screen
    Me.Cls
    
    ' Draw spaceship (simple triangle pointing up)
    Me.ForeColor = vbWhite
    Me.Line (spaceshipX, spaceshipY + spaceshipHeight)-(spaceshipX + spaceshipWidth / 2, spaceshipY)
    Me.Line (spaceshipX + spaceshipWidth / 2, spaceshipY)-(spaceshipX + spaceshipWidth, spaceshipY + spaceshipHeight)
    Me.Line (spaceshipX, spaceshipY + spaceshipHeight)-(spaceshipX + spaceshipWidth, spaceshipY + spaceshipHeight)
    
    ' Draw asteroids
    Call DrawAsteroids
End Sub

Private Sub RestartGame()
    ' Reset game state
    gameRunning = True
    score = 0
    lblScore.Caption = "Score: " & score
    lblGameOver.Visible = False
    
    ' Reset spaceship position
    spaceshipX = Me.ScaleWidth / 2 - SPACESHIP_WIDTH / 2
    
    ' Reset asteroids
    Call InitializeAsteroids
    
    ' Restart timer
    tmrGameLoop.Enabled = True
End Sub

Public Sub GameOver()
    gameRunning = False
    tmrGameLoop.Enabled = False
    lblGameOver.Visible = True
End Sub

Private Sub UpdateScore()
    score = score + 1
    lblScore.Caption = "Score: " & score
End Sub