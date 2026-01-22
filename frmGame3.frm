VERSION 5.00
Begin VB.Form frmGame3 
   BackColor       =   &H80000013&
   Caption         =   "Form2"
   ClientHeight    =   10065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17070
   LinkTopic       =   "Form2"
   ScaleHeight     =   10065
   ScaleWidth      =   17070
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   16200
      Top             =   120
   End
   Begin VB.Image imgShip 
      Height          =   960
      Left            =   7440
      Picture         =   "frmGame3.frx":0000
      Top             =   7080
      Width           =   960
   End
End
Attribute VB_Name = "frmGame3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AsteroidX() As Single
Private AsteroidY() As Single
Private AsteroidSize() As Single
Private AsteroidSpeed() As Single
Private NumAsteroids As Integer
Private justLeveledUp As Boolean

' For double buffering to prevent flickering
Private AutoRedrawBackup As Boolean

Private Sub Form_Load()
    Me.KeyPreview = True
    Me.ScaleMode = vbPixels
    Me.BackColor = vbBlack
    
    ' Enable double buffering to prevent flickering
    AutoRedrawBackup = Me.AutoRedraw
    Me.AutoRedraw = True
    
    imgShip.Visible = True
    
    InitGameVariables
    
    imgShip.Left = (Me.ScaleWidth - imgShip.Width) / 2
    imgShip.Top = Me.ScaleHeight - imgShip.Height - 100
    
    InitLevel
    
    DrawFullFrame
    
    ' Update form caption
    Me.Caption = "Asteroid Dodge - Level " & Level
End Sub

Public Sub StartGame()
    Timer1.Enabled = True
    DrawFullFrame
End Sub

Private Sub InitLevel()
    ' Number of asteroids increases with level
    NumAsteroids = 8 + (Level * 3) ' Start with 8, add 3 per level
    If NumAsteroids > 40 Then NumAsteroids = 40 ' Maximum limit
    
    ReDim AsteroidX(1 To NumAsteroids)
    ReDim AsteroidY(1 To NumAsteroids)
    ReDim AsteroidSize(1 To NumAsteroids)
    ReDim AsteroidSpeed(1 To NumAsteroids)
    
    ' Create new asteroids
    Dim i As Integer, j As Integer
    Dim overlap As Boolean
    Dim attempts As Integer
    
    For i = 1 To NumAsteroids
        attempts = 0
        Do
            overlap = False
            AsteroidX(i) = Rnd * (Me.ScaleWidth - 150)
            AsteroidY(i) = -Rnd * 800 - 200 ' Start well above screen
            
            ' Create varied asteroid sizes
            Dim sizeType As Integer
            sizeType = Int(Rnd * 10) + 1
            
            Select Case sizeType
                Case 1 To 3 ' 30% small asteroids - FASTER
                    AsteroidSize(i) = 30 + (Rnd * 20)  ' 30-50 pixels
                    AsteroidSpeed(i) = 6 + Rnd * 3 + (Level * 1) ' Faster
                Case 4 To 7 ' 40% medium asteroids
                    AsteroidSize(i) = 50 + (Rnd * 30)  ' 50-80 pixels
                    AsteroidSpeed(i) = 4 + Rnd * 2 + (Level * 0.8)
                Case 8 To 10 ' 30% large asteroids - SLOWER BUT BIGGER
                    AsteroidSize(i) = 80 + (Rnd * 40)  ' 80-120 pixels
                    AsteroidSpeed(i) = 3 + Rnd * 2 + (Level * 0.6)
            End Select
            
            ' Ensure minimum size and speed
            If AsteroidSize(i) < 25 Then AsteroidSize(i) = 25
            If AsteroidSpeed(i) < 3 Then AsteroidSpeed(i) = 3
            
            ' Check for overlap with existing asteroids using circular collision
            For j = 1 To i - 1
                If CheckAsteroidToAsteroidCollision(i, j) Then
                    overlap = True
                    Exit For
                End If
            Next j
            
            attempts = attempts + 1
            If attempts > 50 Then Exit Do
        Loop While overlap
    Next i
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer, j As Integer
    
    If justLeveledUp Then
        justLeveledUp = False
        Exit Sub
    End If
    
    ' Ship movement
    If MoveLeft And imgShip.Left > 0 Then imgShip.Left = imgShip.Left - 10
    If MoveRight And imgShip.Left < (Me.ScaleWidth - imgShip.Width) Then imgShip.Left = imgShip.Left + 10
    If MoveUp And imgShip.Top > 0 Then imgShip.Top = imgShip.Top - 10
    If MoveDown And imgShip.Top < (Me.ScaleHeight - imgShip.Height) Then imgShip.Top = imgShip.Top + 10
    
    ' Move asteroids and check for collisions
    For i = 1 To NumAsteroids
        ' Store original position
        Dim originalY As Single, originalX As Single
        originalY = AsteroidY(i)
        originalX = AsteroidX(i)
        
        ' Move asteroid
        AsteroidY(i) = AsteroidY(i) + AsteroidSpeed(i)
        
        ' Check for asteroid-asteroid collisions after movement
        Dim collisionOccurred As Boolean
        collisionOccurred = False
        
        For j = 1 To NumAsteroids
            If i <> j Then
                If CheckAsteroidToAsteroidCollision(i, j) Then
                    AsteroidY(i) = originalY
                    AsteroidX(i) = originalX
                    
                    Dim newX As Single
                    If AsteroidX(i) < AsteroidX(j) Then
                        newX = AsteroidX(i) - 10
                    Else
                        newX = AsteroidX(i) + 10
                    End If
                    
                    If newX >= 0 And newX <= Me.ScaleWidth - AsteroidSize(i) Then
                        AsteroidX(i) = newX
                    End If
                    
                    collisionOccurred = True
                    Exit For
                End If
            End If
        Next j
        
        If Not collisionOccurred And AsteroidY(i) > Me.ScaleHeight Then
        
            Dim safePositionFound As Boolean
            Dim respawnAttempts As Integer
            
            safePositionFound = False
            respawnAttempts = 0
            
            Do While Not safePositionFound And respawnAttempts < 20
                safePositionFound = True
                AsteroidY(i) = -AsteroidSize(i) - 200
                AsteroidX(i) = Rnd * (Me.ScaleWidth - AsteroidSize(i))
                

                For j = 1 To NumAsteroids
                    If i <> j Then
                        If CheckAsteroidToAsteroidCollision(i, j) Then
                            safePositionFound = False
                            Exit For
                        End If
                    End If
                Next j
                
                respawnAttempts = respawnAttempts + 1
            Loop
            
            ' Update speed for new level
            AsteroidSpeed(i) = 3 + Rnd * 3 + (Level * 0.8)
        End If
        
               ' Check for ship collision using circular collision
        If CheckShipToAsteroidCollision(i) Then
            Timer1.Enabled = False
            GameOver = True
            
            ' Show custom game over message with replay and quit buttons
            Dim response As Integer
            response = MsgBox("Game Over!" & vbCrLf & "Final Level: " & Level & vbCrLf & "Final Score: " & Score & vbCrLf & vbCrLf & "Play Again?", _
                             vbQuestion + vbYesNo, "Game Over")
            
            If response = vbYes Then
                ' Replay
                Level = 1
                Score = 0
                GameOver = False
                
                ' Reset ship position
                imgShip.Left = (Me.ScaleWidth - imgShip.Width) / 2
                imgShip.Top = Me.ScaleHeight - imgShip.Height - 100
                
            
                MoveUp = False
                MoveDown = False
                MoveLeft = False
                MoveRight = False
                
                ' Create new asteroids
                InitLevel
                
                ' Update form caption
                Me.Caption = "Asteroid Dodge - Level " & Level
                
                ' Redraw and restart timer
                DrawFullFrame
                Timer1.Enabled = True
            Else
                ' Quit
                Me.AutoRedraw = AutoRedrawBackup ' Restore AutoRedraw setting
                frmStart3.Show
                Unload Me
            End If
            Exit Sub
        End If
    Next i
    
    ' Check finish line
    If imgShip.Top <= 50 Then
        Level = Level + 1
        Score = Score + (Level * 100)
        MsgBox "Level " & Level & " completed!" & vbCrLf & "Score: +" & (Level * 100) & vbCrLf & "Get ready for more asteroids!", vbInformation, "Level Up!"
        
        ' Reset ship position
        imgShip.Left = (Me.ScaleWidth - imgShip.Width) / 2
        imgShip.Top = Me.ScaleHeight - imgShip.Height - 100
    
        justLeveledUp = True
        
        MoveUp = False
        MoveDown = False
        MoveLeft = False
        MoveRight = False
        
        ' Create new asteroids
        InitLevel
        
        ' Update form caption
        Me.Caption = "Asteroid Dodge - Level " & Level
    End If
    
    ' Draw everything
    DrawFullFrame
End Sub

Private Function CheckAsteroidToAsteroidCollision(ast1 As Integer, ast2 As Integer) As Boolean
    ' Asteroid 1 center and radius
    Dim ast1CenterX As Single, ast1CenterY As Single, ast1Radius As Single
    ast1CenterX = AsteroidX(ast1) + AsteroidSize(ast1) / 2
    ast1CenterY = AsteroidY(ast1) + AsteroidSize(ast1) / 2
    ast1Radius = AsteroidSize(ast1) / 2
    
    ' Asteroid 2 center and radius
    Dim ast2CenterX As Single, ast2CenterY As Single, ast2Radius As Single
    ast2CenterX = AsteroidX(ast2) + AsteroidSize(ast2) / 2
    ast2CenterY = AsteroidY(ast2) + AsteroidSize(ast2) / 2
    ast2Radius = AsteroidSize(ast2) / 2
    
    ' Calculate distance between centers
    Dim dx As Single, dy As Single, distance As Single
    dx = ast1CenterX - ast2CenterX
    dy = ast1CenterY - ast2CenterY
    distance = Sqr(dx * dx + dy * dy)
    
    If distance < (ast1Radius + ast2Radius - 2) Then ' -2 pixel buffer to prevent sticking
        CheckAsteroidToAsteroidCollision = True
    Else
        CheckAsteroidToAsteroidCollision = False
    End If
End Function

Private Function CheckShipToAsteroidCollision(asteroidIndex As Integer) As Boolean
    ' Asteroid center and radius
    Dim astCenterX As Single, astCenterY As Single, astRadius As Single
    astCenterX = AsteroidX(asteroidIndex) + AsteroidSize(asteroidIndex) / 2
    astCenterY = AsteroidY(asteroidIndex) + AsteroidSize(asteroidIndex) / 2
    astRadius = AsteroidSize(asteroidIndex) / 2
    
    Dim shipCenterX As Single, shipCenterY As Single, shipRadius As Single
    shipCenterX = imgShip.Left + imgShip.Width / 2
    shipCenterY = imgShip.Top + imgShip.Height / 2
    
    shipRadius = IIf(imgShip.Width < imgShip.Height, imgShip.Width / 2, imgShip.Height / 2)
    shipRadius = shipRadius * 0.8 '
    
    ' Calculate distance between centers
    Dim dx As Single, dy As Single, distance As Single
    dx = astCenterX - shipCenterX
    dy = astCenterY - shipCenterY
    distance = Sqr(dx * dx + dy * dy)
    
    ' Check if distance is less than sum of radii
    If distance < (astRadius + shipRadius) Then
        CheckShipToAsteroidCollision = True
    Else
        CheckShipToAsteroidCollision = False
    End If
End Function

Private Sub DrawFullFrame()
    Dim i As Integer
    
    Me.Cls
    
    ' Draw finish line (green)
    Me.Line (0, 50)-(Me.ScaleWidth, 50), vbGreen
    Me.Line (0, 51)-(Me.ScaleWidth, 51), vbGreen
    Me.Line (0, 52)-(Me.ScaleWidth, 52), vbGreen
    
    ' Draw finish line text
    Me.CurrentX = Me.ScaleWidth / 2 - 50
    Me.CurrentY = 20
    Me.ForeColor = vbGreen
    Me.Print "FINISH LINE"
    
    ' Draw asteroids
    Me.FillColor = vbRed
    Me.FillStyle = vbFSSolid
    For i = 1 To NumAsteroids
        ' Only draw if asteroid is visible on screen
        If AsteroidY(i) + AsteroidSize(i) > -100 And AsteroidY(i) < Me.ScaleHeight + 100 Then
            Me.Circle (AsteroidX(i) + AsteroidSize(i) / 2, AsteroidY(i) + AsteroidSize(i) / 2), AsteroidSize(i) / 2, vbRed
        End If
    Next i
    
    ' Draw score and level
    Me.CurrentX = 20
    Me.CurrentY = 20
    Me.ForeColor = vbYellow
    Me.Print "Level: " & Level & "  Score: " & Score
    
    ' Draw instructions
    Me.CurrentX = Me.ScaleWidth - 300
    Me.CurrentY = Me.ScaleHeight - 80
    Me.ForeColor = vbWhite
    Me.Print "Arrow Keys: Move Ship"
    Me.CurrentX = Me.ScaleWidth - 300
    Me.CurrentY = Me.ScaleHeight - 60
    Me.Print "ESC: Quit Game"
    Me.CurrentX = Me.ScaleWidth - 300
    Me.CurrentY = Me.ScaleHeight - 40
    Me.Print "Reach the green line!"
    
    Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft: MoveLeft = True
        Case vbKeyRight: MoveRight = True
        Case vbKeyUp: MoveUp = True
        Case vbKeyDown: MoveDown = True
        Case vbKeyEscape
            Timer1.Enabled = False
            If MsgBox("Quit game?", vbYesNo + vbQuestion, "Quit") = vbYes Then
                Me.AutoRedraw = AutoRedrawBackup ' Restore AutoRedraw setting
                frmStart3.Show
                Unload Me
            Else
                Timer1.Enabled = True
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft: MoveLeft = False
        Case vbKeyRight: MoveRight = False
        Case vbKeyUp: MoveUp = False
        Case vbKeyDown: MoveDown = False
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    GameOver = True
    Me.AutoRedraw = AutoRedrawBackup ' Restore AutoRedraw setting
End Sub

Private Sub Form_Paint()
    DrawFullFrame
End Sub
