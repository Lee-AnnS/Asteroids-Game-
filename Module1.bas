Attribute VB_Name = "Module1"
Option Explicit

' Asteroid structure
Public Type Asteroid
    X As Single
    Y As Single
    Width As Single
    Height As Single
    Speed As Single
    Active As Boolean
End Type

' Global asteroid array
Public asteroids(1 To 10) As Asteroid

' Constants for asteroids
Public Const ASTEROID_MIN_SIZE As Single = 20
Public Const ASTEROID_MAX_SIZE As Single = 50
Public Const ASTEROID_MIN_SPEED As Single = 2
Public Const ASTEROID_MAX_SPEED As Single = 6

Public Sub InitializeAsteroids()
    Dim i As Integer
    
    ' Initialize all asteroids as inactive
    For i = 1 To UBound(asteroids)
        asteroids(i).Active = False
        asteroids(i).X = 0
        asteroids(i).Y = -100 ' Start above screen
        asteroids(i).Width = ASTEROID_MIN_SIZE + Rnd * (ASTEROID_MAX_SIZE - ASTEROID_MIN_SIZE)
        asteroids(i).Height = asteroids(i).Width ' Square asteroids
        asteroids(i).Speed = ASTEROID_MIN_SPEED + Rnd * (ASTEROID_MAX_SPEED - ASTEROID_MIN_SPEED)
    Next i
    
    ' Activate first few asteroids
    Call SpawnNewAsteroid
    Call SpawnNewAsteroid
End Sub

Public Sub UpdateAsteroids()
    Dim i As Integer
    
    ' Update existing asteroids
    For i = 1 To UBound(asteroids)
        If asteroids(i).Active Then
            ' Move asteroid down
            asteroids(i).Y = asteroids(i).Y + asteroids(i).Speed
            
            ' Check if asteroid has moved off screen
            If asteroids(i).Y > Form1.ScaleHeight + 50 Then
                asteroids(i).Active = False
                ' Spawn a new asteroid to replace it
                Call SpawnNewAsteroid
            End If
        End If
    Next i
    
    ' Randomly spawn new asteroids
    If Rnd < 0.02 Then ' 2% chance each frame
        Call SpawnNewAsteroid
    End If
End Sub

Public Sub SpawnNewAsteroid()
    Dim i As Integer
    Dim foundSlot As Boolean
    
    foundSlot = False
    
    ' Find an inactive asteroid slot
    For i = 1 To UBound(asteroids)
        If Not asteroids(i).Active Then
            ' Activate this asteroid
            asteroids(i).Active = True
            asteroids(i).X = Rnd * (Form1.ScaleWidth - asteroids(i).Width)
            asteroids(i).Y = -asteroids(i).Height
            asteroids(i).Width = ASTEROID_MIN_SIZE + Rnd * (ASTEROID_MAX_SIZE - ASTEROID_MIN_SIZE)
            asteroids(i).Height = asteroids(i).Width
            asteroids(i).Speed = ASTEROID_MIN_SPEED + Rnd * (ASTEROID_MAX_SPEED - ASTEROID_MIN_SPEED)
            foundSlot = True
            Exit For
        End If
    Next i
End Sub

Public Sub DrawAsteroids()
    Dim i As Integer
    
    ' Draw all active asteroids
    For i = 1 To UBound(asteroids)
        If asteroids(i).Active Then
            ' Draw asteroid as a filled circle/ellipse
            Form1.ForeColor = RGB(139, 69, 19) ' Brown color for asteroids
            Form1.FillColor = RGB(160, 82, 45)  ' Lighter brown fill
            Form1.FillStyle = 0 ' Solid fill
            
            ' Draw the asteroid as a circle
            Form1.Circle (asteroids(i).X + asteroids(i).Width / 2, asteroids(i).Y + asteroids(i).Height / 2), asteroids(i).Width / 2
            
            ' Add some detail lines to make it look more like an asteroid
            Form1.ForeColor = RGB(101, 67, 33) ' Darker brown for details
            Form1.Line (asteroids(i).X + 5, asteroids(i).Y + 5)-(asteroids(i).X + asteroids(i).Width - 5, asteroids(i).Y + asteroids(i).Height - 5)
            Form1.Line (asteroids(i).X + asteroids(i).Width - 5, asteroids(i).Y + 5)-(asteroids(i).X + 5, asteroids(i).Y + asteroids(i).Height - 5)
        End If
    Next i
End Sub

Public Function CheckCollisions() As Boolean
    Dim i As Integer
    Dim spaceshipLeft As Single, spaceshipRight As Single
    Dim spaceshipTop As Single, spaceshipBottom As Single
    Dim asteroidLeft As Single, asteroidRight As Single
    Dim asteroidTop As Single, asteroidBottom As Single
    
    ' Get spaceship bounds from Form1
    spaceshipLeft = Form1.spaceshipX
    spaceshipRight = Form1.spaceshipX + Form1.spaceshipWidth
    spaceshipTop = Form1.spaceshipY
    spaceshipBottom = Form1.spaceshipY + Form1.spaceshipHeight
    
    ' Check collision with each active asteroid
    For i = 1 To UBound(asteroids)
        If asteroids(i).Active Then
            asteroidLeft = asteroids(i).X
            asteroidRight = asteroids(i).X + asteroids(i).Width
            asteroidTop = asteroids(i).Y
            asteroidBottom = asteroids(i).Y + asteroids(i).Height
            
            ' Check for overlap (collision)
            If spaceshipRight > asteroidLeft And _
               spaceshipLeft < asteroidRight And _
               spaceshipBottom > asteroidTop And _
               spaceshipTop < asteroidBottom Then
                
                ' Collision detected!
                Form1.GameOver
                CheckCollisions = True
                Exit Function
            End If
        End If
    Next i
    
    CheckCollisions = False
End Function