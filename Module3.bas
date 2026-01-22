Attribute VB_Name = "Module3"
Option Explicit

Public Level As Integer
Public Score As Long
Public GameOver As Boolean

Public MoveLeft As Boolean
Public MoveRight As Boolean
Public MoveUp As Boolean
Public MoveDown As Boolean

Public AsteroidX() As Single
Public AsteroidY() As Single
Public AsteroidSize() As Single
Public AsteroidSpeed() As Single
Public NumAsteroids As Integer

Public Sub InitGameVariables()
    Randomize
    Level = 1
    Score = 0
    GameOver = False
    MoveLeft = False
    MoveRight = False
    MoveUp = False
    MoveDown = False
End Sub
