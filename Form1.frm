VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   15000
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   9375
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9315
      ScaleWidth      =   14835
      TabIndex        =   0
      Top             =   0
      Width           =   14895
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   7080
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   7080
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
FirstdrawFlag = 1
ScaleFactor(0) = 1
ScaleFactor(1) = 1
2 ShapeStart = 10
MaxObjects = 1000
Detail = 200
VelocityMod = 0.5

TimeKeeper(0) = GetTickCount
TimeKeeper(3) = GetTickCount

Picture1.Top = 0
Picture1.Left = 0
Picture1.Width = Screen.Width
Picture1.Height = Screen.Height - 850
Picture1.BackColor = 0
Picture1.ForeColor = RGB(128, 128, 128)

'dimention 1 = object number
'dimention 2 = position and details on how to draw the object
'              val0 = does object exist? 1= yes 0 =no
'              val1 = xcoord'
'              val2 = ycoord
'              val3 = orientation
'              val4 = velocity

'              val5 = direction of movement
               ' val6 = Mass
               'val7 = rate of uncontrolled rotation
               'val8 = detect collision?
               
'              val10-20 = xcoord eg 5 = degrees, 6 = distance from centre



ReDim SOLast(MaxObjects, 1)
ReDim CollideRecord(MaxObjects), SpaceObject(MaxObjects, Detail), LSpaceObject(MaxObjects, Detail)
ReDim MDC(MaxObjects), LastCoords(MaxObjects, Detail, 1), LastCoords2(MaxObjects, Detail, 1)



'Space Ship 1

SpaceObject(0, 0) = 1                ' This object exists
SpaceObject(0, 1) = Picture1.Width / 0.5
SpaceObject(0, 2) = Picture1.Height / 0.5
SpaceObject(0, 6) = 1000              ' Mass = 1 ton
SpaceObject(0, 8) = 1                 ' Detect collisions between this object and all others
SpaceObject(0, 9) = 2                 ' Except this object

SpaceObject(0, 10) = 0
SpaceObject(0, 11) = 100

SpaceObject(0, 12) = 0
SpaceObject(0, 13) = 500

SpaceObject(0, 14) = 45
SpaceObject(0, 15) = 500

SpaceObject(0, 16) = 90
SpaceObject(0, 17) = 500

SpaceObject(0, 18) = 135
SpaceObject(0, 19) = 500

SpaceObject(0, 20) = 180
SpaceObject(0, 21) = 500

SpaceObject(0, 22) = 225
SpaceObject(0, 23) = 500

SpaceObject(0, 24) = 270
SpaceObject(0, 25) = 500

SpaceObject(0, 26) = 315
SpaceObject(0, 27) = 500

SpaceObject(0, 28) = 360
SpaceObject(0, 29) = 500

SpaceObject(0, 30) = -1



'Player 2 ship
SpaceObject(3, 0) = 1 'this object exists
SpaceObject(3, 1) = Picture1.Width / 4
SpaceObject(3, 2) = Picture1.Height / 4
SpaceObject(3, 6) = 10008 'Mass = 1 ton
SpaceObject(3, 8) = 1 'detect collisions between this object and all others
SpaceObject(3, 9) = 4 'except this object

SpaceObject(3, 10) = 0
SpaceObject(3, 11) = 100

SpaceObject(3, 12) = 0
SpaceObject(3, 13) = 500

SpaceObject(3, 14) = 45
SpaceObject(3, 15) = 500

SpaceObject(3, 16) = 90
SpaceObject(3, 17) = 500
8
SpaceObject(3, 18) = 135
SpaceObject(3, 19) = 500

SpaceObject(3, 20) = 180
SpaceObject(3, 21) = 500

SpaceObject(3, 22) = 225
SpaceObject(3, 23) = 500

SpaceObject(3, 24) = 270
SpaceObject(3, 25) = 500

SpaceObject(3, 26) = 315
SpaceObject(3, 27) = 500

SpaceObject(3, 28) = 360
SpaceObject(3, 29) = 500

SpaceObject(3, 30) = -1

'dimention 1 = object number
'dimention 2 = position and details on how to draw the object
'              val0 = does object exist? 1= yes 0 =no
'              val1 = xcoord'
'              val2 = ycoord
'              val3 = orientation
'              val4 = velocity

'              val5 = direction of movement
               ' val6 = Mass
               'val7 = rate of uncontrolled rotation
               'val8 = detect collision?
               
'              val10-20 = xcoord eg 5 = degrees, 6 = distance from centre

'Asteroids
Rnd (3)
Dim BasicSize
For Z = 5 To 20
    BasicSize = 50 + CLng(1000 * Rnd)
    SpaceObject(Z, 0) = 1 'this object exists
    SpaceObject(Z, 1) = CLng(Rnd * Picture1.Width)
    SpaceObject(Z, 2) = CLng(Rnd * Picture1.Width)
    SpaceObject(Z, 4) = 100 'CLng(Rnd * 50)
    SpaceObject(Z, 5) = CLng(Rnd * 350)
    SpaceObject(Z, 6) = 100 * BasicSize 'Mass = 1 ton
    SpaceObject(Z, 7) = 0 '(Int((5 * Rnd) + 1) * (1050 / BasicSize)) / 10 'rotation speed
    If Int((6 * Rnd) + 1) >= 3 Then SpaceObject(Z, 7) = SpaceObject(Z, 7) * -1
    SpaceObject(Z, 8) = 1 'detect collisions between this object and all others
    SpaceObject(Z, 9) = -1 'except this object
    X = 10
    
    For Y = 0 To 340 Step 20
        SpaceObject(Z, X) = Y 'the angle
        SpaceObject(Z, X + 1) = BasicSize + Int((100 * Rnd) + 1)
        X = X + 2
        
        
    Next Y
    
    SpaceObject(Z, X) = -1
Next Z



If X = 12345 Then
    SpaceObject(3, 1) = SpaceObject(0, 1) + 1000
    SpaceObject(3, 2) = SpaceObject(0, 2) + 0
    
    SpaceObject(0, 2) = 1000
    SpaceObject(0, 3) = 0
    SpaceObject(0, 4) = 20
    SpaceObject(3, 3) = 90
    
    SpaceObject(5, 1) = 14900
    SpaceObject(5, 2) = 0
    SpaceObject(5, 4) = 0
    SpaceObject(5, 5) = 90
    SpaceObject(6, 1) = 10000
    SpaceObject(6, 2) = 5000
    SpaceObject(6, 4) = 0
    SpaceObject(6, 5) = 270
End If


SpaceObject(1, 0) = 0
SpaceObject(1, 1) = SpaceObject(0, 1)
SpaceObject(1, 0) = SpaceObject(0, 2) + 50
ReDim MaxProx(MaxObjects), MinProx(MaxObjects)
For X = 0 To MaxObjects
    MinProx(X) = 1000000
Next X
For X = 0 To MaxObjects
    If SpaceObject(X, 0) = 1 Then
        MDC(X) = 0
        
        For Y = 10 To Detail - 1 Step 2
            If SpaceObject(X, Y + 1) > 0 Then
                If SpaceObject(X, Y + 1) > MDC(X) Then MDC(X) = SpaceObject(X, Y + 1) + 1
                If SpaceObject(X, Y + 1) > MaxProx(X) Then MaxProx(X) = SpaceObject(X, Y + 1) + 1
                If SpaceObject(X, Y + 1) < MinProx(X) Then
                    MinProx(X) = SpaceObject(X, Y + 1) + 1
                End If
            Else
                Exit For
            End If
            'If SpaceObject(X, Y) = -1 Then Exit For
            
            
        Next Y
    End If
Next X

ReDim MaxProx2D(MaxObjects, MaxObjects), MinProx2D(MaxObjects, MaxObjects)
For X = 0 To MaxObjects
    For Y = 0 To MaxObjects
        MaxProx2D(X, Y) = MaxProx(X) + MaxProx(Y)
        MinProx2D(X, Y) = MinProx(X) + MinProx(Y)
    Next Y
Next X

End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
KP(KeyCode) = 1

Exit Sub
'If KeyCode = 37 Then KP = -1
'If KeyCode = 39 Then KP = 1
'If KeyCode = 40 Then KP = 10
'If KeyCode = 17 Then KP = 20


'If KeyCode = 65 Then KP = -2
'If KeyCode = 68 Then KP = 2
'If KeyCode = 83 Then KP = 11
'If KeyCode = 32 Then KP = 21


End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 44 Then KP = -1
'If KeyAscii = 46 Then KP = 1
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
KP(KeyCode) = 0 = 0
ExitE = 1
End Sub

Private Sub Timer1_Timer()

Dim Direction As Double, X As Long, Y As Long, Degrees As Long, DegreesP As Long, XCentre As Long, YCentre As Long, Velocity As Double, TD As Double, S As Double
Dim xoff As Double, yoff As Double, XOff2 As Double, YOff2 As Double, XOff3 As Double, YOff3 As Double, XOff4 As Double, YOff4 As Double, NewDirection As Double





If KP(37) = 1 Then
    SpaceObject(0, 7) = 0
    SpaceObject(0, 3) = SpaceObject(0, 3) - 5
End If
If KP(39) = 1 Then
    SpaceObject(0, 7) = 0
    SpaceObject(0, 3) = SpaceObject(0, 3) + 5
End If
If SpaceObject(0, 3) < 0 Then SpaceObject(0, 3) = SpaceObject(0, 3) + 360
If SpaceObject(0, 3) > 359 Then SpaceObject(0, 3) = SpaceObject(0, 3) - 360

If KP(65) = 1 Then
    SpaceObject(3, 3) = SpaceObject(3, 3) - 5
    SpaceObject(3, 7) = 0
End If
If KP(68) = 1 Then
    SpaceObject(3, 3) = SpaceObject(3, 3) + 5
    SpaceObject(3, 7) = 0
End If
    

Call UpdateRotation

If SpaceObject(3, 3) < 0 Then SpaceObject(3, 3) = SpaceObject(3, 3) + 360
If SpaceObject(3, 3) > 359 Then SpaceObject(3, 3) = SpaceObject(3, 3) - 360

For X = 0 To MaxObjects
    If SpaceObject(X, 0) = 1 Then
        For Y = 0 To Detail
            LSpaceObject(X, Y) = SpaceObject(X, Y)
            If LSpaceObject(X, Y) = -1 Then Exit For
        Next Y
    End If
Next X

If KP(40) = 1 Then
    Call DoThrust(0, SpaceObject(0, 3), 1)
    'do the thruster
    Call DoThruster(0)
Else
    SpaceObject(1, 0) = 0
End If

If KP(83) = 1 Then
    Call DoThrust(3, SpaceObject(3, 3), 1)
    'do the thruster
    Call DoThruster(3)
Else
    SpaceObject(4, 0) = 0
End If

If KP(17) = 1 Then
    Call Shoot(0)
    'KP(17) = 0
End If
If KP(71) = 1 Then
    
    Call Shoot(3)
End If




Dim ModX(2) As Long, ModY(2) As Long
ModX(1) = -Picture1.Width
ModX(2) = Picture1.Width

ModY(1) = -Picture1.Height
ModY(2) = Picture1.Height

Picture1.AutoRedraw = True
Picture1.Cls
Dim DoTwice() As Byte
ReDim DoTwice(MaxObjects, 1)




For X = 0 To MaxObjects
    
    If SpaceObject(X, 0) = 1 Then
        'For Y = 0 To Detail
            If Abs(SpaceObject(X, 1)) < 200000 Then
                XCentre = CLng(SpaceObject(X, 1))
                YCentre = CLng(SpaceObject(X, 2))
            Else
                XCentre = -20000
                yxentre = -20000
            End If
            SOLast(X, 0) = XCentre
            SOLast(X, 1) = YCentre
            Degrees = (SpaceObject(X, 3))
            Velocity = (SpaceObject(X, 4)) * VelocityMod
            Direction = (SpaceObject(X, 5))
            If Direction = 360 Then Direction = 0
            'first work out change in centre based on velocity
            If Velocity > 0 Then
                
                If Direction > 0 And Direction < 90 Then
                    TD = 90 - Direction
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = -S * Velocity
                    S = Cos(TD)
                    xoff = S * Velocity
                ElseIf Direction = 0 Then
                    yoff = -Velocity
                    xoff = 0
                ElseIf Direction = 90 Then
                    yoff = 0
                    xoff = Velocity
                ElseIf Direction > 90 And Direction < 180 Then
                    TD = Direction - 90
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = S * Velocity
                    S = Cos(TD)
                    xoff = S * Velocity
                    X = X
                ElseIf Direction = 180 Then
                    yoff = Velocity
                    xoff = 0
                ElseIf Direction > 180 And Direction < 270 Then
                    TD = Direction - 180
                    TD = TD * (Pi / 180)
                    S = Cos(TD)
                    yoff = S * Velocity
                    S = Sin(TD)
                    xoff = -S * Velocity
                ElseIf Direction = 270 Then
                    yoff = 0
                    xoff = -Velocity
                ElseIf Direction > 270 And Direction < 360 Then
                    TD = Direction - 270
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = -S * Velocity
                    S = Cos(TD)
                    xoff = -S * Velocity
                    X = X
                End If
                XCentre = XCentre + xoff '7394,4604:7394,4604
                YCentre = YCentre + yoff
                
                If XCentre + MaxProx(X) * 2 > Picture1.Width Then
                    DoTwice(X, 0) = 1
                End If
                If XCentre > Picture1.Width Then
                    XCentre = XCentre - Picture1.Width
                End If
                
                If XCentre < MaxProx(X) * 2 Then
                    DoTwice(X, 0) = 2
                End If
                If XCentre < 0 Then
                    XCentre = Picture1.Width + XCentre
                End If
                If YCentre + MaxProx(X) * 2 > Picture1.Height Then
                    DoTwice(X, 1) = 1
                End If
                If YCentre > Picture1.Height Then
                    YCentre = YCentre - Picture1.Height
                End If
                
                If YCentre < MaxProx(X) * 2 Then
                    DoTwice(X, 1) = 2
                End If
                If YCentre < 0 Then
                    YCentre = Picture1.Height + YCentre
                End If
                
                SpaceObject(X, 1) = XCentre
                SpaceObject(X, 2) = YCentre
            End If
            
            
        'Next Y
    End If
Next X

'Detect collisions and modify positions as necessary
'Call DetectCollide

ReDim DrawObject(MaxObjects, Detail), FinalPoint(MaxObjects)

Call MakeCoords(SpaceObject(), DrawObject(), FinalPoint(), 0, MaxObjects)

For X = 0 To MaxObjects
    For Y = 0 To FinalPoint(X) Step 2
        LastCoords2(X, Y, 0) = LastCoords(X, Y, 0)
        LastCoords2(X, Y + 1, 0) = LastCoords(X, Y + 1, 0)
        
    Next Y
Next X

For X = 0 To MaxObjects
    For Y = 0 To FinalPoint(X) Step 2
        LastCoords(X, Y, 0) = DrawObject(X, Y)
        LastCoords(X, Y + 1, 0) = DrawObject(X, Y + 1)
        
    Next Y
Next X

For X = 0 To MaxObjects
    For Y = 0 To FinalPoint(X) Step 2
        LastCoords(X, Y, 0) = DrawObject(X, Y)
        LastCoords(X, Y + 1, 0) = DrawObject(X, Y + 1)
        
    Next Y
Next X
'this is where detectcollide used to be

Dim XMod As Long, YMod As Long

xx = SpaceObject(6, 5) = 0
oc = frmGame.Picture1.ForeColor
Dim ColAdj As Double, LCX1 As Double, LCX2 As Double, LCY1 As Double, LCY2 As Double, LCA1 As Double, LCA2 As Double, LCB1 As Double, LCB2 As Double
Dim GhostNumX As Double, GhostNumY As Double, GhostNum As Long, ColMapUp As Long
If FirstdrawFlag = 0 Then
    Call DetectCollide4(FinalPoint())
End If

'Sleep 100
For X = 0 To MaxObjects
    If SpaceObject(X, 0) > 0 Then
        ' If X >= 50 Then
             'Form1.Picture1.ForeColor = RGB(64, 64, 64)
        ' End If
       
        frmGame.Picture1.ForeColor = RGB(255, 255, 255)
        XMod = ModX(DoTwice(X, 0))
        YMod = ModY(DoTwice(X, 1))
        
        If SpaceObject(X, 6) = 1 Then
           
            LCX2 = CLng(SpaceObject(X, 1))
            LCX1 = CLng(SOLast(X, 0))
            LCY2 = CLng(SpaceObject(X, 2))
            LCY1 = CLng(SOLast(X, 1))
            
            If Abs(LCX1 - LCX2) > 10000 Then
                LCX2 = LCX2 - Picture1.Width
            End If
            
            If Abs(LCY1 - LCY2) > 10000 Then
                LCY2 = LCY2 - Picture1.Height
            End If
            
           ' Form1.Picture1.Circle (LCX2, LCY2), 30, RGB(255, 0, 0)
           ' Form1.Picture1.Circle (LCX1, LCY1), 30, RGB(255, 255, 0)
          '  Form1.Picture1.Circle (LastCoords(X, 0, 0), LastCoords(X, 1, 0)), 30, RGB(0, 255, 0)
            xx = LastCoords(X, 0, 0)
            xx = LastCoords(X, 1, 0)
            Dim GapPieces As Integer, ColMap As Long
            GapPieces = 6
            If (Abs(LCX1 - LCX2) >= Screen.TwipsPerPixelX * 2) Or (Abs(LCY1 - LCY2) > Screen.TwipsPerPixelY * 2) Then
                'Form1.Picture1.Circle (LCX2, LCY2), 30, RGB(255, 0, 0)
                'work out how many ghosts need to be printed
2                'If Abs(LCX1 - LCX2) > Abs(LCY1 - LCY2) Then
                GhostNumX = LCX1 - LCX2
                
                'Else
                GhostNumY = LCY1 - LCY2
                If Abs(GhostNumX) < 5000 And Abs(GhostNumY) < 5000 Then
                    ' End If
                     If Abs(GhostNumY) > Abs(GhostNumX) Then
                         GapPieces = CLng(Abs((GhostNumY / Screen.TwipsPerPixelY) / 2))
                         GhostNum = Abs(CLng(GhostNumY) / GapPieces)
                         
                     Else
                         GapPieces = CLng(Abs((GhostNumX / Screen.TwipsPerPixelX) / 2))
                         GhostNum = Abs(CLng(GhostNumX) / GapPieces)
                     End If
                     GhostNumX = GhostNumX / GapPieces ' CLng(Ghostnum / 12)
                     GhostNumY = GhostNumY / GapPieces 'GhostNum 'CLng(Ghostnum / 12)
                     'LCY1 = LCY2
                     'LCX1 = LCX2
                     ColMapUp = CInt(255 / GapPieces)
                     ColMap = 255
                     For a = 0 To GapPieces
                        ColMap = ColMap - ColMapUp
                        ColAdj = ColAdj + ColMap
                     Next a
                     ColAdj = ColAdj / 255
                     ColAdj = ColAdj / 2
                     ColMap = 255
                     For a = 0 To GapPieces 'To 0 Step -1
                         LCY2 = LCY2 + GhostNumY
                         LCX2 = LCX2 + GhostNumX
                         ColMap = ColMap - ColMapUp
                         If ColMap < 0 Then ColMap = 0
                         'ColMap = 255
                         frmGame.Picture1.PSet (LCX2, LCY2), RGB(ColMap / ColAdj, ColMap / ColAdj, ColMap / ColAdj)
                         'Form1.Picture1.PSet (LCX2, LCY2), RGB(ColMap / ColAdj, ColMap / ColAdj, ColMap / ColAdj)

                         X = X
                     Next a
                Else
                
                    frmGame.Picture1.PSet (LCX2, LCY2), RGB(255, 255, 255)
                End If 'X = X
                
            Else
                frmGame.Picture1.PSet (LCX2, LCY2), RGB(255, 255, 255)
            End If
        
        
            
        Else
            For Y = 0 To FinalPoint(X) - 2 Step 2
                 LCX2 = CLng(LastCoords2(X, Y, 0))
                 LCX1 = CLng(DrawObject(X, Y))
                 LCY2 = CLng(LastCoords2(X, Y + 1, 0))
                 LCY1 = CLng(DrawObject(X, Y + 1))
                 LCA2 = CLng(LastCoords2(X, Y + 2, 0))
                 LCA1 = CLng(DrawObject(X, Y + 2))
                 LCB2 = CLng(LastCoords2(X, Y + 3, 0))
                 LCB1 = CLng(DrawObject(X, Y + 3))
                 
                 If LCX1 <= 1 And LCY1 <= 1 Then
                    LCX1 = LCX2
                    LCY1 = LCY2
                 End If
                 If LCX2 <= 1 And LCY2 <= 1 Then
                    LCX2 = LCX1
                    LCY2 = LCY1
                 End If
                 GhostNumX = LCX1 - LCX2
                 GhostNumY = LCY1 - LCY2
                 GapPieces = 6
                If X = 12345 Then 'X <> 1 And X <> 4 And Abs(GhostNumX) < 5000 And Abs(GhostNumY) < 5000 And (Abs(LCX1 - LCX2) >= Screen.TwipsPerPixelX * 4 Or Abs(LCY1 - LCY2) >= Screen.TwipsPerPixelY * 4) And (LCX1 <> 0 Or LCY1 <> 0) Then
                    'Form1.Picture1.Circle (LCX2, LCY2), 30, RGB(255, 0, 0)
                    'work out how many ghosts need to be printed
                    'If Abs(LCX1 - LCX2) > Abs(LCY1 - LCY2) Then
                   
                    
                    ' End If
                     If Abs(GhostNumY) > Abs(GhostNumX) Then
                         GapPieces = (Abs((GhostNumY / Screen.TwipsPerPixelY) / 4))
                         GhostNum = Abs(CLng(GhostNumY) / GapPieces)
                     Else
                         GapPieces = (Abs((GhostNumX / Screen.TwipsPerPixelX) / 4))
                         GhostNum = Abs(CLng(GhostNumX) / GapPieces)
                     End If
                     Ghos2tNumX = GhostNumX / GapPieces ' CLng(Ghostnum / 12)
                     GhostNumY = GhostNumY / GapPieces 'GhostNum 'CLng(Ghostnum / 12)
                     'LCY1 = LCY2
                     'LCX1 = LCX2
                     ColMapUp = CInt(255 / GapPieces)
                     
                     
                     ColMap = 0
                     ColAdj = 0
                     For a = 0 To GapPieces
                        ColMap = ColMap + ColMapUp
                        ColAdj = ColAdj + ColMap
                     Next a
                     ColAdj = ColAdj / 255
                     ColMap = 0
                     For a = 0 To GapPieces 'To 0 Step -1
                         LCX1 = LCX1 + GhostNumX
                         LCY1 = LCY1 + GhostNumY
                         LCA1 = LCA1 + GhostNumX
                         LCB21 = LCB1 + GhostNumY
                         ColMap = ColMap + ColMapUp
                         If ColMap > 255 Then ColMap = 255
                         frmGame.Picture1.ForeColor = RGB(ColMap / ColAdj, ColMap / ColAdj, ColMap / ColAdj)
                         Picture1.Line (LCX1, LCY1)-(LCA1, LCB1)
                        If YMod <> 0 And YMod <> 0 Then
                           Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                        End If
                        If XMod <> 0 Then
                            Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                        End If
                        If YMod <> 0 Then
                           Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                        End If
                     Next a
                     X = X
                Else
                 
                 
                    Picture1.Line (LCX1, LCY1)-(LCA1, LCB1)
                    If YMod <> 0 And YMod <> 0 Then
                       Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                    End If
                    If XMod <> 0 Then
                        Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                    End If
                    If YMod <> 0 Then
                       Picture1.Line (LCX1 + XMod, LCY1 + YMod)-(LCA1 + XMod, LCB1 + YMod)
                    End If
                End If
             Next Y
             

        End If
         xx = X
    End If
Next X



For X = 0 To MaxObjects
    If SpaceObject(X, 8) < -1 Then
        SpaceObject(X, 8) = SpaceObject(X, 8) + 1
    ElseIf SpaceObject(X, 8) = -1 Then
        SpaceObject(X, 8) = 1
    
    
    End If
    For Y = 0 To FinalPoint(X) Step 2
        LastCoords(X, Y, 1) = LastCoords(X, Y, 0)
        LastCoords(X, Y + 1, 1) = LastCoords(X, Y + 1, 0)
    Next Y
Next X
xx = SpaceObject(0, 5)
FirstdrawFlag = 0
frmGame.Picture1.ForeColor = oc


If LSpaceObject(BulletNo, 0) > 0 Then

    If SpaceObject(BulletNo, 9) > 0 Then SpaceObject(BulletNo, 9) = 0

End If
oc = Picture1.ForeColor
Picture1.ForeColor = RGB(255, 255, 255)
Picture1.DrawMode = 12
Picture1.FontSize = 20
Picture1.FontBold = True

Picture1.CurrentX = 1000
Picture1.CurrentY = 50
Picture1.Print "Player 1 Score: " & Format(Scores(1), "00")

Picture1.CurrentX = Picture1.Width - Picture1.TextWidth(Str(Scores(0))) - 5000
Picture1.CurrentY = 50
Picture1.Print "Player 2 Score: " & Format(Scores(0), "00")


Picture1.ForeColor = oc
Picture1.DrawMode = 13

DoEvents

Picture1.Refresh
'Picture1.SetFocus

End Sub
