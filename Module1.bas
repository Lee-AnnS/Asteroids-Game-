Attribute VB_Name = "Module1"
Public SOLast() As Long, ExitE As Byte, FinalPoint() As Long, DrawObject() As Long, FirstdrawFlag, LastCoords() As Double, LastCoords2() As Double, BulletNo As Long, Scores(1) As Long, MDC() As Long, LSpaceObject() As Double, TimeKeeper(10) As Long
Public NumObjects, SpaceObject() As Double, MaxObjects As Long, Detail As Long, ScaleFactor(1) As Double, KP(255) As Long, ShapeStart As Long, VelocityMod As Double
Public CollideRecord() As Byte
Public MaxProx() As Single, MinProx() As Single
Public MaxProx2D() As Single, MinProx2D() As Single
Public Const Pi = 3.14159265358979

Type POINTAPI
    XPos As Long
    Y As Long
End Type

Declare Function GetTickCount Lib "kernel32" () As Long
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Sub DetectCollide2()
For X = 0 To MaxObjects
    If SpaceObject(X, 0) = 1 Then
        For Z = X + 1 To MaxObjects
            
            
        Next Z
    End If
Next X
End Sub
Public Sub MakeCoords(SpaceObject() As Double, DrawObject() As Long, FinalPoint() As Long, SX, EX)
For X = SX To EX
    If SpaceObject(X, 0) = 1 And SpaceObject(X, 6) > 1 Then
        'For Y = 0 To Detail

            'now translate all the degrees and distances from centre into x-y coordinates
            
            FinalPoint(X) = 0
            XCentre = (SpaceObject(X, 1))
            YCentre = (SpaceObject(X, 2))
            Degrees = (SpaceObject(X, 3))
            Velocity = (SpaceObject(X, 4)) * VelocityMod
            Direction = (SpaceObject(X, 5))
            For Y = ShapeStart To Detail - 1 Step 2 'Detail Step 2
                If SpaceObject(X, Y) = -1 Then Exit For
                DegreesP = SpaceObject(X, Y) + Degrees
                If DegreesP >= 360 Then DegreesP = (DegreesP - 360)
                If DegreesP > 0 And DegreesP < 90 Then
                    TD = 90 - DegreesP
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = -S * SpaceObject(X, Y + 1)
                    S = Cos(TD)
                    xoff = S * SpaceObject(X, Y + 1)
                ElseIf DegreesP = 0 Then
                    yoff = -SpaceObject(X, Y + 1)
                    xoff = 0
                ElseIf DegreesP = 90 Then
                    yoff = 0
                    xoff = SpaceObject(X, Y + 1)
                ElseIf DegreesP > 90 And DegreesP < 180 Then
                    TD = DegreesP - 90
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = S * SpaceObject(X, Y + 1)
                    S = Cos(TD)
                    xoff = S * SpaceObject(X, Y + 1)
                    X = X
                ElseIf DegreesP = 180 Then
                    yoff = SpaceObject(X, Y + 1)
                    xoff = 0
                ElseIf DegreesP > 180 And DegreesP < 270 Then
                    TD = DegreesP - 180
                    TD = TD * (Pi / 180)
                    S = Cos(TD)
                    yoff = S * SpaceObject(X, Y + 1)
                    S = Sin(TD)
                    xoff = -S * SpaceObject(X, Y + 1)
                ElseIf DegreesP = 270 Then
                    yoff = 0
                    xoff = -SpaceObject(X, Y + 1)
                ElseIf DegreesP > 270 And DegreesP < 360 Then
                    TD = DegreesP - 270
                    TD = TD * (Pi / 180)
                    S = Sin(TD)
                    yoff = -S * SpaceObject(X, Y + 1) '-83
                    S = Cos(TD)
                    xoff = -S * SpaceObject(X, Y + 1) '-54
2                    X = X
                End If
                DrawObject(X, Y - ShapeStart) = XCentre + xoff '7394,4604:7394,4604
                DrawObject(X, (Y - ShapeStart + 1)) = YCentre + yoff
                FinalPoint(X) = (Y - ShapeStart + 1)
            Next Y
            DrawObject(X, FinalPoint(X) + 1) = DrawObject(X, 0)
            DrawObject(X, FinalPoint(X) + 2) = DrawObject(X, 1)
            FinalPoint(X) = FinalPoint(X) + 2
            'SpaceObject(X, 1) = XCentre
            'SpaceObject(X, 2) = YCentre
        'Next Y
    End If
Next X
End Sub
Public Sub FindDirection(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length As Double, Gradient As Double)

Length = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
'Length = Length * VelocityMod
'X1 = 10
'Y1 = 0
'X2 = 9
'Y2 = 10
'gradient  = change in  y/change in x
If (X1 - X2) <> 0 Then
    If (Y1 - Y2) <> 0 Then
        Gradient = (Y1 - Y2) / (X1 - X2)
        
        Gradient = Atn(Gradient) * 180 / Pi
        If Y1 <= Y2 And X1 <= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 180
            Else
                Gradient = Gradient + 90
            End If
        ElseIf Y1 >= Y2 And X1 >= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 360
            Else
                Gradient = Gradient + 270
            End If
        ElseIf Y1 <= Y2 And X1 >= X2 Then
           X = X
           If Gradient < 0 Then
                Gradient = 270 + Gradient
            Else
                Gradient = 180 + Gradient
            End If
        ElseIf Y1 >= Y2 And X1 <= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 90
            Else
                Gradient = Gradient
            End If
        End If
        
2        X = X
    Else
        
        If X1 <= X2 Then
             Gradient = 90
             
        ElseIf X1 >= X2 Then
            
            Gradient = 270
            
        
        End If
        
    End If
Else
     If (Y1 - Y2) <> 0 Then
        'Gradient = (Y1 - Y2) / (0.000000001)
        'Gradient = Atn(Gradient) * 180 / Pi
        If Y1 <= Y2 Then
            Gradient = 180
            
        ElseIf Y1 >= Y2 Then
           Gradient = 0
        End If
    Else
        Gradient = 0 / 0.000000001
        Gradient = Atn(Gradient) * 180 / Pi
        If Y1 <= Y2 And X1 <= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 180
            Else
                Gradient = Gradient + 90
            End If
        ElseIf Y1 >= Y2 And X1 >= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 360
            Else
                Gradient = Gradient + 270
            End If
        ElseIf Y1 <= Y2 And X1 >= X2 Then
           X = X
           If Gradient < 0 Then
                Gradient = 270 + Gradient
            Else
                Gradient = 180 + Gradient
            End If
        ElseIf Y1 >= Y2 And X1 <= X2 Then
            If Gradient < 0 Then
                Gradient = Gradient + 90
            Else
                Gradient = Gradient
            End If
        End If
        
    End If
End If
If Gradient < 0 Then Gradient = Gradient + 360
If Gradient >= 360 Then Gradient = Gradient - 360

End Sub
Public Sub DetectCollide(FinalPoint() As Long)
Dim A1 As Double, A2 As Double, CollideA, CollideB, ProxY(1, 4) As Long, ProxZ(1, 4) As Long, Z As Long, Extremes(1, 4), XOL As Long, YOL As Long, V(1)


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


For X = 0 To MaxObjects
    If CollideRecord(X) = 0 Or X = X Then
        If SpaceObject(X, 0) = 1 And SpaceObject(X, 8) = 1 Then
            V(0) = SpaceObject(X, 4)
            
            If Abs(LSpaceObject(X, 1) - SpaceObject(X, 1)) > V(0) Then
                If LSpaceObject(X, 1) > SpaceObject(X, 1) Then
                    LSpaceObject(X, 1) = LSpaceObject(X, 1) - Form1.Picture1.Width
                Else
                    LSpaceObject(X, 1) = Form1.Picture1.Width + LSpaceObject(X, 1)
                End If
            End If
            If Abs(LSpaceObject(X, 2) - SpaceObject(X, 2)) > V(0) Then
                If LSpaceObject(X, 2) > SpaceObject(X, 2) Then
                    LSpaceObject(X, 2) = LSpaceObject(X, 2) - Form1.Picture1.Height
                Else
                    LSpaceObject(X, 2) = Form1.Picture1.Height + LSpaceObject(X, 2)
                End If
            End If
            ProxY(0, 0) = SpaceObject(X, 1) - MDC(X)
            ProxY(0, 1) = SpaceObject(X, 1) + MDC(X)
            ProxY(0, 2) = SpaceObject(X, 2) - MDC(X)
            ProxY(0, 3) = SpaceObject(X, 2) + MDC(X)
            ProxY(1, 0) = LSpaceObject(X, 1) - MDC(X)
            ProxY(1, 1) = LSpaceObject(X, 1) + MDC(X)
            ProxY(1, 2) = LSpaceObject(X, 2) - MDC(X)
            ProxY(1, 3) = LSpaceObject(X, 2) + MDC(X)
            
            
            If ProxY(0, 0) < ProxY(1, 0) Then
                Extremes(0, 0) = ProxY(0, 0)
            Else
                Extremes(0, 0) = ProxY(1, 0)
            End If
            If ProxY(0, 1) > ProxY(1, 1) Then
                Extremes(0, 1) = ProxY(0, 1)
            Else
                Extremes(0, 1) = ProxY(1, 1)
            End If
            If ProxY(0, 2) < ProxY(1, 2) Then
                Extremes(0, 2) = ProxY(0, 2)
            Else
                Extremes(0, 2) = ProxY(1, 2)
            End If
            If ProxY(0, 3) > ProxY(1, 3) Then
                Extremes(0, 3) = ProxY(0, 3)
            Else
                Extremes(0, 3) = ProxY(1, 3)
            End If
            For Z = 0 To MaxObjects
                If CollideRecord(Z) = 0 Or X = X Then
                    If Z <> X And SpaceObject(X, 9) <> Z + 1 And SpaceObject(Z, 9) <> X + 1 Then
                         
                         If SpaceObject(Z, 0) = 1 And LSpaceObject(Z, 0) = 1 Then
                            V(1) = SpaceObject(Z, 4)
                            If Abs(LSpaceObject(Z, 1) - SpaceObject(Z, 1)) > V(1) Then
                                If LSpaceObject(Z, 1) > SpaceObject(Z, 1) Then
                                    o = LSpaceObject(Z, 1)
                                    LSpaceObject(Z, 1) = LSpaceObject(Z, 1) - Form1.Picture1.Width
                                Else
                                    o = LSpaceObject(Z, 1)
                                    LSpaceObject(Z, 1) = Form1.Picture1.Width + LSpaceObject(Z, 1)
                                End If
                            End If
                            
                            
                            
                            If Abs(LSpaceObject(Z, 2) - SpaceObject(Z, 2)) > V(1) Then
                                If LSpaceObject(Z, 2) > SpaceObject(Z, 2) Then
                                    LSpaceObject(Z, 2) = LSpaceObject(Z, 2) - Form1.Picture1.Height
                                Else
                                    LSpaceObject(Z, 2) = Form1.Picture1.Height + LSpaceObject(Z, 2)
                                End If
                            End If
                            
                            ProxZ(0, 0) = SpaceObject(Z, 1) - MDC(Z)
                            ProxZ(0, 1) = SpaceObject(Z, 1) + MDC(Z)
                            ProxZ(0, 2) = SpaceObject(Z, 2) - MDC(Z)
                            ProxZ(0, 3) = SpaceObject(Z, 2) + MDC(Z)
                            ProxZ(1, 0) = LSpaceObject(Z, 1) - MDC(Z)
                            ProxZ(1, 1) = LSpaceObject(Z, 1) + MDC(Z)
                            ProxZ(1, 2) = LSpaceObject(Z, 2) - MDC(Z)
                            ProxZ(1, 3) = LSpaceObject(Z, 2) + MDC(Z)
                            If ProxZ(0, 0) < ProxZ(1, 0) Then
                                Extremes(1, 0) = ProxZ(0, 0)
                            Else
                                Extremes(1, 0) = ProxZ(1, 0)
                            End If
                            If ProxZ(0, 1) > ProxZ(1, 1) Then
                                Extremes(1, 1) = ProxZ(0, 1)
                            Else
                                Extremes(1, 1) = ProxZ(1, 1)
                            End If
                            If ProxZ(0, 2) < ProxZ(1, 2) Then
                                Extremes(1, 2) = ProxZ(0, 2)
                            Else
                                Extremes(1, 2) = ProxZ(1, 2)
                            End If
                            If ProxZ(0, 3) > ProxZ(1, 3) Then
                                Extremes(1, 3) = ProxZ(0, 3)
                            Else
                                Extremes(1, 3) = ProxZ(1, 3)
                            End If
                            XOL = 0
                            YOL = 0
                            If Extremes(0, 0) >= Extremes(1, 0) And Extremes(0, 0) <= Extremes(1, 1) Then
                                XOL = 1
                            ElseIf Extremes(0, 1) >= Extremes(1, 0) And Extremes(0, 1) <= Extremes(1, 1) Then
                                XOL = 1
                            ElseIf Extremes(0, 0) <= Extremes(1, 0) And Extremes(0, 1) >= Extremes(1, 1) Then
                                XOL = 1
                            End If
                            If XOL = 1 Then
                                If Extremes(0, 2) >= Extremes(1, 2) And Extremes(0, 2) <= Extremes(1, 3) Then
                                    YOL = 1
                                ElseIf Extremes(0, 3) >= Extremes(1, 2) And Extremes(0, 3) <= Extremes(1, 3) Then
                                    YOL = 1
                                ElseIf Extremes(0, 2) <= Extremes(1, 2) And Extremes(0, 3) >= Extremes(1, 3) Then
                                    YOL = 1
                                End If
                                If YOL = 1 Then
                                    'Form1.Picture1.ForeColor = RGB(128, 128, 128)
                                    Form1.Picture1.DrawMode = 9
                                    
                                    'Form1.Picture1.Line (Extremes(0, 0), Extremes(0, 2))-(Extremes(0, 1), Extremes(0, 3)), RGB(128, 128, 128), BF
                                    'Form1.Picture1.Line (Extremes(1, 0), Extremes(1, 2))-(Extremes(1, 1), Extremes(1, 3)), RGB(128, 128, 128), BF
                                    
                                    
                                    Dim LCoord() As Long, Coord() As Long, FinalPos() As Long
                                    
                                    Dim mcVals() As Double
                                    'Test to see if lines in objects cross (1) at this time
                                    '(2) check to see if the lines between last coord and this coord for any points overlap
                                    
                                    
                                    ReDim mcVals(3, Detail, 1)
                                    
                                    
                                    'work out m and c for all timelapse points of x
                                    For a = 0 To FinalPoint(X) - 2 Step 2
                                        cx = LastCoords(X, a, 0) - LastCoords(X, a, 1)
                                        cy = LastCoords(X, a + 1, 0) - LastCoords(X, a + 1, 1)
                                        If cx <> 0 Then
                                            mcVals(0, a, 0) = cy / cx
                                            mcVals(0, a, 1) = LastCoords(X, a + 1, 0) - mcVals(0, a, 0) * LastCoords(X, a, 0)
                                            X = X
                                        Else
                                            mcVals(0, a, 0) = 1000000000
                                            mcVals(0, a, 1) = 1000000000
                                        End If
                                    Next a
                                    'work out m and c for all current points of x
                                    For a = 0 To FinalPoint(X) - 2 Step 2
                                        cx = LastCoords(X, a, 0) - LastCoords(X, a + 2, 0)
                                        cy = LastCoords(X, a + 1, 0) - LastCoords(X, a + 3, 0)
                                        If cx <> 0 Then
                                            mcVals(1, a, 0) = cy / cx
                                            mcVals(1, a, 1) = LastCoords(X, a + 1, 0) - mcVals(1, a, 0) * LastCoords(X, a, 0)
                                        Else
                                            mcVals(1, a, 0) = 1000000000
                                            mcVals(1, a, 1) = 1000000000
                                        End If
                                        X = X
                                    Next a
                                    'work out m and c for all timelapse points of z
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        cx = LastCoords(Z, a, 0) - LastCoords(Z, a, 1)
                                        cy = LastCoords(Z, a + 1, 0) - LastCoords(Z, a + 1, 1)
                                        If cx <> 0 Then
                                            mcVals(2, a, 0) = cy / cx
                                            mcVals(2, a, 1) = LastCoords(Z, a + 1, 0) - mcVals(2, a, 0) * LastCoords(Z, a, 0)
                                        Else
                                            mcVals(2, a, 0) = 1000000000
                                            mcVals(2, a, 1) = 1000000000
                                        End If
                                        X = X
                                    Next a
                                    'work out m and c for all current points of z
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        cx = LastCoords(Z, a, 0) - LastCoords(Z, a + 2, 0)
                                        cy = LastCoords(Z, a + 1, 0) - LastCoords(Z, a + 3, 0)
                                        If cx <> 0 Then
                                            mcVals(3, a, 0) = cy / cx
                                            mcVals(3, a, 1) = LastCoords(Z, a + 1, 0) - mcVals(3, a, 0) * LastCoords(Z, a, 0)
                                        Else
                                            mcVals(3, a, 0) = 1000000000
                                            mcVals(3, a, 1) = 1000000000
                                        End If
                                    Next a
                                    
                                    
                                    'solve equations to see where lines overlap.
                                    
                                    Dim XV, YV
                                    'taplapse z vs current x
                                    collide = 0
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        'If LastCoords(Z, a + 3, 0) > 0 And LastCoords(Z, a + 2, 0) > 0 Then
                                            For b = 0 To FinalPoint(X) - 2 Step 2
                                                'If LastCoords(X, b + 3, 0) > 0 And LastCoords(X, b + 2, 0) > 0 Then
                                                    'solve for x
                                                    
                                                    If (mcVals(1, b, 0) - mcVals(2, a, 0)) <> 0 Then
                                                        If mcVals(1, b, 1) < 100000000 And mcVals(2, a, 1) < 100000000 Then
                                                            XV = (mcVals(1, b, 1) - mcVals(2, a, 1)) / (mcVals(2, a, 0) - mcVals(1, b, 0))
                                                            YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                        Else
                                                            If mcVals(1, b, 1) > 100000000 And mcVals(2, a, 1) < 100000000 Then
                                                                XV = LastCoords(X, b, 0)
                                                                YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                            ElseIf mcVals(1, b, 1) < 100000000 And mcVals(2, a, 1) > 100000000 Then
                                                                XV = LastCoords(Z, a, 0)
                                                                YV = mcVals(1, b, 0) * XV + mcVals(1, b, 1)
                                                            Else
                                                                XV = -1
                                                                YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                            End If
                                                        End If
                                                    'solve y
                                                        
                                                        'is this point of the line in x current
                                                        CollideA = 0
                                                        If LastCoords(X, b, 0) >= XV And LastCoords(X, b + 2, 0) <= XV Then
                                                            If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                CollideA = 1
                                                            ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                CollideA = 1
                                                            End If
                                                        ElseIf LastCoords(X, b, 0) <= XV And LastCoords(X, b + 2, 0) >= XV Then
                                                            If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                CollideA = 1
                                                            ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                CollideA = 1
                                                            End If
                                                        End If
                                                        CollideB = 0
                                                        If LastCoords(Z, a, 0) >= XV And LastCoords(Z, a, 1) <= XV Then
                                                            If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                CollideB = 1
                                                            ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                CollideB = 1
                                                            End If
                                                        ElseIf LastCoords(Z, a, 0) <= XV And LastCoords(Z, a, 1) >= XV Then
                                                            If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                CollideB = 1
                                                            ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                CollideB = 1
                                                            End If
                                                        End If
                                                        
                                                        
                                                        If CollideA = 1 And CollideB = 1 Then
                                                            Form1.Picture1.DrawMode = 13
                                                            Form1.Picture1.Circle (XV, YV), 30, RGB(255, 255, 0)
                                                            Form1.Picture1.Circle (LastCoords(X, b, 0), LastCoords(X, b + 1, 0)), 20, RGB(0, 255, 0)
                                                            Form1.Picture1.Circle (LastCoords(X, b + 2, 0), LastCoords(X, b + 3, 0)), 20, RGB(0, 255, 255)
                                                            Form1.Picture1.Circle (LastCoords(Z, a, 0), LastCoords(Z, a + 1, 0)), 20, RGB(255, 0, 255)
                                                            Form1.Picture1.Circle (LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1)), 20, RGB(255, 255, 255)
                                                            
                                                            SpaceObject(X, 8) = -100
                                                            SpaceObject(Z, 8) = -100
                                                            
                                                            
                                                            
                                                            'work out the angle of incidence of z to x
                                                            'mcvals(1,b,0) = slope of x
                                                            'mcvals(2,a,0) = slope of z
                                                            If mcVals(1, b, 0) < 10000000 Then
                                                                A1 = Atn(mcVals(1, b, 0)) * 180 / Pi
                                                            Else
                                                                A1 = 90
                                                            End If
                                                            
                                                            If mcVals(2, a, 0) < 10000000 Then
                                                                A2 = Atn(mcVals(2, a, 0)) * 180 / Pi
                                                            Else
                                                                A2 = 90
                                                            End If
                                                            
                                                            If A1 < 0 Then A1 = Abs(90 + A1) + 90
                                                            If A2 < 0 Then A2 = Abs(90 + A2) + 90
                                                            'angle of incidence = a2-a1
                                                            ai = Abs(A1 - A2)
                                                            'If ai > 90 Then ai = 180 - ai
                                                            'new angle of a = a2-2*ai
                                                            'ar = A2 - (180 + A1)
                                                            'ar = A1 + Ai '2 * (180 - ai)
                                                            'ar = ar - 90
                                                            'now work out whether comming from bottom or top
                                                            
                                                            If A1 >= 90 Then
                                                                If LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) Then
                                                                    ar = A1 + ai + 180
                                                                ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) Then
                                                                    ar = A1 + ai
                                                                Else
                                                                    If LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                                        ar = A1 + ai
                                                                    ElseIf LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                                        ar = A1 + ai + 180
                                                                    Else
                                                                        'If SpaceObject(Z, 5) > 0 Then
                                                                            ar = SpaceObject(Z, 5) - 180
                                                                        
                                                                        
                                                                        'End If
                                                                    End If
                                                                End If
                                                            Else
                                                                If LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) Then
                                                                     ar = A1 - ai + 180
                                                                ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) Then
                                                                    ar = (A1 - ai)
                                                                Else
                                                                    If LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                                        ar = (A1 - ai)
                                                                    ElseIf LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                                         ar = A1 - ai + 180
                                                                    Else
                                                                        'If LastCoords(X, a, 0) < LastCoords(X, a, 1) Then
                                                                        ar = SpaceObject(Z, 5) - 180
                                                                        'End If
                                                                        
                                                                    End If
                                                                End If
                                                            
                                                            End If
                                                            X = X
                                                            'If LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                            '    ar = A1 + ai + 180
                                                            'ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                            '
                                                            '    ar = A1 - ai
                                                            'ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                            '    ar = A1 + ai
                                                           '
                                                           ' ElseIf LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                           '     ar = A1 - ai + 180
                                                           '
                                                           ' End If
                                                            
                                                            ar = ar - 90 ' puts it back into the game mode -i.e with 0 degrees straight up
                                                            
                                                            If ar >= 360 Then ar = ar - 360
                                                            If ar < 0 Then ar = ar + 360
                                                            xc = XV - LastCoords(Z, a, 0)
                                                            yc = YV - LastCoords(Z, a + 1, 0)
                                                            If xc < 0 Then
                                                                xc = xc - 5
                                                            ElseIf xc > 0 Then
                                                                xc = xc + 5
                                                            Else
                                                                If ar > 0 And ar < 180 Then
                                                                    xc = xc + 5
                                                                ElseIf ar > 180 Then
                                                                    xc = xc - 5
                                                                
                                                                End If
                                                            End If
8
                                                            If yc > 0 Then
                                                                yc = yc + 5
                                                            ElseIf yc < 1 Then
                                                                yc = yc - 5
                                                            Else
                                                                If ar > 270 Or ar < 90 Then
                                                                    yc = yc - 5
                                                                ElseIf ar < 270 And ar > 90 Then
                                                                    yc = yc + 5
                                                                End If
                                                            End If
                                                            
                                                            'THIS DOESN't WORK IF COLLIDING OBJECTS ARE MOVING IN THE SAME DIRECTION
                                                            For C = 0 To FinalPoint(Z) - 1 Step 2
                                                                LastCoords(Z, C, 0) = LastCoords(Z, C, 0) + xc
                                                                LastCoords(Z, C + 1, 0) = LastCoords(Z, C + 1, 0) + yc
                                                            Next C
                                                           ' SpaceObject(Z, 1) = SpaceObject(Z, 1) + xc
                                                           ' SpaceObject(Z, 2) = SpaceObject(Z, 2) + yc
                                                            
                                                            SpaceObject(Z, 5) = ar
                                                            
                                                            
                                                            'now do the x objext
                                                            If X = 12345 Then
                                                                If A2 >= 90 Then
                                                                 If LastCoords(X, a + 1, 0) > LastCoords(X, a + 1, 1) Then
                                                                     ar = A2 + ai + 180
                                                                 ElseIf LastCoords(X, a + 1, 0) < LastCoords(X, a + 1, 1) Then
                                                                     ar = A2 + ai
                                                                 Else
                                                                     If LastCoords(X, a, 0) > LastCoords(X, a, 1) Then
                                                                         ar = A2 + ai
                                                                     ElseIf LastCoords(X, a, 0) < LastCoords(X, a, 1) Then
                                                                         ar = A2 + ai + 180
                                                                     Else
                                                                         'If SpaceObject(x, 5) > 0 Then
                                                                             ar = SpaceObject(X, 5) - 180
                                                                         
                                                                         
                                                                         'End If
                                                                     End If
                                                                 End If
                                                             Else
                                                                 If LastCoords(X, a + 1, 0) > LastCoords(X, a + 1, 1) Then
                                                                      ar = A2 - ai + 180
                                                                 ElseIf LastCoords(X, a + 1, 0) < LastCoords(X, a + 1, 1) Then
                                                                     ar = (A2 - ai)
                                                                 Else
                                                                     If LastCoords(X, a, 0) > LastCoords(X, a, 1) Then
                                                                         ar = (A2 - ai)
                                                                     ElseIf LastCoords(X, a, 0) < LastCoords(X, a, 1) Then
                                                                          ar = A2 - ai + 180
                                                                     Else
                                                                         'If LastCoords(X, a, 0) < LastCoords(X, a, 1) Then
                                                                         ar = SpaceObject(X, 5) - 180
                                                                         'End If
                                                                         
                                                                     End If
                                                                 End If
                                                             
                                                             End If
                                                             X = X
                                                             'If LastCoords(x, a + 1, 0) > LastCoords(x, a + 1, 1) And LastCoords(x, a, 0) > LastCoords(x, a, 1) Then
                                                             '    ar = A2 + ai + 180
                                                             'ElseIf LastCoords(x, a + 1, 0) < LastCoords(x, a + 1, 1) And LastCoords(x, a, 0) > LastCoords(x, a, 1) Then
                                                             '
                                                             '    ar = A2 - ai
                                                             'ElseIf LastCoords(x, a + 1, 0) < LastCoords(x, a + 1, 1) And LastCoords(x, a, 0) < LastCoords(x, a, 1) Then
                                                             '    ar = A2 + ai
                                                            '
                                                            ' ElseIf LastCoords(x, a + 1, 0) > LastCoords(x, a + 1, 1) And LastCoords(x, a, 0) < LastCoords(x, a, 1) Then
                                                            '     ar = A2 - ai + 180
                                                            '
                                                            ' End If
                                                             
                                                             ar = ar - 90 ' puts it back into the game mode -i.e with 0 degrees straight up
                                                             
                                                             If ar >= 360 Then ar = ar - 360
                                                             If ar < 0 Then ar = ar + 360
                                                             xc = XV - LastCoords(X, a, 0)
                                                             yc = YV - LastCoords(X, a + 1, 0)
                                                             If xc < 0 Then
                                                                 xc = xc - 5
                                                             ElseIf xc > 0 Then
                                                                 xc = xc + 5
                                                             Else
                                                                 If ar > 0 And ar < 180 Then
                                                                     xc = xc + 5
                                                                 ElseIf ar > 180 Then
                                                                     xc = xc - 5
                                                                 
                                                                 End If
                                                             End If
                                                             
                                                             If yc > 0 Then
                                                                 yc = yc + 5
                                                             ElseIf yc < 1 Then
                                                                 yc = yc - 5
                                                             Else
                                                                 If ar > 270 Or ar < 90 Then
                                                                     yc = yc - 5
                                                                 ElseIf ar < 270 And ar > 90 Then
                                                                     yc = yc + 5
                                                                 End If
                                                             End If
                                                             
                                                             'THIS DOESN't WORK IF COLLIDING OBJECTS ARE MOVING IN THE SAME DIRECTION
                                                             For C = 0 To FinalPoint(X) - 1 Step 2
                                                                 LastCoords(X, C, 0) = LastCoords(X, C, 0) + xc
                                                                 LastCoords(X, C + 1, 0) = LastCoords(X, C + 1, 0) + yc
                                                             Next C
                                                            ' SpaceObject(x, 1) = SpaceObject(x, 1) + xc
                                                            ' SpaceObject(x, 2) = SpaceObject(x, 2) + yc
                                                             
                                                             SpaceObject(X, 5) = ar
                                                            
                                                            
                                                            End If
                                                            
                                                            
                                                            
                                                            'xc = ((LastCoords(X, b, 0) - LastCoords(X, b, 1)) + (LastCoords(X, b + 2, 0) - LastCoords(X, b + 2, 1))) / 2
                                                            'yc = ((LastCoords(X, b + 1, 0) - LastCoords(X, b + 1, 1)) + (LastCoords(X, b + 3, 0) - LastCoords(X, b + 3, 1))) / 2
                                                            
                                                            X1 = (LastCoords(X, b, 0) + LastCoords(X, b + 2, 0)) / 2
                                                            Y1 = (LastCoords(X, b + 1, 0) + LastCoords(X, b + 3, 0)) / 2
                                                            
                                                            X2 = (LastCoords(X, b, 1) + LastCoords(X, b + 2, 1)) / 2
                                                            Y2 = (LastCoords(X, b + 1, 1) + LastCoords(X, b + 3, 1)) / 2
                                                            
                                                            'gradient of this line is the direction of the "hit" and the length is the velocity of the hit
                                                            Length = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
                                                            'gradient  = change in  y/change in x
                                                            If (X1 - X2) > 0 Then
                                                                If (Y1 - Y2) > 0 Then
                                                                    Gradient = (Y1 - Y2) / (X1 - X2)
                                                                    Gradient = Atn(Gradient) * 180 / Pi
                                                                    Gradient = Gradient - 90
                                                                    If Gradient < 0 Then Gradient = Gradient + 360
                                                                Else
                                                                    Gradient = 90
                                                                End If
                                                            Else
                                                                Gradient = 0
                                                                'If mcVals(1, b, 0) < 10000000 Then
                                                                '    A1 = Atn(mcVals(1, b, 0)) * 180 / Pi
                                                                'Else
                                                                '    A1 = 90
                                                                'End If
                                                            End If
                                                            
                                                           'NEED TO DO A FORCE EQUATION WITH TRIANGLES
                                                           'BUT USING MOMENTUM TO WORK OUT CHANGES IN
                                                           'DIRCTION.
                                                            
                                                            'For c = 0 To FinalPoint(X) - 1 Step 2
                                                            '    LastCoords(X, c, 0) = LastCoords(X, c, 0) - xc
                                                            '    LastCoords(X, c + 1, 0) = LastCoords(X, c + 1, 0) - yc
                                                            'Next c
                                                            'SpaceObject(X, 1) = SpaceObject(X, 1) - xc
                                                            'SpaceObject(X, 2) = SpaceObject(X, 2) - yc
                                                           
                                                            ' Form1.Picture1.Circle (LastCoords(X, b, 0), LastCoords(X, b + 1, 0)), 20, RGB(0, 255, 0)
                                                            'Form1.Picture1.Circle (LastCoords(X, b + 2, 0), LastCoords(X, b + 3, 0)), 20, RGB(0, 255, 255)
                                                            
                                                            
                                                            
                                                            'work out velocities
                                                            m1 = SpaceObject(X, 6) * Length 'length = SpaceObject(X, 4) + spin effect
                                                            m2 = SpaceObject(Z, 6) * SpaceObject(Z, 4) 'iv2 'iv2 = SpaceObject(z, 4) + spin effect
                                                            tm = m1 + m2
                                                            If SpaceObject(X, 6) + SpaceObject(Z, 6) > 0 Then
                                                                m1 = tm * SpaceObject(X, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6))
                                                            
                                                                m2 = tm * SpaceObject(Z, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6))
                                                            End If
                                                            If SpaceObject(X, 6) > 0 Then
                                                                SpaceObject(X, 4) = m1 / SpaceObject(X, 6)
                                                            End If
8                                                            If SpaceObject(Z, 6) > 0 Then
                                                                SpaceObject(Z, 4) = m2 / SpaceObject(Z, 6)
                                                            End If
                                                            
                                                            
                                                            collide = 1
                                                            
                                                            'a = 0
                                                            'b = 0
                                                        End If
                                                        
                                                        
                                                    End If
                                                    If collide = 0 And X = 12345 Then
                                                        If (mcVals(3, b, 0) - mcVals(0, a, 0)) <> 0 Then
                                                            If mcVals(3, b, 1) < 100000000 And mcVals(0, a, 1) < 100000000 Then
                                                                XV = (mcVals(3, b, 1) - mcVals(0, a, 1)) / (mcVals(0, a, 0) - mcVals(3, b, 0))
                                                                YV = mcVals(0, a, 0) * XV + mcVals(0, a, 1)
                                                            Else
                                                                If mcVals(3, b, 1) > 100000000 And mcVals(0, a, 1) < 100000000 Then
                                                                    XV = LastCoords(X, b, 0)
                                                                    YV = mcVals(0, a, 0) * XV + mcVals(0, a, 1)
                                                                ElseIf mcVals(3, b, 1) < 100000000 And mcVals(0, a, 1) > 100000000 Then
                                                                    XV = LastCoords(Z, a, 0)
                                                                    YV = mcVals(3, b, 0) * XV + mcVals(3, b, 1)
                                                                Else
                                                                    XV = -1
                                                                    YV = mcVals(0, a, 0) * XV + mcVals(0, a, 1)
                                                                End If
                                                            End If
                                                        'solve y
                                                            
                                                            'is this point of the line in x current
                                                            CollideA = 0
                                                            If LastCoords(X, b, 0) >= XV And LastCoords(X, b + 2, 0) <= XV Then
                                                                If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                    CollideA = 1
                                                                ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                    CollideA = 1
                                                                End If
                                                            ElseIf LastCoords(X, b, 0) <= XV And LastCoords(X, b + 2, 0) >= XV Then
                                                                If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                    CollideA = 1
                                                                ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                    CollideA = 1
                                                                End If
                                                            End If
                                                            CollideB = 0
                                                            If LastCoords(Z, a, 0) >= XV And LastCoords(Z, a, 1) <= XV Then
                                                                If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                    CollideB = 1
                                                                ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                    CollideB = 1
                                                                End If
                                                            ElseIf LastCoords(Z, a, 0) <= XV And LastCoords(Z, a, 1) >= XV Then
                                                                If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                    CollideB = 1
                                                                ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                    CollideB = 1
                                                                End If
                                                            End If
                                                            
                                                            
                                                            If CollideA = 1 And CollideB = 1 Then
                                                                Form1.Picture1.DrawMode = 13
                                                                'Form1.Picture1.Circle (XV, YV), 30, RGB(255, 255, 0)
                                                                'Form1.Picture1.Circle (LastCoords(X, b, 0), LastCoords(X, b + 1, 0)), 20, RGB(0, 255, 0)
                                                                'Form1.Picture1.Circle (LastCoords(X, b + 2, 0), LastCoords(X, b + 3, 0)), 20, RGB(0, 255, 255)
                                                                'Form1.Picture1.Circle (LastCoords(Z, a, 0), LastCoords(Z, a + 1, 0)), 20, RGB(255, 0, 255)
                                                                'Form1.Picture1.Circle (LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1)), 20, RGB(255, 255, 255)
                                                                xx = SpaceObject(Z, 3)
                                                                'work out the angle of incidence of z to x
                                                                'mcvals(1,b,0) = slope of x
                                                                'mcvals(0,a,0) = slope of z
8                                                                If mcVals(3, b, 0) < 10000000 Then
                                                                    A1 = Atn(mcVals(3, b, 0)) * 180 / Pi
                                                                Else
                                                                    A1 = 90
                                                                End If
                                                                
                                                                If mcVals(0, a, 0) < 10000000 Then
                                                                    A2 = Atn(mcVals(0, a, 0)) * 180 / Pi
                                                                Else
                                                                    A2 = 90
                                                                End If
                                                                
                                                                If A1 < 0 Then A1 = Abs(90 + A1) + 90
                                                                If A2 < 0 Then A2 = Abs(90 + A2) + 90
                                                                'angle of incidence = a2-a1
                                                                ai = Abs(A1 - A2)
                                                                'If ai > 90 Then ai = 180 - ai
                                                                'new angle of a = a2-2*ai
                                                                'ar = A2 - (180 + A1)
                                                                'ar = A1 + Ai '2 * (180 - ai)
                                                                'ar = ar - 90
                                                                'now work out whether comming from bottom or top
                                                                
                                                                If A1 >= 90 Then
                                                                    If LastCoords(Z, a + 1, 0) >= LastCoords(Z, a + 1, 1) Then
                                                                        ar = A1 + ai + 180
                                                                    Else
                                                                        ar = A1 + ai
                                                                    End If
                                                                Else
                                                                    If LastCoords(Z, a + 1, 0) >= LastCoords(Z, a + 1, 1) Then
                                                                         ar = A1 - ai + 180
                                                                    Else
                                                                        ar = A1 - ai
                                                                    End If
                                                                
                                                                End If
                                                                'If LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                                '    ar = A1 + ai + 180
                                                                'ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) > LastCoords(Z, a, 1) Then
                                                                '
                                                                '    ar = A1 - ai
                                                                'ElseIf LastCoords(Z, a + 1, 0) < LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                                '    ar = A1 + ai
                                                               '
                                                               ' ElseIf LastCoords(Z, a + 1, 0) > LastCoords(Z, a + 1, 1) And LastCoords(Z, a, 0) < LastCoords(Z, a, 1) Then
                                                               '     ar = A1 - ai + 180
                                                               '
                                                               ' End If
                                                                
                                                                ar = ar - 90 ' puts it back into the game mode -i.e with 0 degrees straight up
                                                                
                                                                If ar >= 360 Then ar = ar - 360
                                                                If ar < 0 Then ar = ar + 360
                                                                xc = XV - LastCoords(Z, a, 0)
                                                                yc = YV - LastCoords(Z, a + 1, 0)
                                                                If xc < 0 Then xc = xc - 5
                                                                If xc > 0 Then xc = xc + 5
                                                                If yc > 0 Then yc = yc + 5
                                                                If yc < 1 Then yc = yc - 5
                                                                
                                                                
                                                                For C = 0 To FinalPoint(Z) - 1 Step 2
                                                                    LastCoords(Z, C, 0) = LastCoords(Z, C, 0) + xc
                                                                    LastCoords(Z, C + 1, 0) = LastCoords(Z, C + 1, 0) + yc
                                                                Next C
                                                                SpaceObject(Z, 1) = SpaceObject(Z, 1) + xc
                                                                SpaceObject(Z, 2) = SpaceObject(Z, 2) + yc
                                                                
                                                                
                                                                
                                                                SpaceObject(Z, 5) = ar
                                                                
                                                                
                                                                For C = 0 To FinalPoint(X) - 1 Step 2
                                                                    LastCoords(X, C, 0) = LastCoords(X, C, 0) - xc
                                                                    LastCoords(X, C + 1, 0) = LastCoords(X, C + 1, 0) - yc
                                                                Next C
                                                                SpaceObject(X, 1) = SpaceObject(X, 1) - xc
                                                                SpaceObject(X, 2) = SpaceObject(X, 2) - yc
                                                                SpaceObject(X, 5) = ar
                                                                
                                                                
                                                                collide = 1
                                                                
                                                                'a = 0
                                                                'b = 0
                                                            End If
                                                            
                                                            
                                                        End If
                                                    
                                                    End If
                                                    
                                                'End If
                                            Next b
                                        'End If
                                    Next a
                                    'do x timelaps vs z current
                                    
                                    
                                    
                                    'do z timelaps vs x current
                                    
                                    
                                    
                                    
                                    
                                    
                                    'If Z > 4 Then
                                    '    SpaceObject(Z, 11) = SpaceObject(Z, 11) - 1
                                   '
                                   ' End If
                                    'get coordinates before and after turn
                                    
                                    'if actual collision
                                      '  If SpaceObject(X, 7) = 0 Then
                                      '      SpaceObject(X, 7) = 1
                                      '  Else
                                      '      SpaceObject(X, 7) = SpaceObject(X, 7) * 1.1
                                      '  End If
                                      
                                      
                                      
                                        If (X = 0 Or Z = 0) And collide = 1 Then
                                            Scores(0) = Scores(0) + 1
                                        End If
                                        If (X = 3 Or Z = 3) And collide = 1 Then
                                            Scores(1) = Scores(1) + 1
                                            
                                        End If
                                        
                                    'end if
                                    If collide = 1 Then
                                        CollideRecord(Z) = CollideRecord(Z) + 1
                                        CollideRecord(X) = CollideRecord(X) + 1
                                    End If
                                    'destroy bullet
                                    If Z >= 50 And Z <= 100 And collide = 1 Then
                                        If CollideRecord(Z) > 5 Then SpaceObject(Z, 0) = 0
                                        
                                        
                                    End If
                                    Form1.Picture1.DrawMode = 13
                                End If
                            End If
                            
                        End If
                    End If
                    'If Z >= 50 And Z <= 100 Then
                    '    SpaceObject(Z, 0) = 0
                    
                End If
            Next Z
            
        End If
    Else
        CollideRecord(X) = CollideRecord(X) + 1
    End If
Next X
For X = 0 To MaxObjects
    If CollideRecord(X) > 10 Then
        CollideRecord(X) = 0
    End If
    
Next X

End Sub
Public Sub DetectCollide3(FinalPoint() As Long)
Dim A1 As Double, A2 As Double, CollideA, CollideB, ProxY(1, 4) As Long, ProxZ(1, 4) As Long, Z As Long, Extremes(1, 4), XOL As Long, YOL As Long, V(1)


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


For X = 0 To MaxObjects
    'If CollideRecord(X) = 0 Or X = X Then
        If SpaceObject(X, 0) = 1 And SpaceObject(X, 8) = 1 Then
            'Exit For
            V(0) = SpaceObject(X, 4)
            
            If Abs(LSpaceObject(X, 1) - SpaceObject(X, 1)) > V(0) Then
                If LSpaceObject(X, 1) > SpaceObject(X, 1) Then
                    LSpaceObject(X, 1) = LSpaceObject(X, 1) - Form1.Picture1.Width
                Else
                    LSpaceObject(X, 1) = Form1.Picture1.Width + LSpaceObject(X, 1)
                End If
            End If
            If Abs(LSpaceObject(X, 2) - SpaceObject(X, 2)) > V(0) Then
                If LSpaceObject(X, 2) > SpaceObject(X, 2) Then
                    LSpaceObject(X, 2) = LSpaceObject(X, 2) - Form1.Picture1.Height
                Else
                    LSpaceObject(X, 2) = Form1.Picture1.Height + LSpaceObject(X, 2)
                End If
            End If
            ProxY(0, 0) = SpaceObject(X, 1) - MDC(X)
            ProxY(0, 1) = SpaceObject(X, 1) + MDC(X)
            ProxY(0, 2) = SpaceObject(X, 2) - MDC(X)
            ProxY(0, 3) = SpaceObject(X, 2) + MDC(X)
            ProxY(1, 0) = LSpaceObject(X, 1) - MDC(X)
            ProxY(1, 1) = LSpaceObject(X, 1) + MDC(X)
            ProxY(1, 2) = LSpaceObject(X, 2) - MDC(X)
            ProxY(1, 3) = LSpaceObject(X, 2) + MDC(X)
            
            
            If ProxY(0, 0) < ProxY(1, 0) Then
                Extremes(0, 0) = ProxY(0, 0)
            Else
                Extremes(0, 0) = ProxY(1, 0)
            End If
            If ProxY(0, 1) > ProxY(1, 1) Then
                Extremes(0, 1) = ProxY(0, 1)
            Else
                Extremes(0, 1) = ProxY(1, 1)
            End If
            If ProxY(0, 2) < ProxY(1, 2) Then
                Extremes(0, 2) = ProxY(0, 2)
            Else
                Extremes(0, 2) = ProxY(1, 2)
            End If
            If ProxY(0, 3) > ProxY(1, 3) Then
                Extremes(0, 3) = ProxY(0, 3)
            Else
                Extremes(0, 3) = ProxY(1, 3)
            End If
            For Z = X + 1 To MaxObjects
                
                    If SpaceObject(X, 9) <> Z + 1 And SpaceObject(Z, 9) <> X + 1 Then
                         
                         If SpaceObject(Z, 0) = 1 And LSpaceObject(Z, 0) = 1 Then
                            V(1) = SpaceObject(Z, 4)
                            If Abs(LSpaceObject(Z, 1) - SpaceObject(Z, 1)) > V(1) Then
                                If LSpaceObject(Z, 1) > SpaceObject(Z, 1) Then
                                    o = LSpaceObject(Z, 1)
                                    LSpaceObject(Z, 1) = LSpaceObject(Z, 1) - Form1.Picture1.Width
                                Else
                                    o = LSpaceObject(Z, 1)
                                    LSpaceObject(Z, 1) = Form1.Picture1.Width + LSpaceObject(Z, 1)
                                End If
                            End If
                            
                            
                            
                            If Abs(LSpaceObject(Z, 2) - SpaceObject(Z, 2)) > V(1) Then
                                If LSpaceObject(Z, 2) > SpaceObject(Z, 2) Then
                                    LSpaceObject(Z, 2) = LSpaceObject(Z, 2) - Form1.Picture1.Height
                                Else
                                    LSpaceObject(Z, 2) = Form1.Picture1.Height + LSpaceObject(Z, 2)
                                End If
                            End If
                            
                            ProxZ(0, 0) = SpaceObject(Z, 1) - MDC(Z)
                            ProxZ(0, 1) = SpaceObject(Z, 1) + MDC(Z)
                            ProxZ(0, 2) = SpaceObject(Z, 2) - MDC(Z)
                            ProxZ(0, 3) = SpaceObject(Z, 2) + MDC(Z)
                            ProxZ(1, 0) = LSpaceObject(Z, 1) - MDC(Z)
                            ProxZ(1, 1) = LSpaceObject(Z, 1) + MDC(Z)
                            ProxZ(1, 2) = LSpaceObject(Z, 2) - MDC(Z)
                            ProxZ(1, 3) = LSpaceObject(Z, 2) + MDC(Z)
                            If ProxZ(0, 0) < ProxZ(1, 0) Then
                                Extremes(1, 0) = ProxZ(0, 0)
                            Else
                                Extremes(1, 0) = ProxZ(1, 0)
                            End If
                            If ProxZ(0, 1) > ProxZ(1, 1) Then
                                Extremes(1, 1) = ProxZ(0, 1)
                            Else
                                Extremes(1, 1) = ProxZ(1, 1)
                            End If
                            If ProxZ(0, 2) < ProxZ(1, 2) Then
                                Extremes(1, 2) = ProxZ(0, 2)
                            Else
                                Extremes(1, 2) = ProxZ(1, 2)
                            End If
                            If ProxZ(0, 3) > ProxZ(1, 3) Then
                                Extremes(1, 3) = ProxZ(0, 3)
                            Else
                                Extremes(1, 3) = ProxZ(1, 3)
                            End If
                            XOL = 0
                            YOL = 0
                            If Extremes(0, 0) >= Extremes(1, 0) And Extremes(0, 0) <= Extremes(1, 1) Then
                                XOL = 1
                            ElseIf Extremes(0, 1) >= Extremes(1, 0) And Extremes(0, 1) <= Extremes(1, 1) Then
                                XOL = 1
                            ElseIf Extremes(0, 0) <= Extremes(1, 0) And Extremes(0, 1) >= Extremes(1, 1) Then
                                XOL = 1
                            End If
                            If XOL = 1 Then
                                If Extremes(0, 2) >= Extremes(1, 2) And Extremes(0, 2) <= Extremes(1, 3) Then
                                    YOL = 1
                                ElseIf Extremes(0, 3) >= Extremes(1, 2) And Extremes(0, 3) <= Extremes(1, 3) Then
                                    YOL = 1
                                ElseIf Extremes(0, 2) <= Extremes(1, 2) And Extremes(0, 3) >= Extremes(1, 3) Then
                                    YOL = 1
                                End If
                                If YOL = 1 Then
                                    'Form1.Picture1.ForeColor = RGB(128, 128, 128)
                                    Form1.Picture1.DrawMode = 9
                                    
                                    'Form1.Picture1.Line (Extremes(0, 0), Extremes(0, 2))-(Extremes(0, 1), Extremes(0, 3)), RGB(128, 128, 128), BF
                                    'Form1.Picture1.Line (Extremes(1, 0), Extremes(1, 2))-(Extremes(1, 1), Extremes(1, 3)), RGB(128, 128, 128), BF
                                    
                                    
                                    Dim LCoord() As Long, Coord() As Long, FinalPos() As Long
                                    
                                    Dim mcVals() As Double
                                    'Test to see if lines in objects cross (1) at this time
                                    '(2) check to see if the lines between last coord and this coord for any points overlap
                                    
                                    
                                    ReDim mcVals(3, Detail, 1)
                                    
                                    
                                    'work out m and c for all timelapse points of x
                                    For a = 0 To FinalPoint(X) - 2 Step 2
                                        cx = LastCoords(X, a, 0) - LastCoords(X, a, 1)
                                        cy = LastCoords(X, a + 1, 0) - LastCoords(X, a + 1, 1)
                                        If cx <> 0 Then
                                            mcVals(0, a, 0) = cy / cx
                                            mcVals(0, a, 1) = LastCoords(X, a + 1, 0) - mcVals(0, a, 0) * LastCoords(X, a, 0)
                                            X = X
                                        Else
                                            mcVals(0, a, 0) = 1000000000
                                            mcVals(0, a, 1) = 1000000000
                                        End If
                                    Next a
                                    'work out m and c for all current points of x
                                    For a = 0 To FinalPoint(X) - 2 Step 2
                                        cx = LastCoords(X, a, 0) - LastCoords(X, a + 2, 0)
                                        cy = LastCoords(X, a + 1, 0) - LastCoords(X, a + 3, 0)
                                        If cx <> 0 Then
                                            mcVals(1, a, 0) = cy / cx
                                            mcVals(1, a, 1) = LastCoords(X, a + 1, 0) - mcVals(1, a, 0) * LastCoords(X, a, 0)
                                        Else
                                            mcVals(1, a, 0) = 1000000000
                                            mcVals(1, a, 1) = 1000000000
                                        End If
                                        X = X
                                    Next a
                                    'work out m and c for all timelapse points of z
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        cx = LastCoords(Z, a, 0) - LastCoords(Z, a, 1)
                                        cy = LastCoords(Z, a + 1, 0) - LastCoords(Z, a + 1, 1)
                                        If cx <> 0 Then
                                            mcVals(2, a, 0) = cy / cx
                                            mcVals(2, a, 1) = LastCoords(Z, a + 1, 0) - mcVals(2, a, 0) * LastCoords(Z, a, 0)
                                        Else
                                            mcVals(2, a, 0) = 1000000000
                                            mcVals(2, a, 1) = 1000000000
                                        End If
                                        X = X
                                    Next a
                                    'work out m and c for all current points of z
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        cx = LastCoords(Z, a, 0) - LastCoords(Z, a + 2, 0)
                                        cy = LastCoords(Z, a + 1, 0) - LastCoords(Z, a + 3, 0)
                                        If cx <> 0 Then
                                            mcVals(3, a, 0) = cy / cx
                                            mcVals(3, a, 1) = LastCoords(Z, a + 1, 0) - mcVals(3, a, 0) * LastCoords(Z, a, 0)
                                        Else
                                            mcVals(3, a, 0) = 1000000000
                                            mcVals(3, a, 1) = 1000000000
                                        End If
                                    Next a
                                    
                                    
                                    'solve equations to see where lines overlap.
                                    
                                    Dim XV, YV
                                    'taplapse z vs current x
                                    collide = 0
                                    GoOn = 0
                                    For a = 0 To FinalPoint(Z) - 2 Step 2
                                        'If LastCoords(Z, a + 3, 0) > 0 And LastCoords(Z, a + 2, 0) > 0 Then
                                            For b = 0 To FinalPoint(X) - 2 Step 2
                                                'If LastCoords(X, b + 3, 0) > 0 And LastCoords(X, b + 2, 0) > 0 Then
                                                    'solve for x
                                                    
                                                    If (mcVals(1, b, 0) - mcVals(2, a, 0)) <> 0 Then
                                                        If mcVals(1, b, 1) < 100000000 And mcVals(2, a, 1) < 100000000 Then
                                                            XV = (mcVals(1, b, 1) - mcVals(2, a, 1)) / (mcVals(2, a, 0) - mcVals(1, b, 0))
                                                            YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                        Else
                                                            If mcVals(1, b, 1) > 100000000 And mcVals(2, a, 1) < 100000000 Then
                                                                XV = LastCoords(X, b, 0)
                                                                YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                            ElseIf mcVals(1, b, 1) < 100000000 And mcVals(2, a, 1) > 100000000 Then
                                                                XV = LastCoords(Z, a, 0)
                                                                YV = mcVals(1, b, 0) * XV + mcVals(1, b, 1)
                                                            Else
                                                                XV = -1
                                                                YV = mcVals(2, a, 0) * XV + mcVals(2, a, 1)
                                                            End If
                                                        End If
                                                    'solve y
                                                        
                                                        'is this point of the line in x current
                                                        CollideA = 0
                                                        If LastCoords(X, b, 0) >= XV And LastCoords(X, b + 2, 0) <= XV Then
                                                            If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                CollideA = 1
                                                            ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                CollideA = 1
                                                            End If
                                                        ElseIf LastCoords(X, b, 0) <= XV And LastCoords(X, b + 2, 0) >= XV Then
                                                            If LastCoords(X, b + 1, 0) >= YV And LastCoords(X, b + 3, 0) <= YV Then
                                                                CollideA = 1
                                                            ElseIf LastCoords(X, b + 1, 0) <= YV And LastCoords(X, b + 3, 0) >= YV Then
                                                                CollideA = 1
                                                            End If
                                                        End If
                                                        CollideB = 0
                                                        If LastCoords(Z, a, 0) >= XV And LastCoords(Z, a, 1) <= XV Then
                                                            If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                CollideB = 1
                                                            ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                CollideB = 1
                                                            End If
                                                        ElseIf LastCoords(Z, a, 0) <= XV And LastCoords(Z, a, 1) >= XV Then
                                                            If LastCoords(Z, a + 1, 0) >= YV And LastCoords(Z, a + 1, 1) <= YV Then
                                                                CollideB = 1
                                                            ElseIf LastCoords(Z, a + 1, 0) <= YV And LastCoords(Z, a + 1, 1) >= YV Then
                                                                CollideB = 1
                                                            End If
                                                        End If
                                                        
                                                        
                                                        If CollideB = 1 And CollideA = 1 Then
                                                            Form1.Picture1.DrawMode = 13
                                                            'Form1.Picture1.Circle (XV, YV), 30, RGB(255, 255, 0)
                                                            'Form1.Picture1.Circle (LastCoords(X, b, 0), LastCoords(X, b + 1, 0)), 20, RGB(0, 255, 0)
                                                            'Form1.Picture1.Circle (LastCoords(X, b + 2, 0), LastCoords(X, b + 3, 0)), 20, RGB(0, 255, 255)
                                                            'Form1.Picture1.Circle (LastCoords(Z, a, 0), LastCoords(Z, a + 1, 0)), 20, RGB(255, 0, 255)
                                                            'Form1.Picture1.Circle (LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1)), 20, RGB(255, 255, 255)
                                                            'treat one object as stationary
                                                            'add x and y velocity to moving object
                                                            'work out angle to the centre of the objects at collision time based on lastcoord z,a+1 and lastcoord z,a
                                                            
                                                            'work out velocity and angle of x
                                                            
                                                            Dim XXVelocity As Double, XYVelocity As Double, ZXVelocity As Double, ZYVelocity As Double
                                                            
                                                            Call MovementCalc(SpaceObject(X, 5), SpaceObject(X, 4), XXVelocity, XYVelocity)
                                                            Call MovementCalc(SpaceObject(Z, 5), SpaceObject(Z, 4), ZXVelocity, ZYVelocity)
                                                            X = X
                                                            
                                                           ' 'Make X stationary
                                                           ' ZXVelocity = ZXVelocity - XXVelocity
                                                           ' ZYVelocity = ZYVelocity - XYVelocity
                                                            
                                                            'work out direction from point of contact to centre of objects
                                                            
                                                            Dim GradX As Double, LenX As Double, GradZ As Double, LenZ As Double
                                                            If CollideA = 1 Then
                                                                Call FindDirection(LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1), SpaceObject(X, 1), SpaceObject(X, 2), LenX, GradX)
                                                                Call FindDirection(LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1), SpaceObject(Z, 1), SpaceObject(Z, 2), LenZ, GradZ)
                                                            ElseIf CollideB = 1 Then
                                                                Call FindDirection(LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1), SpaceObject(X, 1), SpaceObject(X, 2), LenX, GradX)
                                                                Call FindDirection(LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1), SpaceObject(Z, 1), SpaceObject(Z, 2), LenZ, GradZ)
                                                            End If
                                                           
                                                            
                                                            
                                                            'work out momentum
                                                            Dim MomentXX As Double, MomentXY As Double, MomentZX As Double, MomentZY As Double
                                                            
                                                            MomentXX = SpaceObject(X, 6) * XXVelocity
                                                            MomentXY = SpaceObject(X, 6) * XYVelocity
                                                            MomentZX = SpaceObject(Z, 6) * ZXVelocity
                                                            MomentZY = SpaceObject(Z, 6) * ZYVelocity
                                                            
                                                            Dim FinalMomentXX As Double, FinalMomentXY As Double, FinalMomentZX As Double, FinalMomentZY As Double
                                                            'If SpaceObject(X, 6) > SpaceObject(Z, 6) Then
                                                                FinalMomentXX = (MomentXX + MomentZX) / 2 'SpaceObject(X, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6))
                                                                FinalMomentXY = (MomentXY + MomentZY) / 2 '(MomentZY + MomentXY) * (SpaceObject(X, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6)))
                                                            'Else
                                                            
                                                            'End If
                                                            
                                                             'get new x and y momentums
                                                            
                                                            
                                                            
                                                            
                                                            FinalMomentZX = MomentZX + MomentXX - FinalMomentXX '(MomentZX + MomentXX) * (SpaceObject(Z, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6)))
                                                            FinalMomentZY = MomentZY + MomentXY - FinalMomentXY '(MomentZY + MomentXY) * (SpaceObject(Z, 6) / (SpaceObject(X, 6) + SpaceObject(Z, 6)))

                                                            'work out final velocity
                                                            
                                                            Dim FinalVelocityXX As Double, FinalVelocityXY As Double, FinalVelocityZX As Double, FinalVelocityZY As Double
                                                            
                                                            FinalVelocityXX = FinalMomentXX / SpaceObject(X, 6)
                                                            FinalVelocityXY = FinalMomentXY / SpaceObject(X, 6)
                                                            FinalVelocityZX = FinalMomentZX / SpaceObject(Z, 6)
                                                            FinalVelocityZY = FinalMomentZY / SpaceObject(Z, 6)
                                                           
                                                           Dim InitMomentum As Double, FinalMomentum As Double, ForceA As Double, ForceB As Double
                                                            
                                                                ForceA = SpaceObject(Z, 4) * (SpaceObject(X, 6) / SpaceObject(Z, 6))
                                                                ForceB = SpaceObject(X, 4) * (SpaceObject(Z, 6) / SpaceObject(X, 6))
                                                                InitMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                                                
                                                                Call DoThrust(X, GradX, ForceB)
                                                                
                                                                Call DoThrust(Z, GradZ, ForceA)
                                                                FinalMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                                                
                                                                
                                                                ForceA = (((FinalMomentum - InitMomentum) / 2) / SpaceObject(Z, 6))
                                                                If ForceA < 0 Then
                                                                    ForceA = ForceA / 2
                                                               
                                                                End If
                                                                SpaceObject(Z, 4) = SpaceObject(Z, 4) - ForceA
                                                                 ForceA = (((FinalMomentum - InitMomentum) / 2) / SpaceObject(X, 6))
                                                                If ForceA < 0 Then
                                                                    ForceA = ForceA / 2
                                                                
                                                                End If
                                                                SpaceObject(X, 4) = SpaceObject(X, 4) - ForceA
                                                                
                                                                FinalMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                                                
                                                                If FinalMomentum > InitMomentum Then
                                                                    X = X
                                                                End If
                                                                
                                                                X = X
                                                          '  Call DoThrust(Z, GradZ, 10)
                                                           
                                                            
                                                            
                                                            
                                                          '
                                                             
                                                      '      Call MovementCalc(GradX, Sqr(FinalVelocityXX ^ 2 + FinalVelocityXY ^ 2), XXVelocity, XYVelocity)
                                                      '      Call MovementCalc(GradZ, Sqr(FinalVelocityZX ^ 2 + FinalVelocityZY ^ 2), ZXVelocity, ZYVelocity)
                                                      '
                                                            'Add back initial velocities
                                                            'FinalVelocityXX = FinalVelocityXX + XXVelocity
                                                            'FinalVelocityXY = FinalVelocityXY + XYVelocity
                                                            'FinalVelocityZX = FinalVelocityZX + XXVelocity
                                                            'FinalVelocityZY = FinalVelocityZY + XYVelocity
                                                            
                                                     '       Call FindDirection(0, 0, XXVelocity, XYVelocity, LenX, GradX)
                                                     '       Call FindDirection(0, 0, ZXVelocity, ZYVelocity, LenZ, GradZ)
                                                            
                                                            
                                                            
                                                            
                                                            
                                                     '       SpaceObject(X, 4) = LenX
                                                     '       SpaceObject(X, 5) = GradX
                                                     '
                                                     '       SpaceObject(Z, 4) = LenZ
                                                     '       SpaceObject(Z, 5) = GradZ
                                                     '
                                                            SpaceObject(Z, 8) = -20
                                                            SpaceObject(X, 8) = -20
                                                            'goon = 1
                                                            Exit For
                                                        
                                                        End If
                                                    End If
                                                    
                                                    
                                                'End If
                                            Next b
                                            If GoOn = 1 Then Exit For
                                        'End If
                                    Next a
                                    'do x timelaps vs z current
                                    
                                    
                                    
                                    'do z timelaps vs x current
                                    
                                    
                                    
                                    
                                    
                                    
                                    'If Z > 4 Then
                                    '    SpaceObject(Z, 11) = SpaceObject(Z, 11) - 1
                                   '
                                   ' End If
                                    'get coordinates before and after turn
                                    
                                    'if actual collision
                                      '  If SpaceObject(X, 7) = 0 Then
                                      '      SpaceObject(X, 7) = 1
                                      '  Else
                                      '      SpaceObject(X, 7) = SpaceObject(X, 7) * 1.1
                                      '  End If
                                      
                                      
                                      
                                        If (X = 0 Or Z = 0) And collide = 1 Then
                                            Scores(0) = Scores(0) + 1
                                        End If
                                        If (X = 3 Or Z = 3) And collide = 1 Then
                                            Scores(1) = Scores(1) + 1
                                            
                                        End If
                                        
                                    'end if
                                    If collide = 1 Then
                                        CollideRecord(Z) = CollideRecord(Z) + 1
                                        CollideRecord(X) = CollideRecord(X) + 1
                                    End If
                                    'destroy bullet
                                    If Z >= 50 And Z <= 100 And collide = 1 Then
                                        If CollideRecord(Z) > 5 Then SpaceObject(Z, 0) = 0
                                        
                                        
                                    End If
                                    Form1.Picture1.DrawMode = 13
                                End If
                            End If
                            
                        End If
                    End If
                    'If Z >= 50 And Z <= 100 Then
                    '    SpaceObject(Z, 0) = 0
                    
                Next Z
            
        End If
    'Else
    '    CollideRecord(X) = CollideRecord(X) + 1
    'End If
Next X
For X = 0 To MaxObjects
    If CollideRecord(X) > 10 Then
        CollideRecord(X) = 0
    End If
    
Next X

End Sub
Public Sub DetectCollide4(FinalPoint() As Long)
Dim A1 As Double, A2 As Double, CollideA, CollideB, ProxY(1, 4) As Long, ProxZ(1, 4) As Long, Z As Long, Extremes(1, 4), XOL As Long, YOL As Long, V(1)


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
Dim ScreenModXX As Double, ScreenModXY As Double, ScreenModZX As Double, ScreenModZY As Double, bModX As Double, bModY As Double, tModX As Double, tModY As Double, D1 As Double, D2 As Double, D3 As Double, D4 As Double, XXOff As Double, XYOff As Double, ZXOff As Double, ZYOff As Double, OpX As Double, OpY As Double, GoOn As Byte, DistBetween As Double, MindistX As Long, MindistX2 As Long, MindistZ As Long, MindistZ2 As Long, CoordMindist(1, 1) As Double, CoordMindist2(1, 1) As Double

For X = 0 To 50

    
    If SpaceObject(X, 0) = 1 And SpaceObject(X, 8) = 1 Then
        
        Call MovementCalc(SpaceObject(X, 5), SpaceObject(X, 4) * VelocityMod, XXOff, XYOff)
        For Z = X + 1 To MaxObjects
            
            If SpaceObject(Z, 0) = 1 And SpaceObject(Z, 8) = 1 Then
               
                Call MovementCalc(SpaceObject(Z, 5), SpaceObject(Z, 4) * VelocityMod, ZXOff, ZYOff)
                
                If SpaceObject(Z, 6) = 1 Or X = X Then
                    bModX = ZXOff 'bullet mods
                    bModY = ZYOff
                    tModX = XXOff
                    tModY = XYOff
                Else
                    bModX = 0 'no bullet mods
                    bModY = 0
                    tModX = 0
                    tModY = 0
                End If
                
                GoOn = 0
                Dim DistXZ As Double, AngXZ As Double, AngXZ2 As Double
                ScreenModXX = 0
                ScreenModXY = 0
                ScreenModZX = 0
                ScreenModZY = 0
                
                
                If (SpaceObject(X, 1) - SpaceObject(Z, 1)) > frmGame.Picture1.Width / 2 Then
                    ScreenModXX = -frmGame.Picture1.Width
                ElseIf (SpaceObject(Z, 1) - SpaceObject(X, 1)) > frmGame.Picture1.Width / 2 Then
                    ScreenModZX = -frmGame.Picture1.Width
                End If
                
                If (SpaceObject(X, 2) - SpaceObject(Z, 2)) > frmGame.Picture1.Height / 2 Then
                    ScreenModXY = -frmGame.Picture1.Height
                ElseIf (SpaceObject(Z, 2) - SpaceObject(X, 2)) > frmGame.Picture1.Height / 2 Then
                    ScreenModZY = -frmGame.Picture1.Height
                End If
                If X = 0 And Z = 5 Then
                    X = X
                End If
                Call FindDirection(SpaceObject(X, 1) + tModX + ScreenModXX, SpaceObject(X, 2) + tModY + ScreenModXY, SpaceObject(Z, 1) + bModX + ScreenModZX, (SpaceObject(Z, 2) + bModY + ScreenModZY), DistXZ, AngXZ)
                
                
                GoOn = 0
                If X = X Then
                    If DistXZ <= MaxProx2D(X, Z) Then
                        
                        GoOn = 1
                    ElseIf DistXZ < MaxProx2D(X, Z) Then
                        If SpaceObject(X, 4) / 2 > MinProx(Z) * 2 Or SpaceObject(Z, 4) / 2 > MinProx(X) * 2 Then  'what happens if the object is small enough for the bullet to jumpright over it
                            Call FindDirection(SpaceObject(X, 1) + ScreenModXX, SpaceObject(X, 2) + ScreenModXY, SpaceObject(Z, 1) + ScreenModZX, SpaceObject(Z, 2) + ScreenModZY, DistXZ, AngXZ2)
                            If Abs(AngXZ - AngXZ2) > 140 And Abs(AngXZ - AngXZ2) < 220 Then
                                GoOn = 1
                            End If
                        
                        End If
                    End If
                    
                Else
                    'this is never executed
                    If Abs(SpaceObject(X, 1) + tModX - (SpaceObject(Z, 1) + bModX)) < MaxProx2D(X, Z) Then 'it is close enough in the x-dim
                        GoOn = 1
                        OpX = 0
                    ElseIf Abs(SpaceObject(X, 1) + tModX - (SpaceObject(Z, 1) + bModX)) > Form1.Picture1.Width - MaxProx2D(X, Z) Then 'check to see if they are on opposite sides of the screen
                    
                        GoOn = 1
                        OpX = -Form1.Picture1.Width
                    End If
                End If
                If GoOn = 1 Then
                    If X = 12345 Then
                        GoOn = 0
                        'this is also never execurted
                        If Abs(SpaceObject(X, 2) + tModY + ScreenModXY - (SpaceObject(Z, 2) + bModY + ScreenModZY)) < MaxProx2D(X, Z) Then 'it is close enough in the y-dim
                            GoOn = 1
                            OpY = 0
                        ElseIf Abs(SpaceObject(X, 2) + tModY + ScreenModXY - (SpaceObject(Z, 2) + bModY + ScreenModZY)) > Form1.Picture1.Height - MaxProx2D(X, Z) Then 'check to see if they are on opposite sides of the screen
                            OpY = -Form1.Picture1.Height
                            GoOn = 1
                        End If
                    End If
                    If GoOn = 1 Then
                        'Sleep 500
                        'Form1.Picture1.Circle (SpaceObject(Z, 1), SpaceObject(Z, 2)), 40, RGB(255, 0, 0)
                        'Form1.Picture1.Circle (SpaceObject(Z, 1) + ZXOff, SpaceObject(Z, 2) + ZYOff), 40, RGB(255, 255, 0)
                        If SpaceObject(Z, 6) = 1 Then ' when you are dealing with a bullet
                            
                            GoOn = 0
                            Dim Point1X As Double, Point2X As Double, Point1Y As Double, Point2Y As Double, IncomingAngle1 As Double, IncomingAngle2 As Double, LenBB2 As Double, LenBB1 As Double, XXNextCoord As Double, XYNextCoord As Double, ZXNextCoord As Double, ZYNextCoord As Double
                            'Call MovementCalc(SpaceObject(X, 5), SpaceObject(X, 4) * VelocityMod, XXNextCoord, XYNextCoord)
                            'Call MovementCalc(SpaceObject(Z, 5), SpaceObject(Z, 4) * VelocityMod, ZXNextCoord, ZYNextCoord)
                            
                            ZXNextCoord = ZXOff
                            ZYNextCoord = ZYOff
                            XXNextCoord = XXOff
                            XYNextCoord = XYOff
                            
                            
                            'find the points on the surface of X with angles encompassing the incoming direction of Z
                            'spaceobject(z,5)= incomming angle of bullet
                            ' the difference between znextcoord, and spaceobject(z,1/2) is the track of the bullet in the next move
                            Call FindDirection(SpaceObject(X, 1) + XXOff, SpaceObject(X, 2) + XYOff, SpaceObject(Z, 1) + ZXOff, SpaceObject(Z, 2) + ZYOff, LenBB1, IncomingAngle1)
                            Call FindDirection(SpaceObject(X, 1), SpaceObject(X, 2), SpaceObject(Z, 1), SpaceObject(Z, 2), LenBB2, IncomingAngle2)
                            If LenBB2 < MinProx(X) Or (LenBB1 < MinProx(X) And (X = 0 Or X = 3)) Or X = 0 Or X = 3 Then
                                'do collide and dissapear
                               ' Form1.Picture1.Circle (SpaceObject(Z, 1), SpaceObject(Z, 2)), 40, RGB(255, 0, 0)
                               ' Form1.Picture1.Circle (SpaceObject(Z, 1) + ZXOff, SpaceObject(Z, 2) + ZYOff), 40, RGB(255, 255, 0)
                                'Form1.Picture1.PSet (SpaceObject(Z, 1), SpaceObject(Z, 2)), RGB(0, 0, 0)
                                
                                
                                'Form1.Picture1.Line (SpaceObject(Z, 1), SpaceObject(Z, 2))-(SOLast(Z, 0), SOLast(Z, 1)), RGB(0, 0, 0)
                                
                                
                                SpaceObject(Z, 0) = 0
                                SpaceObject(Z, 1) = 0
                                SpaceObject(Z, 2) = 0
                                
                                If X = 0 Then
                                    Scores(1) = Scores(1) + 1
                                ElseIf X = 3 Then
                                    Scores(0) = Scores(0) + 1
                                End If
                                
                                
                                IncomingAngle2 = IncomingAngle2 + 180
                                If IncomingAngle2 > 359 Then IncomingAngle2 = IncomingAngle2 - 359
                                If X = X Then ' reinstate to reintroduce bullet impacts
                                    Call DoThrust(X, IncomingAngle2, SpaceObject(Z, 4) * (1 / SpaceObject(X, 6)))
                                End If
                                
                                
                                'Sleep 1000
                                'Form1.Picture1.PSet (SpaceObject(Z, 1), SpaceObject(Z, 2)), RGB(255, 0, 0)
                                'Form1.Picture1.Circle (SpaceObject(Z, 1), SpaceObject(Z, 2)), 40, RGB(255, 0, 0)
                                'maybe make the object that got hit slightly heavier or effcet its movement here
                                'X = X
                                'Sleep 500
                            Else 'If LenBB1 < MaxProx(X) Then
                                'Sleep 2000
                            'Form1.Picture1.AutoRedraw = True
                               ' Sleep 1000
                               ' Form1.Picture1.Circle (SpaceObject(Z, 1), SpaceObject(Z, 2)), 40, RGB(255, 255, 0)
                               ' Sleep 1000
                                'Form1.Picture1.AutoRedraw = False
                                Call FindDirection(SpaceObject(X, 1) + ScreenModXX + XXNextCoord, SpaceObject(X, 2) + ScreenModXY + XYNextCoord, SpaceObject(Z, 1) + ZXNextCoord + ScreenModZX, SpaceObject(Z, 2) + ScreenModZY + ZYNextCoord, LenBB2, IncomingAngle2)
                                'SpaceObject(Z, 5) + 180 'to get the angle relative to the angles in object x
                                If IncomingAngle1 > 359 Then IncomingAngle1 = IncomingAngle1 - 359
                                For a = ShapeStart + 2 To Detail Step 2 'the angle info starts at 10 and moves clockwise
                                    'have to make sure that it wraps
                                    If SpaceObject(X, a) > -1 Then
                                        If (SpaceObject(X, a - 2) + SpaceObject(X, 3)) <= IncomingAngle1 And (SpaceObject(X, a) + SpaceObject(X, 3)) >= IncomingAngle1 Then
                                            Dim LenX As Double, AX As Double, AY As Double, BX As Double, BY As Double
                                            If (SpaceObject(X, a - 2) + SpaceObject(X, 3)) >= IncomingAngle2 Then 'the object could hit to the left
                                                Point1X = DrawObject(X, a - ShapeStart - 2)
                                                Point1Y = DrawObject(X, a - ShapeStart - 1)
                                                If DrawObject(X, a - ShapeStart + 2) <> 0 Then
                                                    Point2X = DrawObject(X, a - ShapeStart + 2)
                                                    Point2Y = DrawObject(X, a - ShapeStart + 3)
                                                Else
                                                    Point2X = DrawObject(X, ShapeStart)
                                                    Point2Y = DrawObject(X, ShapeStart + 1)
                                                End If
                                                
                                                
                                            ElseIf (SpaceObject(X, a) + SpaceObject(X, 3)) <= IncomingAngle2 Then 'the object could hit to the right
                                                
                                                
                                                If a >= ShapeStart + 4 Then
                                                    Point1X = DrawObject(X, a - ShapeStart - 4)
                                                    Point1Y = DrawObject(X, a - ShapeStart - 3)
                                                Else
                                                    For b = Detail To ShapeStart Step -1
                                                        If DrawObject(X, b) > 0 Then
                                                            X = X
                                                            Point1Y = DrawObject(X, b)
                                                            Point1X = DrawObject(X, b - 1)
                                                            Exit For
                                                        End If
                                                    Next b
                                                End If
                                                Point2X = DrawObject(X, a - ShapeStart)
                                                Point2Y = DrawObject(X, a - ShapeStart + 1)
                                            Else
                                            
                                                
                                                Point1X = DrawObject(X, a - ShapeStart - 2)
                                                Point1Y = DrawObject(X, a - ShapeStart - 1)
                                                Point2X = DrawObject(X, a - ShapeStart)
                                                Point2Y = DrawObject(X, a - ShapeStart + 1)
                                                'Call FindDirection(SpaceObject(X, 1), SpaceObject(X, 2), SpaceObject(Z, 1), SpaceObject(Z, 2), LenX, GradX)
                                                'X = X
                                            End If
                                            
                                            Dim XInterceptEarly As Double, XInterceptLate As Double, YInterceptEarly As Double, YInterceptLate As Double
                                            Dim XCValEarly As Double, XMValEarly As Double, XCValLate As Double, XMValLate As Double, ZCVal As Double, ZMVal As Double
                                            If (Point1X - Point2X) <> 0 Then
                                                XMValEarly = (Point1Y - Point2Y) / (Point1X - Point2X)
                                            Else
                                                Point2X = Point2X + 1
                                                XMValEarly = (Point1Y - Point2Y) / (Point1X - Point2X)
                                            End If
                                            XMValLate = (Point1Y + XYNextCoord - Point2Y + XYNextCoord) / (Point1X + XXNextCoord - Point2X + XXNextCoord)
                                            
                                            XCValEarly = (Point2X * Point1Y - Point1X * Point2Y) / (Point2X - Point1X)
                                            XCValLate = ((Point2X + XXNextCoord) * (Point1Y + XYNextCoord) - (Point1X + XXNextCoord) * (Point2Y + XYNextCoord)) / ((Point2X + XXNextCoord) - (Point1X + XXNextCoord))
                                            
                                            
                                            If ZXNextCoord <> 0 And ZYNextCoord <> 0 Then
                                                ZMVal = (SpaceObject(Z, 2) - (SpaceObject(Z, 2) + ZYNextCoord)) / (SpaceObject(Z, 1) - (SpaceObject(Z, 1) + ZXNextCoord))
                                                ZCVal = ((SpaceObject(Z, 1) + ZXNextCoord) * SpaceObject(Z, 2) - SpaceObject(Z, 1) * (SpaceObject(Z, 2) + ZYNextCoord)) / ((SpaceObject(Z, 1) + ZXNextCoord) - SpaceObject(Z, 1))
                                            
                                                XInterceptEarly = (XCValEarly - ZCVal) / (ZMVal - XMValEarly)
                                                YInterceptEarly = (ZCVal * XMValEarly - XCValEarly * ZMVal) / (XMValEarly - ZMVal)
                                                
                                                XInterceptLate = (XCValLate - ZCVal) / (ZMVal - XMValLate)
                                                YInterceptLate = (ZCVal * XMValLate - XCValLate * ZMVal) / (XMValLate - ZMVal)
                                                X = X
                                            ElseIf ZXNextCoord = 0 Then
                                                XInterceptEarly = SpaceObject(Z, 1)
                                                XInterceptLate = SpaceObject(Z, 1) + ZXNextCoord
                                                YInterceptEarly = XMValEarly * XInterceptEarly + XCValEarly
                                                YInterceptLate = XMValLate * XInterceptLate + XCValLate
                                                'ZMVal = (SpaceObject(Z, 2) - (SpaceObject(Z, 2) + ZYNextCoord)) / (SpaceObject(Z, 1) - (SpaceObject(Z, 1) + 0.0000000001))
                                                'ZCVal = ((SpaceObject(Z, 1) + 0.0000000001) * SpaceObject(Z, 2) - SpaceObject(Z, 1) * (SpaceObject(Z, 2) + ZYNextCoord)) / ((SpaceObject(Z, 1) + 0.000000000000001) - SpaceObject(Z, 1))
                                            Else
                                                YInterceptEarly = SpaceObject(Z, 2)
                                                YInterceptLate = SpaceObject(Z, 2) + ZYNextCoord
                                                XInterceptEarly = (YInterceptEarly - XCValEarly) / XMValEarly
                                                XInterceptLate = (YInterceptLate - XCValLate) / XMValLate
                                            End If
                                            
                                            'has the bullet actually hit the expected panel
                                            Dim GoOn2 As Byte
                                            'Sleep 1000
                                            GoOn = 0: GoOn2 = 0
                                            If Point1X < Point2X Then
                                                If XInterceptEarly >= Point1X - 1 And XInterceptEarly <= Point2X + 1 Then
                                                    GoOn = 1
                                                ElseIf XInterceptLate >= Point1X - 1 + XXNextCoord And XInterceptLate <= Point2X + 1 + XXNextCoord Then
                                                    GoOn = 1
                                                End If
                                            Else
                                                If XInterceptEarly >= Point2X - 1 And XInterceptEarly <= Point1X + 1 Then
                                                    GoOn = 1
                                                ElseIf XInterceptLate >= Point2X - 1 + XXNextCoord And XInterceptLate <= Point1X + 1 + XXNextCoord Then
                                                    GoOn = 1
                                                End If
                                            End If
                                            
                                            If GoOn = 1 Then
                                                If Point1Y <= Point2Y Then
                                                    If YInterceptEarly >= Point1Y - 1 And YInterceptEarly <= Point2Y + 1 Then
                                                        GoOn2 = 1
                                                    ElseIf YInterceptLate >= Point1Y - 1 + XYNextCoord And YInterceptLate <= Point2Y + 1 + XYNextCoord Then
                                                        GoOn2 = 1
                                                    End If
                                                Else
                                                    If YInterceptEarly >= Point2Y - 1 And YInterceptEarly <= Point1Y + 1 Then
                                                        GoOn2 = 1
                                                    ElseIf YInterceptLate >= Point2Y - 1 + XYNextCoord And YInterceptLate <= Point1Y + 1 + XYNextCoord Then
                                                        GoOn2 = 1
                                                    End If
                                                End If
                                            End If
                                            
                                            xx = X
                                            
                                            'find out exact
                                            If GoOn2 = 1 And GoOn = 1 Then
                                                If X = 0 Then
                                                    Scores(1) = Scores(1) + 1
                                                ElseIf X = 3 Then
                                                    Scores(0) = Scores(0) + 1
                                                End If
                                
                                                'work out point of deflection and rebound position
                                                Dim XLineDir As Double, ZLen As Double, ZDir As Double, AngleDiff As Double, ZLineDir As Double
                                                If Point1X > Point2X Then
                                                    Call FindDirection(Point2X, Point2Y, Point1X, Point1Y, LenX, XLineDir)
                                                Else
                                                    Call FindDirection(Point1X, Point1Y, Point2X, Point2Y, LenX, XLineDir)
                                                End If
                                                ZLen = SpaceObject(Z, 4)
                                                ZDir = SpaceObject(Z, 5)
                                                
                                                X = X
                                                'still need need to work out exact interception point for moving objects
                                                
                                                'Call FindDirection(Point1X, Point1Y, Point2X, Point2Y, ZLen, ZDir)
                                                'clockwise difference in angles
                                                
                                                AngleDiff = ZDir - XLineDir
                                                XLineDir = XLineDir + 180
                                                ZLineDir = XLineDir - AngleDiff
                                                
                                                If ZLineDir > 359 Then ZLineDir = ZLineDir - 359
                                                
                                                'for now I'm just going to use xyinterceptearly
                                                Dim TempV As Double, XOffZ As Double, YOffZ As Double
                                                'tempv is the distance from currentz pos to pint of impact
                                                Call FindDirection(XInterceptEarly, YInterceptEarly, SpaceObject(Z, 1), SpaceObject(Z, 2), TempV, XLineDir)
                                                Call MovementCalc(ZLineDir, TempV * VelocityMod, XOffZ, YOffZ)
                                                
                                                'this puts the bullet inside the asteroid
                                               ' SpaceObject(Z, 1) = XInterceptEarly + XOffZ
                                               ' SpaceObject(Z, 2) = YInterceptEarly + YOffZ
                                                
                                                
                                                'this puts the bullet at the surface of the asteroid
                                                SpaceObject(Z, 1) = XInterceptEarly
                                                SpaceObject(Z, 2) = YInterceptEarly
                                                
                                                ZLineDir = ZLineDir + 180
                                                If ZLineDir > 359 Then ZLineDir = ZLineDir - 359
                                                SpaceObject(Z, 5) = ZLineDir
                                                SpaceObject(Z, 8) = -2
                                                
                                                
                                                
                                                
                                                
                                                'this removes the bullet if it penetrates too deep
                                                Call FindDirection(SpaceObject(Z, 1), SpaceObject(Z, 2), SpaceObject(X, 1), SpaceObject(X, 2), TempV, XLineDir)
                                                If X = X Then ' reinstate to reintroduce bullet impacts
                                                    Call DoThrust(X, XLineDir, SpaceObject(Z, 4) * (1 / SpaceObject(X, 6)))
                                                    SpaceObject(Z, 4) = SpaceObject(Z, 4) - 5
                                                End If
                                                
                                                 If TempV <= MinProx(X) Then
                                                    SpaceObject(Z, 0) = 0
                                                    SpaceObject(Z, 1) = 0
                                                    SpaceObject(Z, 2) = 0
                                                    
                                                End If
                                                
                                                Call MovementCalc(SpaceObject(X, 5), SpaceObject(X, 4) * VelocityMod, XXNextCoord, XYNextCoord)
                                                Call MovementCalc(SpaceObject(Z, 5), SpaceObject(Z, 4) * VelocityMod, ZXNextCoord, ZYNextCoord)
                                                
                                                Call FindDirection(SpaceObject(Z, 1) + ZXNextCoord, SpaceObject(Z, 2) + ZYNextCoord, SpaceObject(X, 1) + XXNextCoord, SpaceObject(X, 2) + XYNextCoord, TempV, XLineDir)
                                                
                                                If X = X Then
                                                    
                                                    
                                                    If TempV <= MinProx(X) Then
                                                       SpaceObject(Z, 0) = 0
                                                        SpaceObject(Z, 1) = 0
                                                        SpaceObject(Z, 2) = 0
                                                        
                                                    End If
                                                End If
                                                'get 2 points on x
                                                
                                                
                                                'work out point of contact
                                            End If
                                            Exit For
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next a
                            End If
                            
                        Else
                        
                        
                            MindistX = 1000000
                            MindistZ = 1000000
                        
                            For a = 0 To FinalPoint(X) Step 2
                                For b = 0 To FinalPoint(Z) Step 2
                                    DistBetween = Sqr(Abs((DrawObject(X, a) - DrawObject(Z, b)) + OpX) ^ 2 + (Abs(DrawObject(X, a + 1) - DrawObject(Z, b + 1)) + OpY) ^ 2)
                                    'D2 = Sqr(Abs((DrawObject(X, A) - DrawObject(Z, B)) + OpX + ZXOff) ^ 2 + (Abs(DrawObject(X, A + 1) - DrawObject(Z, B + 1)) + OpY + ZYOff) ^ 2)
                                    'D3 = Sqr(Abs((DrawObject(X, A) + XXOff - DrawObject(Z, B)) + OpX) ^ 2 + (Abs(DrawObject(X, A + 1) + XYOff - DrawObject(Z, B + 1)) + OpY) ^ 2)
                                    D2 = Sqr(Abs((DrawObject(X, a) + XXOff - DrawObject(Z, b)) + OpX + ZXOff) ^ 2 + (Abs(DrawObject(X, a + 1) + XYOff - DrawObject(Z, b + 1)) + OpY + ZYOff) ^ 2)
                                    If DistBetween > D2 Then
                                        DistBetween = D2
                                    End If
                                    
                                    If DistBetween < MindistX Then
                                         MindistX2 = MindistX
                                         CoordMindist2(0, 0) = CoordMindist(0, 0)
                                         CoordMindist2(0, 1) = CoordMindist(0, 1)
                                         
                                         
                                         MindistX = DistBetween
                                         CoordMindist(0, 0) = DrawObject(X, a)
                                         CoordMindist(0, 1) = DrawObject(X, a + 1)
                                         
                                    ElseIf DistBetween < MindistX2 Then
                                         MindistX2 = DistBetween
                                         CoordMindist2(0, 0) = DrawObject(X, a)
                                         CoordMindist2(0, 1) = DrawObject(X, a + 1)
                                         
                                    End If
                                    If DistBetween < MindistZ Then
                                         MindistZ2 = MindistZ
                                         
                                         CoordMindist2(1, 0) = CoordMindist(1, 0)
                                         CoordMindist2(1, 1) = CoordMindist(1, 1)
                                         
                                         MindistZ = DistBetween
                                         
                                         CoordMindist(1, 0) = DrawObject(Z, b)
                                         CoordMindist(1, 1) = DrawObject(Z, b + 1)
                                    ElseIf DistBetween < MindistZ2 Then
                                         MindistZ2 = DistBetween
                                        
                                         CoordMindist2(1, 0) = DrawObject(Z, b)
                                         CoordMindist2(1, 1) = DrawObject(Z, b + 1)
                                    End If
                                Next b
                            Next a
                            'check and see if the mincoords of z are closer to x than the mincoord of x (collidea)
                            CollideA = 0
                            CollideB = 0
                            CollideC = 0
                            'If X = X Then
                                'what is the angle of the closest point in object x to z
                                Dim Ang1 As Double, Ang2 As Double, Ang3 As Double, Len1 As Double, Len2 As Double, Len3 As Double
                                Call FindDirection(SpaceObject(Z, 1) + ZXOff + ScreenModZX, SpaceObject(Z, 2) + ZYOff + ScreenModZY, CoordMindist(0, 0) + XXOff, CoordMindist(0, 1) + XYOff, Len1, Ang1)
                                'find the two points on object z that encompass ang1
                                For a = 0 To FinalPoint(Z) Step 2
                                    Ang3 = SpaceObject(Z, a + 10) + SpaceObject(Z, 3)
                                    
                                    If Ang3 >= Ang1 Then
                                        Len3 = SpaceObject(Z, a + 11)
                                    
                                        If a > 0 Then
                                            Ang2 = SpaceObject(Z, a + 8) + SpaceObject(Z, 3)
                                            Len2 = SpaceObject(Z, a + 9) + SpaceObject(Z, 3)
                                        Else
                                            Ang2 = SpaceObject(Z, 10 + FinalPoint(Z)) + SpaceObject(Z, 3)
                                            Len2 = SpaceObject(Z, FinalPoint(Z) + 11) + SpaceObject(Z, 3)
                                        End If
                                        Exit For
                                    End If
                                    
                                Next a
                                If GoOn = 1 Then
                                    If Len1 <= Len2 Or Len1 <= Len3 Then
                                        CollideC = 1
                                        
                                    
                                    End If
                                End If
                            'Else 'old collide detection works OK but with glitches
                            
                                If ((CoordMindist(0, 0) - (SpaceObject(Z, 1) + ScreenModZX)) ^ 2 + ((CoordMindist(0, 1) - (SpaceObject(Z, 2) + ScreenModZY)) ^ 2)) <= (((CoordMindist(1, 0) - (SpaceObject(Z, 1) + ScreenModZX)) ^ 2 + (CoordMindist(1, 1) - (SpaceObject(Z, 2) + ScreenModZY)) ^ 2)) Then
                                    CollideA = 1
                                ElseIf OpX <> 0 Or OpY <> 0 Then
                                    If Abs(CoordMindist(0, 0) - (SpaceObject(Z, 1) + ScreenModZX)) < Abs(Abs(CoordMindist(0, 0) - (SpaceObject(Z, 1) + ScreenModZY)) + OpX) Then
                                        V1 = Abs(CoordMindist(0, 0) - (SpaceObject(Z, 1) + ScreenModZX))
                                    Else
                                        V1 = Abs(Abs(CoordMindist(0, 0) - (SpaceObject(Z, 1) + ScreenModZX)) + OpX)
                                    End If
                                    If Abs((CoordMindist(0, 1) - SpaceObject(Z, 2))) < Abs(Abs(CoordMindist(0, 1) - SpaceObject(Z, 2)) + OpY) Then
                                        V2 = Abs((CoordMindist(0, 1) - (SpaceObject(Z, 2) + ScreenModZY)))
                                    Else
                                        V2 = Abs(Abs(CoordMindist(0, 1) - SpaceObject(Z, 2)) + OpY)
                                    End If
                                    If Abs(CoordMindist(1, 0) - SpaceObject(Z, 1)) < Abs(Abs(CoordMindist(1, 0) - SpaceObject(Z, 1)) + OpX) Then
                                        V3 = Abs(CoordMindist(1, 0) - SpaceObject(Z, 1))
                                    Else
                                        V3 = Abs((Abs(CoordMindist(1, 0) - SpaceObject(Z, 1)) + OpX))
                                    End If
                                    If Abs(CoordMindist(1, 1) - SpaceObject(Z, 2)) < Abs(Abs(CoordMindist(1, 1) - SpaceObject(Z, 2) + OpY)) Then
                                        V4 = Abs(CoordMindist(1, 1) - SpaceObject(Z, 2))
                                    Else
                                        V4 = Abs(Abs(CoordMindist(1, 1) - SpaceObject(Z, 2) + OpY))
                                    End If
                                    If V1 ^ 2 + V2 ^ 2 <= V3 ^ 2 + V4 ^ 2 Then
                                        CollideA = 1
                                    End If
                                End If
                            
                            
                            
                            
                            
                                If (CoordMindist(1, 0) - (SpaceObject(X, 1) + ScreenModXX)) ^ 2 + (CoordMindist(1, 1) - (SpaceObject(X, 2) + ScreenModXY)) ^ 2 <= (CoordMindist(0, 0) - (SpaceObject(X, 1) + ScreenModXX)) ^ 2 + (CoordMindist(0, 1) - (SpaceObject(X, 2) + ScreenModXY)) ^ 2 Then
                                    CollideB = 1
                                ElseIf OpY <> 0 Or OpX <> 0 Then
                                    If Abs(CoordMindist(1, 0) - SpaceObject(X, 1)) < Abs(Abs(CoordMindist(1, 0) - SpaceObject(X, 1)) + OpX) Then
                                        V1 = Abs(CoordMindist(1, 0) - SpaceObject(X, 1))
                                    Else
                                        V1 = Abs(Abs(CoordMindist(1, 0) - SpaceObject(X, 1)) + OpX)
                                    End If
                                    
                                    If Abs(CoordMindist(1, 1) - SpaceObject(X, 2)) < Abs(Abs(CoordMindist(1, 1) - SpaceObject(X, 2)) + OpY) Then
                                        V2 = Abs(CoordMindist(1, 1) - SpaceObject(X, 2))
                                    Else
                                        V2 = Abs(Abs(CoordMindist(1, 1) - SpaceObject(X, 2)) + OpY)
                                    End If
                                    
                                    If Abs(CoordMindist(0, 0) - SpaceObject(X, 1)) < Abs(Abs(CoordMindist(0, 0) - SpaceObject(X, 1)) + OpX) Then
                                        V3 = Abs(CoordMindist(0, 0) - SpaceObject(X, 1))
                                    Else
                                        V3 = Abs(Abs(CoordMindist(0, 0) - SpaceObject(X, 1)) + OpX)
                                    End If
                                    
                                    If Abs(CoordMindist(0, 1) - SpaceObject(X, 2)) < Abs(Abs(CoordMindist(0, 1) - SpaceObject(X, 2)) + OpY) Then
                                        V4 = Abs(CoordMindist(0, 1) - SpaceObject(X, 2))
                                    Else
                                        V4 = Abs(Abs(CoordMindist(0, 1) - SpaceObject(X, 2)) + OpY)
                                    End If
                                    
                                    If V1 ^ 2 + V2 ^ 2 <= V3 ^ 2 + V4 ^ 2 Then
                                        CollideB = 1
                                    End If
                                End If
                            'End If
                            If CollideB = 1 Or CollideA = 1 Or CollideC = 1 Or SpaceObject(Z, 6) = 1 Then
                               'averaging distances between the nearest 2 points
                               
                              ' Do
                                   'DoEvents
                                   'Sleep 100
                                '   If ExitE = 1 Then
                                '       ExitE = 0
                                '       Exit Do
                                '  End If
                               'Loop
                               
                               
                              ' CoordMindist(0, 0) = (CoordMindist(0, 0) + CoordMindist2(0, 0)) / 2
                              ' CoordMindist(0, 1) = (CoordMindist(0, 1) + CoordMindist2(0, 1)) / 2
                               CoordMindist(1, 0) = (CoordMindist(1, 0) + CoordMindist2(1, 0)) / 2
                               CoordMindist(1, 1) = (CoordMindist(1, 1) + CoordMindist2(1, 1)) / 2
                               
                               frmGame.Picture1.DrawMode = 13
                               'frmGame.Picture1.Circle (XV, YV), 30, RGB(255, 255, 0)
                               'frmGame.Picture1.Circle (LastCoords(X, b, 0), LastCoords(X, b + 1, 0)), 20, RGB(0, 255, 0)
                               'frmGame.Picture1.Circle (LastCoords(X, b + 2, 0), LastCoords(X, b + 3, 0)), 20, RGB(0, 255, 255)
                               'frmGame.Picture1.Circle (LastCoords(Z, a, 0), LastCoords(Z, a + 1, 0)), 20, RGB(255, 0, 255)
                               'frmGame.Picture1.Circle (LastCoords(Z, a, 1), LastCoords(Z, a + 1, 1)), 20, RGB(255, 255, 255)
                               'treat one object as stationary
                               'add x and y velocity to moving object
                               'work out angle to the centre of the objects at collision time based on lastcoord z,a+1 and lastcoord z,a
                               
                               'work out velocity and angle of x
                               
                               
                               
                               
                              ' 'Make X stationary
                              ' ZXVelocity = ZXVelocity - XXVelocity
                              ' ZYVelocity = ZYVelocity - XYVelocity
                               
                               'work out direction from point of contact to centre of objects
                               
                               Dim G1 As Double, G2 As Double, G3 As Double, G4 As Double, GradX As Double, GradZ As Double, LenZ As Double, GradC As Double, LenC As Double
                               Call ModCoords(OpX, OpY, G1, G2, G3, G4, CoordMindist(1, 0), CoordMindist(1, 1), SpaceObject(X, 1) + ScreenModXX, SpaceObject(X, 2) + ScreenModXY)
                               
                               Call FindDirection(G1, G2, G3, G4, LenX, GradX)
                               
                               
                               Dim E1 As Double, E2 As Double, E3 As Double, E4 As Double
                               Call ModCoords(OpX, OpY, E1, E2, E3, E4, CoordMindist(0, 0), CoordMindist(0, 1), SpaceObject(Z, 1) + ScreenModZX, SpaceObject(Z, 2) + ScreenModZY)
                               Call FindDirection(E1, E2, E3, E4, LenZ, GradZ)
                               
                               
                               
                              ' If Abs(GradX - GradZ) = X Then
                              '
                              ' End If
                               'this is to check whether the inferred angle for x is towards z (it should not be - if it is there was an error with teh collision detection)
                               Call FindDirection(E3, E4, G3, G4, LenC, GradC)
                               If Abs(GradC - GradX) > 90 And Abs(GradC - GradX) < 270 Then
                                   GradX = GradC
                               End If
                               GradC = GradC + 180
                               If GradC > 360 Then GradC = GradC - 360
                               If Abs(GradC - GradZ) > 90 And Abs(GradC - GradZ) < 270 Then
                                   GradZ = GradC
                               End If
                               X = X
                               
                              
                               'work out final velocity
                               
                              Dim InitVelocity As Double
                              
                              If SpaceObject(Z, 4) > SpaceObject(X, 4) Then
                                   InitVelocity = SpaceObject(Z, 4)
                               Else
                                   InitVelocity = SpaceObject(X, 4)
                               End If
                              Dim InitMomentum As Double, FinalMomentum As Double, ForceA As Double, ForceB As Double, ForceC As Double, ForceD As Double
                               
                                   ForceA = SpaceObject(Z, 4) * (SpaceObject(Z, 6) / SpaceObject(X, 6))
                                   ForceB = SpaceObject(X, 4) * (SpaceObject(X, 6) / SpaceObject(Z, 6))
                                   
                                   If ForceA > SpaceObject(Z, 4) Then ForceA = SpaceObject(Z, 4)
                                   If ForceB > SpaceObject(X, 4) Then ForceB = SpaceObject(X, 4)
                                   
                                   InitMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                   
                                   Call DoThrust(X, GradX, ForceA)
                                   Call DoThrust(Z, GradZ, ForceB)
                                   
                                   ForceC = ForceA * (SpaceObject(X, 6) / SpaceObject(Z, 6))
                                   ForceD = ForceB * (SpaceObject(Z, 6) / SpaceObject(X, 6))
                                   
                                   If SpaceObject(Z, 6) > 1 Then
                                       If ForceC > ForceA Then ForceC = ForceA
                                       If ForceD > ForceB Then ForceD = ForceB
                                   End If
                                   
                                   
                                   
                                   'equal and opposite push
                                   
                                   Call DoThrust(Z, GradZ, ForceC)
                                   Call DoThrust(X, GradX, ForceD)
                                   
                                   FinalMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                   
                                   
                                   ForceA = (((FinalMomentum - InitMomentum) / 2) / SpaceObject(Z, 6))
                                   
                                   If ForceA < 0 Then
                                       ForceA = ForceA / 2
                                  
                                   End If
                                   SpaceObject(Z, 4) = SpaceObject(Z, 4) - ForceA
                                    ForceA = (((FinalMomentum - InitMomentum) / 2) / SpaceObject(X, 6))
                                   If ForceA < 0 Then
                                       ForceA = ForceA / 2
                                   
                                   End If
                                   SpaceObject(X, 4) = SpaceObject(X, 4) - ForceA
                                   
                                   FinalMomentum = SpaceObject(X, 4) * SpaceObject(X, 6) + SpaceObject(Z, 4) * SpaceObject(Z, 6)
                                   
                                   If SpaceObject(X, 4) > InitVelocity Then SpaceObject(X, 4) = InitVelocity
                                   If SpaceObject(Z, 4) > InitVelocity Then SpaceObject(Z, 4) = InitVelocity
                                   If SpaceObject(X, 6) > 5 Then
                                       If SpaceObject(Z, 4) <= 5 Then SpaceObject(Z, 4) = Rnd * 5
                                   End If
                                   If SpaceObject(Z, 6) > 5 Then
                                       If SpaceObject(X, 4) <= 5 Then SpaceObject(X, 4) = Rnd * 5
                                   End If
                                   xx = SpaceObject(X, 5)
                                   xx = SpaceObject(Z, 5)
                                   X = X
                             '  Call DoThrust(Z, GradZ, 10)
                              
                                    If X = 0 Then
                                        Scores(1) = Scores(1) + CLng(ForceD)
                                    ElseIf X = 3 Then
                                        Scores(0) = Scores(0) + CLng(ForceD)
                                    End If
                               
                               
                             '
                                
                         '      Call MovementCalc(GradX, Sqr(FinalVelocityXX ^ 2 + FinalVelocityXY ^ 2), XXVelocity, XYVelocity)
                         '      Call MovementCalc(GradZ, Sqr(FinalVelocityZX ^ 2 + FinalVelocityZY ^ 2), ZXVelocity, ZYVelocity)
                         '
                               'Add back initial velocities
                               'FinalVelocityXX = FinalVelocityXX + XXVelocity
                               'FinalVelocityXY = FinalVelocityXY + XYVelocity
                               'FinalVelocityZX = FinalVelocityZX + XXVelocity
                               'FinalVelocityZY = FinalVelocityZY + XYVelocity
                               
                        '       Call FindDirection(0, 0, XXVelocity, XYVelocity, LenX, GradX)
                        '       Call FindDirection(0, 0, ZXVelocity, ZYVelocity, LenZ, GradZ)
                               
                               
                               
                               
                               
                        '       SpaceObject(X, 4) = LenX
                        '       SpaceObject(X, 5) = GradX
                        '
                        '       SpaceObject(Z, 4) = LenZ
                        '       SpaceObject(Z, 5) = GradZ
                        '
                               ''SpaceObject(Z, 8) = -1
                               'SpaceObject(X, 8) = -1
                               'goon = 1
                               Exit For
                           
                           End If
                                                    
                        End If
                    End If
                End If
            End If
                
        Next Z
        
    End If
   
Next X
For X = 0 To MaxObjects
    If CollideRecord(X) > 10 Then
        CollideRecord(X) = 0
    End If
    
Next X

End Sub

Public Sub UpdateRotation()
For X = 0 To MaxObjects
    If SpaceObject(X, 0) = 1 Then
        If SpaceObject(X, 7) <> 0 Then
            SpaceObject(X, 3) = SpaceObject(X, 3) + SpaceObject(X, 7)
            If SpaceObject(X, 3) > 359 Then SpaceObject(X, 3) = SpaceObject(X, 3) - 360
            If SpaceObject(X, 3) < 0 Then SpaceObject(X, 3) = SpaceObject(X, 3) + 360
        End If
    End If
Next X
End Sub
Public Sub DoThruster(X)
If X = 0 Then
    Y = 1: Z = 0
Else
    Y = 4: Z = 3
End If
    LSpaceObject(Y, 0) = 0
    SpaceObject(Y, 0) = 1
    
    SpaceObject(Y, 1) = SpaceObject(Z, 1)
    SpaceObject(Y, 2) = SpaceObject(Z, 2)
    SpaceObject(Y, 3) = SpaceObject(Z, 3)
    SpaceObject(Y, 4) = SpaceObject(Z, 4)
    SpaceObject(Y, 5) = SpaceObject(Z, 5)
    SpaceObject(Y, 6) = SpaceObject(Z, 6)
    
    SpaceObject(Y, 10) = 180
    SpaceObject(Y, 11) = 20
    
    SpaceObject(Y, 12) = 135
    SpaceObject(Y, 13) = 100
    
    SpaceObject(Y, 14) = 150
    SpaceObject(Y, 15) = 200
    
    SpaceObject(Y, 16) = 160
    SpaceObject(Y, 17) = 150
    
    
    SpaceObject(Y, 18) = 180
    SpaceObject(Y, 19) = 400
    
    SpaceObject(Y, 20) = 200
    SpaceObject(Y, 21) = 150
     
    SpaceObject(Y, 22) = 210
    SpaceObject(Y, 23) = 200
    
    
    
    SpaceObject(Y, 24) = 225
    SpaceObject(Y, 25) = 100
    
    'SpaceObject(y, 18) = 180
    'SpaceObject(y, 19) = 20
    
    SpaceObject(Y, 26) = -1
End Sub
Public Sub Shoot(SON)
Dim Timer As Long, Direction As Long, NewDirsction As Long

Timer = GetTickCount
'If Abs(Timer - TimeKeeper(SON)) < 100 Then Exit Sub
TimeKeeper(SON) = Timer

NumObjects = NumObjects + 1
If NumObjects > 200 Then NumObjects = 0
BulletNo = 50 + NumObjects

'Do Player 1
Direction = (SpaceObject(SON, 5))
Degrees = (SpaceObject(SON, 3))
Velocity = (SpaceObject(SON, 4))

If Velocity > 0 And Direction <> Degrees Then


    Call MovementCalc(Direction, Velocity, xoff, yoff)
    Call MovementCalc(Degrees, 500, XOff2, YOff2)
    XOff3 = XOff2 + xoff
    YOff3 = YOff2 + yoff
    SpaceObject(BulletNo, 4) = Sqr(XOff3 ^ 2 + YOff3 ^ 2)
    XOff4 = Abs(XOff3)
    YOff4 = Abs(YOff3)
    If XOff4 > 0 And YOff4 > 0 Then
        NewDirection = Atn(YOff4 / XOff4)
        NewDirection = NewDirection * 180 / Pi
        If XOff3 > 0 And YOff3 > 0 Then
            NewDirection = 90 + NewDirection
        
        ElseIf XOff3 > 0 And YOff3 < 0 Then
            NewDirection = 90 - NewDirection
        ElseIf XOff3 < 0 And YOff3 > 0 Then
            NewDirection = 270 - NewDirection
        ElseIf XOff3 < 0 And YOff3 < 0 Then
            NewDirection = 270 + NewDirection
        End If
        If NewDirection = 360 Then NewDirection = 0
        
    Else
        If XOff3 = 0 And YOff3 > 0 Then
            NewDirection = 180
        ElseIf XOff3 = 0 And YOff3 < 0 Then
            NewDirection = 0
        ElseIf XOff3 > 0 And YOff3 = 0 Then
             NewDirection = 90
        ElseIf XOff3 < 0 And YOff3 = 0 Then
            NewDirection = 270
        End If
    End If
    SpaceObject(BulletNo, 5) = NewDirection
Else
    SpaceObject(BulletNo, 4) = SpaceObject(SON, 4) + 500
    SpaceObject(BulletNo, 5) = SpaceObject(SON, 3)

End If

SpaceObject(BulletNo, 0) = 1
    
SpaceObject(BulletNo, 1) = SpaceObject(SON, 1)
SpaceObject(BulletNo, 2) = SpaceObject(SON, 2)

Dim Scatter As Double




SpaceObject(BulletNo, 3) = SpaceObject(SON, 3)
SpaceObject(BulletNo, 9) = SON + 1
SpaceObject(BulletNo, 8) = 1
SpaceObject(BulletNo, 6) = 1
LSpaceObject(BulletNo, 0) = 0


If X = X Then 'make bullets less predictable
    Scatter = 1
    
    SpaceObject(BulletNo, 5) = SpaceObject(BulletNo, 5) + ((Rnd * Scatter) - Scatter / 2)
    
    
    Scatter = 30
    SpaceObject(BulletNo, 4) = SpaceObject(BulletNo, 4) + ((Rnd * Scatter) - Scatter / 2)
    
    
    If SpaceObject(BulletNo, 5) < 0 Then SpaceObject(BulletNo, 5) = 359 + SpaceObject(BulletNo, 5)
    If SpaceObject(BulletNo, 5) > 359 Then SpaceObject(BulletNo, 5) = SpaceObject(BulletNo, 5) - 359
    
    Scatter = 500
    Call MovementCalc(SpaceObject(BulletNo, 5), Rnd * Scatter, xoff, yoff)
End If


SpaceObject(BulletNo, 1) = SpaceObject(BulletNo, 1) + xoff
SpaceObject(BulletNo, 2) = SpaceObject(BulletNo, 2) + yoff
SpaceObject(BulletNo, 10) = 0
SpaceObject(BulletNo, 11) = 100


SpaceObject(BulletNo, 12) = 0
SpaceObject(BulletNo, 13) = 110
MDC(BulletNo) = SpaceObject(BulletNo, 13) - SpaceObject(BulletNo, 11) + 1
MaxProx(BulletNo) = 1
MinProx(BulletNo) = 1
For X = 0 To MaxObjects
    
        MaxProx2D(BulletNo, X) = MaxProx(X)
        MinProx2D(BulletNo, X) = MinProx(X)
    
Next X



SpaceObject(BulletNo, 14) = -1
End Sub
Public Sub ModCoords(OpX As Double, OpY As Double, G1 As Double, G2 As Double, G3 As Double, G4 As Double, V1 As Double, V2 As Double, V3 As Double, V4 As Double)
If OpX = 0 And OpY = 0 Then
    G1 = V1
    G2 = V2
    G3 = V3
    G4 = V4
Else
    If Abs(V1 - V3) < Abs(Abs(V1 - V3) + OpX) Then
        G1 = V1
        G3 = V3
    Else
        If V1 < V3 Then
            G1 = V1
            G3 = V3 + OpX
        Else
            G1 = V1 + OpX
            G3 = V3
        End If
    End If
    
    If Abs(V2 - V4) < Abs(Abs(V2 - V4) + OpY) Then
        G2 = V2
        G4 = V4
    Else
        If V2 < V4 Then
            G2 = V2
            G4 = V4 + OpY
        Else
            G2 = V2 + OpY
            G4 = V4
        End If
    End If
    
End If
End Sub
Public Sub DoThrust(SO, Degrees As Double, Force As Double)

Dim Velocity As Double, Direction As Double, Nedirection As Double, xoff As Double, yoff As Double, XOff2 As Double, YOff2 As Double, XOff3 As Double, YOff3 As Double, XOff4 As Double, YOff4 As Double
Velocity = (SpaceObject(SO, 4))
If Velocity > 0 Or X = X Then
    Direction = (SpaceObject(SO, 5))
    If Velocity > 0 Then
        Call MovementCalc(Direction, Velocity, xoff, yoff)
    End If
    'Degrees = SpaceObject(SO, 3)
    Call MovementCalc(Degrees, Force, XOff2, YOff2)
    XOff3 = XOff2 + xoff
    YOff3 = YOff2 + yoff
    SpaceObject(SO, 4) = Sqr(XOff3 ^ 2 + YOff3 ^ 2)
    XOff4 = Abs(XOff3)
    YOff4 = Abs(YOff3)
    If XOff4 > 0 And YOff4 > 0 Then 'And YOff4 > 0 Then
        NewDirection = Atn(YOff4 / XOff4)
        NewDirection = NewDirection * 180 / Pi
        If XOff3 > 0 And YOff3 > 0 Then
            NewDirection = 90 + NewDirection
        ElseIf XOff3 > 0 And YOff3 < 0 Then
            NewDirection = 90 - NewDirection
        ElseIf XOff3 < 0 And YOff3 > 0 Then
            NewDirection = 270 - NewDirection
        ElseIf XOff3 < 0 And YOff3 < 0 Then
            NewDirection = 270 + NewDirection
        End If
        If NewDirection = 360 Then NewDirection = 0
        
    Else
        If XOff3 = 0 And YOff3 > 0 Then
            NewDirection = 180
        ElseIf XOff3 = 0 And YOff3 < 0 Then
            NewDirection = 0
        ElseIf XOff3 > 0 And YOff3 = 0 Then
             NewDirection = 90
        ElseIf XOff3 < 0 And YOff3 = 0 Then
            NewDirection = 270
        End If
    End If
    SpaceObject(SO, 5) = NewDirection
    
Else
    SpaceObject(SO, 4) = SpaceObject(SO, 4) + 1
    SpaceObject(SO, 5) = SpaceObject(SO, 3)
End If
X = X
End Sub
Public Sub MovementCalc(Direction, Velocity, xoff, yoff)
    Dim TD As Double, S As Double
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
End Sub

