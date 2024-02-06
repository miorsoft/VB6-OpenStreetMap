Attribute VB_Name = "modCARs"
Option Explicit

Public Type tFuture
    X()           As Double
    Y()           As Double
    D()           As Double
    DX()          As Double
    DY()          As Double
    N             As Long
End Type


Public CAR()      As clsCAR
Public Ncars      As Long

Public CNT        As Long

Public DoSaveFrame As Boolean
Public ShowPath   As Boolean

Private Frame     As Long

'Public GRID       As cSpatialGrid
Public QT         As cQuadTree



Public Const CarWidth As Double = 2    '1.8
Public Const CarInterasse As Double = 2.2    '4.8 - CarWidth

Public Const CarImageRes As Double = 50
Public Const InvCarImageRes As Double = 1 / CarImageRes
Public Const CarImageResHalf As Double = CarImageRes * 0.5


'Public Const DistFromCenterRoadLine As Double = 1.5 + 0.5 - 0.25
Public Const kRoadW As Double = 6.4    '6.5
Public Const DistFromCenterRoadLine As Double = kRoadW * 0.265

Public Follow     As Long
Public DoFollow   As Boolean


Public DoLOOP     As Boolean

'---------- COLLISION
Private pCount&
Private RP1()     As Long
Private RP2()     As Long
Private Rdx()     As Double
Private Rdy()     As Double
Private rDD()     As Double
Private maxPairs  As Long



Public Function Max(ByVal A#, ByVal B#) As Double
    If A > B Then Max = A Else: Max = B
End Function


Public Function calcStreetSegmentCOST(ByVal A As Long, ByVal B As Long) As Double
    Dim DX        As Double
    Dim DY        As Double
    DX = Node(A).X - Node(B).X
    DY = Node(A).Y - Node(B).Y

    calcStreetSegmentCOST = Round(Sqr(DX * DX + DY * DY), 2)    ' _
                                                                + 1000 * Node(A).Traffic _
                                                                + 1000 * Node(B).Traffic
    '    If Node(A).Traffic Then Stop

End Function

Public Sub CARSSETUP()
    Dim I&
    '    frmMain.Caption = "Create Cars and Find Shortest Paths...   "
    SETPROGRESS "Enter N of Cars"
    DoEvents
    Ncars = 1

    Ncars = Val(InputBox("How many cars to put?" & vbCrLf & vbCrLf & _
                         "On Large Maps don't put too much Cars. Pathfinding algorithm is not so fast", "Number of Cars to put", 100))
    '    Ncars = 100

    ReDim CAR(Ncars)

    PhysicClearAll

    For I = 1 To Ncars
        If ((I - 1) And 3&) = 0& Then
            '        frmMain.Caption = "Create Cars and Find Shortest Paths...  " & I & " / " & Ncars
            SETPROGRESS "Creating Cars and their Shortest Paths...   " & I & " / " & Ncars, I / Ncars
        End If

        Set CAR(I) = New clsCAR

        CAR(I).MYIDX = I

        '-------- Predispone punti fisica
        PhysicAddPoint 0, 0
        PhysicAddPoint 0, 0
        CAR(I).PhysicRearIndex = PhysicNP - 1
        CAR(I).PhysicFrontIndex = PhysicNP
        '--------------------------------
        CAR(I).RANDOMstart
        '-------- Predispone fisica
        PhysicAddLine PhysicNP - 1, PhysicNP
        CAR(I).PhysicLineIdx = PhysicNL
        '--------------------------------


        CAR(I).RANDOMend
    Next
    frmMain.Caption = "GO!    (" & Ncars & " Cars)"

    DoEvents

    Follow = 1
    frmMain.txtFollow = Follow
    frmMain.PIC.MousePointer = 1


End Sub
Public Sub MAINLOOP()
    Dim I         As Long

    DoLOOP = True

    Do
        Zoom = Zoom * 0.995 + ZoomToGo * 0.005
        InvZoom = 1# / Zoom

        For I = 1 To Ncars
            CAR(I).DRIVE3
        Next
        Line_Update
        Points_Update

        '     If CNT Mod 12 Then carscollision2
        'If CNT Mod 24 = 0 Then '25
        If (CNT Mod 25&) = 0& Then    '25
            carscollision2


            DRAWMAPandCARS
            '        DRAWMAP
            '            For I = 1 To Ncars
            '                CAR(I).DRAWCC
            '            Next
            '


            'frmMain.PIC = SRF.Picture
            SRF.DrawToDC frmMainPIChDC



            If DoFollow Then
                '                PanX = PanX * 0.85 + CAR(Follow).Xfront * 0.15
                '                PanY = PanY * 0.85 + CAR(Follow).Yfront * 0.15
                PanX = PanX * 0.82 + CAR(Follow).Xfront * 0.18
                PanY = PanY * 0.82 + CAR(Follow).Yfront * 0.18
            Else
                PanX = PanX * 0.87 + PanXtoGo * 0.13
                PanY = PanY * 0.87 + PanYtoGo * 0.13
            End If


            PanZoomChanged = True

            If DoSaveFrame Then
                ' SaveJPG frmMain.PIC.Image, App.Path & "\Frames\" & Format(Frame, "0000") & ".jpg", 95
                ' SRF.WriteContentToJpgFile App.Path & "\Frames\" & Format(Frame, "0000") & ".jpg", 100
                SRF.WriteContentToPngFile App.Path & "\Frames\" & Format(Frame, "0000") & ".png"

                Frame = Frame + 1
            End If



            If (CNT And 63&) = 0& Then
                scr2WorldX0 = xfromScreen(0): scr2WorldY0 = yfromScreen(0)
                scr2WorldX1 = xfromScreen(scrMaxX): scr2WorldY1 = yfromScreen(scrMaxY)
                frmMain.labSTAT = "Visible area: " & Round(scr2WorldX1 - scr2WorldX0) & " X " & Round(scr2WorldY1 - scr2WorldY0) & " meters." & vbCrLf & vbCrLf & _
                                  "DRAWN" & vbCrLf & _
                                  "Road Segments :  " & STATSegs & vbCrLf & _
                                  "Buildings :  " & STATbuild & _
                                  " polygons made of " & STATPolyLines & " Lines. " & vbCrLf & vbCrLf & _
                                  "Cars: " & STATCars & " of " & Ncars
            End If





            DoEvents
        End If


        '        If (CNT And 2047&) = 0 Then
        '           UPDATETRAFFIC
        '        End If



        CNT = CNT + 1
    Loop While DoLOOP


End Sub





Public Sub LineLine(ByVal X1#, ByVal Y1#, ByVal X2#, ByVal Y2#, _
                    ByVal X3#, ByVal Y3#, ByVal X4#, ByVal Y4#, _
                    ByRef RX#, ByRef RY#)
    'https://www.youtube.com/watch?v=bvlIYX9cgls
    Dim D#, X21#, X43#, Y21#, Y43#, A#, B#, X31#, Y31#
    RX = 0: RY = 0
    X21 = X2 - X1
    Y21 = Y2 - Y1
    X43 = X4 - X3
    Y43 = Y4 - Y3
    ' Denominator for A and B are the same, so store this calculation
    D = X43 * Y21 - Y43 * X21
    If D = 0 Then Exit Sub        'Parallel
    X31 = X3 - X1
    Y31 = Y3 - Y1
    D = 1# / D
    A = (X43 * Y31 - Y43 * X31) * D
    If A >= 0# Then
        If A <= 1# Then
            B = (X21 * Y31 - Y21 * X31) * D
            If B >= 0 Then
                If B <= 1 Then
                    RX = X1 + A * X21
                    RY = Y1 + A * Y21
                End If
            End If
        End If
    End If
End Sub


'float sdLine( in vec2 p, in vec2 a, in vec2 b )
'{
'    vec2 pa = p-a, ba = b-a;
'    float H = clamp( dot(pa,ba)/dot(ba,ba), 0.0, 1.0 );
'    return length( pa - ba*H );
'}

Private Function DistFromLine(ByVal PX#, ByVal PY#, ByVal Ax, ByVal Ay#, ByVal Bx#, ByVal By#) As Double
    Dim paX#, BAX#, H#
    Dim paY#, BAY#, DX#, DY#
    paX = PX - Ax: BAX = Bx - Ax
    paY = PY - Ay: BAY = By - Ay
    H = (paX * BAX + paY * BAY) / (BAX * BAX + BAY * BAY)
    If H < 0# Then H = 0
    If H > 1# Then H = 1
    DX = paX - BAX * H
    DY = paY - BAY * H
    DistFromLine = DX * DX + DY * DY    'Squared distance from line
End Function

'
'Public Sub carscollision()
'    Dim I         As Long
'    Dim J         As Long
'    Dim dx        As Double
'    Dim dy        As Double
'    Dim D         As Double
'
'    Dim cDXi      As Double
'    Dim cDYi      As Double
'    Dim cDXj      As Double
'    Dim cDYj      As Double
'
'    Dim V         As Double
'
'    Dim pCount&
'    Dim RP1()     As Long
'    Dim RP2()     As Long
'    Dim Rdx()     As Double
'    Dim Rdy()     As Double
'    Dim rDD()     As Double
'    Dim K         As Long
'
'    GRID.ResetPoints
'
'    For I = 1 To Ncars
'        CAR(I).CanRun = True
'        GRID.InsertPoint CAR(I).X2, CAR(I).Y2
'        CAR(I).FuturePosComputed = False
'    Next
'
'    GRID.GetPairsWDist RP1, RP2, Rdx, Rdy, rDD, pCount
'
'
'
'    '        For I = 1 To Ncars - 1
'    '            For J = I + 1 To Ncars
'    For K = 1 To pCount
'
'        D = rDD(K)
'        '        dx = CAR(I).X2 - CAR(J).X2
'        '        dy = CAR(I).Y2 - CAR(J).Y2
'        '        D = dx * dx + dy * dy
'        If D < 144 Then           '225
'
'            I = RP1(K)
'            J = RP2(K)
'            dx = Rdx(K)
'            dy = Rdy(K)
'
'            '            cDXi = CAR(I).X1 - CAR(I).X0
'            '            cDYi = CAR(I).Y1 - CAR(I).y0
'            '            cDXj = CAR(J).X1 - CAR(J).X0
'            '            cDYj = CAR(J).Y1 - CAR(J).y0
'            cDXi = Cos(CAR(I).ANG)
'            cDYi = Sin(CAR(I).ANG)
'            cDXj = Cos(CAR(J).ANG)
'            cDYj = Sin(CAR(J).ANG)
'
'            If cDXi * cDXj + cDYi * cDYj > -0# Then    'Same Direction
'
'                ' If I = 1 Or J = 1 Then Stop
'
'
'                dx = CAR(J).X1 - CAR(I).X1
'                dy = CAR(J).Y1 - CAR(I).Y1
'                D = dx * dx + dy * dy
'                If D Then
'                    D = 1 / D
'                    dx = dx * D
'                    dy = dy * D
'                End If
'                V = cDXi * dx + cDYi * dy    'FRONT
'                If V > 0 Then
'                    With CAR(I)
'                        .VELtoGO = .VELtoGO * 0.75
'                        .CanRun = False
'                        '                        .VELtoGO = .VELtoGO - 0.01
'                    End With
'                End If
'                V = cDXj * dx + cDYj * dy
'                If V < 0 Then
'                    With CAR(J)
'                        .VELtoGO = .VELtoGO * 0.75
'                        .CanRun = False
'                        '                        .VELtoGO = .VELtoGO - 0.01
'                    End With
'                End If
'
'            End If
'        End If
'
'        '            Next
'        '        Next
'    Next
'
'End Sub



Public Sub carscollision2()
    Dim I         As Long
    Dim J         As Long
    Dim DX        As Double
    Dim DY        As Double




    Dim K         As Long

    Dim JJ&, ii&
    Dim Fi        As tFuture
    Dim Fj        As tFuture

    Dim RX#, RY#

    Dim CiX1#, CiY1#, CiX2#, CiY2#
    Dim CjX1#, CjY1#, CjX2#, CjY2#

    Dim IdireX#, IDireY#
    Dim JdireX#, JDireY#

    Dim CenIx#, CenIy#
    Dim CenJx#, CenJy#
    Dim Collide   As Boolean


    'GRID.ResetPoints
    QT.Reset


    For I = 1 To Ncars
        'GRID.InsertPoint CAR(I).Xfront, CAR(I).Yfront
        QT.InsertSinglePoint CAR(I).Xfront, CAR(I).Yfront, I

        CAR(I).FuturePosComputed = False
    Next

    '    GRID.GetPairsWDist rP1, rP2, rDX, rDY, rDD, pCount
    pCount = 0
    QT.GetPairsWDist 30, RP1, RP2, Rdx, Rdy, rDD, pCount, maxPairs
    frmMain.Caption = pCount & "   " & maxPairs

    For K = 1 To pCount

        If rDD(K) < 900 Then      '225
            I = RP1(K)
            J = RP2(K)
            If CAR(I).Layer = CAR(J).Layer Then
                CAR(I).RRX = 0
                CAR(J).RRX = 0

                Collide = False

                With CAR(I)
                    CiX1 = .Xrear: CiY1 = .yrear: CiX2 = .Xfront: CiY2 = .Yfront
                    CenIx = (.Xfront + .Xrear) * 0.5
                    CenIy = (.Yfront + .yrear) * 0.5
                    IdireX = .direX * 2: IDireY = .direY * 2
                End With
                With CAR(J)
                    CjX1 = .Xrear: CjY1 = .yrear: CjX2 = .Xfront: CjY2 = .Yfront
                    CenJx = (.Xfront + .Xrear) * 0.5
                    CenJy = (.Yfront + .yrear) * 0.5
                    JdireX = .direX * 2: JDireY = .direY * 2
                End With


                If CAR(I).FuturePosComputed = False Then CAR(I).CoumputeFuturepos
                If CAR(J).FuturePosComputed = False Then CAR(J).CoumputeFuturepos

                Fi = CAR(I).GetFuture
                Fj = CAR(J).GetFuture

                'If I = Follow Then Stop

                '------------------------------- I>J

                RX = 0: RY = 0
                For ii = 1 To Fi.N - 1

                    LineLine CjX1 - JdireX, CjY1 - JDireY, CjX2 + JdireX, CjY2 + JDireY, _
                             Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), RX, RY
                    '
                    '
                    If Not (RX <> 0 Or RY <> 0) Then
                        LineLine CjX1 - JDireY, CjY1 + JdireX, CjX1 + JDireY, CjY1 - JdireX, _
                                 Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), RX, RY
                    End If


                    If RX <> 0 Or RY <> 0 Then
                        '                     If I = Follow Then Stop
                        '
                        'CAR(I).VELtoGO = CAR(I).VELtoGO - 0.001    ' CAR(I).VELtoGO * 0#
                        If CAR(I).VELtoGO > 0.01 Then CAR(I).VELtoGO = CAR(I).VELtoGO * 0.8    ' 0.8
                        CAR(J).VELtoGO = CAR(J).VELtoGO * 1.1

                        CAR(I).RRX = RX    ' CenJx    'Fi.X(II) 'RX
                        CAR(I).RRY = RY    'CenJy    'Fi.Y(II) 'RY
                        Exit For
                    End If

                Next

                '                If RX = 0 Or RY = 0 Then
                '                    For ii = 1 To Fi.N - 1
                '                        For JJ = 1 To Fj.N - 1
                '                            LineLine Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), _
                                             '                                     Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), RX, RY
                '                            If RX <> 0 Or RY <> 0 Then
                '
                '                                DX = RX - CenIx
                '                                DY = RY - CenIy
                '                                If (DX * IdireX + DY * IDireY) > 0 Then
                '
                '                                    ii = Fi.N
                '                                    JJ = Fj.N
                '                                    If CAR(I).VELtoGO > 0.001 Then CAR(I).VELtoGO = CAR(I).VELtoGO * 0.1    ' 0.8
                '                                    CAR(J).VELtoGO = CAR(J).VELtoGO * 1.1
                '                                    CAR(I).RRX = RX    ' CenJx    'Fi.X(II) 'RX
                '                                    CAR(I).RRY = RY    'CenJy    'Fi.Y(II) 'RY
                '                                Else
                '
                '
                '                                End If
                '                            End If
                '
                '                        Next
                '                    Next
                '                End If



                '------------------------------- J>I



                RX = 0: RY = 0
                For JJ = 1 To Fj.N - 1
                    LineLine CiX1 - IdireX, CiY1 - IDireY, CiX2 + IdireX, CiY2 + IDireY, _
                             Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), RX, RY
                    '
                    '
                    If Not (RX <> 0 Or RY <> 0) Then
                        LineLine CiX1 - IDireY, CiY1 + IdireX, CiX1 + IDireY, CiY1 - IdireX, _
                                 Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), RX, RY
                    End If


                    If RX <> 0 Or RY <> 0 Then
                        '                    CAR(J).VELtoGO = CAR(J).VELtoGO - 0.001    'CAR(J).VELtoGO * 0#
                        If CAR(J).VELtoGO > 0.01 Then CAR(J).VELtoGO = CAR(J).VELtoGO * 0.8    '0.8
                        CAR(I).VELtoGO = CAR(I).VELtoGO * 1.1


                        CAR(J).RRX = RX    ' CenIx    'Fj.X(JJ) 'RX
                        CAR(J).RRY = RY    'CenIy    'Fj.Y(JJ) 'RY
                        Exit For
                    End If
                Next




                '                If RX = 0 Or RY = 0 Then
                '                    For JJ = 1 To Fj.N - 1
                '                        For ii = 1 To Fi.N - 1
                '                            LineLine Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), _
                                             '                                     Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), RX, RY
                '                            If RX <> 0 Or RY <> 0 Then
                '
                '                                DX = RX - CenJx
                '                                DY = RY - CenJy
                '                                If (DX * JdireX + DY * JDireY) > 0 Then
                '
                '                                    ii = Fi.N
                '                                    JJ = Fj.N
                '                                    If CAR(J).VELtoGO > 0.001 Then CAR(J).VELtoGO = CAR(J).VELtoGO * 0.1    ' 0.8
                '                                    CAR(I).VELtoGO = CAR(I).VELtoGO * 1.1
                '                                    CAR(J).RRX = RX    ' CenJx    'Fi.X(II) 'RX
                '                                    CAR(J).RRY = RY    'CenJy    'Fi.Y(II) 'RY
                '                                Else
                '
                '                                End If
                '                            End If
                '                        Next
                '                    Next
                '                End If




                If rDD(K) < 144 Then
                    Collide = LINELINECOLL(CAR(I).PhysicLineIdx, CAR(J).PhysicLineIdx)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicRearIndex, CAR(J).PhysicRearIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicFrontIndex, CAR(J).PhysicRearIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicRearIndex, CAR(J).PhysicFrontIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicFrontIndex, CAR(J).PhysicFrontIndex)
                    If Collide Then
                        CAR(I).VELtoGO = CAR(I).VELtoGO - 0.0005
                        CAR(J).VELtoGO = CAR(J).VELtoGO - 0.0005
                    End If
                End If

            End If
        End If


    Next

End Sub











Public Sub carscollision3()
    Dim I         As Long
    Dim J         As Long
    Dim DX        As Double
    Dim DY        As Double




    Dim K         As Long

    Dim JJ&, ii&
    Dim Fi        As tFuture
    Dim Fj        As tFuture

    Dim RX#, RY#

    Dim CiX1#, CiY1#, CiX2#, CiY2#
    Dim CjX1#, CjY1#, CjX2#, CjY2#

    Dim IdireX#, IDireY#
    Dim JdireX#, JDireY#

    Dim CenIx#, CenIy#
    Dim CenJx#, CenJy#
    Dim Collide   As Boolean


    'GRID.ResetPoints
    QT.Reset


    For I = 1 To Ncars
        '        GRID.InsertPoint CAR(I).Xfront, CAR(I).Yfront
        QT.InsertSinglePoint CAR(I).Xfront, CAR(I).Yfront, I

        CAR(I).FuturePosComputed = False
    Next

    '    GRID.GetPairsWDist rP1, rP2, rDX, rDY, rDD, pCount
    pCount = 0
    QT.GetPairsWDist 100, RP1, RP2, Rdx, Rdy, rDD, pCount, maxPairs


    For K = 1 To pCount

        If rDD(K) < 900 Then      '225
            I = RP1(K)
            J = RP2(K)
            If CAR(I).Layer = CAR(J).Layer Then
                CAR(I).RRX = 0
                CAR(J).RRX = 0

                Collide = False

                With CAR(I)
                    CiX1 = .Xrear: CiY1 = .yrear: CiX2 = .Xfront: CiY2 = .Yfront
                    CenIx = (.Xfront + .Xrear) * 0.5
                    CenIy = (.Yfront + .yrear) * 0.5
                    IdireX = .direX * 2: IDireY = .direY * 2
                End With
                With CAR(J)
                    CjX1 = .Xrear: CjY1 = .yrear: CjX2 = .Xfront: CjY2 = .Yfront
                    CenJx = (.Xfront + .Xrear) * 0.5
                    CenJy = (.Yfront + .yrear) * 0.5
                    JdireX = .direX * 2: JDireY = .direY * 2
                End With


                If CAR(I).FuturePosComputed = False Then CAR(I).CoumputeFuturepos
                If CAR(J).FuturePosComputed = False Then CAR(J).CoumputeFuturepos
                Fi = CAR(I).GetFuture
                Fj = CAR(J).GetFuture



                '------------------------------- I>J
                '                RX = 0: RY = 0
                '                For ii = 1 To Fi.N - 1
                '                    LineLine CjX1 - JdireX, CjY1 - JDireY, CjX2 + JdireX, CjY2 + JDireY, _
                                     '                             Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), RX, RY
                '                    If Not (RX <> 0 Or RY <> 0) Then
                '                        LineLine CjX1 - JDireY, CjY1 + JdireX, CjX1 + JDireY, CjY1 - JdireX, _
                                         '                                 Fi.X(ii), Fi.Y(ii), Fi.X(ii + 1), Fi.Y(ii + 1), RX, RY
                '                    End If
                '                    If RX <> 0 Or RY <> 0 Then
                '                        If CAR(I).VELtoGO > 0.01 Then CAR(I).VELtoGO = CAR(I).VELtoGO * 0.8    ' 0.8
                '                        CAR(J).VELtoGO = CAR(J).VELtoGO * 1.1
                '                        CAR(I).RRX = RX    ' CenJx    'Fi.X(II) 'RX
                '                        CAR(I).RRY = RY    'CenJy    'Fi.Y(II) 'RY
                '                        Exit For
                '                    End If
                '                Next




                '------------------------------- J>I
                '                RX = 0: RY = 0
                '                For JJ = 1 To Fj.N - 1
                '                    LineLine CiX1 - IdireX, CiY1 - IDireY, CiX2 + IdireX, CiY2 + IDireY, _
                                     '                             Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), RX, RY
                '                    If Not (RX <> 0 Or RY <> 0) Then
                '                        LineLine CiX1 - IDireY, CiY1 + IdireX, CiX1 + IDireY, CiY1 - IdireX, _
                                         '                                 Fj.X(JJ), Fj.Y(JJ), Fj.X(JJ + 1), Fj.Y(JJ + 1), RX, RY
                '                    End If
                '                    If RX <> 0 Or RY <> 0 Then
                '                        If CAR(J).VELtoGO > 0.01 Then CAR(J).VELtoGO = CAR(J).VELtoGO * 0.8    '0.8
                '                        CAR(I).VELtoGO = CAR(I).VELtoGO * 1.1
                '                        CAR(J).RRX = RX    ' CenIx    'Fj.X(JJ) 'RX
                '                        CAR(J).RRY = RY    'CenIy    'Fj.Y(JJ) 'RY
                '                        Exit For
                '                    End If
                '                Next


                Dim X#, Y#
                Dim D#
                Dim dm#
                dm = 0
                '---------------------------------------------------------------------

                X = CiX2 + (IdireX * CAR(I).VEL - JdireX * CAR(J).VEL) * 55
                Y = CiY2 + (IDireY * CAR(I).VEL - JDireY * CAR(J).VEL) * 55
                DX = X - CjX2
                DY = Y - CjY2: D = DX * DX + DY * DY
                If D < 12 Then
                    D = Sqr(D): DX = DX / D: DY = DY / D
                    If (DX * JdireX + DY * JDireY) > 0.5 Then
                        D = D / 4 - 0.1
                        CAR(J).VELtoGO = CAR(J).VELtoGO * D
                    End If
                End If


                X = CjX2 + (JdireX * CAR(J).VEL - IdireX * CAR(I).VEL) * 55
                Y = CjY2 + (JDireY * CAR(J).VEL - IDireY * CAR(I).VEL) * 55
                DX = X - CiX2
                DY = Y - CiY2:: D = DX * DX + DY * DY
                If D < 12 Then
                    D = Sqr(D): DX = DX / D: DY = DY / D
                    If (DX * IdireX + DY * IDireY) > 0.5 Then
                        D = D / 4 - 0.1
                        CAR(I).VELtoGO = CAR(I).VELtoGO * D
                    End If
                End If







                If rDD(K) < 144 Then
                    Collide = LINELINECOLL(CAR(I).PhysicLineIdx, CAR(J).PhysicLineIdx)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicRearIndex, CAR(J).PhysicRearIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicFrontIndex, CAR(J).PhysicRearIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicRearIndex, CAR(J).PhysicFrontIndex)
                    Collide = Collide Or POINTPOINTCOLL(CAR(I).PhysicFrontIndex, CAR(J).PhysicFrontIndex)
                    If Collide Then
                        CAR(I).VELtoGO = CAR(I).VELtoGO - 0.0005
                        CAR(J).VELtoGO = CAR(J).VELtoGO - 0.0005
                    End If
                End If

            End If
        End If


    Next

End Sub




