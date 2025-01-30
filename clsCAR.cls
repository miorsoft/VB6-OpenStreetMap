VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private DJ     As clsDijkstra
Private dijPATH   As tRoad



''Public X0         As Double 'REAR
''Public y0         As Double
''Public X1         As Double 'FRONT
''Public Y1         As Double

Public Xrear      As Double             'REAR
Public yrear      As Double
Public Xfront     As Double             'FRONT
Public Yfront     As Double


Private FutureX#
Private FutureY#

Public direX#
Public direY#


Public ANG        As Double
Public Steer      As Double
Public ANGtoGO    As Double

Public VEL        As Double
Public VELtoGO    As Double

Private STARTnode As Long
Public NEXTnode   As Long
Private PREVnode  As Long

Private ENDnode   As Long

Private PATHnextNode As Long

Private rANG      As Double
Private COSrANG   As Double
Private SINrANG   As Double
Private NextrANG  As Double


Public CanRun     As Double

Public MyColor    As Long

Public FuturePosComputed As Boolean




'Private FutN      As Long
Private MaxFutN   As Long
'Private FutX()    As Double
'Private FutY()    As Double
'Private FutD()    As Double

Private FUTURE    As tFuture

Public RRX#
Public RRY#

' For Physics
Public PhysicRearIndex As Long
Public PhysicFrontIndex As Long
Public PhysicLineIdx As Long

Private MySRF     As cCairoSurface

Public MYIDX      As Long

Public Layer      As Long


Friend Property Get GetPATHNN() As Long
    GetPATHNN = dijPATH.NN

End Property
Friend Function GetFuture() As tFuture
    GetFuture = FUTURE
End Function


Private Function calcROADang() As Double
    Dim NodeFROM  As Long
    Dim NodeTO    As Long

    Dim DX        As Double
    Dim DY        As Double

    If dijPATH.NN <> 0 Then
        NodeFROM = dijPATH.N(PATHnextNode - 1)
        NodeTO = dijPATH.N(PATHnextNode)
        DX = Node(NodeTO).X - Node(NodeFROM).X
        DY = Node(NodeTO).Y - Node(NodeFROM).Y

        rANG = Atan2(DX, DY)

        COSrANG = Cos(rANG)
        SINrANG = Sin(rANG)
        '-------------------------
        If PATHnextNode + 1 <= dijPATH.NN Then

            NodeFROM = dijPATH.N(PATHnextNode)
            NodeTO = dijPATH.N(PATHnextNode + 1)

            If Node(NodeTO).NNext > 1 Then

                DX = Node(NodeTO).X - Node(NodeFROM).X
                DY = Node(NodeTO).Y - Node(NodeFROM).Y

                NextrANG = Atan2(DX, DY)

                NextrANG = Abs(AngleDIFF(NextrANG, rANG))

                NextrANG = NextrANG * InvPIh

                If NextrANG > 1 Then NextrANG = 1
                '''NextrANG = 1 - NextrANG
            Else

            End If
        End If
    End If

End Function
Friend Sub SetEndNode(TargetNode&)
    ENDnode = TargetNode
    STARTnode = PREVnode
    dijPATH = DJ.DIJKSTRA(STARTnode, ENDnode)
    PATHnextNode = 2
    calcROADang
End Sub
Friend Sub RANDOMstart()
    Dim I         As Long


    Do
        STARTnode = Int(1 + Rnd * NNode)
    Loop While Node(STARTnode).dIsWay = False


    Xrear = Node(STARTnode).X
    yrear = Node(STARTnode).Y

    Xfront = Xrear + CarInterasse
    Yfront = yrear

    PATHnextNode = 2

    '-----------------------------
    Points(PhysicRearIndex).X = Xrear
    Points(PhysicRearIndex).Y = yrear + Rnd

    Points(PhysicFrontIndex).X = Xfront
    Points(PhysicFrontIndex).Y = Yfront + Rnd

    For I = PhysicRearIndex To PhysicFrontIndex    'Front-Rear contiguous
        With Points(I)
            .xPrev = .X
            .yPrev = .Y
            .u = 0
            .v = 0
        End With
    Next



End Sub


Friend Sub RANDOMend()
    Dim Try       As Long

AG:
    Try = 0
    Do
        Do
            ENDnode = Int(1 + Rnd * NNode)
        Loop While Node(ENDnode).dIsWay = False    'Or Node(ENDnode).NNext < 1    '2

        dijPATH = DJ.DIJKSTRA(STARTnode, ENDnode)

        Try = Try + 1

    Loop While dijPATH.NN = 0 And Try < 2

    If dijPATH.NN = 0 Then RANDOMstart: GoTo AG

    PATHnextNode = 2
    calcROADang

End Sub
Private Function CarDistFromNode(ByVal ii As Long) As Double
    Dim DX        As Double
    Dim DY        As Double
    DX = Node(ii).X - Xfront
    DY = Node(ii).Y - Yfront
    CarDistFromNode = (DX * DX + DY * DY)

End Function


Friend Sub DRAWCC()
    Dim X1        As Double
    Dim Y1        As Double
    Dim X2        As Double
    Dim Y2        As Double
    Dim I&
    Dim J&

    Dim X#, Y#, DX#, DY#

    '------------------------------------------------------------------------------


    '    X1 = XtoScreen(Xrear)
    '    Y1 = YtoScreen(yrear)
    '    If InsideScreen(X1, Y1) Then
    'Avoid XToScreen :
    If BBOverlapping(Xfront, Yfront, _
                     Xfront, Yfront, _
                     scr2WorldX0, scr2WorldY0, _
                     scr2WorldX1, scr2WorldY1) Then


        STATCars = STATCars + 1



        '-------------- SHOW PATH ---------------
        If MYIDX = Follow Then
            If ShowPath Then
                'DRAW DIJKSTRA PATH
                CC.SetSourceRGBA 1, 1, 0, 0.75
                '                CC.SetLineCap CAIRO_LINE_CAP_ROUND
                X = Zoom * 4#: If X < 2# Then X = 2#
                CC.SetLineWidth X
                CC.SetDashes 0, 1E+99
                With dijPATH
                    X1 = XtoScreen(Node(.N(1)).X)
                    Y1 = YtoScreen(Node(.N(1)).Y)
                    CC.MoveTo X1, Y1
                    For J = 2 To .NN
                        X2 = XtoScreen(Node(.N(J)).X)
                        Y2 = YtoScreen(Node(.N(J)).Y)
                        CC.LineTo X2, Y2
                        X1 = X2
                        Y1 = Y2
                    Next
                    CC.Stroke
                End With
            End If
        End If




        '        '        ' JUST LINE
        ''        X1 = XtoScreen(Xrear)
        ''        Y1 = YtoScreen(yrear)
        ''        X2 = XtoScreen(Xfront)
        ''        Y2 = YtoScreen(Yfront)
        ''        CC.DrawLine X1, Y1, X2, Y2, , CarWidth * Zoom, MyColor
        '        X = (Xrear + Xfront) * 0.5
        '        Y = (yrear + Yfront) * 0.5
        '        X = XtoScreen(X)
        '        Y = YtoScreen(Y)
        '        DX = direX * (CarInterasse) * Zoom * 0.5
        '        DY = direY * (CarInterasse) * Zoom * 0.5
        '        CC.DrawLine X - DX, Y - DY, X + DX, Y + DY, , CarWidth * Zoom, MyColor





        Dim CS#
        Dim SS#
        Dim hcI#
        Dim hcW#

        hcI = CarInterasse * 0.5
        hcW = CarWidth * 0.5

        CS = Cos(Steer)
        SS = Sin(Steer)


        X1 = (Xrear + Xfront) * 0.5
        Y1 = (yrear + Yfront) * 0.5
        X = XtoScreen(X1)
        Y = YtoScreen(Y1)
        CC.save
        CC.TranslateDrawings X, Y
        CC.RotateDrawings ANG
        CC.ScaleDrawings Zoom, Zoom

        'rear
        CC.DrawLine -0.23 - hcI, -hcW, 0.23 - hcI, -hcW, , 0.3, vbBlack
        CC.DrawLine -0.23 - hcI, hcW, 0.23 - hcI, hcW, , 0.3, vbBlack
        'Front
        CC.DrawLine CS * -0.23 + hcI, -hcW + SS * -0.23, _
                    CS * 0.23 + hcI, -hcW + SS * 0.23, , 0.3, vbBlack
        CC.DrawLine CS * -0.23 + hcI, hcW + SS * -0.23, _
                    CS * 0.23 + hcI, hcW + SS * 0.23, , 0.3, vbBlack

        CC.DrawLine -hcI, 0, hcI, 0, , CarWidth, MyColor    'Body


        CC.SetSourceColor vbCyan, 0.5
        CC.SetLineWidth 0.5
        CC.Arc -hcI, 0, CarInterasse * 0.8, -0.2, 0.2
        CC.Stroke


        CC.Restore



        ' USING PNG
        '        X = (Xrear + Xfront) * 0.5
        '        Y = (yrear + Yfront) * 0.5
        '        X = XtoScreen(X)
        '        Y = YtoScreen(Y)
        '        CC.save
        '        CC.TranslateDrawings X, Y
        '        CC.RotateDrawings ANG
        '        CC.ScaleDrawings Zoom * InvCarImageRes * 1.05, Zoom * InvCarImageRes * 1.25
        '        CC.RenderSurfaceContent MySRF, -(CarInterasse + CarWidth) * CarImageResHalf, -CarWidth * CarImageResHalf
        '        CC.Restore
        '------------------------------



        ' TARGET
        X1 = XtoScreen(Node(ENDnode).X)
        Y1 = YtoScreen(Node(ENDnode).Y)
        CC.SetSourceColor MyColor
        CC.Arc X1, Y1, 2
        CC.Fill






        '  ------------------------------------- FOR DEBUG
        '                If FuturePosComputed Then
        '
        '                    CC.SetLineWidth 2
        '                    CC.SetSourceColor vbYellow
        '                    CC.SetDashes 0, 1e+99
        '
        '                    With FUTURE
        '                        CC.MoveTo XtoScreen(.X(1)), YtoScreen(.Y(1))
        '                        For I = 2 To .N
        '                            CC.LineTo XtoScreen(.X(I)), YtoScreen(.Y(I))
        '                        Next
        '                        CC.Stroke
        '
        '                        CC.SetSourceColor vbRed
        '                        For I = 1 To .N
        '                            CC.Arc XtoScreen(.X(I)), YtoScreen(.Y(I)), 0.1 * Zoom
        '                            CC.Stroke
        '                        Next
        '                    End With
        '
        '                    CC.SetSourceColor MyColor    'Last Point
        '                    CC.Arc XtoScreen(FutureX), YtoScreen(FutureY), Zoom * 0.25
        '                    CC.Fill
        '                    ' Intersect
        '                    CC.Arc XtoScreen(RRX), YtoScreen(RRY), Zoom * 3
        '                    CC.Stroke
        '                End If






    End If


End Sub
Friend Sub DRIVE()
    Dim DX        As Double
    Dim DY        As Double
    Dim Dx2       As Double
    Dim Dy2       As Double

    Dim D         As Double

    Dim NEXTnode  As Long
    Dim P0        As Double
    Dim P1        As Double

    Dim D1        As Double
    Dim D2        As Double

    Dim X2        As Double
    Dim Y2        As Double




    If dijPATH.NN = 0 Then Exit Sub

    If PATHnextNode <> 0 Then

        NEXTnode = dijPATH.N(PATHnextNode)

        ''        Dx = Node(NEXTnode).x - .xfront
        ''        Dy = Node(NEXTnode).y - .yfront
        ''        ANGtoGO = Atan2(Dx, Dy)


        Dx2 = Node(NEXTnode).X - SINrANG * 1.3 - Xfront
        Dy2 = Node(NEXTnode).Y + COSrANG * 1.3 - Yfront
        D2 = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
        DX = Node(dijPATH.N(PATHnextNode - 1)).X - SINrANG * 1.3 - Xfront
        DY = Node(dijPATH.N(PATHnextNode - 1)).Y + COSrANG * 1.3 - Yfront
        D1 = Sqr(DX * DX + DY * DY) + VEL * 10


        P0 = 1 - D1 / (D1 + D2)         ': If P0 < 0 Then P0 = 0
        P1 = 1 - P0


        DX = DX * P0 + Dx2 * P1
        DY = DY * P0 + Dy2 * P1

        ANGtoGO = Atan2(DX, DY)



        '--------------------------------------------

        DX = Xfront - Xrear
        DY = Yfront - yrear
        ANG = Atan2(DX, DY)

        'Steer = Steer * 0.93 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.07
        Steer = Steer * 0.92 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.08



        If Abs(Steer) > 1 Then Stop: Steer = Sgn(Steer)

        P1 = (1 - NextrANG) * (1 - D2 * 0.001)


        If P1 < 0 Then P1 = 0

        P0 = 1 - P1


        VEL = VEL * 0.99 + (3 * VELtoGO * P0 + VELtoGO * P1) * 0.01

        Xfront = Xfront + Cos(ANG + Steer) * VEL
        Yfront = Yfront + Sin(ANG + Steer) * VEL

        '-----------------------------------------------------
        DX = Xfront - Xrear
        DY = Yfront - yrear
        D = Sqr(DX * DX + DY * DY)
        DX = DX / D
        DY = DY / D
        D = (D - CarInterasse) * 0.2
        DX = DX * D
        DY = DY * D
        Xrear = Xrear + DX
        yrear = yrear + DY
        Xfront = Xfront - DX
        Yfront = Yfront - DY
        '-----------------------------------------------------


        X2 = Xfront + Cos(ANG) * VEL * 100 * VEL
        Y2 = Yfront + Sin(ANG) * VEL * 100 * VEL



        If CarDistFromNode(NEXTnode) < 5 Then

            PATHnextNode = PATHnextNode + 1

            If PATHnextNode > dijPATH.NN Then

                STARTnode = NEXTnode
                RANDOMend

            End If

            calcROADang

        End If




    End If



End Sub

'''
'''Friend Sub DRIVE2()
'''    Dim dx        As Double
'''    Dim dy        As Double
'''    Dim Dx2       As Double
'''    Dim Dy2       As Double
'''
'''    Dim D         As Double
'''
'''    Dim NEXTnode  As Long
'''    Dim P0        As Double
'''    Dim P1        As Double
'''
'''    Dim D1        As Double
'''    Dim Dist2NEXT As Double
'''
'''
'''    Dim v         As Double
'''Dim X2 As Double
'''Dim Y2 As Double
'''
'''
'''
'''
'''
'''
'''    If dijPATH.NN = 0 Then Exit Sub
'''
'''    If PATHnextNode <> 0 Then
'''
'''        NEXTnode = dijPATH.N(PATHnextNode)
'''
'''        ''        Dx = Node(NEXTnode).x - Xfront
'''        ''        Dy = Node(NEXTnode).y - Yfront
'''        ''        ANGtoGO = Atan2(Dx, Dy)
'''
'''        If Node(NEXTnode).OneWay Then
'''
'''            Dx2 = Node(NEXTnode).X - SINrANG * DistFromCenterRoadLine * 0.75 - Xfront
'''            Dy2 = Node(NEXTnode).Y + COSrANG * DistFromCenterRoadLine * 0.75 - Yfront
'''            Dist2NEXT = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
'''            dx = Node(dijPATH.N(PATHnextNode - 1)).X - SINrANG * DistFromCenterRoadLine * 0.75 - Xfront
'''            dy = Node(dijPATH.N(PATHnextNode - 1)).Y + COSrANG * DistFromCenterRoadLine * 0.75 - Yfront
'''            D1 = Sqr(dx * dx + dy * dy) + VEL * 50
'''
'''        Else
'''            Dx2 = Node(NEXTnode).X - SINrANG * DistFromCenterRoadLine - Xfront
'''            Dy2 = Node(NEXTnode).Y + COSrANG * DistFromCenterRoadLine - Yfront
'''            Dist2NEXT = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
'''            dx = Node(dijPATH.N(PATHnextNode - 1)).X - SINrANG * DistFromCenterRoadLine - Xfront
'''            dy = Node(dijPATH.N(PATHnextNode - 1)).Y + COSrANG * DistFromCenterRoadLine - Yfront
'''            D1 = Sqr(dx * dx + dy * dy) + VEL * 50
'''        End If
'''
'''        P1 = D1 / (D1 + Dist2NEXT)    ': If P0 < 0 Then P0 = 0
'''        P0 = 1 - P1
'''
'''        dx = dx * P0 + Dx2 * P1
'''        dy = dy * P0 + Dy2 * P1
'''
'''        ANGtoGO = Atan2(dx, dy)
'''
'''        '--------------------------------------------
'''
'''        direX = Xfront - Xrear
'''        direY = Yfront - yrear
'''        ANG = Atan2(direX, direY)
'''
'''        'Steer = Steer * 0.93 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.07
'''        Steer = Steer * 0.92 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.08
'''
'''        If Abs(Steer) > 1 Then Steer = Sgn(Steer)
'''
'''
'''
'''        If Dist2NEXT < 40 Then    '30
'''            v = NextrANG * 0.001
'''            VELtoGO = VELtoGO - v
'''            VELtoGO = (1 - NextrANG) * 0.12
'''            If VELtoGO < 0.05 Then VELtoGO = 0.05
'''        End If
'''
'''
''''        VEL = VEL * 0.99 + VELtoGO * 0.01
'''        VEL = VEL * 0.988 + VELtoGO * 0.012
'''
'''        VELtoGO = VELtoGO + 0.00005
'''        If VELtoGO > 0.2 Then VELtoGO = 0.2
'''
'''
'''        Xfront = Xfront + Cos(ANG + Steer) * VEL
'''        Yfront = Yfront + Sin(ANG + Steer) * VEL
'''
'''
'''        '-----------------------------------------------------
'''        direX = Xfront - Xrear
'''        direY = Yfront - yrear
'''        D = Sqr(direX * direX + direY * direY)
'''        dx = direX / D
'''        dy = direY / D
'''        D = (D - CarInterasse) * 0.2
'''        dx = dx * D
'''        dy = dy * D
'''
'''        Xrear = Xrear + dx
'''        yrear = yrear + dy
'''        Xfront = Xfront - dx
'''        Yfront = Yfront - dy
'''        '-----------------------------------------------------
'''
'''        X2 = Xfront + Cos(ANG) * (VEL - 0.025) * 200
'''        Y2 = Yfront + Sin(ANG) * (VEL - 0.025) * 200
'''
'''
'''
'''
'''        dx = Node(NEXTnode).X - Xfront
'''        dy = Node(NEXTnode).Y - Yfront
'''        D = (dx * dx + dy * dy)
'''        If D < 5 Then
'''            PATHnextNode = PATHnextNode + 1
'''            If PATHnextNode > dijPATH.NN Then
'''                STARTnode = NEXTnode
'''                RANDOMend
'''            End If
'''            calcROADang
'''        End If
'''    End If
'''
'''
'''
'''End Sub


Friend Sub DRIVE3()
    Dim DX        As Double
    Dim DY        As Double
    Dim Dx2       As Double
    Dim Dy2       As Double

    Dim D         As Double

    Dim P0        As Double
    Dim P1        As Double

    Dim D1        As Double
    Dim Dist2NEXT As Double

    Dim v         As Double

    Dim X2        As Double
    Dim Y2        As Double

    Dim CA#, SA#


    If dijPATH.NN = 0 Then Exit Sub

    If PATHnextNode <> 0 Then
        PREVnode = NEXTnode: If PREVnode = 0 Then PREVnode = dijPATH.N(1)

        NEXTnode = dijPATH.N(PATHnextNode)

        Layer = Max(Node(NEXTnode).Layer, Node(PREVnode).Layer)

        ''        Dx = Node(NEXTnode).x - Xfront
        ''        Dy = Node(NEXTnode).y - Yfront
        ''        ANGtoGO = Atan2(Dx, Dy)



        '        If Node(NEXTnode).ONEWAY Then '-1, +1
        '            Dx2 = Node(NEXTnode).X - SINrANG * DistFromCenterRoadLine * 0.075 - Xfront
        '            Dy2 = Node(NEXTnode).Y + COSrANG * DistFromCenterRoadLine * 0.075 - Yfront
        '            Dist2NEXT = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
        '            DX = Node(dijPATH.N(PATHnextNode - 1)).X - SINrANG * DistFromCenterRoadLine * 0.075 - Xfront
        '            DY = Node(dijPATH.N(PATHnextNode - 1)).Y + COSrANG * DistFromCenterRoadLine * 0.075 - Yfront
        '            D1 = Sqr(DX * DX + DY * DY) + VEL * 50
        '        Else
        Dx2 = Node(NEXTnode).X - SINrANG * DistFromCenterRoadLine - Xfront
        Dy2 = Node(NEXTnode).Y + COSrANG * DistFromCenterRoadLine - Yfront
        Dist2NEXT = Sqr(Dx2 * Dx2 + Dy2 * Dy2)
        DX = Node(dijPATH.N(PATHnextNode - 1)).X - SINrANG * DistFromCenterRoadLine - Xfront
        DY = Node(dijPATH.N(PATHnextNode - 1)).Y + COSrANG * DistFromCenterRoadLine - Yfront
        D1 = Sqr(DX * DX + DY * DY) + VEL * 50
        '        End If

        P1 = D1 / (D1 + Dist2NEXT)      ': If P0 < 0 Then P0 = 0
        P0 = 1 - P1

        DX = DX * P0 + Dx2 * P1
        DY = DY * P0 + Dy2 * P1

        ANGtoGO = Atan2(DX, DY)

        '--------------------------------------------

        direX = Xfront - Xrear
        direY = Yfront - yrear
        ANG = Atan2(direX, direY)

        direX = Cos(ANG)
        direY = Sin(ANG)
        Points(PhysicRearIndex).WAdirX = -direY
        Points(PhysicRearIndex).WAdirY = direX



        'Steer = Steer * 0.93 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.07
        '        Steer = Steer * 0.92 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.08





        Steer = Steer * 0.96 + AngleDIFF(ANG + Steer, ANGtoGO) * 0.04



        If Abs(Steer) > 0.5 Then Steer = 0.5 * Sgn(Steer)



        If Dist2NEXT < 40 Then          '30
            '            v = NextrANG * 0.001
            '            VELtoGO = VELtoGO - v
            '            '                     VELtoGO = (1 - NextrANG) * 0.12
            '            VELtoGO = VELtoGO * 0.95 + 0.05 * (1 - NextrANG) * 0.12
            '            If VELtoGO < 0.05 Then VELtoGO = 0.05
            '-------------------------------
            '
            ''VELtoGO = VELtoGO * 0.9 + VELtoGO * 0.1 * (1 - NextrANG) * 0.001 * (1 - Dist2NEXT * 0.025)
            'VELtoGO = VELtoGO * 0.96 + 0.04 * 0.13 * (1 - NextrANG) '* (0.2 + Dist2NEXT * 0.025)
            'If VELtoGO < 0.05 Then VELtoGO = 0.05

            If VELtoGO > 0.075 Then VELtoGO = VELtoGO - NextrANG * 0.01
        End If


        '        VEL = VEL * 0.99 + VELtoGO * 0.01
        VEL = VEL * 0.988 + VELtoGO * 0.012

        VELtoGO = VELtoGO + 0.00005
        If VELtoGO > 0.2 Then VELtoGO = 0.2


        '        'COME Prima ma penso inutile con Physics:
        '        Xfront = Xfront + Cos(ANG + Steer) * VEL
        '        Yfront = Yfront + Sin(ANG + Steer) * VEL


        '  PHYSICS UPDATE
        '----------------------- MODE 1
        'Points(PhysicfrontIndex).x = Xfront
        'Points(PhysicfrontIndex).y = Yfront
        'Points(PhysicfrontIndex).xPrev = Xfront - Cos(ANG + Steer) * VEL * 0.5
        'Points(PhysicfrontIndex).yPrev = Yfront - Sin(ANG + Steer) * VEL * 0.5

        '----------------------- MODE 2
        CA = Cos(ANG + Steer)
        SA = Sin(ANG + Steer)
        Points(PhysicFrontIndex).X = Xfront + CA * VEL * 0.025
        Points(PhysicFrontIndex).Y = Yfront + SA * VEL * 0.025
        Points(PhysicFrontIndex).WAdirX = -SA
        Points(PhysicFrontIndex).WAdirY = CA

        '-----------------------------------------------------
        '        direX = Xfront - Xrear
        '        direY = Yfront - yrear
        '        D = Sqr(direX * direX + direY * direY)
        '        dx = direX / D
        '        dy = direY / D
        '        D = (D - CarInterasse) * 0.2
        '        dx = dx * D
        '        dy = dy * D
        '
        '        Xrear = Xrear + dx
        '        yrear = yrear + dy
        '        Xfront = Xfront - dx
        '        Yfront = Yfront - dy
        '-----------------------------------------------------

        '        X2 = Xfront + Cos(ANG) * (VEL - 0.025) * 200
        '        Y2 = Yfront + Sin(ANG) * (VEL - 0.025) * 200




        DX = Node(NEXTnode).X - Xfront
        DY = Node(NEXTnode).Y - Yfront
        D = (DX * DX + DY * DY)
        If D < 15 Then                  ' 5
            PATHnextNode = PATHnextNode + 1
            If PATHnextNode > dijPATH.NN Then
                STARTnode = NEXTnode
                RANDOMend
            End If
            calcROADang
        End If
    End If



End Sub
Private Sub Class_Initialize()
    VELtoGO = 0.07
    'MyColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    MyColor = RGB(Rnd * 150, Rnd * 150, Rnd * 150)

    'Set DJ = New clsDijkstra


    Set MySRF = Cairo.ImageList("C" & Int(Rnd * 5))

End Sub

Private Sub Class_Terminate()
    '    Set DJ = Nothing

End Sub



Public Sub CoumputeFuturepos()
    Dim NodeFROM&, NodeTO&
    Dim PATHnextNodeTMP&
    Dim X#, Y#

    Dim LineLen   As Double
    Dim DX#, DY#
    Dim PrimoTratto As Double

    Dim TOTPercorsa As Double
    Dim DaPercorere As Double
    Dim SA#
    Dim CA#
    Dim ftpANG#

    '    DaPercorere = VEL * 175 '150
    '    DaPercorere = Abs(VEL) ^ 0.5 * 50
    DaPercorere = 7 + VEL * VEL * 800

    If DaPercorere < 0 Then DaPercorere = 1

    X = Xfront
    Y = Yfront

    With FUTURE
        .N = 1
        If .N > MaxFutN Then
            MaxFutN = .N + 2
            ReDim .X(MaxFutN)
            ReDim .Y(MaxFutN)
            ReDim .D(MaxFutN)
            ReDim .DX(MaxFutN)
            ReDim .DY(MaxFutN)

        End If
        .X(1) = Xfront
        .Y(1) = Yfront


    End With



    PATHnextNodeTMP = PATHnextNode

    NodeTO = dijPATH.N(PATHnextNodeTMP)
    NodeFROM = dijPATH.N(PATHnextNodeTMP - 1)


    DX = Node(NodeTO).X - Node(NodeFROM).X
    DY = Node(NodeTO).Y - Node(NodeFROM).Y
    ftpANG = Atan2(DX, DY)
    DX = Node(NodeTO).X - X - Sin(ftpANG) * DistFromCenterRoadLine
    DY = Node(NodeTO).Y - Y + Cos(ftpANG) * DistFromCenterRoadLine

    PrimoTratto = Sqr(DX * DX + DY * DY)

    TOTPercorsa = PrimoTratto

    If TOTPercorsa < DaPercorere Then

        PATHnextNodeTMP = PATHnextNodeTMP + 1
        If PATHnextNodeTMP <= dijPATH.NN Then
            Do

                NodeFROM = dijPATH.N(PATHnextNodeTMP - 1)
                NodeTO = dijPATH.N(PATHnextNodeTMP)

                DX = Node(NodeTO).X - Node(NodeFROM).X
                DY = Node(NodeTO).Y - Node(NodeFROM).Y
                LineLen = Sqr(DX * DX + DY * DY)
                TOTPercorsa = TOTPercorsa + LineLen
                PATHnextNodeTMP = PATHnextNodeTMP + 1




                With FUTURE
                    ftpANG = Atan2(DX, DY)
                    CA = Cos(ftpANG)
                    SA = Sin(ftpANG)
                    .N = .N + 1
                    If .N > MaxFutN Then
                        MaxFutN = .N + 2
                        ReDim Preserve .X(MaxFutN)
                        ReDim Preserve .Y(MaxFutN)
                        ReDim Preserve .D(MaxFutN)
                        ReDim Preserve .DX(MaxFutN)
                        ReDim Preserve .DY(MaxFutN)

                    End If
                    .X(.N) = Node(NodeFROM).X - SA * DistFromCenterRoadLine
                    .Y(.N) = Node(NodeFROM).Y + CA * DistFromCenterRoadLine
                    .D(.N - 1) = LineLen
                    .DX(.N - 1) = CA
                    .DY(.N - 1) = SA
                End With
                If TOTPercorsa > DaPercorere Then Exit Do

            Loop While PATHnextNodeTMP <= dijPATH.NN    'And TOTPercorsa < DaPercorere
            ftpANG = Atan2(DX, DY)
            CA = Cos(ftpANG)
            SA = Sin(ftpANG)

            If TOTPercorsa > DaPercorere Then
                DX = DX / LineLen
                DY = DY / LineLen
                DX = DX * (TOTPercorsa - DaPercorere)
                DY = DY * (TOTPercorsa - DaPercorere)

                FutureX = Node(NodeTO).X - DX - SA * DistFromCenterRoadLine
                FutureY = Node(NodeTO).Y - DY + CA * DistFromCenterRoadLine
            Else
                FutureX = Node(NodeTO).X - SA * DistFromCenterRoadLine    'FINE CORSA
                FutureY = Node(NodeTO).Y + CA * DistFromCenterRoadLine
            End If

            With FUTURE
                .N = .N + 1
                If .N > MaxFutN Then
                    MaxFutN = .N + 2
                    ReDim Preserve .X(MaxFutN)
                    ReDim Preserve .Y(MaxFutN)
                    ReDim Preserve .D(MaxFutN)
                    ReDim Preserve .DX(MaxFutN)
                    ReDim Preserve .DY(MaxFutN)

                End If
                .X(.N) = FutureX
                .Y(.N) = FutureY
                .DX(.N - 1) = CA
                .DY(.N - 1) = SA
                If TOTPercorsa > DaPercorere Then
                    .D(.N - 1) = LineLen - (TOTPercorsa - DaPercorere)
                Else                    'FineCorsa
                    .D(.N - 1) = LineLen
                End If
            End With



        End If
    Else

        ftpANG = Atan2(DX, DY)
        CA = Cos(ftpANG)
        SA = Sin(ftpANG)

        DX = DX / PrimoTratto
        DY = DY / PrimoTratto
        DX = DX * (TOTPercorsa - DaPercorere)
        DY = DY * (TOTPercorsa - DaPercorere)

        FutureX = Node(NodeTO).X - DX - SA * DistFromCenterRoadLine
        FutureY = Node(NodeTO).Y - DY + CA * DistFromCenterRoadLine

        '
        With FUTURE
            .N = .N + 1
            .X(.N) = FutureX
            .Y(.N) = FutureY
            .D(.N - 1) = DaPercorere
            .DX(.N - 1) = CA
            .DY(.N - 1) = SA

        End With

    End If
    FuturePosComputed = True

End Sub
