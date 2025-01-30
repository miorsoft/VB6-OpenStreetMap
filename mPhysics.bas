Attribute VB_Name = "mPhysics"
Option Explicit

' https://github.com/kotsoft/Line-Physics/blob/master/physics.js
' VB6 PORT by Roberto Mior (reexre,miorsoft)

Private Const CollisionResp As Double = 0.33    ' for CARS

Private Type tLine
    P1            As Long
    P2            As Long
    solid         As Long
    restLength    As Double

    Length        As Double

    DX            As Double
    DY            As Double
    massDiff      As Double
    totalMass     As Double
    mulA          As Double
    mulB          As Double
End Type

Private Type tPoint
    X             As Double
    Y             As Double
    xPrev         As Double
    yPrev         As Double

    u             As Double
    v             As Double

    dispX         As Double
    dispY         As Double

    mass          As Double
    numConstraints As Double

    Fixed         As Long
    solid         As Long

    WAdirX        As Double             'Wheel axis direction
    WAdirY        As Double

End Type


Public Lines()    As tLine
Public Points()   As tPoint
Attribute Points.VB_VarUserMemId = 1073741824
Public PhysicNP   As Long
Public PhysicNL   As Long

Public Const PointSize As Double = 2    'CarWidth
Public Const PointSize2 As Double = PointSize * PointSize

Public Const kGravity As Double = 0.0015
Public GravityX   As Double
Public GravityY   As Double

Public PhysicRunning As Boolean


Public Sub PhysicUpdate()
    Dim I         As Long

    '    For I = 1 To PhysicNL
    '        Line_Update I
    '    Next
    '
    '    For I = 1 To PhysicNP
    '        Points_CheckCollisions I
    '    Next
    '
    '    For I = 1 To PhysicNP
    '        Points_Update I
    '    Next


End Sub


Public Sub PhysicDraw(Optional PointsToo As Long = -1)
    ' ctx.lineCap = 'round';
    '    for (var i in this.lines) {
    '      this.lines[i].PhysicDraw(ctx);
    '    }
    '
    '    for (var i in this.points) {
    '      ctx.beginPath();
    '      ctx.arc(this.points[i].x, this.points[i].y, this.pointSize / 2, 0, 2 * Math.PI, false);
    '      ctx.fill();
    '}
    Dim P1        As Long
    Dim P2        As Long
    Dim I         As Long
    Dim lw        As Double

    lw = PointSize * 0.9

    With CC
        .SetSourceColor RGB(20, 20, 20)
        .Paint
        .SetSourceColor vbGreen

        For I = 1 To PhysicNL
            P1 = Lines(I).P1
            P2 = Lines(I).P2

            .DrawLine Points(P1).X, Points(P1).Y, Points(P2).X, Points(P2).Y, , lw, vbGreen
        Next

        If PointsToo Then
            For I = 1 To PhysicNP
                .Arc Points(I).X, Points(I).Y, lw * 0.5
                .Fill
            Next
        End If

    End With

    SRF.DrawToDC frmMain.PIC.hDC
    DoEvents
End Sub

Public Sub PhysicClearAll()
    PhysicNP = 0
    PhysicNL = 0
End Sub
Public Sub PhysicAddLine(P1 As Long, P2 As Long, Optional solid As Long = -1)
    Dim DX        As Double
    Dim DY        As Double

    PhysicNL = PhysicNL + 1
    ReDim Preserve Lines(PhysicNL)
    With Lines(PhysicNL)
        .P1 = P1
        .P2 = P2
        .solid = solid
        DX = Points(P2).X - Points(P1).X
        DY = Points(P2).Y - Points(P1).Y
        .restLength = Sqr(DX * DX + DY * DY)


        Points(P1).mass = Points(P1).mass + .restLength * IIf(solid, 1, 0.25)
        Points(P2).mass = Points(P2).mass + .restLength * IIf(solid, 1, 0.25)

    End With

End Sub


Public Sub PhysicAddPoint(X As Double, Y As Double, Optional Fixed As Long = 0)
    PhysicNP = PhysicNP + 1
    ReDim Preserve Points(PhysicNP)
    With Points(PhysicNP)
        .X = X
        .Y = Y
        .xPrev = X
        .yPrev = Y
        .Fixed = Fixed
    End With

End Sub


Public Sub Line_Update()
    Dim LLL&

    Dim DiffX     As Double
    Dim DiffY     As Double

    Dim disp      As Double
    Dim dispX     As Double
    Dim dispY     As Double
    Dim P1        As Long
    Dim P2        As Long

    Dim I         As Long
    Dim point     As tPoint
    Dim interpolatedMass As Double
    Dim totalMass As Double
    Dim mulPoint  As Double
    Dim mulAB     As Double
    Dim mulSq     As Double
    Dim mulA2     As Double
    Dim mulB2     As Double

    Dim invL      As Double

    Dim mulA      As Double
    Dim mulB      As Double


    Const stiff   As Double = 0.5


    For LLL = 1 To PhysicNL

        With Lines(LLL)


            'this.updateVectors();<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            P1 = .P1
            P2 = .P2

            DiffX = Points(P2).X - Points(P1).X
            DiffY = Points(P2).Y - Points(P1).Y
            .Length = Sqr(DiffX * DiffX + DiffY * DiffY)
            invL = 1 / .Length
            .DX = DiffX * invL
            .DY = DiffY * invL
            .massDiff = Points(P2).mass - Points(P1).mass
            .totalMass = Points(P2).mass + Points(P1).mass
            .mulA = Points(P2).mass / .totalMass
            .mulB = 1# - .mulA
            mulA = .mulA
            mulB = .mulB

            'this.preserveLength();<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            disp = .restLength - .Length
            dispX = .DX * disp * stiff
            dispY = .DY * disp * stiff


            With Points(.P1)
                .dispX = .dispX - mulA * dispX
                .dispY = .dispY - mulA * dispY
                .numConstraints = .numConstraints + 1#
            End With
            With Points(.P2)
                .dispX = .dispX + mulB * dispX
                .dispY = .dispY + mulB * dispY
                .numConstraints = .numConstraints + 1#
            End With


            '''        'this.checkCollisions();<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            '''        If (.solid) Then
            '''            For I = 1 To PhysicNP
            '''                If (I <> .P2) Then
            '''                    If (I <> .P1) Then
            '''                        DiffX = Points(I).x - Points(.P1).x
            '''                        DiffY = Points(I).y - Points(.P1).y
            '''                        dp = DiffX * .dx + DiffY * .dy
            '''                        If (dp > 0 And dp < .length) Then
            '''                            cp = DiffX * .dy - DiffY * .dx
            '''
            '''                            If (cp > -PointSize) Then
            '''                                If (cp < PointSize) Then
            '''
            '''
            '''                                    '                                    point = Points(I)
            '''                                    f = dp / .length
            '''                                    interpolatedMass = 2# * (Points(.P1).mass + f * .massDiff)
            '''                                    totalMass = interpolatedMass + Points(I).mass
            '''                                    mulPoint = interpolatedMass / totalMass
            '''                                    mulAB = 1# - mulPoint
            '''
            '''                                    If cp > 0 Then
            '''                                        disp = PointSize - cp
            '''                                    Else
            '''                                        disp = -PointSize - cp
            '''                                    End If
            '''
            '''                                    dispX = disp * .dy
            '''                                    dispY = disp * -.dx
            '''                                    With Points(I)
            '''                                        .dispX = .dispX + mulPoint * dispX
            '''                                        .dispY = .dispY + mulPoint * dispY
            '''                                        .numConstraints = .numConstraints + 1#
            '''                                    End With
            '''
            '''                                    mulSq = mulAB / (.mulA * .mulA + .mulB * .mulB)
            '''                                    mulA2 = .mulA * mulSq
            '''                                    mulB2 = .mulB * mulSq
            '''                                    '                            this.pointA.dispX -= mulA2 * dispX;
            '''                                    '                            this.pointA.dispY -= mulA2 * dispY;
            '''                                    '                            this.pointB.dispX -= mulB2 * dispX;
            '''                                    '                            this.pointB.dispY -= mulB2 * dispY;
            '''
            '''                                    With Points(.P1)
            '''                                        .dispX = .dispX - mulA2 * dispX
            '''                                        .dispY = .dispY - mulA2 * dispY
            '''                                        .numConstraints = .numConstraints + 1#
            '''                                    End With
            '''
            '''                                    With Points(.P2)
            '''                                        .dispX = .dispX - mulB2 * dispX
            '''                                        .dispY = .dispY - mulB2 * dispY
            '''                                        .numConstraints = .numConstraints + 1#
            '''                                    End With
            '''
            '''
            '''                                    '                                    Points(I) = point
            '''
            '''                                End If
            '''                            End If
            '''
            '''                        End If
            '''                    End If
            '''                End If
            '''            Next
            '''        End If

        End With
    Next

End Sub
Private Sub Points_CheckCollisions(wp As Long)
    Dim I         As Long
    Dim point     As tPoint
    Dim DiffX     As Double
    Dim DiffY     As Double
    Dim Length    As Double
    Dim disp      As Double
    Dim DX        As Double
    Dim DY        As Double
    Dim totalMass As Double
    Dim mulA      As Double
    Dim mulB      As Double



    With Points(wp)
        For I = wp + 1 To PhysicNP
            DiffX = Points(I).X - .X
            If DiffX > -PointSize Then
                If DiffX < PointSize Then
                    DiffY = Points(I).Y - .Y
                    If DiffY > -PointSize Then
                        If DiffY < PointSize Then
                            Length = (DiffX * DiffX + DiffY * DiffY)
                            If Length < PointSize2 Then
                                'point = Points(I)

                                Length = Sqr(Length)
                                disp = (PointSize - Length) / Length
                                DX = DiffX * disp
                                DY = DiffY * disp

                                totalMass = .mass + Points(I).mass
                                mulA = Points(I).mass / totalMass
                                mulB = 1# - mulA

                                .dispX = .dispX - mulA * DX
                                .dispY = .dispY - mulA * DY
                                .numConstraints = .numConstraints + 1#

                                With Points(I)
                                    .dispX = .dispX + mulB * DX
                                    .dispY = .dispY + mulB * DY
                                    .numConstraints = .numConstraints + 1#
                                End With

                                ' Points(I) = point
                            End If
                        End If
                    End If
                End If
            End If
        Next

    End With

End Sub


Public Sub Points_Update()
    Dim I         As Long
    Dim D#
    Dim IC#
    For I = 1 To PhysicNP
        With Points(I)

            If Not (.Fixed) Then
                '            If .numConstraints = 0 Then .numConstraints = 1
                If .numConstraints Then
                    IC = 1# / .numConstraints
                    .X = .X + .dispX * IC
                    .Y = .Y + .dispY * IC
                End If
                '            If (.x < 0) Then
                '                .xPrev = .x: .x = 0
                '            ElseIf (.x > MaxW) Then
                '                .xPrev = .x: .x = MaxW
                '            End If
                '            If (.y < 0) Then
                '                .yPrev = .y: .y = 0
                '            ElseIf (.y > maxH) Then
                '                .yPrev = .y: .y = maxH
                '            End If
                .u = .X - .xPrev        '+ GravityX    '+ This.world.gravX
                .v = .Y - .yPrev        '+ GravityY    '+ This.world.gravY
                .u = .u * 0.97
                .v = .v * 0.97

                '                '------------------- TENUTA STRANA (RUOTA)
                D = .u * .WAdirX + .v * .WAdirY
                .u = .u - .WAdirX * D * 0.0125
                .v = .v - .WAdirY * D * 0.0125
                '                '------------------------------------


                .xPrev = .X
                .yPrev = .Y
                .X = .X + .u
                .Y = .Y + .v
            End If

            .dispX = 0#
            .dispY = 0#
            .numConstraints = 0#

        End With
    Next
    For I = 1 To Ncars
        With CAR(I)
            .Xrear = Points(.PhysicRearIndex).X
            .yrear = Points(.PhysicRearIndex).Y
            .Xfront = Points(.PhysicFrontIndex).X
            .Yfront = Points(.PhysicFrontIndex).Y
        End With
    Next

End Sub




Public Function LINELINECOLL(Line1&, Line2&) As Boolean

    Dim I         As Long
    Dim DiffX     As Double
    Dim DiffY     As Double

    Dim disp      As Double
    Dim dispX     As Double
    Dim dispY     As Double

    Dim point     As tPoint
    Dim DP        As Double
    Dim CP        As Double
    Dim F         As Double
    Dim interpolatedMass As Double
    Dim totalMass As Double
    Dim mulPoint  As Double
    Dim mulAB     As Double
    Dim mulSq     As Double
    Dim mulA2     As Double
    Dim mulB2     As Double



    With Lines(Line1)

        If (.solid) Then


            For I = Lines(Line2).P1 To Lines(Line2).P1 + 1
                If (I <> .P2) Then
                    If (I <> .P1) Then
                        DiffX = Points(I).X - Points(.P1).X
                        DiffY = Points(I).Y - Points(.P1).Y
                        DP = DiffX * .DX + DiffY * .DY
                        If (DP > 0 And DP < .Length) Then
                            CP = DiffX * .DY - DiffY * .DX

                            If (CP > -PointSize) Then
                                If (CP < PointSize) Then



                                    '                                    point = Points(I)
                                    F = DP / .Length
                                    interpolatedMass = 2# * (Points(.P1).mass + F * .massDiff)
                                    totalMass = interpolatedMass + Points(I).mass
                                    mulPoint = interpolatedMass / totalMass
                                    mulAB = 1# - mulPoint

                                    If CP > 0 Then
                                        disp = PointSize - CP
                                    Else
                                        disp = -PointSize - CP
                                    End If

                                    disp = disp * CollisionResp

                                    dispX = disp * .DY
                                    dispY = disp * -.DX
                                    With Points(I)
                                        .dispX = .dispX + mulPoint * dispX
                                        .dispY = .dispY + mulPoint * dispY
                                        .numConstraints = .numConstraints + 1#
                                    End With

                                    mulSq = mulAB / (.mulA * .mulA + .mulB * .mulB)
                                    mulA2 = .mulA * mulSq
                                    mulB2 = .mulB * mulSq
                                    '                            this.pointA.dispX -= mulA2 * dispX;
                                    '                            this.pointA.dispY -= mulA2 * dispY;
                                    '                            this.pointB.dispX -= mulB2 * dispX;
                                    '                            this.pointB.dispY -= mulB2 * dispY;

                                    With Points(.P1)
                                        .dispX = .dispX - mulA2 * dispX
                                        .dispY = .dispY - mulA2 * dispY
                                        .numConstraints = .numConstraints + 1#
                                    End With

                                    With Points(.P2)
                                        .dispX = .dispX - mulB2 * dispX
                                        .dispY = .dispY - mulB2 * dispY
                                        .numConstraints = .numConstraints + 1#
                                    End With

                                    LINELINECOLL = True


                                    '                                    Points(I) = point

                                End If
                            End If

                        End If
                    End If
                End If
            Next
        End If
    End With

End Function


Public Function POINTPOINTCOLL(J&, I&) As Boolean
    Dim point     As tPoint
    Dim DiffX     As Double
    Dim DiffY     As Double
    Dim Length    As Double
    Dim disp      As Double
    Dim DX        As Double
    Dim DY        As Double
    Dim totalMass As Double
    Dim mulA      As Double
    Dim mulB      As Double


    With Points(J)
        DiffX = Points(I).X - .X
        If DiffX > -PointSize Then
            If DiffX < PointSize Then
                DiffY = Points(I).Y - .Y
                If DiffY > -PointSize Then
                    If DiffY < PointSize Then
                        Length = (DiffX * DiffX + DiffY * DiffY)
                        If Length < PointSize2 Then
                            'point = Points(I)
                            If Length Then

                                Length = Sqr(Length)
                                disp = (PointSize - Length) / Length

                                disp = disp * CollisionResp
                                DX = DiffX * disp
                                DY = DiffY * disp

                                totalMass = .mass + Points(I).mass
                                mulA = Points(I).mass / totalMass
                                mulB = 1# - mulA

                                .dispX = .dispX - mulA * DX
                                .dispY = .dispY - mulA * DY
                                .numConstraints = .numConstraints + 1#

                                With Points(I)
                                    .dispX = .dispX + mulB * DX
                                    .dispY = .dispY + mulB * DY
                                    .numConstraints = .numConstraints + 1#
                                End With

                                POINTPOINTCOLL = True

                                ' Points(I) = point
                            End If

                        End If
                    End If
                End If
            End If
        End If

    End With
End Function



Public Function LineLineTest(TX1#, TY1#, TX2#, TY2#, F As tFuture, Idx As Long) As Boolean

    Dim DX#, DY#
    Dim DiffX#, DiffY#, DP#, CP#
    DX = F.DX(Idx)
    DY = F.DY(Idx)

    DiffX = TX1 - F.X(Idx)
    DiffY = TY1 - F.Y(Idx)
    DP = DiffX * DX + DiffY * DY
    If (DP > 0 And DP < F.D(Idx)) Then
        CP = DiffX * DY - DiffY * DX

        If (CP > -2) Then
            If (CP < 2) Then
                LineLineTest = True
            End If
        End If
    End If

    If Not (LineLineTest) Then
        DiffX = TX2 - F.X(Idx)
        DiffY = TY2 - F.Y(Idx)
        DP = DiffX * DX + DiffY * DY
        If (DP > 0 And DP < F.D(Idx)) Then
            CP = DiffX * DY - DiffY * DX

            If (CP > -2) Then
                If (CP < 2) Then
                    LineLineTest = True
                End If
            End If
        End If
    End If

End Function




' NOT USED AND TESTED YET !

' https://arrowinmyknee.com/2021/03/15/some-math-about-capsule-collision/

'// Computes closest points C1 and C2 of S1(s)=P1+s*(Q1-P1) and
'// S2(t)=P2+t*(Q2-P2), returning s and t. Function result is squared
'// distance between between S1(s) and S2(t)
Public Function ClosestPtSegmentSegment(p1X#, p1Y#, q1X#, q1Y#, _
                                        p2X#, p2Y#, q2X#, q2Y#) As Double
    Const EPSILON As Double = 0.0001

    'float &s, float &t, Point &c1, Point &c2)

    Dim c1X#, c1Y#
    Dim c2X#, c2Y#



    Dim S#, T#
    Dim D1x#, D1y#, D2x#, D2y#
    Dim RX#, RY#
    Dim A#, E#, F#, C#, B#, denom#
    Dim DX#, DY#
    '    Vector d1 = q1 - p1; // Direction vector of segment S1
    '    Vector d2 = q2 - p2; // Direction vector of segment S2
    '    Vector r = p1 - p2;

    D1x = q1X - p1X
    D1y = q1Y - p1Y
    D2x = q2X - p2X
    D2y = q2Y - p2Y
    RX = p1X - p2X
    RY = p1Y - p2Y

    '    float a = Dot(d1, d1); // Squared length of segment S1, always nonnegative
    '    float e = Dot(d2, d2); // Squared length of segment S2, always nonnegative
    '    float f = Dot(d2, r);

    A = D1x * D1x + D1y * D1y
    E = D2x * D2x + D2y * D2y
    F = D2x * RX + D2y * RY


    '// Check if either or both segments degenerate into points
    'if (a <= EPSILON && e <= EPSILON) {
    If A <= EPSILON And E <= EPSILON Then
        '// Both segments degenerate into points
        S = 0: T = 0
        c1X = p1X: c1Y = p1Y
        c2X = p2X: c2Y = p2Y
        '        return Dot(c1 - c2, c1 - c2);
        DX = c1X - c2X: DY = c1Y - c2Y
        ClosestPtSegmentSegment = DX * DX + DY * DY

        Exit Function
    End If

    If (A <= EPSILON) Then
        '// First segment degenerates into a point
        S = 0#
        T = F / E                       '; // s = 0 => t = (b*s + f) / e = f / e
        '        t = Clamp(t, 0.0f, 1.0f);
        If T < 0 Then T = 0
        If T > 1 Then T = 1
    Else
        'c = Dot(d1, r);
        C = D1x * RX + D1y * RY
        If (E <= EPSILON) Then
            '// Second segment degenerates into a point
            T = 0#
            's = Clamp(-c / a, 0.0f, 1.0f); // t = 0 => s = (b*t - c) / a = -c / a
            S = -C / A
            If S < 0 Then S = 0
            If S > 1 Then S = 1
        Else
            '// The general nondegenerate case starts here
            'float b = Dot(d1, d2);
            B = D1x * D2x + D1y * D2y
            denom = A * E - B * B       '; // Always nonnegative
            '// If segments not parallel, compute closest point on L1 to L2 and
            '// clamp to segment S1. Else pick arbitrary s (here 0)
            If (denom) Then
                S = (B * F - C * E) / denom
                If S < 0 Then S = 0
                If S > 1 Then S = 1
            Else
                S = 0
            End If
            '// Compute point on L2 closest to S1(s) using
            '// t = Dot((P1 + D1*s) - P2,D2) / Dot(D2,D2) = (b*s + f) / e
            T = (B * S + F) / E
            '// If t in [0,1] done. Else clamp t, recompute s for the new value
            '// of t using s = Dot((P2 + D2*t) - P1,D1) / Dot(D1,D1)= (t*b - c) / a
            '// and clamp s to [0, 1]
            If (T < 0#) Then
                T = 0#
                S = -C / A
                If S < 0 Then S = 0
                If S > 1 Then S = 1
            ElseIf (T > 1#) Then
                T = 1#
                S = (B - C) / A
                If S < 0 Then S = 0
                If S > 1 Then S = 1
            End If
        End If
    End If
    '    c1 = p1 + d1 * s;
    '    c2 = p2 + d2 * t;

    c1X = p1X + D1x * S
    c1Y = p1Y + D1y * S
    c2X = p2X + D2x * T
    c2Y = p2Y + D2y * T

    DX = c1X - c2X: DY = c1Y - c2Y

    'return Dot(c1 - c2, c1 - c2);
    ClosestPtSegmentSegment = DX * DX + DY * DY

End Function

