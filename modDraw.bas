Attribute VB_Name = "modDraw"
Option Explicit

Private mapMaxX   As Double
Private mapMaxY   As Double

Private KX        As Double
Private KY        As Double
Public frmMainPIChDC As Long
Public pHDC2      As Long

Public scrMaxX    As Double
Public scrMaxY    As Double


Public ClipLEFT   As Double
Public ClipTOP    As Double
Public ClipRIGHT  As Double
Public ClipBOTTOM As Double


Public PanX       As Double
Public PanY       As Double
Public PanXtoGo   As Double
Public PanYtoGo   As Double

Public CENPanX    As Double
Public CENPanY    As Double


Public CenX       As Double
Public CenY       As Double

Public ZoomToGo   As Double

Public Zoom       As Double
Public InvZoom    As Double
Public Navigating As Boolean
Public PanZoomChanged As Boolean

Public Const PI   As Double = 3.14159265358979
Public Const PI2  As Double = 6.28318530717959
Public Const PIh  As Double = 1.5707963267949
Public Const InvPI As Double = 1 / PI
Public Const InvPIh As Double = 1 / PIh


Public SRF        As cCairoSurface
Public CC         As cCairoContext

Public scr2WorldX0#
Public scr2WorldY0#
Public scr2WorldX1#
Public scr2WorldY1#


Public STATSegs   As Long
Public STATbuild  As Long
Public STATPolyLines As Long
Public STATCars   As Long

Private BuildingPolygon As cArrayList

Private CameraAng As Double

Public Function Atan2(ByVal DX As Double, ByVal DY As Double) As Double


    If DX Then Atan2 = Atn(DY / DX) + PI * (DX < 0#) Else: Atan2 = -PIh - (DY > 0#) * PI
    '    While Atan2 < 0: Atan2 = Atan2 + PI2: Wend
    '    While Atan2 > PI2: Atan2 = Atan2 - PI2: Wend
End Function
Public Function AngleDIFF(ByRef A1 As Double, ByRef A2 As Double) As Double

    AngleDIFF = A2 - A1
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend
End Function

Public Function AngleDIFF180(ByRef A1 As Double, ByRef A2 As Double) As Double

    AngleDIFF180 = A2 - A1
    While AngleDIFF180 < -PIh
        AngleDIFF180 = AngleDIFF180 + PI
    Wend
    While AngleDIFF180 > PIh
        AngleDIFF180 = AngleDIFF180 - PI
    Wend
End Function

Public Function InsideScreen(ByVal X As Double, ByVal Y As Double) As Boolean
    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > scrMaxX Then Exit Function
    If Y > scrMaxY Then Exit Function
    InsideScreen = True
End Function



Public Function XtoScreen(ByVal X As Double) As Double    'Long
    XtoScreen = Zoom * (X - PanX) + CenX
End Function
Public Function YtoScreen(ByVal Y As Double) As Double    'Long
    YtoScreen = Zoom * (Y - PanY) + CenY
End Function

Public Function xfromScreen(ByVal X As Double) As Double
    xfromScreen = (X - CenX) * InvZoom + PanX
End Function
Public Function yfromScreen(ByVal Y As Double) As Double
    yfromScreen = (Y - CenY) * InvZoom + PanY
End Function


Public Sub InitDraw()
    Dim I         As Long



    Set BuildingPolygon = New_c.ArrayList(vbDouble)


    Zoom = 1
    InvZoom = 1
    ZoomToGo = 1


    ClipLEFT = -25
    ClipTOP = -25
    ClipRIGHT = scrMaxX + 25
    ClipBOTTOM = scrMaxY + 25
    'Show/Debug Clipline:
    '    ClipLEFT = 150
    '    ClipTOP = 150
    '    ClipRIGHT = scrMaxX - 150
    '    ClipBOTTOM = scrMaxY - 150




    CenX = scrMaxX \ 2
    CenY = scrMaxY \ 2

    '    PanX = CenX
    '    PanY = CenY

    mapMaxX = -1E+55
    mapMaxY = -1E+55

    For I = 1 To NNode
        If Node(I).dIsWay = True Then
            If Node(I).X > mapMaxX Then mapMaxX = Node(I).X
            If Node(I).Y > mapMaxY Then mapMaxY = Node(I).Y
        End If
    Next


    KX = scrMaxX / (mapMaxX)            '- MinX)
    KY = scrMaxY / (mapMaxY)            ' - MinY)

    'PanX = XtoScreen((mapMaxX + MinX) * 0.5)
    'PanY = YtoScreen((mapMaxY + MinY) * 0.5)

    ' V1 2024
    PanX = XtoScreen(mapMaxX * 0.5)
    PanY = YtoScreen(mapMaxY * 0.5)
    CENPanX = PanX
    CENPanY = PanY




    If KX < KY Then KY = KX Else: KX = KY




End Sub


Public Sub DRAWMAP()
    Dim I         As Long
    Dim J As Long, K As Long

    Dim X1        As Double
    Dim Y1        As Double

    Dim X2        As Double
    Dim Y2        As Double

    Dim cX1       As Double
    Dim cY1       As Double
    Dim cX2       As Double
    Dim cY2       As Double

    Dim Dash1#, Dash2#
    Dim IsROAD    As Boolean

    Dim DX#, DY#
    Dim Xa#, Ya#
    Dim Xc#, Yc#

    Dim ANG#, CosA#, SinA#
    Dim Ang2#, CosA2#, SinA2#
    Dim Ang3#, CosA3#, SinA3#
    Dim D#
    Dim L#


    Dim C         As Long

    Dim RoadW     As Double

    Dim SolidPatternLongBLACK As cCairoPattern

    Dim Color     As Long

    Dim WhiteLineWidth As Double

    C = vbRed


    Dim X#, Y#




    WhiteLineWidth = Zoom * 0.3


    If PanZoomChanged Then

        Set SolidPatternLongBLACK = Cairo.CreateSolidPatternLng(vbBlack)


        STATSegs = 0
        STATbuild = 0
        STATPolyLines = 0
        STATCars = 0


        Dash1 = Zoom * 4                '4          '2.5
        Dash2 = Zoom * 8                '15         '*8


        scr2WorldX0 = xfromScreen(ClipLEFT): scr2WorldY0 = yfromScreen(ClipTOP)
        scr2WorldX1 = xfromScreen(ClipRIGHT): scr2WorldY1 = yfromScreen(ClipBOTTOM)



        'CC.SetSourceColor 0: CC.Paint
        '        CC.SetSourceRGB 0.23, 0.35, 0.23: CC.Paint  'GREEM
        CC.SetSourceRGB 0.1, 0.43, 0.1: CC.Paint    'GREEM




        'DRAW BULIDINGS / BLOCKS /AREAS ------------------------------------------
        CC.SetLineWidth 0.35 * Zoom
        CC.SetDashes 0, 1E+99
        '        CC.SetLineCap CAIRO_LINE_CAP_ROUND

        For I = 1 To NWay
            With WAY(I)
                If .IsBuilding Or .IsLeisure Or .IsAmenity Or .IsWater Or .IsShop Or _
                   .IsNatural Then
                    If .IsBuilding Then
                        'CC.SetSourceRGB 0.6, 0.3, 0.2 ' BUILDING
                        CC.SetSourceRGB 0.5 + .R * 0.08, 0.2 + .G * 0.08, 0.1 + .B = 0.08

                    End If
                    If .IsLeisure Then
                        If .Name = "Swimming Pool" Then
                            CC.SetSourceRGB 0.1, 1, 1
                        Else
                            CC.SetSourceRGB 0.3, 0.7, 0.3
                        End If
                    End If
                    If .IsWater Then
                        CC.SetSourceRGB 0, 0.5, 1    '.2, 0.2, 1
                    End If

                    If .IsAmenity Then
                        CC.SetSourceRGB 0.15, 0.15, 0.15
                    End If

                    If .IsShop Then
                        CC.SetSourceRGB 0.65, 0.65, 0.65
                    End If
                    '
                    '                    If .IsNatural Then
                    '                        CC.SetSourceRGB 1, 0, 0
                    '                    End If

                    '
                    '                    ' If InsideScreenBuilding(XtoScreen(.cX), YtoScreen(.cY)) Then
                    '                    If BBOverlapping(XtoScreen(.BBx1), YtoScreen(.BBY1), _
                                         '                                     XtoScreen(.BBx2), YtoScreen(.BBY2), _
                                         '                                     0, 0, _
                                         '                                     scrMaxX, scrMaxY) Then


                    If BBOverlapping(.BBx1, .BBY1, _
                                     .BBx2, .BBY2, _
                                     scr2WorldX0, scr2WorldY0, _
                                     scr2WorldX1, scr2WorldY1) Then

                        K = 0
                        ReDim PTS(.NN * 2 - 1)
                        X = 0
                        Y = 0
                        STATbuild = STATbuild + 1
                        'https://www.vbforums.com/showthread.php?902225-Stretch-a-Polygon-region&p=5627915&viewfull=1#post5627915

                        BuildingPolygon.removeAll

                        For J = 1 To .NN
                            X1 = XtoScreen(Node(.N(J)).X)
                            Y1 = YtoScreen(Node(.N(J)).Y)
                            'PTS(K) = X1
                            'PTS(K + 1) = Y1
                            BuildingPolygon.Add X1
                            BuildingPolygon.Add Y1
                            X = X + X1
                            Y = Y + Y1
                            K = K + 2
                        Next
                        '                        CC.PolygonSingle PTS, False, splNone, , True
                        CC.PolygonPtr BuildingPolygon.DataPtr, BuildingPolygon.Count \ 2, False, splNone

                        If Zoom > 0.85 Then
                            CC.Fill True
                            CC.Stroke , SolidPatternLongBLACK
                        Else
                            CC.Fill
                        End If
                        .CenterX = X * .invNN    ' Per Stampa Names
                        .CenterY = Y * .invNN    ' CC.TextOut .CX, .CY, .NAME

                        STATPolyLines = STATPolyLines + BuildingPolygon.Count \ 2
                    End If


                End If
            End With
        Next
        '----------------------------------------------------------------------------



        '---------------------------------DRAW ROADS-----------------------------------------
        '---------------------------------Black Border
        CC.SetDashes 0, 1E+99
        If Zoom > 2 Then
            For I = 1 To NWay
                With WAY(I)
                    If .IsArea Or .IsBuilding Or .IsWater Or _
                       .IsAmenity Or .IsLeisure Or .IsShop Or .IsNatural Then
                    Else
                        If LenB(.wayType) Then
                            If BBOverlapping(.BBx1, .BBY1, _
                                             .BBx2, .BBY2, _
                                             scr2WorldX0, scr2WorldY0, _
                                             scr2WorldX1, scr2WorldY1) Then

                                X1 = XtoScreen(Node(.N(1)).X)
                                Y1 = YtoScreen(Node(.N(1)).Y)
                                cX1 = X1
                                cY1 = Y1

                                RoadW = .RoadWidth * Zoom
                                RoadW = RoadW * 0.5 * .Lanes + Zoom    '1 meter larger

                                For J = 2 To .NN
                                    X2 = XtoScreen(Node(.N(J)).X)
                                    Y2 = YtoScreen(Node(.N(J)).Y)
                                    cX2 = X2
                                    cY2 = Y2


                                    If CLIPLINEcc(cX1, cY1, cX2, cY2) Then
                                        IsROAD = Node(.N(J)).dIsWay And Node(.N(J - 1)).dIsWay
                                        If IsROAD Then
                                            CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, 3289650    'vbBlack
                                        End If
                                    End If
                                    X1 = X2
                                    Y1 = Y2
                                    cX1 = X2
                                    cY1 = Y2
                                Next
                            End If

                        End If
                    End If
                End With
            Next
        End If
        '---------- ASFALTO---- ROADS -----------------
        For I = 1 To NWay
            With WAY(I)
                If .IsArea Or .IsBuilding Or .IsWater Or _
                   .IsAmenity Or .IsLeisure Or .IsShop Or .IsNatural Then
                Else
                    If LenB(.wayType) Then
                        If BBOverlapping(.BBx1, .BBY1, _
                                         .BBx2, .BBY2, _
                                         scr2WorldX0, scr2WorldY0, _
                                         scr2WorldX1, scr2WorldY1) Then

                            X1 = XtoScreen(Node(.N(1)).X)
                            Y1 = YtoScreen(Node(.N(1)).Y)
                            cX1 = X1
                            cY1 = Y1

                            RoadW = .RoadWidth * Zoom
                            RoadW = RoadW * 0.5 * .Lanes

                            For J = 2 To .NN

                                STATSegs = STATSegs + 1

                                X2 = XtoScreen(Node(.N(J)).X)
                                Y2 = YtoScreen(Node(.N(J)).Y)

                                cX2 = X2
                                cY2 = Y2

                                CLIPLINEcc cX1, cY1, cX2, cY2

                                If cX1 <> -100 Then

                                    CC.SetDashes 0, 1E+99
                                    IsROAD = Node(.N(J)).dIsWay And Node(.N(J - 1)).dIsWay

                                    '                                    CC.SetLineCap CAIRO_LINE_CAP_ROUND
                                    If IsROAD Then    'MAIN ROAD
                                        '                                        If .isNotAsphalt Then
                                        '                                            CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, 6579300    '100
                                        '                                        Else


                                        If .Layer Then
                                            D = .Layer * 30
                                            Color = RGB(180 + D, 180 + D, 180 + D)
                                        Else
                                            Color = 11842740
                                        End If

                                        '                                        If .Layer <> 0 Then
                                        '                                            If J <> 2 And J <> .NN Then
                                        '                                                CC.DrawLine cX1 - .nDY(J) * RoadW * 0.5, cY1 + .nDX(J) * RoadW * 0.5, _
                                                                                         '                                                            cX2 - .nDY(J) * RoadW * 0.5, cY2 + .nDX(J) * RoadW * 0.5, , Zoom, vbBlack
                                        '                                                CC.DrawLine cX1 + .nDY(J) * RoadW * 0.5, cY1 - .nDX(J) * RoadW * 0.5, _
                                                                                         '                                                            cX2 + .nDY(J) * RoadW * 0.5, cY2 - .nDX(J) * RoadW * 0.5, , Zoom, vbBlack
                                        '                                            End If
                                        '                                        End If

                                        CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, Color    '180




                                        '                                        End If

                                        'Else    'PEDONALE O CICLABILE
                                        '   CC.DrawLine cX1, cY1, cX2, cY2, , RoadW * 0.1, 6579400, 0.5
                                    End If


                                    ''If Node(.N(J)).dCost <100 Then
                                    'CC.SetSourceRGBA 1, 1, 0, 0.5
                                    'CC.Arc X2, Y2, 1 + Node(.N(J)).Traffic * Zoom
                                    'CC.Fill
                                    ''End If


                                    If Zoom > 2# Then    '''' WHATE LINES

                                        If IsROAD Then
                                            '                                            CC.SetLineCap CAIRO_LINE_CAP_BUTT

                                            If .OneWayDirections = 0 Then    'Central Line

                                                CC.SetDashes Zoom * 10, Dash1, Dash2
                                                CC.DrawLine X1, Y1, X2, Y2, , WhiteLineWidth, vbWhite


                                            Else    'SENSO UNICO
                                                '                                                CC.SetLineCap CAIRO_LINE_CAP_ROUND
                                                CC.SetDashes 0, 1E+99

                                                DX = X2 - X1
                                                DY = Y2 - Y1

                                                ANG = .SegAngle(J - 1)
                                                CosA = .SegDX(J - 1)
                                                SinA = .SegDY(J - 1)

                                                D = Sqr(DX * DX + DY * DY)

                                                If .OneWayDirections = 1 Then
                                                    Ang2 = ANG - PIh * 1.7
                                                    CosA2 = Cos(Ang2) * RoadW * 0.5
                                                    SinA2 = Sin(Ang2) * RoadW * 0.5

                                                    Ang3 = ANG + PIh * 1.7
                                                    CosA3 = Cos(Ang3) * RoadW * 0.5
                                                    SinA3 = Sin(Ang3) * RoadW * 0.5

                                                ElseIf .OneWayDirections = -1 Then
                                                    Ang2 = ANG - PIh * 1.7 + PI
                                                    CosA2 = Cos(Ang2) * RoadW * 0.5
                                                    SinA2 = Sin(Ang2) * RoadW * 0.5

                                                    Ang3 = ANG + PIh * 1.7 + PI
                                                    CosA3 = Cos(Ang3) * RoadW * 0.5
                                                    SinA3 = Sin(Ang3) * RoadW * 0.5
                                                End If
                                                C = 0
                                                For L = 1 To D Step Zoom * 3.5
                                                    C = C + 1
                                                    If (C And 7&) = 0 Then
                                                        Xa = X1 + L * CosA
                                                        Ya = Y1 + L * SinA
                                                        Xc = Xa + CosA2
                                                        Yc = Ya + SinA2
                                                        CC.DrawLine Xa, Ya, Xc, Yc, , WhiteLineWidth, vbWhite
                                                        Xc = Xa + CosA3
                                                        Yc = Ya + SinA3
                                                        CC.DrawLine Xa, Ya, Xc, Yc, , WhiteLineWidth, vbWhite
                                                    End If
                                                Next

                                            End If

                                        End If

                                    End If

                                End If

                                X1 = X2
                                Y1 = Y2
                                cX1 = X2
                                cY1 = Y2

                            Next
                        End If
                    End If
                End If
            End With
        Next
        '--------------------------------------------------------------------------------------










        'Dim TN1&
        'Dim TN2&
        '
        '
        '        'STAMPA NOMI--------------------------------------------------------------
        '        If Zoom > 3 Then
        '            CC.SetSourceColor vbWhite
        '            For I = 1 To NWay
        '                With Way(I)
        '                    If LenB(.NAME) Then
        '                        If .IsLeisure Or .IsBuilding Or .wayType <> "" Then
        '                            If .IsLeisure Then
        '                                X3 = .screenCX
        '                                Y3 = .screenCY
        '                            Else
        '                            TN1 = .N(.NN \ 2)
        '                            TN2 = .N(.NN \ 2 + 1)
        '
        '                                X3 = XtoScreen(Node(TN1).X)
        '                                Y3 = YtoScreen(Node(TN1).Y)
        '                                X3 = (X3 + XtoScreen(Node(TN2).X)) * 0.5
        '                                Y3 = (Y3 + YtoScreen(Node(TN2).Y)) * 0.5
        '                            End If
        '
        '                            If InsideScreen(X3, Y3) Then
        '                                CC.TextOut X3, Y3, .NAME
        '                            End If
        '
        '                        End If
        '                    End If
        '                End With
        '            Next
        '        End If



        '        BitBlt pHDC2, 0, 0, scrMaxX, scrMaxY, frmMainPIChDC, 0, 0, vbSrcCopy













        Dim TW#, TH#
        Dim TTW#
        Dim txtL  As Long

        '***************************** NAMES ********************************
        If Zoom > 2.5 Then

            '            CC.SetLineCap CAIRO_LINE_CAP_ROUND

            'DRAW BULIDINGS / BLOCKS /AREAS ---NAME NAMES---------------------------------------
            CC.SetSourceRGBA 0.5, 1, 0.5, 0.4


            For I = 1 To NWay
                With WAY(I)
                    If LenB(.Name) Then
                        If .IsBuilding Or .IsLeisure Or .IsAmenity Or .IsWater Or .IsShop Or .IsNatural Then
                            If BBOverlapping(.BBx1, .BBY1, _
                                             .BBx2, .BBY2, _
                                             scr2WorldX0, scr2WorldY0, _
                                             scr2WorldX1, scr2WorldY1) Then
                                '                            CC.TextOut .CX, .CY, .NAME

                                txtL = Len(.Name)
                                'If .NAME = "Arhena5" Then Stop

                                '                                TW = 8.5 * txtL   ' FONT 10
                                '                                TH = 19
                                '                                TTW = TW
                                '                                While TW > 100
                                '                                    TTW = TW
                                '                                    TW = TW - 100
                                '                                    TH = TH + 17.5
                                '                                Wend
                                '                                TW = TTW
                                '                                If TW > 100 Then TW = 100


                                TW = txtL * 10    ' 8.5
                                TH = 22    '19
                                TTW = TW
                                While TW > 120    '100
                                    TTW = TW
                                    TW = TW - 120    '100
                                    TH = TH + 24    '17.5
                                Wend
                                TW = TTW
                                If TW > 120 Then TW = 120


                                '                                CC.RoundedRect (.CenterX - TW * 0.5), (.CenterY - TH * 0.5), TW, TH, 10
                                '                                CC.Fill

                                CC.DrawText (.CenterX - TW * 0.5), (.CenterY - TH * 0.5), TW, TH, _
                                            .Name, False, vbCenter, 2, 1


                            End If
                        End If
                    End If
                End With
            Next
        End If









        CC.SetDashes 0, 1E+99












    Else
        '     BitBlt frmMainPIChDC, 0, 0, scrMaxX, scrMaxY, pHDC2, 0, 0, vbSrcCopy
    End If
    PanZoomChanged = False

End Sub


''Private Sub DRAWroadCC(ByVal X1 As Double, ByVal Y1 As Double, _
 ''                       ByVal X2 As Double, ByVal Y2 As Double, ByVal w As Double, ByVal OneWay As Long, ByVal IsROAD As Boolean)
''
''    Dim ANG       As Double
''    Dim Ang2      As Double
''    Dim Ang3      As Double
''
''
''    Dim L         As Double
''    Dim dx        As Double
''    Dim dy        As Double
''    Dim D         As Double
''    Dim C         As Long
''    Dim CosA      As Double
''    Dim SinA      As Double
''
''    Dim CosA2     As Double
''    Dim SinA2     As Double
''    Dim CosA3     As Double
''    Dim SinA3     As Double
''
''    Dim Xa        As Double
''    Dim Ya        As Double
''    Dim Xb        As Double
''    Dim Yb        As Double
''    Dim Xc        As Double
''    Dim Yc        As Double
''
''    Dim Color     As Long
''
''    Dim whiteWidth As Double
''    Dim drawWidth As Double
''
''
''
''    If IsROAD Then
''        Color = 11842740          'Asfalto
''        whiteWidth = Zoom * 0.25
''        drawWidth = w
''    Else
''        Color = RGB(90, 60, 60)   '6579400  'Pedonale o ciclabile '200,100,100
''        whiteWidth = Zoom * 0.05
''        drawWidth = w * 0.3
''    End If
''
''    '    Stop
''
''    CC.DrawLine X1, Y1, X2, Y2, , drawWidth, Color    'STRADA GRIGIA
''
''    '   If Zoom > 1 Then
''    '        If IsROAD Then
''    '            DX = X2 - X1
''    '            DY = Y2 - Y1
''    '            Ang = Atan2(DX, DY)
''    '            CosA = Cos(Ang)
''    '            SinA = Sin(Ang)
''    '
''    '            D = Sqr(DX * DX + DY * DY)
''    '
''    '            If OneWay = 1 Then
''    '                Ang2 = Ang - PIh * 1.7
''    '                CosA2 = Cos(Ang2) * w * 1
''    '                SinA2 = Sin(Ang2) * w * 1
''    '
''    '                Ang3 = Ang + PIh * 1.7
''    '                CosA3 = Cos(Ang3) * w * 1
''    '                SinA3 = Sin(Ang3) * w * 1
''    '
''    '            ElseIf OneWay = -1 Then
''    '                Ang2 = Ang - PIh * 1.7 + PI
''    '                CosA2 = Cos(Ang2) * w * 1
''    '                SinA2 = Sin(Ang2) * w * 1
''    '
''    '                Ang3 = Ang + PIh * 1.7 + PI
''    '                CosA3 = Cos(Ang3) * w * 1
''    '                SinA3 = Sin(Ang3) * w * 1
''    '            End If
''
''
''    '            For L = 1 To D Step Zoom * 3
''    '                C = C + 1
''    '                Xb = Xa
''    '                Yb = Ya
''    '                Xa = X1 + L * CosA
''    '                Ya = Y1 + L * SinA
''    '                If C Mod 4 = 0 Then
''    '                    If OneWay = 0 Then
''    '                        '                        FastLine frmMainPIChDC, Xa, Ya, Xb, Yb, whiteWidth, vbWhite
''    '                        CC.DrawLine Xa, Ya, Xb, Yb, , whiteWidth, vbWhite
''    '
''    '                    Else
''    '                        Xc = Xa + CosA2
''    '                        Yc = Ya + SinA2
''    '                        '                        FastLine frmMainPIChDC, Xa, Ya, Xc, Yc, whiteWidth, vbWhite
''    '                        CC.DrawLine Xa, Ya, Xc, Yc, , whiteWidth, vbWhite
''    '                        Xc = Xa + CosA3
''    '                        Yc = Ya + SinA3
''    '                        '                        FastLine frmMainPIChDC, Xa, Ya, Xc, Yc, whiteWidth, vbWhite
''    '                        CC.DrawLine Xa, Ya, Xc, Yc, , whiteWidth, vbWhite
''    '                    End If
''    '
''    '                End If
''    '            Next
''    '        End If
''    '    End If
''
''
''
''End Sub
'Public Function BoundsOverlaps(BoundsA As tBounds, boundsB As tBounds) As Boolean
'
'    If BoundsA.Max.X < boundsB.Min.X Then Exit Function
'    If BoundsA.Max.Y < boundsB.Min.Y Then Exit Function
'    If BoundsA.Min.X > boundsB.Max.X Then Exit Function
'    If BoundsA.Min.Y > boundsB.Max.Y Then Exit Function
'
'    BoundsOverlaps = True
'
'End Function

Public Function BBOverlapping(ByVal AminX#, ByVal AminY#, ByVal AMaxX#, ByVal AMaxY#, _
                              ByVal BminX#, ByVal BminY#, ByVal BMaxX#, ByVal BMaxY#)


    If AMaxX < BminX Then Exit Function
    If AMaxY < BminY Then Exit Function
    If AminX > BMaxX Then Exit Function
    If AminY > BMaxY Then Exit Function
    BBOverlapping = True

End Function




Public Sub LoadImages()

    Dim S         As String
    Dim SP        As String
    Dim I         As Long
    SP = App.Path & "\Images\"


    S = Dir(SP)
    I = 0
    Do
        Cairo.ImageList.AddImage "C" & CStr(I), SP & S, (CarInterasse + CarWidth) * CarImageRes, CarWidth * CarImageRes, False

        S = Dir
        I = I + 1
    Loop While Len(S)


End Sub



Public Sub SETPROGRESS(Text As String, Optional Perc#)
    Dim X1#, Y1#, X2#, Y2

    X1 = scrMaxX * 0.5 - 200: X2 = scrMaxX * 0.5 + 200
    Y1 = scrMaxY * 0.5 - 40: Y2 = scrMaxY * 0.5 + 40

    With CC
        .SetSourceRGB 0.2, 0.7, 0.2
        .RoundedRect X1, Y1, X2 - X1, Y2 - Y1, 15: .Fill
        .DrawText X1, Y1, X2 - X1, Y2 - Y1 - 35, Text, , vbCenter, , 1

        If Perc Then
            .SetSourceRGB 0.5, 0.5, 0.5
            .RoundedRect X1 + 10, Y2 - 30, (X2 - X1) - 20, 20, 20: .Fill
            .SetSourceRGB 0.1, 0.5, 0.1
            .RoundedRect X1 + 10, Y2 - 30, ((X2 - X1) - 20) * Perc, 20, 20: .Fill


            .DrawText X1 + 10, Y2 - 30, (X2 - X1) - 20, 20, Format$(Perc * 100, "0.0") & "%", , vbCenter
        End If

    End With

    SRF.DrawToDC frmMainPIChDC


End Sub



























Public Sub DRAWMAPandCARS()
    Dim I         As Long
    Dim J As Long, K As Long

    Dim X1        As Double
    Dim Y1        As Double

    Dim X2        As Double
    Dim Y2        As Double

    Dim cX1       As Double
    Dim cY1       As Double
    Dim cX2       As Double
    Dim cY2       As Double

    Dim Dash1#, Dash2#
    Dim IsROAD    As Boolean

    Dim DX#, DY#
    Dim Xa#, Ya#
    Dim Xc#, Yc#

    Dim ANG#, CosA#, SinA#
    Dim Ang2#, CosA2#, SinA2#
    Dim Ang3#, CosA3#, SinA3#
    Dim D#
    Dim L#


    Dim C         As Long

    Dim RoadW     As Double

    Dim SolidPatternLongBLACK As cCairoPattern

    Dim Color     As Long

    Dim WhiteLineWidth As Double

    C = vbRed


    Dim X#, Y#

    '    Dim PTS()     As Single       'POINTAPI


    WhiteLineWidth = Zoom * 0.3





    Set SolidPatternLongBLACK = Cairo.CreateSolidPatternLng(vbBlack)


    STATSegs = 0
    STATbuild = 0
    STATPolyLines = 0
    STATCars = 0


    Dash1 = Zoom * 4                    '4          '2.5
    Dash2 = Zoom * 8                    '15         '*8


    scr2WorldX0 = xfromScreen(ClipLEFT): scr2WorldY0 = yfromScreen(ClipTOP)
    scr2WorldX1 = xfromScreen(ClipRIGHT): scr2WorldY1 = yfromScreen(ClipBOTTOM)



    'CC.SetSourceColor 0: CC.Paint
    '        CC.SetSourceRGB 0.23, 0.35, 0.23: CC.Paint
    CC.SetSourceRGB 0.1, 0.43, 0.1: CC.Paint


    Dim Layer     As Long




    'DRAW BULIDINGS / BLOCKS /AREAS ------------------------------------------
    CC.SetLineWidth 0.35 * Zoom
    CC.SetDashes 0, 1E+99


    '''CameraAng = CameraAng * 0.98 - 0.02 * AngleDIFF((-CAR(Follow).ANG - PIh), 0)
    ''CameraAng = CameraAng * 0.99 - 0.01 * AngleDIFF180((-CAR(Follow).ANG), PIh)
    ''CC.save
    ''CC.TranslateDrawings scrMaxX * 0.5, scrMaxY * 0.5
    ''CC.RotateDrawings CameraAng
    ''CC.TranslateDrawings -scrMaxX * 0.5, -scrMaxY * 0.5



    For I = 1 To NWay
        With WAY(I)
            If .IsBuilding Or .IsLeisure Or .IsAmenity Or .IsWater Or .IsShop Or _
               .IsNatural Or .IsBoundary Then
                If .IsBuilding Then
                    'CC.SetSourceRGB 0.6, 0.3, 0.2 ' BUILDING
                    CC.SetSourceRGB 0.5 + .R * 0.08, 0.2 + .G * 0.08, 0.1 + .B = 0.08

                End If
                If .IsLeisure Then
                    If .Name = "Swimming Pool" Then
                        CC.SetSourceRGB 0.1, 1, 1
                    Else
                        CC.SetSourceRGB 0.3, 0.7, 0.3
                    End If
                End If
                If .IsWater Then
                    CC.SetSourceRGB 0, 0.5, 1    '.2, 0.2, 1
                End If

                If .IsAmenity Then
                    CC.SetSourceRGB 0.15, 0.15, 0.15
                End If

                If .IsShop Then
                    CC.SetSourceRGB 0.65, 0.65, 0.65
                End If
                '
                '                    If .IsNatural Then
                '                        CC.SetSourceRGB 1, 0, 0
                '                    End If

                If BBOverlapping(.BBx1, .BBY1, _
                                 .BBx2, .BBY2, _
                                 scr2WorldX0, scr2WorldY0, _
                                 scr2WorldX1, scr2WorldY1) Then

                    K = 0
                    ReDim PTS(.NN * 2 - 1)
                    X = 0
                    Y = 0
                    STATbuild = STATbuild + 1
                    'https://www.vbforums.com/showthread.php?902225-Stretch-a-Polygon-region&p=5627915&viewfull=1#post5627915

                    BuildingPolygon.removeAll

                    For J = 1 To .NN
                        X1 = XtoScreen(Node(.N(J)).X)
                        Y1 = YtoScreen(Node(.N(J)).Y)
                        'PTS(K) = X1
                        'PTS(K + 1) = Y1
                        BuildingPolygon.Add X1
                        BuildingPolygon.Add Y1
                        X = X + X1
                        Y = Y + Y1
                        K = K + 2
                    Next
                    '                        CC.PolygonSingle PTS, False, splNone, , True
                    CC.PolygonPtr BuildingPolygon.DataPtr, BuildingPolygon.Count \ 2, False, splNone

                    If .IsBoundary Then
                        CC.save
                        CC.SetLineWidth Zoom * 1.5
                        CC.Stroke , Cairo.CreateSolidPatternLng(vbYellow)
                        'Stop
                        CC.Restore

                    Else                'Building /Area / Block

                        If Zoom > 0.85 Then
                            CC.Fill True
                            CC.Stroke , SolidPatternLongBLACK
                        Else
                            CC.Fill
                        End If

                    End If

                    .CenterX = X * .invNN    ' Per Stampa Names
                    .CenterY = Y * .invNN    ' CC.TextOut .CX, .CY, .NAME

                    STATPolyLines = STATPolyLines + BuildingPolygon.Count \ 2
                End If


            End If

        End With
    Next
    '----------------------------------------------------------------------------






    '---------------------------------DRAW ROADS-----------------------------------------
    '---------------------------------Black Border
    ''    CC.SetDashes 0, 1e+99
    ''    If Zoom > 2 Then
    ''        For I = 1 To NWay
    ''            With WAY(I)
    ''                If .IsArea Or .IsBuilding Or .IsWater Or _
     ''                   .IsAmenity Or .IsLeisure Or .IsShop Or .IsNatural Then
    ''                Else
    ''                    If LenB(.wayType) Then
    ''                        If BBOverlapping(.BBx1, .BBY1, _
     ''                                         .BBx2, .BBY2, _
     ''                                         scr2WorldX0, scr2WorldY0, _
     ''                                         scr2WorldX1, scr2WorldY1) Then
    ''
    ''                            X1 = XtoScreen(Node(.N(1)).X)
    ''                            Y1 = YtoScreen(Node(.N(1)).Y)
    ''                            cX1 = X1
    ''                            cY1 = Y1
    ''                            If .RoadWidth = 0 Then
    ''                                Stop
    ''
    ''                                RoadW = kRoadW * Zoom
    ''                            Else
    ''                                RoadW = .RoadWidth * Zoom
    ''                            End If
    ''                            RoadW = RoadW * 0.5 * .Lanes + Zoom    '1 meter larger
    ''
    ''                            For J = 2 To .NN
    ''                                X2 = XtoScreen(Node(.N(J)).X)
    ''                                Y2 = YtoScreen(Node(.N(J)).Y)
    ''                                cX2 = X2
    ''                                cY2 = Y2
    ''
    ''                                CLIPLINEcc cX1, cY1, cX2, cY2
    ''
    ''                                If cX1 <> -100 Then
    ''                                    IsROAD = Node(.N(J)).dIsWay And Node(.N(J - 1)).dIsWay
    ''                                    If IsROAD Then
    ''                                        CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, 3289650    'vbBlack
    ''                                    End If
    ''                                End If
    ''                                X1 = X2
    ''                                Y1 = Y2
    ''                                cX1 = X2
    ''                                cY1 = Y2
    ''                            Next
    ''                        End If
    ''
    ''                    End If
    ''                End If
    ''
    ''            End With
    ''        Next
    ''    End If
    '---------- Side Parking -----------------

    For Layer = minLayer To maxLayer
        For I = 1 To NWay
            With WAY(I)
                If .Layer = Layer Then
                    If .IsStreetSideParkable Then
                        If .isDRIVEABLE Then
                            If BBOverlapping(.BBx1, .BBY1, _
                                             .BBx2, .BBY2, _
                                             scr2WorldX0, scr2WorldY0, _
                                             scr2WorldX1, scr2WorldY1) Then

                                X1 = XtoScreen(Node(.N(2)).X)
                                Y1 = YtoScreen(Node(.N(2)).Y)
                                cX1 = X1
                                cY1 = Y1
                                RoadW = .RoadWidth * Zoom
                                RoadW = RoadW * .Lanes
                                For J = 3 To .NN - 1

                                    cX2 = XtoScreen(Node(.N(J)).X)
                                    cY2 = YtoScreen(Node(.N(J)).Y)
                                    If CLIPLINEcc(cX1, cY1, cX2, cY2) Then
                                        STATSegs = STATSegs + 1
                                        CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, &H303030

                                    End If
                                    cX1 = cX2
                                    cY1 = cY2
                                Next
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next

    '---------- ASFALTO---- ROADS -----------------
    For Layer = minLayer To maxLayer
        For I = 1 To NWay
            With WAY(I)
                If .Layer = Layer Then
                    If .IsArea Or .IsBuilding Or .IsWater Or _
                       .IsAmenity Or .IsLeisure Or .IsShop Or .IsNatural Then
                    Else

                        If LenB(.wayType) Then
                            If BBOverlapping(.BBx1, .BBY1, _
                                             .BBx2, .BBY2, _
                                             scr2WorldX0, scr2WorldY0, _
                                             scr2WorldX1, scr2WorldY1) Then

                                X1 = XtoScreen(Node(.N(1)).X)
                                Y1 = YtoScreen(Node(.N(1)).Y)
                                cX1 = X1
                                cY1 = Y1
                                RoadW = .RoadWidth * Zoom
                                RoadW = RoadW * 0.5 * .Lanes

                                For J = 2 To .NN



                                    X2 = XtoScreen(Node(.N(J)).X)
                                    Y2 = YtoScreen(Node(.N(J)).Y)

                                    cX2 = X2
                                    cY2 = Y2

                                    '                                    CLIPLINEcc cX1, cY1, cX2, cY2

                                    If CLIPLINEcc(cX1, cY1, cX2, cY2) Then

                                        CC.SetDashes 0, 1E+99
                                        IsROAD = Node(.N(J)).dIsWay And Node(.N(J - 1)).dIsWay

                                        '                                    CC.SetLineCap CAIRO_LINE_CAP_ROUND
                                        If IsROAD Then    'MAIN ROAD
                                            '                                        If .isNotAsphalt Then
                                            '                                            CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, 6579300    '100
                                            '                                        Else


                                            If .Layer Then
                                                D = .Layer * 30
                                                Color = RGB(180 + D, 180 + D, 180 + D)
                                            Else
                                                Color = 11842740
                                            End If

                                            If .isNotAsphalt Then Color = RGB(180, 100, 50)


                                            '                                        If .Layer <> 0 Then
                                            '                                            If J <> 2 And J <> .NN Then
                                            '                                                CC.DrawLine cX1 - .nDY(J) * RoadW * 0.5, cY1 + .nDX(J) * RoadW * 0.5, _
                                                                                             '                                                            cX2 - .nDY(J) * RoadW * 0.5, cY2 + .nDX(J) * RoadW * 0.5, , Zoom, vbBlack
                                            '                                                CC.DrawLine cX1 + .nDY(J) * RoadW * 0.5, cY1 - .nDX(J) * RoadW * 0.5, _
                                                                                             '                                                            cX2 + .nDY(J) * RoadW * 0.5, cY2 - .nDX(J) * RoadW * 0.5, , Zoom, vbBlack
                                            '                                            End If
                                            '                                        End If
                                            STATSegs = STATSegs + 1
                                            CC.DrawLine cX1, cY1, cX2, cY2, , RoadW, Color    '180




                                            '                                        End If




                                            ''If Node(.N(J)).dCost <100 Then
                                            'CC.SetSourceRGBA 1, 1, 0, 0.5
                                            'CC.Arc X2, Y2, 1 + Node(.N(J)).Traffic * Zoom
                                            'CC.Fill
                                            ''End If


                                            If Zoom > 2# Then    '''' WHITE LINES
                                                If Not (.isNotAsphalt) Then
                                                    '                                            CC.SetLineCap CAIRO_LINE_CAP_BUTT

                                                    If .OneWayDirections = 0 Then    'Central Line

                                                        CC.SetDashes Zoom * 10, Dash1, Dash2
                                                        CC.DrawLine X1, Y1, X2, Y2, , WhiteLineWidth, vbWhite

                                                    Else    'SENSO UNICO
                                                        '                                                CC.SetLineCap CAIRO_LINE_CAP_ROUND
                                                        CC.SetDashes 0, 1E+99

                                                        DX = X2 - X1
                                                        DY = Y2 - Y1

                                                        ANG = .SegAngle(J - 1)
                                                        CosA = .SegDX(J - 1)
                                                        SinA = .SegDY(J - 1)

                                                        D = Sqr(DX * DX + DY * DY)

                                                        If .OneWayDirections = 1 Then
                                                            Ang2 = ANG - PIh * 1.7
                                                            CosA2 = Cos(Ang2) * RoadW * 0.5
                                                            SinA2 = Sin(Ang2) * RoadW * 0.5

                                                            Ang3 = ANG + PIh * 1.7
                                                            CosA3 = Cos(Ang3) * RoadW * 0.5
                                                            SinA3 = Sin(Ang3) * RoadW * 0.5

                                                        ElseIf .OneWayDirections = -1 Then
                                                            Ang2 = ANG - PIh * 1.7 + PI
                                                            CosA2 = Cos(Ang2) * RoadW * 0.5
                                                            SinA2 = Sin(Ang2) * RoadW * 0.5

                                                            Ang3 = ANG + PIh * 1.7 + PI
                                                            CosA3 = Cos(Ang3) * RoadW * 0.5
                                                            SinA3 = Sin(Ang3) * RoadW * 0.5
                                                        End If
                                                        C = 0
                                                        For L = 1 To D Step Zoom * 3.5
                                                            C = C + 1
                                                            If (C And 7&) = 0 Then
                                                                Xa = X1 + L * CosA
                                                                Ya = Y1 + L * SinA
                                                                Xc = Xa + CosA2
                                                                Yc = Ya + SinA2
                                                                CC.DrawLine Xa, Ya, Xc, Yc, , WhiteLineWidth, vbWhite
                                                                Xc = Xa + CosA3
                                                                Yc = Ya + SinA3
                                                                CC.DrawLine Xa, Ya, Xc, Yc, , WhiteLineWidth, vbWhite
                                                            End If
                                                        Next


                                                    End If
                                                End If
                                            End If

                                        Else    ' NOT ROAD
                                            If .IsRailWay Then    'TRAIN path
                                                STATSegs = STATSegs + 1
                                                CC.DrawLine cX1, cY1, cX2, cY2, , RoadW * 0.5, 0
                                            Else    'PEDONALE O CICLABILE
                                                STATSegs = STATSegs + 1
                                                CC.DrawLine cX1, cY1, cX2, cY2, , RoadW * 0.08, 6579400    ', 0.5
                                            End If
                                        End If

                                    End If

                                    X1 = X2
                                    Y1 = Y2
                                    cX1 = X2
                                    cY1 = Y2

                                Next
                            End If
                        End If
                    End If
                End If

            End With
        Next

        CC.SetDashes 0, 1E+99


        For I = 1 To Ncars
            If CAR(I).Layer = Layer Then CAR(I).DRAWCC
        Next

    Next Layer

    '--------------------------------------------------------------------------------------




    'For I = 1 To NNode
    'X = XtoScreen(Node(I).X)
    'Y = YtoScreen(Node(I).Y)
    'If InsideScreen(X, Y) Then
    'CC.TextOut X, Y, CStr(Node(I).ONEWAY)
    'End If
    'Next










    'Dim TN1&
    'Dim TN2&
    '
    '
    '        'STAMPA NOMI--------------------------------------------------------------
    '        If Zoom > 3 Then
    '            CC.SetSourceColor vbWhite
    '            For I = 1 To NWay
    '                With Way(I)
    '                    If LenB(.NAME) Then
    '                        If .IsLeisure Or .IsBuilding Or .wayType <> "" Then
    '                            If .IsLeisure Then
    '                                X3 = .screenCX
    '                                Y3 = .screenCY
    '                            Else
    '                            TN1 = .N(.NN \ 2)
    '                            TN2 = .N(.NN \ 2 + 1)
    '
    '                                X3 = XtoScreen(Node(TN1).X)
    '                                Y3 = YtoScreen(Node(TN1).Y)
    '                                X3 = (X3 + XtoScreen(Node(TN2).X)) * 0.5
    '                                Y3 = (Y3 + YtoScreen(Node(TN2).Y)) * 0.5
    '                            End If
    '
    '                            If InsideScreen(X3, Y3) Then
    '                                CC.TextOut X3, Y3, .NAME
    '                            End If
    '
    '                        End If
    '                    End If
    '                End With
    '            Next
    '        End If


    ''    Dim TN1&, TN2&
    ''Dim A#
    ''
    ''    If Zoom > 3 Then
    ''    CC.SetLineWidth 1
    ''    CC.SetSourceColor 0
    ''
    ''        For I = 1 To NWay
    ''            With WAY(I)
    ''                If LenB(.NAME) Then
    ''                    'If .IsLeisure Or .IsBuilding Or .wayType <> "" Then
    ''                    If LenB(.wayType) Then
    ''
    ''                        TN1 = .N(.NN \ 2)
    ''                        TN2 = .N(.NN \ 2 + 1)
    ''
    ''                        X3 = XtoScreen(Node(TN1).X)
    ''                        Y3 = YtoScreen(Node(TN1).Y)
    ''                        X3 = (X3 + XtoScreen(Node(TN2).X)) * 0.5
    ''                        Y3 = (Y3 + YtoScreen(Node(TN2).Y)) * 0.5
    ''
    ''                        If InsideScreen(X3, Y3) Then
    ''
    ''                        CC.save
    ''
    ''                        CC.TranslateDrawings X3, Y3
    ''                        A = .NA(.NN \ 2)
    ''                        While A < -PIh: A = A + PI: Wend
    ''                        While A > PIh: A = A - PI: Wend
    ''
    ''
    ''                        CC.RotateDrawings A
    ''                            CC.DrawText 0, -Zoom * kRoadW * 0.5, 200, 30, .NAME, True, , , , , , True
    ''                            CC.Stroke
    ''                        CC.Restore
    ''
    ''
    ''                        End If
    ''
    ''                    End If
    ''                End If
    ''            End With
    ''        Next
    ''
    ''    End If
    ''












    Dim TW#, TH#
    Dim TTW#
    Dim txtL      As Long

    '***************************** NAMES ********************************
    If Zoom > 2.5 Then

        '            CC.SetLineCap CAIRO_LINE_CAP_ROUND

        'DRAW BULIDINGS / BLOCKS /AREAS ---NAME NAMES---------------------------------------
        '        CC.SetSourceRGBA 0.5, 1, 0.5, 0.4

        '        CC.SetLineWidth 0.7
        For I = 1 To NWay
            With WAY(I)
                If LenB(.Name) Then
                    If .IsBuilding Or .IsLeisure Or .IsAmenity Or .IsWater Or .IsShop Or .IsNatural Then
                        If BBOverlapping(.BBx1, .BBY1, _
                                         .BBx2, .BBY2, _
                                         scr2WorldX0, scr2WorldY0, _
                                         scr2WorldX1, scr2WorldY1) Then

                            txtL = Len(.Name)

                            TW = txtL * 10    ' 8.5
                            TH = 22     '19
                            TTW = TW
                            While TW > 120    '100
                                TTW = TW
                                TW = TW - 120    '100
                                TH = TH + 24    '17.5
                            Wend
                            TW = TTW
                            If TW > 120 Then TW = 120

                            'CC.SetSourceRGBA 0.5, 1, 0.5, 0.4
                            '                            CC.RoundedRect (.CenterX - TW * 0.5), (.CenterY - TH * 0.5), TW, TH, 10
                            '                            CC.Fill
                            'CC.SetSourceColor vbWhite
                            CC.DrawText (.CenterX - TW * 0.5), (.CenterY - TH * 0.5), TW, TH, _
                                        .Name, False, vbCenter, 2, 1    ', , , True
                            'CC.Stroke

                        End If
                    End If
                End If
            End With
        Next
    End If

    'CC.Restore


End Sub

