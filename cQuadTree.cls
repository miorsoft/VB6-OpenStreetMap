VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQuadTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X0         As Double        'Quad Boundary
Public Y0         As Double
Public X1         As Double
Public Y1         As Double

Private MX        As Double        'Quad Center
Private MY        As Double

Private Q1        As cQuadTree
Private Q2        As cQuadTree
Private Q3        As cQuadTree
Private Q4        As cQuadTree

'Private QUA(1 To 4) As cQuadTree

Private Subdivided As Long

Public Capacity   As Long
Public NP         As Long          ' N of Points
Private MaxNP     As Long          ' (Obsolte)

Private px()      As Double        'Points Coordinates
Private py()      As Double
Private pIDX()    As Double        'Points Indexes
'"Ghost Points" used in GetPairsWDist
Private GPX()     As Double
Private GPY()     As Double
Public GhostNP    As Long
Private GhostmaxNP As Long




Friend Sub Setup(ByVal Xleft#, ByVal Ytop#, ByVal Xright#, ByVal Ydown#, ByVal MyCapacity As Long)
    X0 = Xleft
    X1 = Xright
    Y0 = Ytop
    Y1 = Ydown
    Capacity = MyCapacity
    Reset
End Sub
Friend Sub Reset()

    ' Set N of point to 0
    ' and continue with child nodes if they have points into them
    ' So it's not needed to go through all level ( when they are empty)
    NP = 0
    GhostNP = 0
    Dim I         As Long
    If Subdivided Then
        If Q1.NP Then Q1.Reset
        If Q2.NP Then Q2.Reset
        If Q3.NP Then Q3.Reset
        If Q4.NP Then Q4.Reset
        '        For I = 1 To 4
        '            If QUA(I).NP Then QUA(I).Reset
        '        Next

    End If
End Sub

Private Sub Subdivide()
    ' Subdivede the Quadtree
    Dim I         As Long

    'Compute Center Point
    MX = (X0 + X1) * 0.5
    MY = (Y0 + Y1) * 0.5

    Subdivided = -1

    Set Q1 = New cQuadTree
    Set Q2 = New cQuadTree
    Set Q3 = New cQuadTree
    Set Q4 = New cQuadTree
    '    If QUA(1) Is Nothing Then
    '        For I = 1 To 4
    '            Set QUA(I) = New cQuadTree
    '        Next
    '    End If

    ' Setup the 4 chiled nodes
    Q1.Setup X0, Y0, MX, MY, Capacity    'TopLeft
    Q2.Setup MX, Y0, X1, MY, Capacity    'TopRight
    Q3.Setup X0, MY, MX, Y1, Capacity    'BottomLeft
    Q4.Setup MX, MY, X1, Y1, Capacity    'BottmRight

    '    QUA(1).Setup X0, Y0, MX, MY, Capacity
    '    QUA(2).Setup MX, Y0, X1, MY, Capacity
    '    QUA(3).Setup X0, MY, MX, Y1, Capacity
    '    QUA(4).Setup MX, MY, X1, Y1, Capacity


End Sub

Friend Sub InsertSinglePoint(ByVal X#, ByVal Y#, Index As Long, Optional isRoot As Boolean = True)

    Dim Quadrant  As Long


' This add a copy of inserted points
' It is used in GetPairsWDist
' Becose today I havent found another way to do GetPairsWDist

    If isRoot Then                 ' FOR GetPairsWDist
        GhostNP = GhostNP + 1
        If GhostNP > GhostmaxNP Then
            GhostmaxNP = GhostNP * 2
            ReDim Preserve GPX(GhostmaxNP)
            ReDim Preserve GPY(GhostmaxNP)
        End If
        GPX(GhostNP) = X
        GPY(GhostNP) = Y
    End If

'---------------------------------------------------------------------
' STANDARD QUADTREE

    If NP < Capacity Then 'IF N of Points less than Capacity insert it
        NP = NP + 1
        If NP > MaxNP Then
        ' Redim to Capacity
        ' I found Better to do it Once instead of redim Often based on NP
            MaxNP = Capacity
            ReDim px(1 To Capacity)
            ReDim py(1 To Capacity)
            ReDim pIDX(1 To Capacity)
        End If
        px(NP) = X 'Inser Point
        py(NP) = Y
        pIDX(NP) = Index
    Else 'N of point is Greater or equal to capacity, so no space to add one

        If Not (Subdivided) Then Subdivide

        'Compute in which child quandrant we will insert the point
        If X > MX Then Quadrant = 1&
        If Y > MY Then Quadrant = Quadrant Or 2&


        '        Select Case Quadrant
        '            Case 0: Q1.InsertSinglePoint X, Y, Index, False
        '            Case 1: Q2.InsertSinglePoint X, Y, Index, False
        '            Case 2: Q3.InsertSinglePoint X, Y, Index, False
        '            Case 3: Q4.InsertSinglePoint X, Y, Index, False
        '        End Select


        'Insert point in Appropriate Quadrant
        If Quadrant = 0& Then
            Q1.InsertSinglePoint X, Y, Index, False
        ElseIf Quadrant = 1& Then
            Q2.InsertSinglePoint X, Y, Index, False
        ElseIf Quadrant = 2& Then
            Q3.InsertSinglePoint X, Y, Index, False
        ElseIf Quadrant = 3& Then
            Q4.InsertSinglePoint X, Y, Index, False
        End If




        '        QUA(Quadrant + 1).InsertSinglePoint X, Y, Index, False

    End If

End Sub



Friend Sub InsertALLpoints(X() As Double, Y() As Double, Optional PTBound As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Reset
    If PTBound = 0 Then PTBound = UBound(X)
    For I = 1 To PTBound
        InsertSinglePoint X(I), Y(I), I, isRoot
    Next

End Sub


Friend Sub QueryAroundPoint(ByVal CurrPointX#, ByVal CurrPointY#, ByVal CurrPointIDX As Long, ByVal R#, _
                            RP1() As Long, _
                            RP2() As Long, _
                            Rdx() As Double, _
                            Rdy() As Double, _
                            rDD() As Double, _
                            rPairsCount As Long, MaxNpairs As Long)

' Point coordinates and index is given as input.
' It returns The Points Indexes
' the DX , DY and DD Distances array of point whithin the R distance from given point

    Dim I         As Long
    Dim DX#, DY#, R2#, DD#
    If CurrPointX + R < X0 Then Exit Sub
    If CurrPointY + R < Y0 Then Exit Sub
    If CurrPointX - R > X1 Then Exit Sub
    If CurrPointY - R > Y1 Then Exit Sub

    R2 = R * R

    For I = 1 To NP
        If pIDX(I) > CurrPointIDX Then
            DX = px(I) - CurrPointX: DY = py(I) - CurrPointY: DD = (DX * DX + DY * DY)
            If DD < R2 Then
                rPairsCount = rPairsCount + 1
                If rPairsCount > MaxNpairs Then
                    MaxNpairs = rPairsCount * 2
                    ReDim Preserve RP1(MaxNpairs)
                    ReDim Preserve RP2(MaxNpairs)
                    ReDim Preserve Rdx(MaxNpairs)
                    ReDim Preserve Rdy(MaxNpairs)
                    ReDim Preserve rDD(MaxNpairs)
                End If
                RP1(rPairsCount) = CurrPointIDX
                RP2(rPairsCount) = pIDX(I)
                Rdx(rPairsCount) = DX
                Rdy(rPairsCount) = DY
                rDD(rPairsCount) = DD
            End If
        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QueryAroundPoint CurrPointX, CurrPointY, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
        If Q2.NP Then Q2.QueryAroundPoint CurrPointX, CurrPointY, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
        If Q3.NP Then Q3.QueryAroundPoint CurrPointX, CurrPointY, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
        If Q4.NP Then Q4.QueryAroundPoint CurrPointX, CurrPointY, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs

        '        For I = 1 To 4
        '            If QUA(I).NP Then QUA(I).QueryAroundPoint CurrPointX, CurrPointY, CurrPointIDX, R, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
        '        Next


    End If

End Sub



Friend Sub QueryPointsInCircle(ByVal cX As Double, ByVal cY As Double, ByVal R As Double, _
                               rdist() As Double, _
                               rIdx() As Long, _
                               rPntCount As Long, MaxrPntCount As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Dim DX#, DY#, R2#, DD#
    If cX + R < X0 Then Exit Sub
    If cY + R < Y0 Then Exit Sub
    If cX - R > X1 Then Exit Sub
    If cY - R > Y1 Then Exit Sub

    R2 = R * R

    If isRoot Then rPntCount = 0   ': MaxrPntCount = 0

    For I = 1 To NP
        DX = px(I) - cX: DY = py(I) - cY: DD = (DX * DX + DY * DY)
        If DD < R2 Then
            rPntCount = rPntCount + 1
            If rPntCount > MaxrPntCount Then
                MaxrPntCount = rPntCount * 2
                ReDim Preserve rIdx(MaxrPntCount)
                ReDim Preserve rdist(MaxrPntCount)
            End If
            rdist(rPntCount) = DD
            rIdx(rPntCount) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QueryPointsInCircle cX, cY, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        If Q2.NP Then Q2.QueryPointsInCircle cX, cY, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        If Q3.NP Then Q3.QueryPointsInCircle cX, cY, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        If Q4.NP Then Q4.QueryPointsInCircle cX, cY, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '
        '        For I = 1 To 4
        '            If QUA(I).NP Then QUA(I).QueryPointsInCircle cX, cY, R, rdist, rIdx, rPntCount, MaxrPntCount, False
        '        Next

    End If

End Sub
Friend Sub QuerySquare(cX As Double, _
                       cY As Double, _
                       halfSide As Double, _
                       rpX() As Double, rpY() As Double, rpIDX() As Long, Found As Long, MaxFound As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Dim DX#, DY#
    If cX + halfSide < X0 Then Exit Sub
    If cY + halfSide < Y0 Then Exit Sub
    If cX - halfSide > X1 Then Exit Sub
    If cY - halfSide > Y1 Then Exit Sub


    If isRoot Then Found = 0       ': MaxFound = 0

    For I = 1 To NP
        DX = px(I) - cX: DY = py(I) - cY
        If Abs(DX) < halfSide And Abs(DY) < halfSide Then
            Found = Found + 1
            If Found > MaxFound Then
                MaxFound = Found * 2
                ReDim Preserve rpIDX(MaxFound)
                ReDim Preserve rpX(MaxFound)
                ReDim Preserve rpY(MaxFound)

            End If
            rpX(Found) = px(I)
            rpY(Found) = py(I)
            rpIDX(Found) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q2.NP Then Q2.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q3.NP Then Q3.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q4.NP Then Q4.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False

        '        For I = 1 To 4
        '            If QUA(I).NP Then QUA(I).QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        '        Next

    End If


End Sub

Friend Sub QueryCircleOLD(ByVal cX As Double, ByVal cY As Double, ByVal R As Double, _
                          rpX() As Double, rpY() As Double, rpIDX() As Long, Found As Long, MaxFound As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Dim DX#, DY#, R2#, DD#
    If cX + R < X0 Then Exit Sub
    If cY + R < Y0 Then Exit Sub
    If cX - R > X1 Then Exit Sub
    If cY - R > Y1 Then Exit Sub

    R2 = R * R

    If isRoot Then Found = 0       ': maxfound = 0

    For I = 1 To NP
        DX = px(I) - cX: DY = py(I) - cY: DD = (DX * DX + DY * DY)
        If DD < R2 Then
            Found = Found + 1
            If Found > MaxFound Then
                MaxFound = Found * 2
                ReDim Preserve rpIDX(MaxFound)
                ReDim Preserve rpX(MaxFound)
                ReDim Preserve rpY(MaxFound)

            End If
            rpX(Found) = px(I)
            rpY(Found) = py(I)
            rpIDX(Found) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q2.NP Then Q2.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q3.NP Then Q3.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q4.NP Then Q4.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        '        For I = 1 To 4
        '            If QUA(I).NP Then QUA(I).QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        '        Next

    End If

End Sub
Friend Sub GetPairsWDist(ByVal Radius As Double, RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, rDD() As Double, rPairsCount As Long, MaxNpairs As Long)
    Dim I         As Long
    '  USING GHOST
    '--------------------------------------------------
    rPairsCount = 0
    For I = 1 To GhostNP
        QueryAroundPoint GPX(I), GPY(I), I, Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    Next
    '--------------------------------------------------
    '   Without GHOST
    ' Miss pairs Between Qaudtree "Lines"  [Can Someone fix it?]
    '    If ImRoot Then rPairsCount = 0
    '    For I = 1 To NP
    '        QueryAroundPoint PX(I), PY(I), pIDX(I), Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    '    Next
    '    If Subdivided Then
    '        If Q1.NP Then Q1.GetPairsWDist Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    '        If Q2.NP Then Q2.GetPairsWDist Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    '        If Q3.NP Then Q3.GetPairsWDist Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    '        If Q4.NP Then Q4.GetPairsWDist Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    '    End If

End Sub








Friend Sub DrawToCC(CC As cCairoContext, Optional Clean As Boolean = False)
    Dim I         As Long

    '    If Clean Then CC.SetSourceColor 0: CC.Paint
    '
    '    CC.SetSourceColor vbGreen
    '
    '    CC.Rectangle X0, Y0, X1 - X0, Y1 - Y0: CC.Stroke
    '
    '    CC.SetSourceColor vbYellow
    '    For I = 1 To NP
    '        CC.Arc px(I), py(I), 1.5: CC.Stroke
    '    Next
    '
    '    If Subdivided Then
    '        If Q1.NP Then Q1.DrawToCC CC
    '        If Q2.NP Then Q2.DrawToCC CC
    '        If Q3.NP Then Q3.DrawToCC CC
    '        If Q4.NP Then Q4.DrawToCC CC
    '    End If

End Sub
