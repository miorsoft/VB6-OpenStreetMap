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

Public X0         As Double
Public Y0         As Double
Public X1         As Double
Public Y1         As Double
Private MidX      As Double
Private MidY      As Double

Private Q1        As cQuadTree
Private Q2        As cQuadTree
Private Q3        As cQuadTree
Private Q4        As cQuadTree

'Private NotSubdivided As Long
Private Subdivided As Long

Public Capacity   As Long
Public NP         As Long
Private MaxNP     As Long



Private PX()      As Double
Private PY()      As Double
Private pIDX()    As Double
Private GPX()     As Double
Private GPY()     As Double
Public GhostNP    As Long
Private GhostmaxNP As Long


''Private Declare Function GetMem4 Lib "MSVBVM60" (ByVal Src As Long, ByVal dest As Long) As Long
''Private Declare Function ArrPtr Lib "MSVBVM60" Alias "VarPtr" (Arr() As Any) As Long
''Private Declare Function vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, Dst As Long, Src As Long) As Long
''
''Private Function pvArrayExists(ByVal ppArray As Long) As Long
''    GetMem4 ppArray, VarPtr(pvArrayExists)
''End Function
''Private Sub pvArrCopy(dest() As Double, Src() As Double)
''    Dim Size   As Long
''    Dim W      As Long
''    Dim H      As Long
''    If pvArrayExists(ArrPtr(Src)) Then
''        W = UBound(Src)
''        Size = (W + 1) * LenB(Src(0))
''
''        If pvArrayExists(ArrPtr(dest)) Then
''            If (W - UBound(dest, 1)) Then
''                ReDim dest(W)
''            End If
''        Else
''            '''Array DEST has No Dimension
''            ReDim dest(W)
''        End If
''        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))
''    End If
''End Sub
''Private Sub pvArrCopyLong(dest() As Long, Src() As Long)
''    Dim Size   As Long
''    Dim W      As Long
''    Dim H      As Long
''    If pvArrayExists(ArrPtr(Src)) Then
''        W = UBound(Src)
''        Size = (W + 1) * LenB(Src(0))
''        If pvArrayExists(ArrPtr(dest)) Then
''            If (W - UBound(dest, 1)) Then
''                ReDim dest(W)
''            End If
''        Else
''            '''Array DEST has No Dimension
''            ReDim dest(W)
''        End If
''        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))
''    End If
''End Sub

Private Sub Class_Initialize()
    '    NotSubdivided = -1
    Subdivided = 0
End Sub



Friend Sub Setup(ByVal Xleft, ByVal Ytop, ByVal Xright, ByVal Ydown, MaxPperQ As Long)
    X0 = Xleft
    X1 = Xright
    Y0 = Ytop
    Y1 = Ydown
    Capacity = MaxPperQ
    Reset
End Sub
Friend Sub Reset()
    NP = 0
    GhostNP = 0
    If Subdivided Then
        If Q1.NP Then Q1.Reset
        If Q2.NP Then Q2.Reset
        If Q3.NP Then Q3.Reset
        If Q4.NP Then Q4.Reset
    End If

End Sub

Private Sub Subdivide()

    MidX = (X0 + X1) * 0.5
    MidY = (Y0 + Y1) * 0.5

    'NotSubdivided = 0
    Subdivided = -1

    Set Q1 = New cQuadTree
    Set Q2 = New cQuadTree
    Set Q3 = New cQuadTree
    Set Q4 = New cQuadTree

    '    Q1.X0 = X0: Q1.Y0 = Y0
    '    Q2.X0 = MidX: Q2.Y0 = Y0
    '    Q3.X0 = X0: Q3.Y0 = MidY
    '    Q4.X0 = MidX: Q4.Y0 = MidY
    '
    '    Q1.X1 = MidX: Q1.Y1 = MidY
    '    Q2.X1 = X1: Q2.Y1 = MidY
    '    Q3.X1 = MidX: Q3.Y1 = Y1
    '    Q4.X1 = X1: Q4.Y1 = Y1
    '
    '    Q1.Capacity = Capacity
    '    Q2.Capacity = Capacity
    '    Q3.Capacity = Capacity
    '    Q4.Capacity = Capacity

    Q1.Setup X0, Y0, MidX, MidY, Capacity
    Q2.Setup MidX, Y0, X1, MidY, Capacity
    Q3.Setup X0, MidY, MidX, Y1, Capacity
    Q4.Setup MidX, MidY, X1, Y1, Capacity


End Sub

Friend Sub InsertSinglePoint(ByVal X#, ByVal Y#, Index As Long, Optional isRoot As Boolean = True)

    Dim Q         As Long

    If isRoot Then
        GhostNP = GhostNP + 1
        If GhostNP > GhostmaxNP Then
            GhostmaxNP = GhostNP * 2
            ReDim Preserve GPX(GhostmaxNP)
            ReDim Preserve GPY(GhostmaxNP)
        End If
        GPX(GhostNP) = X
        GPY(GhostNP) = Y
    End If

    If NP < Capacity Then
        NP = NP + 1
        If NP > MaxNP Then
            MaxNP = Capacity
            ReDim PX(MaxNP)
            ReDim PY(MaxNP)
            ReDim pIDX(MaxNP)
        End If
        PX(NP) = X
        PY(NP) = Y
        pIDX(NP) = Index

    Else

        If Not (Subdivided) Then Subdivide

        If X > MidX Then Q = 1&
        If Y > MidY Then Q = Q Or 2&

        Select Case Q
        Case 0: Q1.InsertSinglePoint X, Y, Index, False
        Case 1: Q2.InsertSinglePoint X, Y, Index, False
        Case 2: Q3.InsertSinglePoint X, Y, Index, False
        Case 3: Q4.InsertSinglePoint X, Y, Index, False
        End Select

    End If

End Sub



Friend Sub InsertALLpoints(X() As Double, Y() As Double, Optional PTBound As Long)
    Dim I         As Long
    Reset
    If PTBound = 0 Then PTBound = UBound(X)
    For I = 1 To PTBound
        InsertSinglePoint X(I), Y(I), I, True
    Next
End Sub


Friend Sub QueryAroundPoint(ByVal CurrPointX#, ByVal CurrPointY#, ByVal CurrPointIDX As Long, ByVal R#, _
                            RP1() As Long, _
                            RP2() As Long, _
                            Rdx() As Double, _
                            Rdy() As Double, _
                            rDD() As Double, _
                            rPairsCount As Long, MaxNpairs As Long)

    Dim I         As Long
    Dim DX#, DY#, R2#, DD#
    If CurrPointX + R < X0 Then Exit Sub
    If CurrPointY + R < Y0 Then Exit Sub
    If CurrPointX - R > X1 Then Exit Sub
    If CurrPointY - R > Y1 Then Exit Sub

    R2 = R * R

    For I = 1 To NP
        If pIDX(I) > CurrPointIDX Then
            DX = PX(I) - CurrPointX: DY = PY(I) - CurrPointY: DD = (DX * DX + DY * DY)
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
    End If

End Sub



Friend Sub QueryPointsInCircle(ByVal cX As Double, ByVal cY As Double, ByVal R As Double, _
                               rDist() As Double, _
                               rIdx() As Long, _
                               rPntCount As Long, MaxrPntCount As Long, Optional isRoot As Boolean = True)
    Dim I         As Long
    Dim DX#, DY#, R2#, DD#
    If cX + R < X0 Then Exit Sub
    If cY + R < Y0 Then Exit Sub
    If cX - R > X1 Then Exit Sub
    If cY - R > Y1 Then Exit Sub

    R2 = R * R

    If isRoot Then rPntCount = 0  ': MaxrPntCount = 0

    For I = 1 To NP
        DX = PX(I) - cX: DY = PY(I) - cY: DD = (DX * DX + DY * DY)
        If DD < R2 Then
            rPntCount = rPntCount + 1
            If rPntCount > MaxrPntCount Then
                MaxrPntCount = rPntCount * 2
                ReDim Preserve rIdx(MaxrPntCount)
                ReDim Preserve rDist(MaxrPntCount)
            End If
            rDist(rPntCount) = DD
            rIdx(rPntCount) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QueryPointsInCircle cX, cY, R, rDist, rIdx, rPntCount, MaxrPntCount, False
        If Q2.NP Then Q2.QueryPointsInCircle cX, cY, R, rDist, rIdx, rPntCount, MaxrPntCount, False
        If Q3.NP Then Q3.QueryPointsInCircle cX, cY, R, rDist, rIdx, rPntCount, MaxrPntCount, False
        If Q4.NP Then Q4.QueryPointsInCircle cX, cY, R, rDist, rIdx, rPntCount, MaxrPntCount, False
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


    If isRoot Then Found = 0      ': MaxFound = 0

    For I = 1 To NP
        DX = PX(I) - cX: DY = PY(I) - cY
        If Abs(DX) < halfSide And Abs(DY) < halfSide Then
            Found = Found + 1
            If Found > MaxFound Then
                MaxFound = Found * 2
                ReDim Preserve rpIDX(MaxFound)
                ReDim Preserve rpX(MaxFound)
                ReDim Preserve rpY(MaxFound)

            End If
            rpX(Found) = PX(I)
            rpY(Found) = PY(I)
            rpIDX(Found) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q2.NP Then Q2.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q3.NP Then Q3.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q4.NP Then Q4.QuerySquare cX, cY, halfSide, rpX, rpY, rpIDX, Found, MaxFound, False
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

    If isRoot Then Found = 0      ': maxfound = 0

    For I = 1 To NP
        DX = PX(I) - cX: DY = PY(I) - cY: DD = (DX * DX + DY * DY)
        If DD < R2 Then
            Found = Found + 1
            If Found > MaxFound Then
                MaxFound = Found * 2
                ReDim Preserve rpIDX(MaxFound)
                ReDim Preserve rpX(MaxFound)
                ReDim Preserve rpY(MaxFound)

            End If
            rpX(Found) = PX(I)
            rpY(Found) = PY(I)
            rpIDX(Found) = pIDX(I)

        End If
    Next

    If Subdivided Then
        If Q1.NP Then Q1.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q2.NP Then Q2.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q3.NP Then Q3.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
        If Q4.NP Then Q4.QueryCircleOLD cX, cY, R, rpX, rpY, rpIDX, Found, MaxFound, False
    End If

End Sub
Friend Sub GetPairsWDist(ByVal Radius As Double, RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, rDD() As Double, rPairsCount As Long, MaxNpairs As Long)
    Dim I         As Long
    '  USING GHOST
    rPairsCount = 0
    For I = 1 To GhostNP
        QueryAroundPoint GPX(I), GPY(I), I, Radius, RP1, RP2, Rdx, Rdy, rDD, rPairsCount, MaxNpairs
    Next

    '   Without GHOST
    ''    For I = 1 To NP
    ''        QueryAroundPoint PX(I), PY(I), Pidx(I), Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''    Next
    ''    If Not (NotSubdivided) Then
    ''        If Q1.NP Then Q1.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q2.NP Then Q2.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q3.NP Then Q3.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''        If Q4.NP Then Q4.GetPairsWDist Radius, rP1, rP2, rDX, rDY, rDD, rPairsCount, MaxNpairs
    ''    End If

End Sub








Friend Sub DrawToCC(CC As cCairoContext, Optional Clean As Boolean = False)
    Dim I         As Long

    If Clean Then CC.SetSourceColor 0: CC.Paint

    CC.SetSourceColor vbGreen

    CC.Rectangle X0, Y0, X1 - X0, Y1 - Y0: CC.Stroke

    CC.SetSourceColor vbYellow
    For I = 1 To NP
        CC.Arc PX(I), PY(I), 1.5: CC.Stroke
    Next

    If Subdivided Then
        If Q1.NP Then Q1.DrawToCC CC
        If Q2.NP Then Q2.DrawToCC CC
        If Q3.NP Then Q3.DrawToCC CC
        If Q4.NP Then Q4.DrawToCC CC
    End If

End Sub