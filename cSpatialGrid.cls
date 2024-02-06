VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSpatialGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type tCELL
    pIDX()        As Long
    NP            As Long
    MaxNP         As Long
End Type


Private mW        As Long
Private mH        As Long
Private GridSize  As Long

Private CellW     As Long
Private CellH     As Long


Private CELL()    As tCELL
Private NP        As Long
Private MaxNP     As Long
Private PX()      As Double
Private PY()      As Double

Private MaxDistance2 As Double


Private PairP1()  As Long
Private PairP2()  As Long
Private PairDX()  As Double
Private PairDY()  As Double
Private PairD()   As Double
Private Npairs    As Long
Private MAXNpair  As Long


'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)

Private Declare Function GetMem4 Lib "MSVBVM60" (ByVal Src As Long, ByVal dest As Long) As Long
Private Declare Function ArrPtr Lib "MSVBVM60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Function vbaCopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, Dst As Long, Src As Long) As Long




Public Property Get NofPoints() As Long
    NofPoints = NP
End Property

Public Property Get getPointX(I&) As Double
    getPointX = PX(I)
End Property
Public Property Get getPointY(I&) As Double
    getPointY = PY(I)
End Property


Private Sub pvArrCopySingle(dest() As Double, Src() As Double)
    Dim Size      As Long
    Dim W         As Long
    Dim H         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))
    If pvArrayExists(ArrPtr(Src)) Then

        W = UBound(Src)
        Size = (W + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (W - UBound(dest, 1)) Then
                'ReDim dest(W, H)
                ReDim dest(W)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(W)
        End If

        '        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        '
        '        CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size


        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))
    End If

End Sub

Private Sub pvArrCopyLong(dest() As Long, Src() As Long)
    Dim Size      As Long
    Dim W         As Long
    Dim H         As Long
    '    W = UBound(Src, 1)
    '    H = UBound(Src, 2)
    '    Size = (W + 1) * (H + 1) * LenB(Src(0, 0))

    If pvArrayExists(ArrPtr(Src)) Then

        W = UBound(Src)
        Size = (W + 1) * LenB(Src(0))

        If pvArrayExists(ArrPtr(dest)) Then
            '        If (W - UBound(dest, 1)) Or (H - UBound(dest, 2)) Then
            If (W - UBound(dest, 1)) Then
                'ReDim dest(W, H)
                ReDim dest(W)
            End If
        Else
            '''Array DEST has No Dimension
            'ReDim dest(W, H)
            ReDim dest(W)
        End If

        '    CopyMemory ByVal VarPtr(dest(0, 0)), ByVal VarPtr(Src(0, 0)), Size
        ' CopyMemory ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0)), Size

        vbaCopyBytes Size, ByVal VarPtr(dest(0)), ByVal VarPtr(Src(0))


    End If

End Sub
Private Function pvArrayExists(ByVal ppArray As Long) As Long
    GetMem4 ppArray, VarPtr(pvArrayExists)
End Function



Friend Sub Init(WorldW As Long, WorldH As Long, MaxDist As Long)
    mW = WorldW
    mH = WorldH

    GridSize = MaxDist
    MaxDistance2 = MaxDist * MaxDist

    If CellW <> mW \ GridSize Or CellH <> mH \ GridSize Then    'Per Mantenere MaxNP
        CellW = mW \ GridSize
        CellH = mH \ GridSize
        ReDim CELL(CellW, CellH)
    End If

    NP = 0
    Npairs = 0


End Sub


Friend Sub ResetPoints()

    Dim X         As Long
    Dim Y         As Long

    For Y = 0 To CellH
        For X = 0 To CellW
            CELL(X, Y).NP = 0&
        Next
    Next
    NP = 0
    '        MaxNP = 0

End Sub
Friend Sub InsertPoint(ByVal X As Double, ByVal Y As Double)


    NP = NP + 1
    If NP > MaxNP Then
        MaxNP = NP * 2
        'If NP > UBound(pX) Then
        MaxNP = NP * 3
        ReDim Preserve PX(MaxNP)
        ReDim Preserve PY(MaxNP)
    End If
    PX(NP) = X
    PY(NP) = Y

    pvAddToCell X \ GridSize, Y \ GridSize, NP

End Sub
'''Friend Sub InsertPointWithIDX(ByVal X As Double, ByVal Y As Double, IDX As Long)
'''
'''
'''    NP = NP + 1
'''    'If NP > MaxNP Then
'''    '    MaxNP = NP * 2
'''    If NP > UBound(pX) Then
'''        MaxNP = NP * 3
'''        ReDim Preserve pX(MaxNP)
'''        ReDim Preserve pY(MaxNP)
'''    End If
'''    pX(NP) = X
'''    pY(NP) = Y
'''
'''    pvAddToCell X \ GridSize, Y \ GridSize, IDX 'NP
'''
'''End Sub
Friend Sub InsertALLpoints(X() As Double, Y() As Double, Optional LimitNofPointTo As Long = 0)
    'Remember to Call ResetPoints First

    Dim I         As Long
    Dim u         As Long



    If LimitNofPointTo Then       ' To avoid pre-Call REDIMing of X Y in some cases
        u = LimitNofPointTo
    Else
        u = UBound(X)
    End If

    For I = 1 To u
        NP = NP + 1
        If NP > MaxNP Then


            '            MaxNP = NP + 64    '* 2
            MaxNP = 1 + NP * 1.5
            '            ReDim Preserve PX(MaxNP)
            '            ReDim Preserve PY(MaxNP)
        End If
        '        PX(NP) = X(I)
        '        PY(NP) = Y(I)
        pvAddToCell X(I) \ GridSize, Y(I) \ GridSize, NP
    Next
    pvArrCopySingle PX, X
    pvArrCopySingle PY, Y

End Sub

Private Sub pvAddToCell(ByVal cellX As Long, ByVal cellY As Long, ByVal currPindex As Long)



    With CELL(cellX, cellY)
        .NP = .NP + 1
        If .NP > .MaxNP Then
            .MaxNP = .NP + 64     '* 2
            ReDim Preserve .pIDX(.MaxNP)
        End If
        .pIDX(.NP) = currPindex
    End With

End Sub


Friend Sub GetPairs(RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim DX        As Double
    Dim DY        As Double
    Dim I         As Long
    Dim J         As Long

    Dim cWm1      As Long
    Dim cHm1      As Long
    Dim Xp1       As Long
    Dim Yp1       As Long
    Dim Ym1       As Long



    Dim IJ        As Long
    Dim ii        As Long

    Dim PXI       As Double
    Dim PYI       As Double



    cWm1 = CellW - 1&
    cHm1 = CellH - 1&

    Npairs = 0&
    '      O
    '     /
    '    X-O
    '    |\
    '    O O


    For Y = 1& To cHm1
        Yp1 = Y + 1&
        Ym1 = Y - 1&
        For X = 0& To cWm1
            Xp1 = X + 1&

            With CELL(X, Y)


                For I = 1& To .NP    '- 1& ' Should be -1& to do only SELF but so we can do even others
                    ii = .pIDX(I)
                    PXI = PX(ii)
                    PYI = PY(ii)
                    For J = I + 1& To .NP    'SELF
                        IJ = .pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                    Next


                    'RIGHT
                    For J = 1& To CELL(Xp1, Y).NP
                        IJ = CELL(Xp1, Y).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                    Next

                    'down
                    For J = 1& To CELL(X, Yp1).NP
                        IJ = CELL(X, Yp1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                    Next

                    'DownRight Diag
                    For J = 1& To CELL(Xp1, Yp1).NP
                        IJ = CELL(Xp1, Yp1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                    Next

                    'upRight Diag
                    For J = 1& To CELL(Xp1, Ym1).NP
                        IJ = CELL(Xp1, Ym1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                    Next
                Next
            End With

        Next
    Next



    '-------- NOW Bottom ROW
    '      O
    '     /
    '    X-O

    Y = CellH
    Ym1 = Y - 1&
    For X = 0& To cWm1
        Xp1 = X + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, Y).NP
                    IJ = CELL(Xp1, Y).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                'upRight Diag
                For J = 1& To CELL(Xp1, Ym1).NP
                    IJ = CELL(Xp1, Ym1).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
            Next

        End With

        '-------- NOW TOP ROW
        '    X-O
        '    |\
        '    O O


        With CELL(X, 0&)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, 0&).NP
                    IJ = CELL(Xp1, 0&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                'DownRight Diag
                For J = 1& To CELL(Xp1, 1&).NP
                    IJ = CELL(Xp1, 1&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                'Down
                For J = 1& To CELL(X, 1&).NP
                    IJ = CELL(X, 1&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
            Next

        End With
    Next X



    ' -----------------------------RIGHT Column
    '    X
    '    |
    '    O
    X = CellW
    For Y = 0& To cHm1
        Yp1 = Y + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1& to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
                ' DOWN  )
                For J = 1& To CELL(X, Yp1).NP
                    IJ = CELL(X, Yp1).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
                Next
            Next
        End With
    Next


    ' -------- FINALLY bottm right cell
    With CELL(CellW, CellH)
        'SELF
        For I = 1& To .NP - 1&
            ii = .pIDX(I)
            PXI = PX(ii)
            PYI = PY(ii)
            For J = I + 1& To .NP
                IJ = .pIDX(J)
                DX = PX(IJ) - PXI
                DY = PY(IJ) - PYI
                If DX * DX + DY * DY < MaxDistance2 Then pvAddPair ii, IJ, DX, DY
            Next
        Next
    End With




    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs
    '    rP1 = PairP1
    '    rP2 = PairP2
    '    rDX = PairDX
    '    rDY = PairDY

    'FASTER
    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY


    '    'Check for ERRORS (Duplicates)
    '        For X = 1& To Npairs - 1&
    '            For Y = X + 1& To Npairs
    '                If (PairP1(X) = PairP1(Y)) And (PairP2(X) = PairP2(Y)) Then Stop
    '                If (PairP1(X) = PairP2(Y)) And (PairP2(X) = PairP1(Y)) Then Stop
    '            Next
    '        Next

End Sub
Friend Sub GetPairsWDist(RP1() As Long, RP2() As Long, Rdx() As Double, Rdy() As Double, RD() As Double, rPairsCount As Long)

    'CORE SUB

    Dim X         As Long
    Dim Y         As Long
    Dim DX        As Double
    Dim DY        As Double
    Dim I         As Long
    Dim J         As Long
    Dim D         As Double

    Dim cWm1      As Long
    Dim cHm1      As Long
    Dim Xp1       As Long
    Dim Yp1       As Long
    Dim Ym1       As Long


    Dim IJ        As Long
    Dim ii        As Long

    Dim PXI       As Double
    Dim PYI       As Double


    cWm1 = CellW - 1&
    cHm1 = CellH - 1&


    Npairs = 0&
    '      O
    '     /
    '    X-O
    '    |\
    '    O O

    For Y = 1& To cHm1
        Yp1 = Y + 1&
        Ym1 = Y - 1&
        For X = 0& To cWm1
            Xp1 = X + 1&
            With CELL(X, Y)

                For I = 1& To .NP    '- 1& ' Should be -1 to do only SELF but so we can do even others
                    ii = .pIDX(I)
                    PXI = PX(ii)
                    PYI = PY(ii)
                    For J = I + 1& To .NP    'SELF
                        IJ = .pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        D = DX * DX + DY * DY
                        If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                    Next

                    'RIGHT
                    For J = 1& To CELL(Xp1, Y).NP
                        IJ = CELL(Xp1, Y).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        D = DX * DX + DY * DY
                        If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                    Next

                    'down
                    For J = 1& To CELL(X, Yp1).NP
                        IJ = CELL(X, Yp1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        D = DX * DX + DY * DY
                        If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                    Next

                    'DownRight Diag
                    For J = 1& To CELL(Xp1, Yp1).NP
                        IJ = CELL(Xp1, Yp1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        D = DX * DX + DY * DY
                        If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                    Next

                    'upRight Diag
                    For J = 1& To CELL(Xp1, Ym1).NP
                        IJ = CELL(Xp1, Ym1).pIDX(J)
                        DX = PX(IJ) - PXI
                        DY = PY(IJ) - PYI
                        D = DX * DX + DY * DY
                        If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                    Next
                Next
            End With

        Next
    Next

    '-------- NOW Bottom ROW
    '      O
    '     /
    '    X-O

    Y = CellH
    Ym1 = Y - 1&
    For X = 0& To cWm1
        Xp1 = X + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, Y).NP
                    IJ = CELL(Xp1, Y).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
                'upRight Diag
                For J = 1& To CELL(Xp1, Ym1).NP
                    IJ = CELL(Xp1, Ym1).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
            Next

        End With


        '-------- NOW TOP ROW
        '    X-O
        '    |\
        '    O O
        With CELL(X, 0&)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
                'RIGHT ( )
                For J = 1& To CELL(Xp1, 0&).NP
                    IJ = CELL(Xp1, 0&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
                'DownRight Diag
                For J = 1& To CELL(Xp1, 1&).NP
                    IJ = CELL(Xp1, 1&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
                'Down
                For J = 1& To CELL(X, 1&).NP
                    IJ = CELL(X, 1&).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
                Next
            Next

        End With
    Next X



    ' RIGHT Column
    '    X
    '    |
    '    O
    X = CellW
    For Y = 0& To cHm1
        Yp1 = Y + 1&
        With CELL(X, Y)
            For I = 1& To .NP     '- 1& ' Should be -1 to do only SELF but so we can do even others
                ii = .pIDX(I)
                PXI = PX(ii)
                PYI = PY(ii)
                For J = I + 1& To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                Next
                ' DOWN  )
                For J = 1& To CELL(X, Yp1).NP
                    IJ = CELL(X, Yp1).pIDX(J)
                    DX = PX(IJ) - PXI
                    DY = PY(IJ) - PYI
                    D = DX * DX + DY * DY
                    If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D

                Next
            Next
        End With
    Next


    ' -------- FINALLY bottm right cell
    With CELL(CellW, CellH)
        'SELF
        For I = 1& To .NP - 1&
            ii = .pIDX(I)
            PXI = PX(ii)
            PYI = PY(ii)
            For J = I + 1& To .NP
                IJ = .pIDX(J)
                DX = PX(IJ) - PXI
                DY = PY(IJ) - PYI
                D = DX * DX + DY * DY
                If D < MaxDistance2 Then pvAddPairWDist ii, IJ, DX, DY, D
            Next
        Next
    End With




    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------

    rPairsCount = Npairs


    pvArrCopyLong RP1, PairP1
    pvArrCopyLong RP2, PairP2
    pvArrCopySingle Rdx, PairDX
    pvArrCopySingle Rdy, PairDY
    pvArrCopySingle RD, PairD



End Sub

Private Sub pvAddPair(ByVal P1 As Long, ByVal P2 As Long, ByVal DX As Double, ByVal DY As Double)
    Stop

    Npairs = Npairs + 1
    If Npairs > MAXNpair Then
        MAXNpair = Npairs + 128   '* 2
        ReDim Preserve PairP1(MAXNpair)
        ReDim Preserve PairP2(MAXNpair)
        ReDim Preserve PairDX(MAXNpair)
        ReDim Preserve PairDY(MAXNpair)
    End If

    PairP1(Npairs) = P1
    PairP2(Npairs) = P2
    PairDX(Npairs) = DX
    PairDY(Npairs) = DY

End Sub

Private Sub pvAddPairWDist(ByVal P1 As Long, ByVal P2 As Long, _
                           ByVal DX As Double, ByVal DY As Double, ByVal D As Double)


    Npairs = Npairs + 1

    If Npairs > MAXNpair Then
        '        MAXNpair = Npairs + 128    '* 2

        If Npairs > 4 Then
            MAXNpair = Npairs * 1.3
        Else
            MAXNpair = Npairs + 8

        End If

        ReDim Preserve PairP1(MAXNpair)
        ReDim Preserve PairP2(MAXNpair)
        ReDim Preserve PairDX(MAXNpair)
        ReDim Preserve PairDY(MAXNpair)
        ReDim Preserve PairD(MAXNpair)
    End If

    PairP1(Npairs) = P1
    PairP2(Npairs) = P2
    PairDX(Npairs) = DX
    PairDY(Npairs) = DY
    PairD(Npairs) = D

End Sub



Friend Sub QueryPointsInCircle(cX As Double, cY As Double, R As Double, rDist() As Double, rIdx() As Long, UBcount As Long)

    Dim X         As Long
    Dim Y         As Long
    Dim xx        As Long
    Dim yy        As Long

    Dim J         As Long
    Dim IJ        As Long

    Dim DX        As Double
    Dim DY        As Double
    Dim D         As Double
    Dim gg        As Long
    Dim R2        As Double


    Dim XF        As Long
    Dim XT        As Long
    Dim YF        As Long
    Dim YT        As Long

    X = cX \ GridSize
    Y = cY \ GridSize


    ReDim rIdx(0)
    Npairs = 0

    gg = R \ GridSize
    If gg < 1 Then gg = 1         'Commented 2023
    'If GG < 1 Then Stop
    'If GG > 2 Then Debug.Print GG & "  " & R

    R2 = R * R

    XF = X - gg: If XF < 0& Then XF = 0&
    XT = X + gg: If XT > CellW Then XT = CellW
    YF = Y - gg: If YF < 0& Then YF = 0&
    YT = Y + gg: If YT > CellH Then YT = CellH

    For yy = YF To YT
        For xx = XF To XT

            With CELL(xx, yy)
                For J = 1 To .NP
                    IJ = .pIDX(J)
                    DX = PX(IJ) - cX
                    DY = PY(IJ) - cY
                    D = DX * DX + DY * DY
                    If D <= R2 Then
                        pvAddPointWDist IJ, D
                    End If
                Next
            End With
        Next
    Next

    pvArrCopySingle rDist, PairD
    pvArrCopyLong rIdx, PairP1
    UBcount = Npairs




End Sub


Private Sub pvAddPointWDist(ByVal P1 As Long, ByVal D As Double)

    Npairs = Npairs + 1

    If Npairs > MAXNpair Then

        If Npairs > 4 Then
            MAXNpair = Npairs * 1.3
        Else

            MAXNpair = Npairs * 1.5    'Npairs + 8
        End If

        ReDim Preserve PairP1(MAXNpair)
        'ReDim Preserve PairP2(MAXNpair)
        'ReDim Preserve PairDX(MAXNpair)
        'ReDim Preserve PairDY(MAXNpair)
        ReDim Preserve PairD(MAXNpair)
    End If

    PairP1(Npairs) = P1
    'PairP2(Npairs) = P2
    'PairDX(Npairs) = dx
    'PairDY(Npairs) = dy
    PairD(Npairs) = D

End Sub
