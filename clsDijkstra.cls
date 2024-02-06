VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDijkstra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private dijPATH   As tRoad

Private ToVisitLIST() As Long

'Private cToVisitLIST As cCollection

Private NTVL      As Long
Private NTVLmax   As Long

Private MaxNN     As Long

Private tmpNodeTarget As Long

Private Function fnFindLowestCostNODE() As Long
    Dim I         As Long
    Dim J         As Long
    Dim LC        As Double

    'Find Unvisited node of ToVisitLIST with lower cost
    '(will be the next node to visit)
    LC = 1E+99

    For J = 1 To NTVL
        If Node(ToVisitLIST(J)).dVisited = 0& Then
            I = ToVisitLIST(J)
            If Node(I).dCost < LC Then LC = Node(I).dCost: fnFindLowestCostNODE = I
        End If
    Next

    'Debug.Print NTVL

    '    For J = 0 To cToVisitLIST.Count - 1
    '        I = cToVisitLIST.ItemByIndex(J)
    '        If Node(I).dVisited = 0& Then
    '            If Node(I).dCost < LC Then LC = Node(I).dCost: fnFindLowestCostNODE = I
    '        End If
    '    Next


End Function

Private Sub DIJKSTRAvisit(ByVal VisitNode As Long)
    Dim ToCost    As Double
    Dim I         As Long
    Dim IdxNextNode As Long
    Dim toTargetDX As Double
    Dim toTargetDY As Double
    Dim ToSegDX   As Double
    Dim ToSegDY   As Double
    Dim D         As Double
    Dim DIRECTIONCost As Double
    Dim SEG       As Long
    Dim WAYto     As Long

    Dim Segno     As Double

    With Node(VisitNode)
        ' VISIT Current node to All possible destination
        For I = 1 To .NNext
            IdxNextNode = .NEXTnode(I)

            If Node(IdxNextNode).dVisited = 0& Then    'Node not Visited (Skip Visited node)

                'ToCost =  .dCost + calcStreetSegmentCOST(VisitNode, IdxNextNode)
                'next node COST = StartNode Cost + distance to NEXT


                ' STANDARD
                ' ToCost = .dCost + .NextNodeCOST(I)


                ' SORT of A*  (directional Cost)
                'Multiply cost by Dot Product between Tergt node direction and Segment direction
                toTargetDX = Node(tmpNodeTarget).X - Node(IdxNextNode).X
                toTargetDY = Node(tmpNodeTarget).Y - Node(IdxNextNode).Y
                D = Sqr(toTargetDX * toTargetDX + toTargetDY * toTargetDY)
                If D Then
                    D = 1# / D
                    toTargetDX = toTargetDX * D
                    toTargetDY = toTargetDY * D

                    WAYto = .NEXTNodeWay(I)
                    SEG = .NextNodeWAYSegment(I)
                    Segno = Sgn(SEG): SEG = Abs(SEG)

                    ToSegDX = WAY(WAYto).SegDX(SEG)
                    ToSegDY = WAY(WAYto).SegDY(SEG)
                    DIRECTIONCost = (1.2 - Segno * (toTargetDX * ToSegDX + toTargetDY * ToSegDY))
                    ToCost = .dCost + .NextNodeCOST(I) * DIRECTIONCost

                Else
                    ToCost = .dCost + .NextNodeCOST(I)
                End If


                'if next node COST < previous nextnode Cost
                'Assign to NextNode the current cost
                'An mark it that best way to arrive to is StartNode
                If ToCost < Node(IdxNextNode).dCost Then
                    Node(IdxNextNode).dCost = ToCost
                    Node(IdxNextNode).dBestfrom = VisitNode
                    '  Node(IdxNextNode).dVisited = 0  '????Maybe
                End If
            End If
        Next
    End With

End Sub
Friend Function DIJKSTRA(ByVal Ifrom As Long, ByVal Ito As Long) As tRoad
    Dim I         As Long
    Dim J         As Long
    Dim Curr      As Long
    Dim T         As Long


    With frmMain
        .PIC.MousePointer = 11
        .Caption = "Finding path from " & Ifrom & " to " & Ito
    End With

    tmpNodeTarget = Ito

    If Ifrom = 0 Or Ito = 0 Then Exit Function
    If Ifrom = Ito Then Exit Function

    For I = 1 To NNode
        With Node(I)
            .dCost = 1E+99
            .dVisited = 0         'False
            .dBestfrom = 0
        End With
    Next

    Node(Ifrom).dCost = 0

    NTVL = 1
    ToVisitLIST(1) = Ifrom
    '    cToVisitLIST.removeAll
    '    cToVisitLIST.Add Ifrom



    Do
        '        cntr = cntr + 1
        '        If cntr Mod 50 = 0 Then
        '            frmMain.Caption = "Visited:" & cntr & "    " & Curr: DoEvents
        '            frmMain.PIC.Refresh
        '        End If
        Curr = fnFindLowestCostNODE    'first time will be Ifrom, since we set  Node(Ifrom).dCost = 0

        Node(Curr).dVisited = -1  'True

        For I = 1 To Node(Curr).NNext

            If Node(Node(Curr).NEXTnode(I)).dVisited = 0 Then    ''??? RIGHT ????
                NTVL = NTVL + 1
                If NTVL > NTVLmax Then
                    NTVLmax = NTVL + 1000
                    ReDim Preserve ToVisitLIST(NTVLmax)
                End If
                ToVisitLIST(NTVL) = Node(Curr).NEXTnode(I)

                '            If cToVisitLIST.ItemExists(Node(Curr).NEXTnode(I)) Then
                '            Else
                '                cToVisitLIST.Add Node(Curr).NEXTnode(I)
                '            End If
            Else
                '                Stop
            End If

        Next


        DIJKSTRAvisit Curr

        ' x = Node(Curr).scrX
        ' y = Node(Curr).scrY
        ' MyCircle frmMainPIChDC, x, y, 1, 3, vbRed

    Loop While Curr <> Ito And Curr <> 0

    '    frmMain.Caption = " ready"

    dijPATH.NN = 0
    '    'If Curr = 0 Then MsgBox "ImpossiblePath!": Exit Function
    If Curr = 0 Then Exit Function


    'REVERSE------------------------------------------
    Do
        With dijPATH              'Walk from end to Start
            .NN = .NN + 1
            If .NN > MaxNN Then MaxNN = 1 + .NN * 1.5: ReDim Preserve .N(MaxNN)
            .N(.NN) = Curr
        End With
        Node(Curr).dCost = -1
        Curr = Node(Curr).dBestfrom
    Loop While Curr <> Ifrom

    With dijPATH
        .NN = .NN + 1: If .NN > MaxNN Then ReDim Preserve .N(.NN)
        .N(.NN) = Curr
    End With

    'REVERSE dijPATH ORDER
    With dijPATH
        For I = 1 To .NN \ 2
            J = .NN + 1 - I
            'SWAP:
            T = .N(I): .N(I) = .N(J): .N(J) = T
        Next
    End With

    DIJKSTRA = dijPATH


    frmMain.Caption = ""
    If DoFollow Then
        frmMain.PIC.MousePointer = 1
    Else
        frmMain.PIC.MousePointer = 11
    End If


End Function




Private Sub Class_Initialize()
    ReDim ToVisitLIST(1)
    '    Set cToVisitLIST = New_c.Collection(False)
End Sub
