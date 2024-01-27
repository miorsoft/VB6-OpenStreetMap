Attribute VB_Name = "modDijkstra"
Option Explicit
Public dijPATH As tRoad

Private Function Dist2(iI As Long, iJ As Long) As Double
    Dim Dx     As Double
    Dim Dy     As Double
    Dx = Node(iI).x - Node(iJ).x
    Dy = Node(iI).Y - Node(iJ).Y
    Dist2 = Sqr(Dx * Dx + Dy * Dy)

End Function
Private Function FindLowestCostNODE() As Long
    Dim I      As Long
    Dim LC     As Double

    'Find Unvisited node with lower cost
    '(will be the next node to visit)
    LC = 1E+99
    For I = 1 To NNode
        If Node(I).dVisited = False Then
            If Node(I).dCost < LC Then LC = Node(I).dCost: FindLowestCostNODE = I
        End If
    Next
End Function
Private Sub DIJKSTRAvisit(ByRef iStartNode As Long)
    Dim ToCost As Double
    Dim I      As Long
    Dim iNextNode As Long


    ' VISIT Current node to All possible destination
    For I = 1 To Node(iStartNode).NNext
        iNextNode = Node(iStartNode).NextNode(I)

        If Node(iNextNode).dVisited = False Then    'Skip Visited node

            'next node COST = StartNode Cost + distance to NEXT
            ToCost = Node(iStartNode).dCost + Dist2(iStartNode, iNextNode)

            'if next node COST < previous nextnode Cost
            'Assign to NextNode the current cost
            'An mark it that best way to arrive to is StartNode
            If ToCost < Node(iNextNode).dCost Then
                Node(iNextNode).dCost = ToCost
                Node(iNextNode).dBestfrom = iStartNode
            End If
        End If
    Next

End Sub
Public Function DIJKSTRA(Ifrom As Long, Ito As Long) As tRoad
    Dim I      As Long
    Dim J      As Long

    Dim cntr   As Long
    Dim Curr   As Long

    Dim x      As Long
    Dim Y      As Long

    Dim T      As Long


If Ifrom = 0 Or Ito = 0 Then Exit Function
If Ifrom = Ito Then Exit Function
    
    
    For I = 1 To NNode
        Node(I).dCost = 1E+99
        Node(I).dVisited = False
        Node(I).dBestfrom = 0
    Next

    Node(Ifrom).dCost = 0


    Do
        cntr = cntr + 1
        If cntr Mod 50 = 0 Then
            frmMain.Caption = "Visited:" & cntr & "    " & Curr: DoEvents
            frmMain.PIC.Refresh

        End If

        Curr = FindLowestCostNODE 'first time will be Ifrom, since we set  Node(Ifrom).dCost = 0
        Node(Curr).dVisited = True
        DIJKSTRAvisit Curr

        x = Node(Curr).scrX
        Y = Node(Curr).scrY
        MyCircle phdc, x, Y, 1, 3, vbYellow

    Loop While Curr <> Ito And Curr <> 0

    frmMain.Caption = " ready"


    If Curr = 0 Then MsgBox "ImpossiblePath!": Exit Function


    'REVERSE------------------------------------------

    dijPATH.NN = 0



    Do
        With dijPATH
            .NN = .NN + 1
            ReDim Preserve .N(.NN)
            .N(.NN) = Curr
        End With
        Node(Curr).dCost = -1
        Curr = Node(Curr).dBestfrom
    Loop While Curr <> Ifrom



    With dijPATH
        .NN = .NN + 1
        ReDim Preserve .N(.NN)
        .N(.NN) = Curr
    End With



    'REVERSE dijPATH
    For I = 1 To dijPATH.NN \ 2
        J = dijPATH.NN + 1 - I
        T = dijPATH.N(I)    'SWAP
        dijPATH.N(I) = dijPATH.N(J)
        dijPATH.N(J) = T

    Next





    DIJKSTRA = dijPATH
    


   '' MsgBox "path found" & vbCrLf & "Visited Nodes:" & cntr
    DRAW
    frmMain.PIC.Refresh
    DoEvents

    
End Function



