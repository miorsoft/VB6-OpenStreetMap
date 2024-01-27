Attribute VB_Name = "modOSM"
Option Explicit

'https://wiki.openstreetmap.org/wiki/Map_features

Public Type tNode
    fileID        As Currency     'could be bigger than LONG 4bytes (currency 8 Bytes)
    id            As Long
    X             As Double
    Y             As Double
    Used          As Boolean
    scrX          As Double       ' Long
    scrY          As Double       ' Long

    'Dijkstra--------------------------------
    NNext         As Long         'N of nodes reachable from this Node
    NEXTnode()    As Long         'List of Nodes reachable from this Node
    NextNodeCOST() As Double
    NEXTNodeWay() As Long
    NextNodeWAYSegment() As Long

    dIsWay        As Boolean
    dCost         As Double
    dBestfrom     As Long
    dVisited      As Long         'Boolean
    ONEWAY        As Long         '??  Boolean 'Long  Maybe Useless
    '------------------------------------------------

    Layer         As Long         'For bridges / tunnels

    Traffic       As Double

End Type

Public Type tRoad
    fileID        As Long
    id            As Long
    NN            As Long         ' Number of nodes of the road
    invNN         As Double       '1/NN
    N()           As Currency     'List of Nodes

    SegDX()       As Double       ' 1 to NN-1
    SegDY()       As Double       'Segment direction and angle
    SegAngle()    As Double

    RoadWidth     As Double

    Name          As String
    wayType       As String
    isDRIVEABLE   As Boolean      ' Can be used by Cars

    IsStreetSideParkable As Boolean


    IsRailWay     As Boolean
    IsBoundary    As Boolean



    IsBuilding    As Boolean
    IsArea        As Boolean
    IsWater       As Boolean
    IsLeisure     As Boolean
    IsAmenity     As Boolean
    IsShop        As Boolean
    IsNatural     As Boolean
    isNotAsphalt  As Boolean

    Layer         As Long
    Lanes         As Long

    OneWayDirections As Long      'Drivable ways 0= Both :  1 = Forward -1=Backward


    'bound Box
    BBx1          As Double
    BBY1          As Double
    BBx2          As Double
    BBY2          As Double

    'World Center
    CenterX       As Double       ' NON SALVARE
    CenterY       As Double

    R             As Double
    G             As Double
    B             As Double

End Type


Public Const MAPSsubFolder As String = "\MAPS\"

Public Node()     As tNode
Public NNode      As Long

Public WAY()      As tRoad
Public NWay       As Long

Public DJ         As clsDijkstra
Public dijPATH    As tRoad

Private ptsUSED   As Long


Public MapOwest   As Double
Public MapNorth   As Double
Public MapEast    As Double
Public MapSouth   As Double
Public EastOwestMapExtension As Currency
Public NorthSouthMapExtension As Currency

Private NODESbyFileID As cCollection


Public minLayer   As Long
Public maxLayer   As Long


Private Function RemoveM(v As String) As String
    Dim I         As Long
    I = InStr(1, v, "m")
    If I Then
        RemoveM = Left$(v, I - 1)
    Else
        RemoveM = v
    End If
End Function
Public Sub ReadOSM(FN As String)
    Dim X         As Double
    Dim Y         As Double
    Dim utmX      As Double
    Dim utmY      As Double

    Dim MaxNways  As Long
    Dim MaxNNodes As Long


    Dim objXML    As MSXML2.DOMDocument60
    Dim objXML2   As New MSXML2.DOMDocument60
    Dim objElem   As MSXML2.IXMLDOMElement
    Dim objElem2  As MSXML2.IXMLDOMElement
    Dim objSub    As MSXML2.IXMLDOMElement
    Dim K         As String
    Dim v         As String

    Set objXML = New MSXML2.DOMDocument60

    MapOwest = 1E+99
    MapNorth = 1E+99
    MapEast = -1E+99
    MapSouth = -1E+99

    minLayer = 0
    maxLayer = 0


    NNode = 0
    NWay = 0
    Erase Node
    Erase WAY
    Erase CAR

    Set DJ = New clsDijkstra
    Set NODESbyFileID = New_c.Collection(False)

    SETPROGRESS "Start reading   " & MAPSsubFolder & frmMain.File1.FileName

    If objXML.Load(FN) Then
        NNode = 0
        For Each objElem In objXML.selectNodes("//node")

            NNode = NNode + 1
            If (NNode And 1023&) = 0& Then
                '                frmMain.Caption = "reading nodes.... " & NNode
                SETPROGRESS "reading nodes.... " & NNode
                DoEvents
            End If
            If NNode >= MaxNNodes Then
                MaxNNodes = NNode + 10000
                ReDim Preserve Node(MaxNNodes)
            End If


            With Node(NNode)

                'Set objElem = objXML.selectSingleNode("//node")
                'Debug.Print objElem.getAttribute("id")
                'Debug.Print objElem.getAttribute("lat")
                'Debug.Print objElem.getAttribute("lon")

                .id = NNode
                .fileID = (objElem.getAttribute("id"))

                NODESbyFileID.Add NNode, .fileID




                X = Val(objElem.getAttribute("lat"))
                Y = Val(objElem.getAttribute("lon"))

                LatLongToUTM X, Y, utmX, utmY

                .X = utmX
                .Y = utmY

                If utmX < MapOwest Then MapOwest = utmX
                If utmX > MapEast Then MapEast = utmX
                If utmY < MapNorth Then MapNorth = utmY
                If utmY > MapSouth Then MapSouth = utmY

                '                ' iterate its sub-nodes
                '                For Each objSub In objElem.childNodes
                '                    Debug.Print objSub.xml
                '                    Stop
                '                Next
            End With
        Next

        EastOwestMapExtension = MapEast - MapOwest
        NorthSouthMapExtension = MapSouth - MapNorth

        '---------------------


        NWay = 0
        For Each objElem In objXML.selectNodes("//way")
            If (NWay And 1023&) = 0& Then
                '                frmMain.Caption = "reading Ways.... " & NWay
                SETPROGRESS "reading Ways.... " & NWay
                DoEvents
            End If

            'Debug.Print "---------------------------------------------------------"
            'Debug.Print "way:   " & objElem.getAttribute("id")

            NWay = NWay + 1
            If NWay > MaxNways Then
                MaxNways = NWay + 2500
                ReDim Preserve WAY(MaxNways)
            End If

            With WAY(NWay)
                .fileID = objElem.getAttribute("id")
                For Each objElem2 In objElem.childNodes    ' objXML2.selectNodes("//nd ref")    'objXML2.childNodes
                    If objElem2.getAttribute("ref") <> vbNull Then
                        'Add nodes
                        .NN = .NN + 1
                        ReDim Preserve .N(.NN)

                        ReDim Preserve .SegAngle(.NN)
                        ReDim Preserve .SegDX(.NN)
                        ReDim Preserve .SegDY(.NN)

                        .N(.NN) = objElem2.getAttribute("ref")
                        'Debug.Print objElem2.getAttribute("ref")
                    End If

                    If objElem2.getAttribute("k") <> vbNull Then
                        K = objElem2.getAttribute("k")
                        v = objElem2.getAttribute("v")
                        'Debug.Print K & vbTab & V

                        .Lanes = 2


                        Select Case K
                        Case "highway", "railway"

                            .wayType = Replace$(v, ",", ".")

                            .isDRIVEABLE = _
                            (.wayType <> "pedestrian" And _
                             .wayType <> "path" And _
                             .wayType <> "steps" And _
                             .wayType <> "footway" And _
                             .wayType <> "cycleway" And _
                             .wayType <> "rail" And _
                             K <> "railway")


                            If v = "mini_roundabout" Then
                                .OneWayDirections = 1
                            End If


                            If K = "railway" And .wayType = "rail" Then .IsRailWay = True


                        Case "name"
                            .Name = v

                        Case "width"
                            .RoadWidth = Val(RemoveM(v))    '6M' '-------------------------------------TO DO




                        Case "building"
                            If v <> "no" Then
                                .IsBuilding = True
                                If .Name = vbNullString Then
                                    If v <> "yes" And v <> "hut" And v <> "roof" And _
                                       v <> "residential" And v <> "detached" And _
                                       v <> "house" And _
                                       v <> "apartments" _
                                       Then .Name = v
                                End If
                            End If

                        Case "area"
                            If v <> "no" Then .IsArea = True

                        Case "waterway"
                            If v = "riverbank" Then .IsWater = True

                        Case "leisure"
                            If v <> "no" Then .IsLeisure = True
                            If v = "swimming_pool" Then .IsLeisure = True: .Name = "Swimming Pool"

                        Case "amenity"
                            .IsAmenity = True
                            If .Name = vbNullString Then .Name = v
                            If .Name = "bench" Then .IsAmenity = False

                        Case "lanes"
                            If v = "1" Then
                                '       .OneWayDirections = 1 '''   ???????????
                            End If
                            .Lanes = Val(v)
                            If .Lanes = 0 Then .Lanes = 2    '''????


                        Case "junction"
                            If v = "roundabout" Then
                                .OneWayDirections = 1
                            End If

                        Case "shop"
                            .IsShop = True
                            If .Name = vbNullString Then .Name = v

                        Case "natural"
                            If v = "water" Then
                                .IsWater = True
                                If .Name = vbNullString Then .Name = v
                            End If
                            '-------------------

                        Case "surface"
                            If v <> "asphalt" Then .isNotAsphalt = True

                        Case "layer"
                            .Layer = Val(v)




                        Case "boundary"

                            .IsBoundary = True








                        Case "oneway"
                            If v = "yes" Or v = "true" Then
                                .OneWayDirections = 1
                            ElseIf v = "-1" Then
                                .OneWayDirections = -1
                            ElseIf v = "no" Then
                                .OneWayDirections = 0
                            End If

                        End Select


                        If InStr(1, K, "parking") Then
                            If v <> "no" Then .IsStreetSideParkable = True
                        End If


                        If LenB(.Name) Then .Name = Replace$(.Name, ",", ".")

                        If .RoadWidth < kRoadW Then .RoadWidth = kRoadW


                    End If

                Next
                .invNN = 1 / .NN
            End With
        Next
    End If

    SETPROGRESS "Sorting Ways..."

    SortWaysByLayer WAY, 1, NWay


    frmMain.Caption = "  " & NNode & "  Nodes and " & NWay & "  Ways  readed."



End Sub


''Private Function FindNodeID(fileID As Currency) As Long    ', PrevI As Long) As Long
''    FindNodeID = NODESbyFileID.Item(fileID)

''    Dim I         As Long
''    Dim J         As Long
''    Static prevI  As Long
''
''    I = prevI
''    J = I
''    Do
''        I = I + 1&
''        I = ((I - 1&) Mod NNode) + 1&
''        J = (NNode + J - 2) Mod NNode + 1
''        If Node(I).fileID = fileID Then FindNodeID = I: prevI = I - 1: Exit Do
''        If Node(J).fileID = fileID Then FindNodeID = J: prevI = J + 1: Exit Do
''    Loop While True
''End Function

Private Sub DIJKSTRAsetupCosts()
    Dim I&, J&
    Dim myNode&


    For I = 1 To NWay
        If WAY(I).isDRIVEABLE Then
            For J = 2 To WAY(I).NN


                If WAY(I).OneWayDirections >= 0 Then    'Street Forward
                    myNode = WAY(I).N(J - 1)
                    With Node(myNode)
                        If .dIsWay Then
                            .NNext = .NNext + 1
                            ReDim Preserve .NEXTnode(.NNext)
                            ReDim Preserve .NextNodeCOST(.NNext)


                            .NEXTnode(.NNext) = WAY(I).N(J)
                            .NextNodeCOST(.NNext) = calcStreetSegmentCOST(myNode, WAY(I).N(J) * 1)


                            ReDim Preserve .NEXTNodeWay(.NNext)
                            ReDim Preserve .NextNodeWAYSegment(.NNext)

                            .NEXTNodeWay(.NNext) = I
                            .NextNodeWAYSegment(.NNext) = J - 1

                        End If
                    End With
                    'Else
                    'MsgBox "node " & myNode & " Not used"
                End If

                If WAY(I).OneWayDirections <= 0 Then    'Street Backward
                    myNode = WAY(I).N(J)
                    With Node(myNode)
                        If .dIsWay Then
                            .NNext = .NNext + 1
                            ReDim Preserve .NEXTnode(.NNext)
                            ReDim Preserve .NextNodeCOST(.NNext)
                            .NEXTnode(.NNext) = WAY(I).N(J - 1)
                            .NextNodeCOST(.NNext) = calcStreetSegmentCOST(myNode, WAY(I).N(J - 1) * 1)


                            ReDim Preserve .NEXTNodeWay(.NNext)
                            ReDim Preserve .NextNodeWAYSegment(.NNext)


                            .NEXTNodeWay(.NNext) = I
                            'Negative Sign to indicate Opposite Direction (for Cost)
                            ' in "DIJKSTRAvisit"
                            .NextNodeWAYSegment(.NNext) = -(J - 1)


                        End If
                    End With

                    'MsgBox "node " & myNode & " Not used"
                End If
            Next
        End If
    Next
End Sub

Private Sub DIJKSTRAsetup()
    Dim I         As Long
    Dim J         As Long

    Dim MinX      As Double
    Dim MinY      As Double
    ''''''''''''''''''''''''''''''''''''for dijkstra
    For I = 1 To NWay
        If (I And 1023&) = 0& Then
            SETPROGRESS "DIJKSTRA setup ...." & I & " / " & NWay, I / NWay: DoEvents
        End If

        With WAY(I)
            If Not (.isDRIVEABLE) Or .IsArea Or .IsBuilding Or .IsWater Or .IsAmenity Then
            Else
                If Len(.wayType) And .wayType <> "pedestrian" _
                   And .wayType <> "path" _
                   And .wayType <> "steps" _
                   And .wayType <> "footway" _
                   And .wayType <> "cycleway" Then
                    For J = 1 To .NN
                        Node(.N(J)).dIsWay = True
                        Node(.N(J)).ONEWAY = .OneWayDirections
                        'Debug.Print .wayType, .NAME: Stop
                    Next
                End If
            End If
        End With
    Next

    '--------------------------------------------------
    '''Setup ONEways (for DIJKSTRA) & Draw
    DIJKSTRAsetupCosts

    '--------------------------   END [DIJKSTRAsetup]


    '---------------------TRANSLATE ALL

    MinX = 1E+99
    MinY = 1E+99
    For I = 1 To NNode
        If Node(I).dIsWay Then    '2024
            If Node(I).X < MinX Then MinX = Node(I).X
            If Node(I).Y < MinY Then MinY = Node(I).Y
        End If
    Next


    MinX = MinX + 20
    MinY = MinY + 20

    For I = 1 To NNode
        Node(I).X = Round(Node(I).X - MinX, 2)
        Node(I).Y = Round(Node(I).Y - MinY, 2)
    Next


    'SET World Coordinates Bounding Boxes
    Dim X#, Y#
    For I = 1 To NWay
        With WAY(I)
            .BBx1 = 1E+99
            .BBY1 = 1E+99
            .BBx2 = -1E+99
            .BBY2 = -1E+99
            For J = 1 To .NN
                X = Node(.N(J)).X
                Y = Node(.N(J)).Y
                If X < .BBx1 Then .BBx1 = X
                If Y < .BBY1 Then .BBY1 = Y
                If X > .BBx2 Then .BBx2 = X
                If Y > .BBY2 Then .BBY2 = Y
            Next
            If .IsBuilding Then   'Some Random color
                .R = Rnd * 2 - 1
                .G = Rnd * 2 - 1
                .B = Rnd * 2 - 1
            End If
        End With
    Next

    frmMain.Caption = "": DoEvents

End Sub
Public Sub PURGEAndSave(FN As String)
    Dim I         As Long
    Dim J         As Long
    Dim K         As Long

    '    Dim PREV   As Long

    Dim NID       As Currency
    Dim DX#, DY#, A#

    ptsUSED = 0

    'Find Used Points---------------------And Convert FileID to ID
    For I = 1 To NWay
        With WAY(I)
            For J = 1 To .NN
                'NID = FindNodeID( .N(J))
                NID = NODESbyFileID.Item(.N(J)) '<----------------------------- !!!

                .N(J) = NID
                '  .N(J) = FindNodeID( .N(J), PREV)
                If NID Then Node(.N(J)).Used = True

                If .isDRIVEABLE Then
                    If .Layer > maxLayer Then maxLayer = .Layer
                    If .Layer < minLayer Then minLayer = .Layer
                End If
            Next



            '-------------------------------------------- way angles
            For J = 1 To .NN - 1
                K = J + 1
                DX = Node(.N(K)).X - Node(.N(J)).X
                DY = Node(.N(K)).Y - Node(.N(J)).Y
                A = Atan2(DX, DY)
                If A < 0 Then A = A + PI2
                .SegAngle(J) = Round(A, 4)
                .SegDX(J) = Round(Cos(A), 4)
                .SegDY(J) = Round(Sin(A), 4)
            Next
            For J = 1 To .NN
                '            For J = 2 To .NN - 1
                Node(.N(J)).Layer = .Layer
            Next
            '-------------------------------------------------------


            If (I And 1023&) = 0& Then
                '            frmMain.Caption = "Elaborating Ways ... " & I & "/" & NWay
                SETPROGRESS "Elaborating Ways ... " & I & "/" & NWay, I / NWay
                DoEvents
            End If
        End With
    Next


    '    CalcWaysAngles

    For I = 1 To NNode
        If Node(I).Used Then ptsUSED = ptsUSED + 1
    Next
    '--------------------------------------------------------------------------------------



    DIJKSTRAsetup                 '<---------------------------------------------

    '    SAVEPURGED

    frmMain.Caption = "ready"
    DoEvents

End Sub

Public Sub SAVEPURGED()
    Dim I         As Long
    Dim J         As Long

    '--------------------------------------------------------------------------------------

    '    frmMain.Caption = "Saving points .... used " & ptsUSED & " of " & NNode
    DoEvents


    Open App.Path & "\OUT.txt" For Output As 1

    Print #1, minLayer
    Print #1, maxLayer

    Print #1, NNode               ' ptsUSED
    For I = 1 To NNode
        If (I And 1023&) = 0 Then SETPROGRESS "Saving Points... " & I, I / NNode
        'If Node(I).Used Then
        Print #1, Replace(CStr(Node(I).X), ",", ".")
        Print #1, Replace(CStr(Node(I).Y), ",", ".")

        Print #1, Node(I).Layer
        'End If
    Next





    Print #1, "---------------------------------------------------"
    Print #1, NWay
    For I = 1 To NWay

        With WAY(I)
            If (I And 1023&) = 0 Then
                '                frmMain.Caption = "Saving Way .... " & I
                DoEvents
                SETPROGRESS "Saving Ways... " & I, I / NWay
            End If

            Print #1, "Way -------------" & I
            Print #1, .NN
            For J = 1 To .NN
                Print #1, Node(.N(J)).id
                Print #1, Replace(CStr(.SegAngle(J)), ",", ".")
                Print #1, Replace(CStr(.SegDX(J)), ",", ".")
                Print #1, Replace(CStr(.SegDY(J)), ",", ".")
            Next

            Print #1, .wayType
            'If LenB(.NAME) Then
            Print #1, .Name
            'Else
            '    Print #1, "no name"
            'End If

            Print #1, IIf(.isDRIVEABLE, "yes", "no")

            Print #1, IIf(.IsRailWay, "yes", "no")
            Print #1, IIf(.IsBoundary, "yes", "no")



            Print #1, IIf(.IsArea, "yes", "no")
            Print #1, IIf(.IsBuilding, "yes", "no")
            Print #1, IIf(.IsWater, "yes", "no")

            Print #1, IIf(.IsLeisure, "yes", "no")
            Print #1, IIf(.IsAmenity, "yes", "no")
            Print #1, IIf(.IsShop, "yes", "no")
            Print #1, IIf(.isNotAsphalt, "yes", "no")


            Print #1, .Layer
            Print #1, .Lanes
            '            Print #1, IIf(.IsNatural, "yes", "no")

            Print #1, .OneWayDirections

            Print #1, Replace(.RoadWidth, ",", ".")
            Print #1, Replace(.invNN, ",", ".")

        End With

    Next

    Close 1

End Sub

Private Function YesNo2Bool(S As String) As Boolean
    If S = "yes" Then YesNo2Bool = True

End Function
Public Sub LOADpurged()
    Dim I         As Currency
    Dim J         As Currency
    Dim S         As String

    '--------------------------------------------------------------------------------------

    '    frmMain.Caption = "LOADING... "
    SETPROGRESS "LOADING... "
    DoEvents

    Open App.Path & "\OUT.txt" For Input As 1

    Input #1, minLayer
    Input #1, maxLayer


    Input #1, ptsUSED

    NNode = ptsUSED

    ReDim Node(NNode)

    MapOwest = 1E+99
    MapNorth = 1E+99
    MapEast = -1E+99
    MapSouth = -1E+99
    For I = 1 To NNode
        Input #1, Node(I).X
        Input #1, Node(I).Y
        If Node(I).X < MapOwest Then MapOwest = Node(I).X
        If Node(I).X > MapEast Then MapEast = Node(I).X
        If Node(I).Y < MapNorth Then MapNorth = Node(I).Y
        If Node(I).Y > MapSouth Then MapSouth = Node(I).Y

        Input #1, Node(I).Layer
    Next
    EastOwestMapExtension = MapEast - MapOwest
    NorthSouthMapExtension = MapSouth - MapNorth




    Input #1, S                   '"---------------------------------------------------"
    Input #1, NWay
    ReDim WAY(NWay)
    For I = 1 To NWay

        With WAY(I)
            If (I And 1023&) = 0& Then
                '                frmMain.Caption = "Loading Way .... " & I


                SETPROGRESS "Loading Way .... ", I / NWay
                DoEvents
            End If

            Input #1, S           '"Way -------" & I
            Input #1, .NN

            ReDim .N(.NN)
            ReDim .SegAngle(.NN)
            ReDim .SegDX(.NN)
            ReDim .SegDY(.NN)
            For J = 1 To .NN
                Input #1, .N(J)
                Input #1, .SegAngle(J)
                Input #1, .SegDX(J)
                Input #1, .SegDY(J)

            Next


            Input #1, .wayType
            Input #1, .Name       ': If LenB(.NAME) = 0 Then Stop

            Input #1, S: .isDRIVEABLE = YesNo2Bool(S)

            Input #1, S: .IsRailWay = YesNo2Bool(S)
            Input #1, S: .IsBoundary = YesNo2Bool(S)


            Input #1, S: .IsArea = YesNo2Bool(S)
            Input #1, S: .IsBuilding = YesNo2Bool(S)
            Input #1, S: .IsWater = YesNo2Bool(S)

            Input #1, S: .IsLeisure = YesNo2Bool(S)
            Input #1, S: .IsAmenity = YesNo2Bool(S)
            Input #1, S: .IsShop = YesNo2Bool(S)
            Input #1, S: .isNotAsphalt = YesNo2Bool(S)

            Input #1, S: .Layer = S
            Input #1, S: .Lanes = S

            Input #1, .OneWayDirections

            Input #1, .RoadWidth

            Input #1, .invNN

        End With

    Next

    Close 1

    DIJKSTRAsetup                 '<---------------------------------------------

    '    frmMain.Caption = "Done": DoEvents

End Sub



Public Sub UPDATETRAFFIC()
    Dim I&, N1&, N2&, J&


    For I = 1 To NNode
        Node(I).Traffic = 0
    Next

    For I = 1 To Ncars
        N1 = CAR(I).NEXTnode
        Node(N1).Traffic = Node(N1).Traffic + 1
        '        For J = 1 To Node(N1).NNext
        '            N2 = Node(N1).NEXTnode(J)
        '            Node(N2).Traffic = Node(N2).Traffic + 1
        '        Next
    Next

    SETPROGRESS "TRAFFIC"

    DIJKSTRAsetupCosts

End Sub

Private Function CalcWaysAngles()
    Dim I         As Long
    Dim J         As Long
    Dim K         As Long
    Dim DX#, DY#, A#
    For I = 1 To NWay
        With WAY(I)
            For J = 1 To .NN - 1
                K = J + 1
                DX = Node(.N(K)).X - Node(.N(J)).X
                DY = Node(.N(K)).Y - Node(.N(J)).Y
                A = Atan2(DX, DY)
                If A < 0 Then A = A + PI2
                .SegAngle(J) = Round(A, 4)
                .SegDX(J) = Round(Cos(A), 4)
                .SegDY(J) = Round(Sin(A), 4)
            Next
            For J = 1 To .NN
                '            For J = 2 To .NN - 1
                Node(.N(J)).Layer = .Layer
            Next
        End With
    Next
End Function

Private Function SortWaysByLayer(WAys() As tRoad, ByVal Min As Long, ByVal Max As Long)
    ' FROM HI to LOW  'https://www.vbforums.com/showthread.php?11192-quicksort
    Dim Low As Long, high As Long, temp As tRoad
    Dim TestDist#
    'Debug.Print min, max
    Low = Min: high = Max
    '    TestDist = (WAys(min).Layer + WAys(max).Layer) * 0.5
    TestDist = WAys((Min + Max) * 0.5).Layer
    Do

        Do While (WAys(Low).Layer < TestDist): Low = Low + 1&: Loop
        Do While (WAys(high).Layer > TestDist): high = high - 1&: Loop

        If (Low <= high) Then
            temp = WAys(Low): WAys(Low) = WAys(high): WAys(high) = temp
            Low = Low + 1&: high = high - 1&
        End If
    Loop While (Low <= high)
    If (Min < high) Then SortWaysByLayer WAys, Min, high
    If (Low < Max) Then SortWaysByLayer WAys, Low, Max

End Function
