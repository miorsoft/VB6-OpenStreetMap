VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Drive OSM"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11370
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   730
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   758
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSTAT 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   11775
      Left            =   120
      ScaleHeight     =   783
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   21
      Top             =   120
      Width           =   2415
      Begin VB.Label labSTAT 
         Caption         =   "--------------------------------------"
         Height          =   3375
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   8400
      ScaleHeight     =   687
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkTargetAll 
         Caption         =   "Same Target for every car"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "On mouse right click set same target for all the cars"
         Top             =   8760
         Width           =   1455
      End
      Begin VB.TextBox txtFollow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         TabIndex        =   24
         Text            =   "0"
         Top             =   5760
         Width           =   735
      End
      Begin VB.CheckBox ChkFollow 
         Caption         =   "Follow Car"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   855
         Left            =   2280
         TabIndex        =   18
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Dijkstra TEST"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2280
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.FileListBox File1 
         Height          =   2790
         Left            =   120
         Pattern         =   "*.OSM"
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   975
         Left            =   2280
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "load purged"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CheckBox chSave 
         Caption         =   "PNG"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   9960
         Width           =   1095
      End
      Begin VB.CommandButton cmdNextCar 
         Caption         =   ">> Next Car"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrevCar 
         Caption         =   "<< Prev Car"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load MAP"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox chkShowPath 
         Caption         =   "ShowPath"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "+"
         Height          =   495
         Index           =   5
         Left            =   1320
         TabIndex        =   8
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "o"
         Height          =   495
         Index           =   6
         Left            =   840
         TabIndex        =   7
         Top             =   6360
         Width           =   375
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "-"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   5
         Top             =   7200
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "^"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   6960
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "v"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   7560
         Width           =   495
      End
      Begin VB.CommandButton cmdNAVIGATE 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "LEFT CLICK to select Car"
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label lSize 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   1815
      End
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3975
      Left            =   3720
      MousePointer    =   1  'Arrow
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private N1        As Long
Private N2        As Long


Private WithEvents ZIP As cZipArchive
Attribute ZIP.VB_VarHelpID = -1

Private ZipName   As String

Private TargetAllCars As Boolean


Private Sub ChkFollow_Click()
    Dim I         As Long

    DoFollow = ChkFollow = vbChecked

    If Not (DoFollow) Then
        PanXtoGo = PanX
        PanYtoGo = PanY
        PIC.MousePointer = 2
    Else
        PIC.MousePointer = 0
    End If

    For I = 0 To 3
        cmdNAVIGATE(I).Enabled = Not (DoFollow)
    Next

End Sub

Private Sub chkShowPath_Click()
    ShowPath = chkShowPath = vbChecked

End Sub

Private Sub chkTargetAll_Click()
    TargetAllCars = chkTargetAll = vbChecked
End Sub

Private Sub chSave_Click()
    DoSaveFrame = (chSave.Value = vbChecked)


End Sub

Private Sub cmdNAVIGATE_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = True

    '' Do

    Select Case Index
        Case 0
            PanZoomChanged = True
            PanYtoGo = PanYtoGo - 500 / ZoomToGo

        Case 1
            PanZoomChanged = True
            PanYtoGo = PanYtoGo + 500 / ZoomToGo


        Case 2
            PanZoomChanged = True
            PanXtoGo = PanXtoGo - 500 / ZoomToGo

        Case 3
            PanZoomChanged = True
            PanXtoGo = PanXtoGo + 500 / ZoomToGo


        Case 4
            PanZoomChanged = True
            ZoomToGo = ZoomToGo / 1.2
        Case 5
            PanZoomChanged = True
            ZoomToGo = ZoomToGo * 1.2   '1.2

        Case 6
            PanZoomChanged = True
            ZoomToGo = 1
            PanX = CENPanX
            PanY = CENPanY
            ZoomToGo = PIC.Height / scrMaxY
    End Select

    '        SW.DRAW
    '        PIC.Refresh

    DoEvents


    InvZoom = 1 / Zoom
    ''  Loop While Navigating


    ''    DRAWCC
    ''    '    PIC.Refresh
    ''
    ''    PIC.Picture = SRF.Picture
    ''    DoEvents

End Sub

Private Sub cmdNAVIGATE_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Navigating = False

End Sub

Private Sub cmdNextCar_Click()
    Follow = Follow Mod Ncars + 1
    txtFollow = Follow
End Sub

Private Sub cmdPrevCar_Click()
    Follow = Follow - 1: If Follow < 1 Then Follow = Ncars
    txtFollow = Follow
End Sub

Private Sub Command1_Click()

    Dim I         As Long

    'ReadOSM (App.Path & "\OSM\chulavista1.osm")

    'ReadOSM (App.Path & "\OSM\Tokyo1.osm")


    'ReadOSM (App.Path & "\OSM\PalmanovaL.osm")
    Command1.Enabled = False

    '  ReadOSM (App.Path & "\OSM\ny1.osm")    '**************************
    'ReadOSM (App.Path & "\OSM\roma1.osm")

    'ReadOSM (App.Path & "\OSM\Lakeplacid.osm")


    'ReadOSM (App.Path & "\OSM\monaco-latest.osm") '''''''''''''''''''''''''''''''''

    'ReadOSM (App.Path & "\OSM\isle-of-man-latest.osm")

    'ReadOSM (App.Path & "\OSM\malta-latest.osm") ''SLOW
    '    ReadOSM (App.Path & "\OSM\liechtenstein-latest.osm")

    'ReadOSM (App.Path & "\OSM\Firenze.osm")



    ReadOSM (App.Path & MAPSsubFolder & File1.FileName)



    Dim sgGirdSize&
    sgGirdSize = EastOwestMapExtension \ 25
    If sgGirdSize > 1000 Then sgGirdSize = 1000
    If sgGirdSize < 250 Then sgGirdSize = 250
    '    GRID.Init EastOwestMapExtension + 1, NorthSouthMapExtension + 1, sgGirdSize
    QT.Setup 0, 0, EastOwestMapExtension, NorthSouthMapExtension, 4



    PURGEAndSave App.Path & "\Out.txt"

    InitDraw

    DRAWMAP

    PIC.Picture = SRF.Picture

    Command3.Enabled = True

    CARSSETUP

    If Not DoLOOP Then MAINLOOP

End Sub

Private Sub Command2_Click()
    ' MsgBox (DownloadFile("http://www.vbforums.com/showthread.php?709829-Can-VB6-download-any-kind-of-file&highlight=download+from+url", App.Path & "\Test.osm"))

    MsgBox (DownloadFile("http://api.openstreetmap.org/api/0.6/map?bbox=11.8705,45.3903,11.9343,45.4256", App.Path & "\Test.osm"))
End Sub

Private Sub Command3_Click()

    ''    Dim X      As Long
    ''    Dim Y      As Long
    ''    Dim I      As Long
    ''
    ''
    ''    If N1 = 0 Or N2 = 0 Then
    ''
    ''
    ''        Do
    ''            N1 = 1 + Rnd * NNode
    ''        Loop While Node(N1).dIsWay = False
    ''        Do
    ''            N2 = 1 + Rnd * NNode
    ''        Loop While Node(N2).dIsWay = False
    ''
    ''    End If
    ''
    ''
    ''
    ''
    ''    For X = 1 To NNode: Node(X).dCost = 0: Next
    ''    DRAWCC
    ''    X = XtoScreen(Node(N1).X)
    ''    Y = YtoScreen(Node(N1).Y)
    ''    MyCircle frmMainPIChDC, X, Y, 4, 8, vbMagenta
    ''    X = XtoScreen(Node(N2).X)
    ''    Y = YtoScreen(Node(N2).Y)
    ''
    ''
    ''
    ''
    ''    MyCircle frmMainPIChDC, X, Y, 4, 8, vbBlue
    ''    frmMain.PIC.Refresh
    ''    DoEvents
    ''
    ''
    ''
    ''    dijPATH = DJ.DIJKSTRA(N1, N2)
    ''
    ''    ''    If dijPATH.NN <> 0 Then
    ''    ''        X = 0
    ''    ''        Y = 0
    ''    ''        For I = 1 To dijPATH.NN
    ''    ''            X = X + Node(dijPATH.N(I)).X
    ''    ''            Y = Y + Node(dijPATH.N(I)).Y
    ''    ''        Next
    ''    ''        X = X / dijPATH.NN
    ''    ''        Y = Y / dijPATH.NN
    ''    ''        PanX = (PanX + X) * 0.5
    ''    ''        PanY = (PanY + Y) * 0.5
    ''    ''    End If
    ''
    ''    DRAW
    ''    frmMain.PIC.Refresh
    ''    DoEvents
    ''


End Sub

Private Sub Command4_Click()
    Dim I         As Long

    ''Ncars = 10
    ''
    ''ReDim CAR(Ncars)
    ''
    ''For I = 1 To Ncars
    ''CAR(I).RANDOMstart
    ''CAR(I).RANDOMend
    ''Next

End Sub



Private Sub File1_Click()
    Dim FF        As Long


    If Len(File1.FileName) = 0 Then Exit Sub
    Command1.Enabled = True

    FF = FreeFile
    Open App.Path & MAPSsubFolder & File1.FileName For Random As FF
    lSize = Int((LOF(FF) / 1024) / 1024 * 10) * 0.1 & " Mbytes"
    Close FF

End Sub

Private Sub Form_Activate()

    If Not (New_c.FSO.FileExists(App.Path & "\ZipsExtracted.txt")) Then
        New_c.FSO.WriteTextContent App.Path & "\ZipsExtracted.txt", "0"
    End If


    If Val(New_c.FSO.ReadTextContent(App.Path & "\ZipsExtracted.txt")) = 0 Then
        Set ZIP = New cZipArchive
        ZipName = Dir(App.Path & "\MAPS\*.Zip")
        Do
            ZIP.OpenArchive App.Path & "\MAPS\" & ZipName
            ZIP.Extract App.Path & "\MAPS\"
            ZipName = Dir
        Loop While ZipName <> ""
        File1.Refresh
        SETPROGRESS "ZIP Maps Extraction complete" & vbCrLf & "Ready to load a Map."

        New_c.FSO.WriteTextContent App.Path & "\ZipsExtracted.txt", "1"

    End If


End Sub

Private Sub Form_Load()
    Randomize Timer


    File1.Path = App.Path & MAPSsubFolder


    PanZoomChanged = True


    If Dir(App.Path & "\Frames", vbDirectory) = vbNullString Then MkDir App.Path & "\Frames"
    If Dir(App.Path & "\Frames\*.*") <> vbNullString Then Kill App.Path & "\Frames\*.*"

    '    Set GRID = New cSpatialGrid
    Set QT = New cQuadTree



    LoadImages

    frmMainPIChDC = frmMain.PIC.hDC


    Label1.Caption = "LEFT Click = Select Car" & vbCrLf & _
                     "RIGHT Click = Set Target" & vbCrLf & _
                     "MouseWHEEL = ZOOM"


    ChkFollow = vbChecked


    Call WheelHook(PIC.hwnd)

End Sub
Private Function NearestTo(X As Double, Y As Double, OnlyWays As Boolean) As Long
    Dim I         As Long
    Dim Dmin      As Double
    Dim DX        As Double
    Dim DY        As Double
    Dim D         As Double

    Dmin = 1E+99
    If OnlyWays Then
        For I = 1 To NNode
            If Node(I).dIsWay Then
                DX = Node(I).X - X
                DY = Node(I).Y - Y
                D = DX * DX + DY * DY
                If D < Dmin Then Dmin = D: NearestTo = I
            End If
        Next
    Else
        For I = 1 To NNode
            DX = Node(I).X - X
            DY = Node(I).Y - Y
            D = DX * DX + DY * DY
            If D < Dmin Then Dmin = D: NearestTo = I
        Next
    End If


End Function

Private Sub Form_Resize()

    'https://stackoverflow.com/questions/70758773/how-to-automatically-resize-or-reposition-controls-on-a-form-when-the-form-is-re
    If WindowState = 1 Then Exit Sub


    PIC.Height = (frmMain.ScaleHeight \ 128) * 128    ' 1024             '1024 '720
    PIC.Width = Int(PIC.Height * 4 / 3 * 1)
    PIC.Width = PIC.Width - (PIC.Width Mod 4)
    scrMaxX = frmMain.PIC.Width
    scrMaxY = frmMain.PIC.Height
    CenX = scrMaxX \ 2
    CenY = scrMaxY \ 2

    Set SRF = Cairo.CreateSurface(scrMaxX, scrMaxY, ImageSurface)
    Set CC = SRF.CreateContext
    CC.AntiAlias = CAIRO_ANTIALIAS_FAST
    CC.SelectFont "Segoe UI", 11, vbWhite
    CC.SetLineCap CAIRO_LINE_CAP_ROUND







    PIC.Left = Me.ScaleWidth * 0.5 - PIC.Width * 0.5
    PIC.Top = 10                        'Me.ScaleHeight * 0.5 - PIC.Height * 0.5

    picPanel.Top = PIC.Top
    picPanel.Left = PIC.Left + PIC.Width + 5
    picPanel.Height = PIC.Height

    picSTAT.Top = PIC.Top
    picSTAT.Height = PIC.Height
    picSTAT.Left = PIC.Left - picSTAT.Width - 5

    labSTAT.Width = picSTAT.Width - 10







End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase Node
    Erase WAY
    Erase CAR

    Call WheelUnHook(PIC.hwnd)

    End

End Sub

Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xx        As Double
    Dim yy        As Double
    Dim I&
    Dim J&, DX#, DY#
    Dim DD#, MDD#
    ''    XX = xfromScreen(X * 1)
    ''    YY = yfromScreen(Y * 1)
    ''
    ''    If Button = 1 Then
    ''        N1 = NearestTo(XX * 1, YY * 1)
    ''
    ''
    ''        XX = XtoScreen(Node(N1).X)
    ''        YY = YtoScreen(Node(N1).Y)
    ''
    ''    ElseIf Button = 2 Then
    ''        N2 = NearestTo(XX * 1, YY * 1)
    ''
    ''        XX = XtoScreen(Node(N2).X)
    ''        YY = YtoScreen(Node(N2).Y)
    ''    Else
    ''
    ''        PanZoomChanged = True
    ''        PanX = XX
    ''        PanY = YY
    ''        DRAW
    ''        frmMain.PIC.Refresh
    ''        DoEvents
    ''
    ''    End If
    ''
    ''
    ''    'test funcitons
    ''    '    xx = XtoScreen(xx * 1)
    ''    '    yy = YtoScreen(yy * 1)
    ''    MyCircle frmMainPIChDC, XX * 1, YY * 1, 5, 10, vbRed
    ''    frmMain.PIC.Refresh
    ''    DoEvents

    xx = xfromScreen(X * 1)
    yy = yfromScreen(Y * 1)


    If Button = 1 Then                  'Set car to Follow
        MDD = 1E+99

        If DoFollow Then
            For I = 1 To Ncars
                DX = xx - CAR(I).Xfront
                DY = yy - CAR(I).Yfront
                DD = DX * DX + DY * DY
                If DD < MDD Then MDD = DD: J = I
            Next
            Follow = J
            txtFollow = Follow

        Else
            J = NearestTo(xx, yy, False)
            PanXtoGo = (Node(J).X)
            PanYtoGo = (Node(J).Y)
        End If

    ElseIf Button = 2 Then              ' SET TARGET
        N1 = NearestTo(xx * 1, yy * 1, True)
        CAR(Follow).SetEndNode N1
        If CAR(Follow).GetPATHNN = 0 Then CAR(Follow).RANDOMend

        If TargetAllCars Then
            For I = 1 To Ncars
                CAR(I).SetEndNode N1
                If CAR(I).GetPATHNN = 0 Then CAR(I).RANDOMend
            Next
        End If

    End If

    'MsgBox Button





End Sub

Private Sub Command5_Click()            ' CMD LOADPURGED
    Dim I         As Long

    LOADpurged

    Dim sgGirdSize&
    sgGirdSize = EastOwestMapExtension \ 25
    If sgGirdSize > 1000 Then sgGirdSize = 1000
    If sgGirdSize < 250 Then sgGirdSize = 250
    '    GRID.Init EastOwestMapExtension + 1, NorthSouthMapExtension + 1, sgGirdSize
    QT.Setup 0, 0, EastOwestMapExtension, NorthSouthMapExtension, 4

    InitDraw
    'DRAW
    'PIC.Refresh
    DRAWMAP
    PIC.Picture = SRF.Picture

    Command3.Enabled = True

    CARSSETUP

    MAINLOOP

End Sub





' Here you can add scrolling support to controls that don't normally respond
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)

    Dim I         As Long
    'Dim ctl       As Control
    'For Each ctl In Me.Controls
    '    If TypeOf ctl Is MSFlexGrid Then
    '      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    '    End If
    'Next ctl

    For I = 1 To Abs(Rotation) * 0.01
        If Rotation > 0 Then
            ZoomToGo = ZoomToGo * 1.15  '1.18
        Else
            ZoomToGo = ZoomToGo / 1.15  '1.18

        End If
    Next




End Sub

Private Sub Picture1_Click()

End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xx        As Double
    Dim yy        As Double
    Dim J&

    xx = xfromScreen(X * 1)
    yy = yfromScreen(Y * 1)


    If Button = 1 Then                  'Set car to Follow
        '        MDD = 1E+99
        '        If DoFollow Then
        '            For I = 1 To Ncars
        '                DX = xx - CAR(I).Xfront
        '                DY = yy - CAR(I).Yfront
        '                DD = DX * DX + DY * DY
        '                If DD < MDD Then MDD = DD: J = I
        '            Next
        '            Follow = J
        '            txtFollow = Follow
        '
        '        Else 'Set PAN
        J = NearestTo(xx, yy, False)
        PanXtoGo = (Node(J).X)
        PanYtoGo = (Node(J).Y)
    End If
End Sub

Private Sub ZIP_Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, Cancel As Boolean)

    With ZIP
        SETPROGRESS ZipName & vbCrLf & "Extracting: " & .FileInfo(FileIdx)(0) & "  " & FileIdx + 1 & "/" & .FileCount, (FileIdx) / .FileCount
        DoEvents
    End With
End Sub
