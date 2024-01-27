Attribute VB_Name = "modLineClipping"
Option Explicit
'COHEN-SUTHERLAND line clipping

Private Enum tOUTcode
    kINSIDE = 0
    kLEFT = 1
    kRIGHT = 2
    kBOTTOM = 4
    kTOP = 8
End Enum

Private Function OutCode(ByVal X As Double, ByVal Y As Double) As tOUTcode
    OutCode = kINSIDE
    If X < ClipLEFT Then
        OutCode = kLEFT
    ElseIf X > ClipRIGHT Then
        OutCode = kRIGHT
    End If
    If Y < ClipTOP Then
        OutCode = OutCode Or kBOTTOM
    ElseIf Y > ClipBOTTOM Then
        OutCode = OutCode Or kTOP
    End If
End Function


Public Function CLIPLINEcc(X0 As Double, y0 As Double, X1 As Double, Y1 As Double) As Boolean
    Dim oCode0    As tOUTcode
    Dim oCode1    As tOUTcode
    Dim oCodeOUT  As tOUTcode
    Dim X         As Double
    Dim Y         As Double


    '    Dim Accept    As Boolean

    oCode0 = OutCode(X0, y0)
    oCode1 = OutCode(X1, Y1)

    Do
        If (oCode0 Or oCode1) = 0 Then
            CLIPLINEcc = True: Exit Do
        ElseIf (oCode0 And oCode1) Then
            Exit Do
        End If


        If oCode0 <> kINSIDE Then
            oCodeOUT = oCode0
        Else
            oCodeOUT = oCode1
        End If

        If (oCodeOUT And kTOP) Then
            X = X0 + (X1 - X0) * (ClipBOTTOM - y0) / (Y1 - y0)
            Y = ClipBOTTOM
        ElseIf (oCodeOUT And kBOTTOM) Then
            X = X0 + (X1 - X0) * (ClipTOP - y0) / (Y1 - y0)
            '            Y = 0
            Y = ClipTOP
        ElseIf (oCodeOUT And kRIGHT) Then
            Y = y0 + (Y1 - y0) * (ClipRIGHT - X0) / (X1 - X0)
            X = ClipRIGHT
            'ElseIf (oCodeOUT And kLEFT) <> 0 Then
        Else
            Y = y0 + (Y1 - y0) * (ClipLEFT - X0) / (X1 - X0)
            '            X = 0
            X = ClipLEFT
        End If

        '------------
        If oCodeOUT = oCode0 Then
            X0 = X
            y0 = Y
            oCode0 = OutCode(X0, y0)
        Else
            X1 = X
            Y1 = Y
            oCode1 = OutCode(X1, Y1)
        End If

    Loop While True


    '    CLIPLINEcc = Accept
End Function
