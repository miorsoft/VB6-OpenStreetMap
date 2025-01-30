Attribute VB_Name = "modLatLonConversion"
Option Explicit
'http://www.movable-type.co.uk/scripts/latlong-gridref.html

Private Const PI  As Double = 3.14159265358979
Private Const rad As Double = PI / 360
Private Const deg2rad As Double = PI / 180
'Private Const rad2deg As Double = 180# / PI

Private Function toRad(deg As Double) As Double
    toRad = deg * rad
End Function
''/*
'' * convert geodesic co-ordinates to OS grid reference
'' */
'Public Function LatLongToOSGrid(ByVal lat As Double, ByVal lon As Double, _
 '                                ByRef retX As Double, ByRef retY As Double)
'
'    Dim a      As Double
'    Dim b      As Double
'    Dim F0     As Double
'    Dim Lat0   As Double
'    Dim Lon0   As Double
'    Dim N0     As Double
'    Dim E      As Double
'    Dim E0     As Double
'    Dim e2     As Double
'    Dim N      As Double
'    Dim n2     As Double
'    Dim n3     As Double
'    Dim CosLat As Double
'    Dim Cos3Lat As Double
'    Dim Cos5Lat As Double
'
'    Dim SinLat As Double
'    Dim Tan2Lat As Double
'    Dim Tan4Lat As Double
'
'    Dim nu     As Double
'    Dim rho    As Double
'    Dim eta2   As Double
'    Dim M      As Double
'    Dim Ma     As Double
'    Dim Mb     As Double
'    Dim Mc     As Double
'    Dim Md     As Double
'    Dim I      As Double
'    Dim iI     As Double
'    Dim III    As Double
'    Dim IIIA   As Double
'    Dim IV     As Double
'    Dim V      As Double
'    Dim VI     As Double
'    Dim Dlon   As Double
'    Dim Dlon1  As Double
'    Dim Dlon2  As Double
'    Dim Dlon3  As Double
'    Dim Dlon4  As Double
'    Dim Dlon5  As Double
'    Dim Dlon6  As Double
'
'
'    lat = toRad(lat)
'    lon = toRad(lon)
'
'
'    a = 6377563.396: b = 6356256.91    ' Airy 1830 major & minor semi-axes
'    F0 = 0.9996012717             ' NatGrid scale factor on central meridian
'    Lat0 = toRad(49): Lon0 = toRad(-2)    ' NatGrid true origin
'    N0 = -100000: E0 = 400000     ' northing & easting of true origin: metres
'    e2 = 1 - (b * b) / (a * a)    ' eccentricity squared
'    N = (a - b) / (a + b): n2 = N * N: n3 = N * N * N
'
'    CosLat = Cos(lat)
'    SinLat = Sin(lat)
'    nu = a * F0 / Sqr(1 - e2 * SinLat * SinLat)    ' transverse radius of curvature
'    rho = a * F0 * (1 - e2) / (1 - e2 * SinLat * SinLat ^ 1.5)    ' meridional radius of curvature
'    eta2 = nu / rho - 1
'
'    Ma = (1 + N + (5 / 4) * n2 + (5 / 4) * n3) * (lat - Lat0)
'    Mb = (3 * N + 3 * N * N + (21 / 8) * n3) * Math.Sin(lat - Lat0) * Math.Cos(lat + Lat0)
'    Mc = ((15 / 8) * n2 + (15 / 8) * n3) * Math.Sin(2 * (lat - Lat0)) * Math.Cos(2 * (lat + Lat0))
'    Md = (35 / 24) * n3 * Math.Sin(3 * (lat - Lat0)) * Math.Cos(3 * (lat + Lat0))
'    M = b * F0 * (Ma - Mb + Mc - Md)    ' meridional arc
'
'    Cos3Lat = CosLat * CosLat * CosLat
'    Cos5Lat = Cos3Lat * CosLat * CosLat
'    Tan2Lat = Tan(lat) * Tan(lat)
'    Tan4Lat = Tan2Lat * Tan2Lat
'
'    I = M + N0
'    iI = (nu / 2) * SinLat * CosLat
'    III = (nu / 24) * SinLat * Cos3Lat * (5 - Tan2Lat + 9 * eta2)
'    IIIA = (nu / 720) * SinLat * Cos5Lat * (61 - 58 * Tan2Lat + Tan4Lat)
'    IV = nu * CosLat
'    V = (nu / 6) * Cos3Lat * (nu / rho - Tan2Lat)
'    VI = (nu / 120) * Cos5Lat * (5 - 18 * Tan2Lat + Tan4Lat + 14 * eta2 - 58 * Tan2Lat * eta2)
'
'    Dlon = lon - Lon0
'    Dlon2 = Dlon * Dlon
'    Dlon3 = Dlon2 * Dlon
'    Dlon4 = Dlon3 * Dlon
'    Dlon5 = Dlon4 * Dlon
'    Dlon6 = Dlon5 * Dlon
'
'    N = I + iI * Dlon2 + III * Dlon4 + IIIA * Dlon6
'    E = E0 + IV * Dlon + V * Dlon3 + VI * Dlon5
'
'    'return gridrefNumToLet(E, N, 8)
'    retX = E
'    retY = -N * 1.2
'
'End Function

Public Function LatLongToUTMold(ByVal LAT As Double, ByVal LON As Double, _
                                ByRef retX As Double, ByRef retY As Double)

    'www.uwgb.edu/dutchs/usefuldata/utmformulas.htm

    Const Knu     As Double = 0.9996
    Const A       As Double = 6378137
    Const B       As Double = 6356752.3142

    Dim LonDEG    As Double
    Dim E         As Double
    Dim E2        As Double
    Dim N         As Double
    Dim Zone      As Double

    Dim P         As Double
    Dim rho       As Double
    Dim SINlat    As Double
    Dim COSlat    As Double
    Dim nu        As Double

    Dim A0        As Double
    Dim B0        As Double
    Dim C0        As Double
    Dim D0        As Double
    Dim E0        As Double
    Dim S         As Double

    Dim N2        As Double
    Dim N3        As Double
    Dim N4        As Double
    Dim N5        As Double

    Dim K1        As Double
    Dim K2        As Double
    Dim K3        As Double
    Dim K4        As Double
    Dim K5        As Double

    LonDEG = LON

    LAT = toRad(LAT)
    LON = toRad(LON)


    'E = Sqr(1 - (B * B) / (A * A))
    E = 8.18191909289064E-02
    E2 = E * E / (1 - E * E)
    E2 = 6.73949675658691E-03

    '    N = (A - B) / (A + B)
    N = 1.67922038993736E-03


    '    Zone = (31 + LonDEG / 6) \ 1
    Zone = Int(31 + LonDEG / 6)

    P = toRad(LonDEG - (6 * Zone - 183))


    SINlat = Sin(LAT)
    COSlat = Cos(LAT)

    rho = A * (1 - E * E) / (1 - E * E * SINlat * SINlat) ^ 1.5
    nu = A / (1 - E * E * SINlat * SINlat) ^ 0.5


    N2 = N * N
    N3 = N2 * N
    N4 = N3 * N
    N5 = N4 * N
    A0 = A * (1 - N + (5 / 4) * (N2 - N3) + (81 / 64) * (N4 - N5))
    B0 = (1.5 * A * N) * (1 - N + (7 / 8) * (N2 - N3) + (55 / 64) * (N4 - N5))
    C0 = (15 * A * N2 / 16) * (1 - N + (3 / 4) * (N2 - N3))
    D0 = (35 * A * N3 / 48) * (1 - N + (11 / 16) * (N2 - N3))
    E0 = (315 * A * N4 / 512) * (1 - N)


    S = A0 * LAT - B0 * Sin(LAT * 2) + C0 * Sin(LAT * 4) - D0 * Sin(LAT * 6) + E0 * Sin(LAT * 8)

    K1 = S * Knu
    K2 = Knu * nu * SINlat * COSlat * 0.5
    K3 = (Knu * nu * SINlat * (COSlat ^ 3) / 24) * _
         (5 - Tan(LAT) ^ 2 + 9 * E2 * COSlat ^ 2 + _
          4 * E2 * E2 * COSlat ^ 4)


    K4 = Knu * nu * COSlat
    K5 = (Knu * nu * COSlat * COSlat * COSlat / 6) * (1 - Tan(LAT) ^ 2 + E2 * COSlat * COSlat)

    retY = -(K1 + K2 * P * P + K3 * P * P * P * P)

    retX = 500000 + K4 * P + K5 * P * P * P

    ' manual adjustment.... Remove it
    retX = retX * 0.91

End Function


Public Function LatLongToUTM(ByVal LAT As Double, ByVal LON As Double, _
                             ByRef retX As Double, ByRef retY As Double)

    'https://github.com/bryanibit/LonLat2UTM/blob/master/utm.cpp

    Const aa      As Double = 6378137
    Const eccSquared As Double = 0.00669438
    Const k0      As Double = 0.9996

    Dim LongOrigin As Double
    Dim eccPrimeSquared As Double
    Dim N#, T#, C#, A#, M#
    Dim Si#
    Dim Co#


    Dim LongTemp#, LatRad#, LongRad#, LongOriginRad#

    Dim ZoneNumber&


    '    //Make sure the longitude is between -180.00 .. 179.9
    LongTemp = (LON + 180) - Int((LON + 180) * 2.77777777777778E-03) * 360 - 180    '; // -180.00 .. 179.9;

    LatRad = LAT * deg2rad
    LongRad = LongTemp * deg2rad

    ZoneNumber = Int((LongTemp + 180) * 0.166666666666667) + 1


    '    If (LAT >= 56# And LAT < 64# And LongTemp >= 3# And LongTemp < 12#) Then
    '    ZoneNumber = 32
    '    Else
    '        Stop
    '        '    // Special zones for Svalbard
    '        '    if (Lat >= 72.0 && Lat < 84.0)
    '        '    {
    '        '        if (LongTemp >= 0.0  && LongTemp <  9.0) ZoneNumber = 31;
    '        '        else if (LongTemp >= 9.0  && LongTemp < 21.0) ZoneNumber = 33;
    '        '        else if (LongTemp >= 21.0 && LongTemp < 33.0) ZoneNumber = 35;
    '        '        else if (LongTemp >= 33.0 && LongTemp < 42.0) ZoneNumber = 37;
    '        '    }
    '    End If

    LongOrigin = (ZoneNumber - 1) * 6 - 180 + 3    '  //+3 puts origin in middle of zone
    LongOriginRad = LongOrigin * deg2rad

    '//compute the UTM Zone from the latitude and longitude
    'sprintf(UTMZone, "%d%c", ZoneNumber, UTMLetterDesignator(Lat));

    eccPrimeSquared = (eccSquared) / (1 - eccSquared)
    Si = Sin(LatRad)
    Co = Cos(LatRad)
    T = Tan(LatRad)
    N = aa / Sqr(1 - eccSquared * Si * Si)
    T = T * T
    C = eccPrimeSquared * Co * Co
    A = Co * (LongRad - LongOriginRad)

    M = aa * ((1 - eccSquared * 0.25 - 3 * eccSquared * eccSquared * 0.015625 - 5 * eccSquared * eccSquared * eccSquared * 0.00390625) * LatRad _
              - (3 * eccSquared * 0.125 + 3 * eccSquared * eccSquared * 0.03125 + 45 * eccSquared * eccSquared * eccSquared * 0.0009765625) * Sin(2 * LatRad) _
              + (15 * eccSquared * eccSquared * 0.00390625 + 45 * eccSquared * eccSquared * eccSquared * 0.0009765625) * Sin(4 * LatRad) _
              - (35 * eccSquared * eccSquared * eccSquared * 3.25520833333333E-04) * Sin(6 * LatRad))

    retX = (k0 * N * (A + (1 - T + C) * A * A * A * 0.166666666666667 _
                      + (5 - 18 * T + T * T + 72 * C - 58 * eccPrimeSquared) * A * A * A * A * A * 8.33333333333333E-03) _
                      + 500000#)

    retY = (k0 * (M + N * Tan(LatRad) * (A * A * 0.5 + (5 - T + 9 * C + 4 * C * C) * A * A * A * A * 4.16666666666667E-02 _
                                         + (61 - 58 * T + T * T + 600 * C - 330 * eccPrimeSquared) * A * A * A * A * A * A * 1.38888888888889E-03)))
    If (LAT < 0) Then retY = retY + 10000000#    ' //10000000 meter offset for southern hemisphere

    retY = -retY


    retX = Round(retX, 2)
    retY = Round(retY, 2)



End Function

