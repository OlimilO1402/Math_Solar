VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SolarCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_GeoPos  As GeoPos
Private m_TZBias  As Double
Private m_Date    As Date
Private m_Time    As Date
Private m_JulDay  As Double
Private m_JulCent As Double
Private m_GeomMeanLongSun   As Angle 'Double
Private m_GeomMeanAnomSun   As Angle 'Double
Private m_EccentEarthOrbit  As Double
Private m_SunEqOfCtr        As Double
Private m_SunTrueLong       As Angle
Private m_SunTrueAnom       As Angle
Private m_SunRadVector      As Angle
Private m_SunAppLong        As Angle
Private m_MeanObliqEcliptic As Angle
Private m_ObliqCorr         As Angle
Private m_SunRtAscending    As Angle
Private m_SunDeclination    As Angle
Private m_VarY              As Double
Private m_EqOfTime          As Double
Private m_HourAngleSunrise  As Angle
Private m_SolarNoon         As Date
Private m_SunriseTime       As Date
Private m_SunsetTime        As Date
Private m_SunlightDuration  As Double
Private m_TrueSolarTime     As Double
Private m_HourAngle         As Angle
Private m_SolarZenithAngle             As Angle
Private m_SolarElevationAngle          As Angle
Private m_ApproxAtmosphericRefraction  As Angle
Private m_SolarElevCorrectedForAtmRefr As Angle
Private m_SolarAzimuthAngle            As Angle

Friend Sub New_(aGPS As GeoPos, ByVal aTZBias As Double, ByVal aDate As Date)
    Set m_GeoPos = aGPS: m_TZBias = aTZBias: m_Date = DateSerial(Year(aDate), Month(aDate), Day(aDate))
    m_Time = TimeSerial(Hour(aDate), Minute(aDate), Second(aDate))
    Calc m_Time
End Sub
'copy to MNew:
'Public Function SolarCalc(aGPS As GeoPos, ByVal aTZBias As Double, ByVal aDate As Date) As SolarCalc
'    Set SolarCalc = New SolarCalc: SolarCalc.New_ aGPS, aTZBias, aDate
'End Function

Private Sub Calc(aTime As Date) 'aDate As Date)
    m_Time = aTime
    Dim trig As New Angle
    Dim A As Double
    '$B$5 = Latitude
    '$B$6 = Longitude
    '$B$7 = TimeZone+(to East)
    '$B$9 = Date
    
    'D = Date
    'E = Time
    
    'F = JulianDay
    'F = D4 + 2415018.5 + E4 - $B$7 / 24
    m_JulDay = MTime.Date_JulianDay(m_Date)
    
    'Dim dt As Date: dt = m_Date '+ aTime
    'G = JulianCentury
    m_JulCent = MTime.Date_JulianCentury(m_Date + m_Time)
    'Debug.Print m_JulCent
    
    'I = REST(280.46646 + JulianCentury * (36000.76983 + JulianCentury * 0.0003032) mod 360#)
    Set m_GeomMeanLongSun = MNew.AngleD((280.46646 + m_JulCent * (36000.76983 + m_JulCent * 0.0003032)) Mod 360#)
    
    'J = =357.52911 + JulianCentury * (35999.05029 - 0.0001537 * JulianCentury)
    Set m_GeomMeanAnomSun = MNew.AngleD(357.52911 + m_JulCent * (35999.05029 - 0.0001537 * m_JulCent))
    
    'K = = 0.016708634 - m_JulCent * (0.000042037 + 0.0000001267 * m_JulCent)
    m_EccentEarthOrbit = 0.016708634 - m_JulCent * (0.000042037 + 0.0000001267 * m_JulCent)
    
    'L = SIN(BOGENMASS(J4)) * (1,914602 - G4 * (0,004817 + 0,000014 * G4)) + SIN(BOGENMASS(2 * J4)) * (0,019993 - 0,000101 * G4) + SIN(BOGENMASS(3 * J4)) * 0,000289
    Dim GeomMeanAnomSun2 As Angle: Set GeomMeanAnomSun2 = m_GeomMeanAnomSun.Mul(2)
    Dim GeomMeanAnomSun3 As Angle: Set GeomMeanAnomSun3 = m_GeomMeanAnomSun.Mul(3)
    m_SunEqOfCtr = m_GeomMeanAnomSun.Sinus * (1.914602 - m_JulCent * (0.004817 + 0.000014 * m_JulCent)) + GeomMeanAnomSun2.Sinus * (0.019993 - 0.000101 * m_JulCent) + GeomMeanAnomSun3.Sinus * 0.000289
    
    'M = I4 + L4
    Set m_SunTrueLong = m_GeomMeanLongSun.Add(MNew.AngleD(m_SunEqOfCtr))
    
    'N = J4 + L4
    Set m_SunTrueAnom = m_GeomMeanAnomSun.Add(MNew.AngleD(m_SunEqOfCtr))
    
    'O = (1.000001018 * (1 - K4 * K4)) / (1 + K4 * COS(BOGENMASS(N4)))
    Set m_SunRadVector = MNew.AngleD((1.000001018 * (1 - m_EccentEarthOrbit * m_EccentEarthOrbit)) / (1 + m_EccentEarthOrbit * m_SunTrueAnom.Cosinus))
    
    'P = M4 - 0.00569 - 0.00478 * SIN(BOGENMASS(125.04 - 1934.136 * G4))
    Set m_SunAppLong = MNew.AngleD(m_SunTrueLong.GradF - 0.00569 - 0.00478 * MNew.AngleD(125.04 - 1934.136 * m_JulCent).Sinus)
    
    'Q = 23 +(26 + ((21.448 - G4 * (46.815 + G4 * (0.00059 - G4 * 0.001813)))) / 60) / 60
    
    'a = 23# + (26# + ((21.448 - m_JulCent * (46.815 + m_JulCent * (0.00059 - m_JulCent * 0.001813)))) / 60#) / 60#
    A = 0.00059 - m_JulCent * 0.001813
    A = 46.815 + m_JulCent * A
    A = 21.448 - m_JulCent * A
    A = 23# + (26# + A / 60#) / 60#
    Set m_MeanObliqEcliptic = MNew.AngleD(A)
    
    'R =23,43841609
    Set m_ObliqCorr = MNew.AngleD(23.43841609)
    
    'S = GRAD(ARCTAN2(COS(BOGENMASS(P4));COS(BOGENMASS(R4))*SIN(BOGENMASS(P4))))
    Set m_SunRtAscending = MNew.Angle(trig.ArcusTangensXYF(m_SunAppLong.Cosinus, m_ObliqCorr.Cosinus * m_SunAppLong.Sinus))
    
    'T = GRAD(ARCSIN(SIN(BOGENMASS(R4))*SIN(BOGENMASS(P4))))
    Set m_SunDeclination = trig.ArcusSinus(m_ObliqCorr.Sinus * m_SunAppLong.Sinus)
    
    'U ==TAN(BOGENMASS(R4/2))*TAN(BOGENMASS(R4/2))
    m_VarY = m_ObliqCorr.Div(2).Tangens * m_ObliqCorr.Div(2).Tangens
        
    'V         = 4 * GRAD(U4 *     SIN(2 * BOGENMASS(I4))         - 2 * K4                 * SIN(BOGENMASS(J4))      + 4 * K4                 * U4     * SIN(BOGENMASS(J4))      * COS(2 * BOGENMASS(I4))           - 0,5 * U4     * U4     * SIN(4 * BOGENMASS(I4))         - 1,25 * K4                 * K4                 * SIN(2 * BOGENMASS(J4)))
    m_EqOfTime = 4 * MNew.Angle(m_VarY * m_GeomMeanLongSun.Mul(2).Sinus - 2 * m_EccentEarthOrbit * m_GeomMeanAnomSun.Sinus + 4 * m_EccentEarthOrbit * m_VarY * m_GeomMeanAnomSun.Sinus * m_GeomMeanLongSun.Mul(2).Cosinus - 0.5 * m_VarY * m_VarY * m_GeomMeanLongSun.Mul(4).Sinus - 1.25 * m_EccentEarthOrbit * m_EccentEarthOrbit * m_GeomMeanAnomSun.Mul(2).Sinus).GradF
    
    'I4: m_GeomMeanLongSun;    'J4: m_GeomMeanAnomSun;    'K4: m_EccentEarthOrbit;    'U4: m_VarY
    
    
    'W = GRAD(ARCCOS(COS(BOGENMASS(90,833))/(COS(BOGENMASS($B$5))*COS(BOGENMASS(T4)))-TAN(BOGENMASS($B$5))*TAN(BOGENMASS(T4))))
    Set m_HourAngleSunrise = trig.ArcusCosinus(MNew.AngleD(90.833).Cosinus / (m_GeoPos.Latitude.Cosinus * m_SunDeclination.Cosinus) - m_GeoPos.Latitude.Tangens * m_SunDeclination.Tangens)
    
    
    'X          = (720  - 4  * $B$6                     - V4         + $B$7     * 60 ) / 1440
    m_SolarNoon = (720# - 4# * m_GeoPos.Longitude.GradF - m_EqOfTime + m_TZBias * 60#) / 1440#
    
    'Y            = X4          - W4                       * 4  / 1440
    m_SunriseTime = m_SolarNoon - m_HourAngleSunrise.GradF * 4# / 1440#
    
    'Z           = X4          + W4                       * 4  / 1440
    m_SunsetTime = m_SolarNoon + m_HourAngleSunrise.GradF * 4# / 1440#
    
    'AA = 8 * W4
    m_SunlightDuration = 8 * m_HourAngleSunrise.GradF
    
    'AB             = REST(E4 * 1440  + V4         + 4 * $B$6                    - 60  * $B$7    ;     1440)
    m_TrueSolarTime = (m_Time * 1440# + m_EqOfTime + 4 * m_GeoPos.Latitude.GradF - 60# * m_TZBias) Mod 1440#
    
    'AC=WENN(AB4/4<0;AB4/4+180;AB4/4-180)
    Set m_HourAngle = MNew.AngleD(IIf(m_TrueSolarTime / 4 < 0, m_TrueSolarTime / 4 + 180, m_TrueSolarTime / 4 - 180))
    
    'AD =GRAD(ARCCOS( SIN(BOGENMASS($B$5)) * SIN(BOGENMASS(T4)) + COS(BOGENMASS($B$5)) * COS(BOGENMASS(T4)) * COS(BOGENMASS(AC4)) ))
    '                                        ARCCOS( m_GeoPos.Latitude.Sinus * m_SunDeclination.sinus + m_GeoPos.Latitude.coSinus * m_SunDeclination.cosinus * m_HourAngle.hourangle)
    Set m_SolarZenithAngle = trig.ArcusCosinus(m_GeoPos.Latitude.Sinus * m_SunDeclination.Sinus + m_GeoPos.Latitude.Cosinus * m_SunDeclination.Cosinus * m_HourAngle.Cosinus)
    
    'AE=90-AD4
    Set m_SolarElevationAngle = MNew.AngleD(90).Subt(m_SolarZenithAngle)
    
    'AF = WENN(AE4 > 85; 0; WENN(AE4 > 5; 58,1 / TAN(BOGENMASS(AE4)) - 0,07 / POTENZ(TAN(BOGENMASS(AE4)); 3) + 0,000086 / POTENZ(TAN(BOGENMASS(AE4)); 5); WENN(AE4 > -0,575; 1735 + AE4 * (-518,2 + AE4 * (103,4 + AE4 * (-12,79 + AE4 * 0,711))); -20,772 / TAN(BOGENMASS(AE4))))) / 3600
    Dim seagf  As Double:  seagf = m_SolarElevationAngle.GradF
    Dim tansea As Double: tansea = m_SolarElevationAngle.Tangens
    'Dim a As Double
    If seagf > 85 Then
        A = 0
    ElseIf seagf > 5 Then
        A = 58.1 / tansea - 0.07 / (tansea ^ 3) + 0.000086 / (tansea ^ 5)
    ElseIf seagf > -0.575 Then
        A = 1735 + seagf * (-518.2 + seagf * (103.4 + seagf * (-12.79 + seagf * 0.711)))
    Else
        A = -20.772 / tansea
    End If
    
    Set m_ApproxAtmosphericRefraction = MNew.AngleD(A / 3600)
    
    'AG=AE4+AF4
    Set m_SolarElevCorrectedForAtmRefr = m_SolarElevationAngle.Add(m_ApproxAtmosphericRefraction)
    
    'AH = WENN(AC4 > 0; REST(GRAD(ARCCOS(((SIN(BOGENMASS($B$5)) * COS(BOGENMASS(AD4))) - SIN(BOGENMASS(T4))) / (COS(BOGENMASS($B$5)) * SIN(BOGENMASS(AD4))))) + 180; 360); REST(540 - GRAD(ARCCOS(((SIN(BOGENMASS($B$5)) * COS(BOGENMASS(AD4))) - SIN(BOGENMASS(T4))) / (COS(BOGENMASS($B$5)) * SIN(BOGENMASS(AD4))))); 360))
    Dim hagf As Double: hagf = m_HourAngle.GradF
    If m_HourAngle.GradF > 0 Then
        'AD4 = m_SolarZenithAngle
        'm_SolarAzimuthAngle = REST(GRAD(ARCCOS((( SIN(BOGENMASS($B$5))  * COS(BOGENMASS(AD4)) ) - SIN(BOGENMASS(T4))) / (COS(BOGENMASS($B$5)) * SIN(BOGENMASS(AD4))))) + 180; 360)
        'a = REST(GRAD(ARCCOS(((m_GeoPos.Latitude.Sinus * m_SolarZenithAngle.coSinus) - SIN(BOGENMASS(T4))) / (m_GeoPos.Latitude.coSinus * m_SolarZenithAngle.Sinus))) + 180; 360)
        A = trig.ArcusCosinus(((m_GeoPos.Latitude.Sinus * m_SolarZenithAngle.Cosinus) - m_SunDeclination.Sinus) / (m_GeoPos.Latitude.Cosinus * m_SolarZenithAngle.Sinus)).GradF + 180
        'Set m_SolarAzimuthAngle = trig.ArcusCosinus(((m_GeoPos.Latitude.Sinus * m_SolarZenithAngle.Cosinus) - m_SunDeclination.Sinus) / (m_GeoPos.Latitude.Cosinus * m_SolarZenithAngle.Sinus)).Add(180)   '; 360)
    Else
        '$B$5 = m_GeoPos.Latitude
        
        'm_SolarAzimuthAngle = REST(540 - GRAD(ARCCOS(((SIN(BOGENMASS($B$5)) * COS(BOGENMASS(AD4))) - SIN(BOGENMASS(T4))) / (COS(BOGENMASS($B$5)) * SIN(BOGENMASS(AD4))))); 360)
        A = 540# - trig.ArcusCosinus(((m_GeoPos.Latitude.Sinus * m_SolarZenithAngle.Cosinus) - m_SunDeclination.Sinus) / (m_GeoPos.Latitude.Cosinus * m_SolarZenithAngle.Sinus)).GradF
        'Set m_SolarAzimuthAngle = MNew.AngleD(540).Subt(trig.ArcusCosinus(((m_GeoPos.Latitude.Sinus * m_SolarZenithAngle.Consinus) - m_SunDeclination.Sinus) / (m_GeoPos.Latitude.Cosinus * m_SolarZenithAngle.Sinus)))
    End If
    Set m_SolarAzimuthAngle = MNew.AngleD(A)
    
End Sub

Public Property Get GeoPos() As GeoPos
    Set GeoPos = m_GeoPos
End Property

Public Property Get TimeZoneBias() As Double
    TimeZoneBias = m_TZBias
End Property
Public Property Let TimeZoneBias(ByVal Value As Double)
    m_TZBias = Value
End Property


Public Property Get DayDate() As Date
    DayDate = m_Date
End Property
Public Property Let DayDate(ByVal Value As Date)
    m_Date = DateSerial(Year(Value), Month(Value), Day(Value))
End Property


Public Property Get DayTime() As Date
    DayTime = m_Time
End Property
Public Property Let DayTime(ByVal Value As Date)
    m_Time = TimeSerial(Hour(Value), Minute(Value), Second(Value))
End Property

Public Function ToStr() As String
    Dim s As String
    s = s & "JulianDay        : " & m_JulDay & vbCrLf
    s = s & "JulianCentury    : " & m_JulCent & vbCrLf
    s = s & "GeomMeanLongSun  : " & m_GeomMeanLongSun.GradF & vbCrLf
    s = s & "GeomMeanAnomSun  : " & m_GeomMeanAnomSun.GradF & vbCrLf
    s = s & "EccentEarthOrbit : " & m_EccentEarthOrbit & vbCrLf
    s = s & "SunEqOfCtr       : " & m_SunEqOfCtr & vbCrLf
    s = s & "SunTrueLong      : " & m_SunTrueLong.GradF & vbCrLf
    s = s & "SunTrueAnom      : " & m_SunTrueAnom.GradF & vbCrLf
    s = s & "SunRadVector     : " & m_SunRadVector.GradF & vbCrLf
    s = s & "SunAppLong       : " & m_SunAppLong.GradF & vbCrLf
    s = s & "MeanObliqEcliptic: " & m_MeanObliqEcliptic.GradF & vbCrLf
    s = s & "ObliqCorr        : " & m_ObliqCorr.GradF & vbCrLf
    s = s & "SunRtAscending   : " & m_SunRtAscending.GradF & vbCrLf
    s = s & "SunDeclination   : " & m_SunDeclination.GradF & vbCrLf
    s = s & "VarY             : " & m_VarY & vbCrLf
    s = s & "EqOfTime         : " & m_EqOfTime & vbCrLf
    s = s & "HourAngleSunrise : " & m_HourAngleSunrise.GradF & vbCrLf
    s = s & "SolarNoon        : " & m_SolarNoon & vbCrLf
    s = s & "SunriseTime      : " & m_SunriseTime & vbCrLf
    s = s & "SunsetTime       : " & m_SunsetTime & vbCrLf
    s = s & "SunlightDuration : " & m_SunlightDuration & vbCrLf
    s = s & "TrueSolarTime    : " & m_TrueSolarTime & vbCrLf
    s = s & "HourAngle        : " & m_HourAngle.GradF & vbCrLf
    s = s & "SolarZenithAngle : " & m_SolarZenithAngle.GradF & vbCrLf
    s = s & "SolarElevationAngle          : " & m_SolarElevationAngle.GradF & vbCrLf
    s = s & "ApproxAtmosphericRefraction  : " & m_ApproxAtmosphericRefraction.GradF & vbCrLf
    s = s & "SolarElevCorrectedForAtmRefr : " & m_SolarElevCorrectedForAtmRefr.GradF & vbCrLf
    s = s & "SolarAzimuthAngle            : " & m_SolarAzimuthAngle.GradF & vbCrLf
    ToStr = s
End Function


Public Property Get JulianDay() As Double
    '=D4+2415018,5+E4-$B$7/24
    JulianDay = m_JulDay
End Property
Public Property Get JulianCentury() As Double
    JulianCentury = m_JulCent
End Property


'=I
Public Function GeomMeanLongSun() As Double
    GeomMeanLongSun = m_GeomMeanLongSun
End Function

'=J
Public Function GeomMeanAnomSun(ByVal dt As Date) As Double
    GeomMeanAnomSun = m_GeomMeanAnomSun
End Function

'=K
Public Function EccentEarthOrbit(ByVal dt As Date) As Double
    EccentEarthOrbit = m_EccentEarthOrbit
End Function

'=L
Public Function SunEqOfCtr(ByVal dt As Date) As Double
    SunEqOfCtr m_SunEqOfCtr
End Function

Public Property Get SunTrueLong() As Angle
    Set SunTrueLong = m_SunTrueLong
End Property

Public Property Get SunTrueAnom() As Angle
    Set SunTrueAnom = m_SunTrueAnom
End Property
Public Property Get SunRadVector() As Angle
    Set SunRadVector = m_SunRadVector
End Property
Public Property Get SunAppLong() As Angle
    Set SunAppLong = m_SunAppLong
End Property
Public Property Get MeanObliqEcliptic() As Angle
    Set MeanObliqEcliptic = m_MeanObliqEcliptic
End Property
Public Property Get ObliqCorr() As Angle
    Set ObliqCorr = m_ObliqCorr
End Property
Public Property Get SunRtAscending() As Angle
    Set SunRtAscending = m_SunRtAscending
End Property
Public Property Get SunDeclination() As Angle
    Set SunDeclination = m_SunDeclination
End Property
Public Property Get VarY() As Double
    VarY = m_VarY
End Property
Public Property Get EqOfTime() As Double
    EqOfTime = m_EqOfTime
End Property
Public Property Get HourAngleSunrise() As Angle
    Set HourAngleSunrise = m_HourAngleSunrise
End Property
Public Property Get SolarNoon() As Date
    SolarNoon = m_SolarNoon
End Property
Public Property Get SunriseTime() As Date
    SunriseTime = m_SunriseTime
End Property
Public Property Get SunsetTime() As Date
    SunsetTime = m_SunsetTime
End Property
Public Property Get SunlightDuration() As Double
    SunlightDuration = m_SunlightDuration
End Property
Public Property Get TrueSolarTime() As Double
    TrueSolarTime = m_TrueSolarTime
End Property
Public Property Get HourAngle() As Angle
    Set HourAngle = m_HourAngle
End Property
Public Property Get SolarZenithAngle() As Angle
    Set SolarZenithAngle = m_SolarZenithAngle
End Property
Public Property Get SolarElevationAngle() As Angle
    Set SolarElevationAngle = m_SolarElevationAngle
End Property
Public Property Get ApproxAtmosphericRefraction() As Angle
    Set ApproxAtmosphericRefraction = m_ApproxAtmosphericRefraction
End Property
Public Property Get SolarElevCorrectedForAtmRefr() As Angle
    Set SolarElevCorrectedForAtmRefr = m_SolarElevCorrectedForAtmRefr
End Property
Public Property Get SolarAzimuthAngle() As Angle
    Set SolarAzimuthAngle = m_SolarAzimuthAngle
End Property

