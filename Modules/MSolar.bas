Attribute VB_Name = "MSolar"
Option Explicit
''dt = D
''jd = F
''jc = G
'
''=I
'Public Function GeomMeanLongSunDeg(ByVal dt As Date) As Double
'    '=REST(280,46646+G2*(36000,76983 + G2*0,0003032);360)
'    Dim jc As Double: jc = MTime.Date_JulianCentury(dt)
'    GeomMeanLongSunDeg = (280.46646 + dt * (36000.76983 + dt * 0.0003032)) Mod 360#
'End Function
'
''=J
'Public Function GeomMeanAnomSunDeg(ByVal dt As Date) As Double
'    Dim jc As Double: jc = MTime.Date_JulianCentury(dt)
'    GeomMeanAnomSunDeg = 357.52911 + jc * (35999.05029 - 0.0001537 * jc)
'End Function
'
''=K
'Public Function EccentEarthOrbit(ByVal dt As Date) As Double
'    Dim jc As Double: jc = MTime.Date_JulianCentury(dt)
'    EccentEarthOrbit = 0.016708634 - jc * (0.000042037 + 0.0000001267 * jc)
'End Function
'
''=L
'Public Function SunEqOfCtr(ByVal dt As Date) As Double
'    Dim gmasd As Double: gmasd = GeomMeanAnomSunDeg(dt) 'j
''=SIN(BOGENMASS(J4))*(1,914602-G4*(0,004817+0,000014*G4))+SIN(BOGENMASS(2*J4))*(0,019993-0,000101*G4)+SIN(BOGENMASS(3*J4))*0,000289
'End Function
