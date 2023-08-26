VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
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
   ScaleHeight     =   9495
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtDayTime 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Text            =   "FMain.frx":0000
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox TxtDayDate 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtTZBias 
      Alignment       =   2  'Zentriert
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton BtnGetGPSCoords 
      Caption         =   "Set GPS Coords"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh:mm:ss"
      Height          =   225
      Left            =   3240
      TabIndex        =   11
      Top             =   1350
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dd.mm.yyyy"
      Height          =   225
      Left            =   3240
      TabIndex        =   10
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   225
      Left            =   3225
      TabIndex        =   9
      Top             =   630
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label LblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   405
   End
   Begin VB.Label LblTZBias 
      AutoSize        =   -1  'True
      Caption         =   "TimeZone (+ to E)"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1425
   End
   Begin VB.Label LblGPS 
      Caption         =   "GPS:"
      Height          =   465
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3360
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Es soll möglich sein:
'GPS-Position:
'* die Position einzugeben, oder
'* vom Betriebssystem auszulesen
'Zeitzone bzw Offset:
'* die Zeitzone bzw den Offset einzugeben, oder
'* vom Betriebssystem auszulesen, oder
'* im Programm zur gegebenen gps-Position ermitteln zu lassen

'https://stackoverflow.com/questions/16086962/how-to-get-a-time-zone-from-a-location-using-latitude-and-longitude-coordinates

'Query
'Get current local time at Statue of Liberty using latitude & longitude:
'http://api.timezonedb.com/v2.1/get-time-zone?key=YOUR_API_KEY&format=json&by=position&lat=40.689247&lng=-74.044502
'Response
'{
'    "status"       : "OK",                 ' Status of the API query. Either OK or FAILED.
'    "message"      : "",                   ' Error message. Empty if no error.
'    "countryCode"  : "US",                 ' Country code of the time zone.
'    "countryName"  : "United States",      ' Country name of the time zone.
'    "regionName"   : "New Jersey",         ' PREMIUM Region / State name of the time zone.
'    "cityName"     : "Jersey City",        ' PREMIUM City / Place name of the time zone.
'    "zoneName"     : "America\/New_York",  ' The time zone name.
'    "abbreviation" : "EDT",                ' Abbreviation of the time zone.
'    "gmtOffset"    : -14400,               ' The time offset in seconds based on UTC time.
'    "dst"          : "1",                  ' Whether Daylight Saving Time (DST) is used. Either 0 (No) or 1 (Yes).
'    "zoneStart"    : 1678604400,           ' The Unix time in UTC when current time zone start.
'    "zoneEnd"      : 1699164000,           ' The Unix time in UTC when current time zone end.
'    "nextAbbreviation": "EST",             '
'    "timestamp"    : 1686466062,           ' Current local time in Unix time. Minus the value with gmtOffset to get UTC time.
'    "formatted"    : "2023-06-11 06:47:42" ' Formatted timestamp in Y-m-d h:i:s format. E.g.: 2023-06-11 10:47:41
'    "totalPage"    : ""                    ' The total page of result when exceed 25 records.
'    "currentPage"  : ""                    ' Current page when navigating.
'}

Private m_Solar As SolarCalc

Private Sub BtnGetGPSCoords_Click()
    Dim gps As GeoPos: Set gps = m_Solar.GeoPos
    If FGeoPos.ShowDialog(gps, Me) = vbCancel Then Exit Sub
    UpdateView
End Sub

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_Solar = MNew.SolarCalc(MNew.GeoPos(MNew.AngleDecS("N 48.010973"), MNew.AngleDecS("E 10.617913"), 625, "86825 Bad Wörishofen, Dominikusstraße 9"), 2, Now)
    UpdateView
End Sub

Sub UpdateView()
    LblGPS.Caption = m_Solar.GeoPos.ToStr
    TxtTZBias.Text = m_Solar.TimeZoneBias
    TxtDayDate.Text = m_Solar.DayDate
    TxtDayTime.Text = m_Solar.DayTime
    Text1.Text = m_Solar.ToStr
End Sub

Private Sub TxtTZBias_LostFocus()
    Dim i As Integer, s As String: s = TxtTZBias.Text
    If Not MString.Integer_TryParse(s, i) Then
        MsgBox "Not a valid int16-value: " & s & vbCrLf & "Please give a value in hours +east"
        Exit Sub
    End If
    m_Solar.TimeZoneBias = i
    UpdateView
End Sub

Private Sub TxtDayDate_LostFocus()
    Dim d As Date, s As String: s = TxtDayDate.Text
    If Not MTime.Date_TryParse(s, d) Then
        MsgBox "Not a valid date-value: " & s & vbCrLf '& "Please give a value ..."
        Exit Sub
    End If
    m_Solar.DayDate = d
    UpdateView
End Sub

Private Sub TxtDayTime_LostFocus()
    Dim t As Date, s As String: s = TxtDayTime.Text
    If Not MTime.Time_TryParse(s, t) Then
        MsgBox "Not a valid time-value: " & s & vbCrLf '& "Please give a value ..."
        Exit Sub
    End If
    m_Solar.DayTime = t
    UpdateView
End Sub
