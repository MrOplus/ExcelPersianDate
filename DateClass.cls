VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Initial()
    Call DateModuleSetup
End Sub

Function JToday(Optional mode As String) As String
Dim GToday As String
GToday = Right(Date$, 4) + Left(Date$, 2) + Mid$(Date$, 4, 2)
JToday = JalaliDate(GToday, mode)
End Function

Function JDiff(Jdate1 As String, JDate2 As String) As Long
Dim JYear1 As Long, JMonth1 As Long, jday1 As Long
Dim JYear2 As Long, JMonth2 As Long, JDay2 As Long
Dim TotalDays1 As Long
Dim TotalDays2 As Long
JYear1 = Year_(Jdate1)
If Len(Jdate1) = 6 Then JYear1 = JYear1 + 1300
JMonth1 = Month_(Jdate1)
jday1 = Day_(Jdate1)
JYear2 = Year_(JDate2)
If Len(JDate2) = 6 Then JYear2 = JYear2 + 1300
JMonth2 = Month_(JDate2)
JDay2 = Day_(JDate2)
TotalDays1 = JalaliDays(JYear1, JMonth1, jday1)
TotalDays2 = JalaliDays(JYear2, JMonth2, JDay2)
JDiff = TotalDays2 - TotalDays1
End Function


Function JSubDay(Jdate As String, JDayOff As Integer, Optional mode As String) As String
Dim jyear As Long, jmonth As Long, jday As Long
Dim TotalDays As Long
jyear = Year_(Jdate)
If Len(Jdate) = 6 Then jyear = jyear + 1300
jmonth = Month_(Jdate)
jday = Day_(Jdate)
TotalDays = JalaliDays(jyear, jmonth, jday) - JDayOff + GYearOff
JalaliYMD TotalDays, jyear, jmonth, jday
JSubDay = YMD2Str(jyear, jmonth, jday, mode)
End Function


Function JWeekDay(Jdate As String, Optional mode As String) As Variant
Dim JWeekDays As Variant
Dim jyear As Long, jmonth As Long, jday As Long
Dim tmp As Long
If mode = "" Then
    mode = "SHORT"
Else
    mode = UCase(mode)
End If
JWeekDays = Array(ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1740) & ChrW(1705) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1583) & ChrW(1608) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1587) & ChrW(1607) & " " & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(1588) & ChrW(1606) & ChrW(1576) & ChrW(1607) _
, ChrW(1580) & ChrW(1605) & ChrW(1593) & ChrW(1607))


jyear = Year_(Jdate)
If Len(Jdate) = 6 Then jyear = jyear + 1300
jmonth = Month_(Jdate)
jday = Day_(Jdate)
tmp = (JalaliDays(jyear, jmonth, jday) - JWkDayOff) Mod 7
Select Case mode
Case "LONG"
    JWeekDay = JWeekDays(tmp)
Case Else
    JWeekDay = tmp
End Select
End Function

Function NormDate(Date_ As String) As String
Dim Y As String, M As String, d As String
Dim Slash1Pos As Integer
Dim Slash2Pos As Integer
Date_ = LTrim(RTrim(Date_))
Slash1Pos = InStr(Date_, "/")
Slash2Pos = InStr(Slash1Pos + 1, Date_, "/")
If Slash1Pos = 0 Or Slash2Pos = 0 Then
    NormDate = "#Error"
    Exit Function
End If
Y = Left(Date_, Slash1Pos - 1)
M = Mid(Date_, Slash1Pos + 1, Slash2Pos - Slash1Pos - 1)
d = Mid(Date_, Slash2Pos + 1)
Y = LTrim(RTrim(Y))
M = LTrim(RTrim(M))
d = LTrim(RTrim(d))
M = IIf(Len(M) < 2, "0" + M, M)
d = IIf(Len(d) < 2, "0" + d, d)
NormDate = Y + M + d
End Function

Function JAddDay(Jdate As String, JDayOff As Integer, Optional mode As String) As String
Dim jyear As Long, jmonth As Long, jday As Long
Dim TotalDays As Long
jyear = Year_(Jdate)
If Len(Jdate) = 6 Then jyear = jyear + 1300
jmonth = Month_(Jdate)
jday = Day_(Jdate)
TotalDays = JalaliDays(jyear, jmonth, jday) + JDayOff + GYearOff
JalaliYMD TotalDays, jyear, jmonth, jday
JAddDay = YMD2Str(jyear, jmonth, jday, mode)
End Function

Function JalalDate(GDate As String, Optional mode As String) As String
JalalDate = JalaliDate(GDate, mode)
End Function



