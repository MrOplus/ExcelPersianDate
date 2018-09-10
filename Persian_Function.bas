Attribute VB_Name = "Persian_Function"
Option Base 0
Global AlphaNumeric1(0 To 19) As String
Global AlphaNumeric2(1 To 9) As String
Global AlphaNumeric3(1 To 9) As String

Public GDayTab(2, 13) As Long
Public JDayTab(2, 13) As Long
Public Const GYearOff = 226894
Public Const Solar = 365.25
Public Const JWkDayOff = 3
Public Const GWkDayOff = 0

Function AbH(Number As String) As String



Dim IsNegative As String
Dim DotPosition As Integer
Dim IntegerSegment As String
Dim DecimalSegment As String
Dim DotTxt, DecimalTxt As String

If val(Number) >= 0 Then IsNegative = "" Else IsNegative = ChrW(1605) & ChrW(1606) & ChrW(1601) & ChrW(1740) & " "
DotPosition = InStr(1, Number, ".")

If Not (DotPosition) = 0 Then
    IntegerSegment = Left(Abs(Number), DotPosition - 1)
    DecimalSegment = Left(Right(Number, Len(Number) - DotPosition), 5)
    
If val(IntegerSegment) <> 0 Then DotTxt = _
    " " & ChrW(1605) & ChrW(1605) & ChrW(1740) & ChrW(1586) & " " _
Else DotTxt = ""

    Select Case Len(DecimalSegment)
    
        Case 1
            DecimalTxt = " " & ChrW(1583) & ChrW(1607) & ChrW(1605)
        Case 2
            DecimalTxt = " " & ChrW(1589) & ChrW(1583) & ChrW(1605)
        Case 3
            DecimalTxt = " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
        Case 4
            DecimalTxt = " " & ChrW(1583) & ChrW(1607) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
        Case 5
            DecimalTxt = " " & ChrW(1589) & ChrW(1583) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
            
    End Select

    If Number < 1 And Number > -1 Then
        AbH = IsNegative & DotTxt & Horof(DecimalSegment) & DecimalTxt
    Else
        AbH = IsNegative & Horof(IntegerSegment) & DotTxt & Horof(DecimalSegment) & DecimalTxt
    End If
    
    Exit Function

End If
    
    
    
AbH = WorksheetFunction.Trim(IsNegative & Horof(Abs(Number)))



End Function

Sub alphaset()
   Dim i%
   AlphaNumeric1(0) = ChrW(1589) & ChrW(1601) & ChrW(1585)
   AlphaNumeric1(1) = ChrW(1740) & ChrW(1705)
   AlphaNumeric1(2) = ChrW(1583) & ChrW(1608)
   AlphaNumeric1(3) = ChrW(1587) & ChrW(1607)
   AlphaNumeric1(4) = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585)
   AlphaNumeric1(5) = ChrW(1662) & ChrW(1606) & ChrW(1580)
   AlphaNumeric1(6) = ChrW(1588) & ChrW(1588)
   AlphaNumeric1(7) = ChrW(1607) & ChrW(1601) & ChrW(1578)
   AlphaNumeric1(8) = ChrW(1607) & ChrW(1588) & ChrW(1578)
   AlphaNumeric1(9) = ChrW(1606) & ChrW(1607)
   AlphaNumeric1(10) = ChrW(1583) & ChrW(1607)
   AlphaNumeric1(11) = ChrW(1740) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(12) = ChrW(1583) & ChrW(1608) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(13) = ChrW(1587) & ChrW(1740) & ChrW(1586) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(14) = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(15) = ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(16) = ChrW(1588) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(17) = ChrW(1607) & ChrW(1601) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(18) = ChrW(1607) & ChrW(1740) & ChrW(1580) & ChrW(1583) & ChrW(1607)
   AlphaNumeric1(19) = ChrW(1606) & ChrW(1608) & ChrW(1586) & ChrW(1583) & ChrW(1607)

   
   
   AlphaNumeric2(1) = ChrW(1583) & ChrW(1607)
   AlphaNumeric2(2) = ChrW(1576) & ChrW(1740) & ChrW(1587) & ChrW(1578)
   AlphaNumeric2(3) = ChrW(1587) & ChrW(1740)
   AlphaNumeric2(4) = ChrW(1670) & ChrW(1607) & ChrW(1604)
   AlphaNumeric2(5) = ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(1575) & ChrW(1607)
   AlphaNumeric2(6) = ChrW(1588) & ChrW(1589) & ChrW(1578)
   AlphaNumeric2(7) = ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1575) & ChrW(1583)
   AlphaNumeric2(8) = ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1575) & ChrW(1583)
   AlphaNumeric2(9) = ChrW(1606) & ChrW(1608) & ChrW(1583)
   
   AlphaNumeric3(1) = ChrW(1740) & ChrW(1705) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(2) = ChrW(1583) & ChrW(1608) & ChrW(1740) & ChrW(1587) & ChrW(1578)
   AlphaNumeric3(3) = ChrW(1587) & ChrW(1740) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(4) = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(5) = ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(6) = ChrW(1588) & ChrW(1588) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(7) = ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(8) = ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1589) & ChrW(1583)
   AlphaNumeric3(9) = ChrW(1606) & ChrW(1607) & ChrW(1589) & ChrW(1583)
    
   
End Sub


Function Horof(Number As String) As String

   alphaset
   
    Dim No As Currency, N As String
    
    On Error GoTo Horoferror
    
    No = CCur(Number)
    N = CStr(No)
    
    Select Case Len(N)
        Case 1 To 3:
                If N < 20 Then
                    Horof = AlphaNumeric1(N)
                ElseIf N < 100 Then
                    If N Mod 10 = 0 Then
                        Horof = AlphaNumeric2(N \ 10)
                    Else
                        Horof = AlphaNumeric2(N \ 10) & " " & ChrW(1608) & " " & Horof(N Mod 10)
                    End If
                ElseIf N < 1000 Then
                    If N Mod 100 = 0 Then
                        Horof = AlphaNumeric3(N \ 100)
                    Else
                        Horof = AlphaNumeric3(N \ 100) & " " & ChrW(1608) & " " & Horof(N Mod 100)
                    End If
                        
                End If
        Case 4 To 6:
                If (Right(N, 3)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 3)) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & " "
                Else
                    Horof = Horof(Left(N, Len(N) - 3)) & " " & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & " " & ChrW(1608) & " " & Horof(Right(N, 3))
                End If
        Case 7 To 9:
                If (Right(N, 6)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 6)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & " "
                Else
                    Horof = Horof(Left(N, Len(N) - 6)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & " " & ChrW(1608) & " " & Horof(Right(N, 6))
                End If
        Case Else:
                If (Right(N, 9)) = 0 Then
                   Horof = Horof(Left(N, Len(N) - 9)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & " "
                Else
                    Horof = Horof(Left(N, Len(N) - 9)) & " " & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & " " & ChrW(1608) & " " & Horof(Right(N, 9))
                End If
            
    End Select
    
    Exit Function
Horoferror:
    Horof = "#Error"
End Function

Function TbH(Jdate As String, Optional mode As Integer)

On Error GoTo ErrorHandler
    
Dim txtYear, txtMonth, txtDay, S As String
Dim StandardDate, StandardYear, StandardMonth, StandardDay As String
Dim DayofWeek As String
Dim x20 As New DateClass

    x20.Initial
 
 If Left(Jdate, 2) = "13" Then StandardDate = x20.NormDate(Jdate) Else StandardDate = x20.NormDate(13 & Jdate)
 
StandardYear = Left(StandardDate, 4)
StandardMonth = Mid(StandardDate, 5, 2)
StandardDay = Right(StandardDate, 2)

Select Case val(StandardMonth)
    Case 1
       S = ChrW(1601) & ChrW(1585) & ChrW(1608) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1606)
    Case 2
       S = ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1740) & ChrW(1576) & ChrW(1607) & ChrW(1588) & ChrW(1578)
    Case 3
       S = ChrW(1582) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
    Case 4
       S = ChrW(1578) & ChrW(1740) & ChrW(1585)
    Case 5
       S = ChrW(1605) & ChrW(1585) & ChrW(1583) & ChrW(1575) & ChrW(1583)
    Case 6
       S = ChrW(1588) & ChrW(1607) & ChrW(1585) & ChrW(1740) & ChrW(1608) & ChrW(1585)
    Case 7
       S = ChrW(1605) & ChrW(1607) & ChrW(1585)
    Case 8
       S = ChrW(1570) & ChrW(1576) & ChrW(1575) & ChrW(1606)
    Case 9
       S = ChrW(1570) & ChrW(1584) & ChrW(1585)
    Case 10
       S = ChrW(1583) & ChrW(1740)
    Case 11
       S = ChrW(1576) & ChrW(1607) & ChrW(1605) & ChrW(1606)
    Case 12
       S = ChrW(1575) & ChrW(1587) & ChrW(1601) & ChrW(1606) & ChrW(1583)
End Select



txtYear = Horof(val(StandardYear))
txtMonth = S

Select Case val(StandardDay)
    Case 3:     txtDay = ChrW(1587) & ChrW(1608) & ChrW(1605)
    Case 23:    txtDay = ChrW(1576) & ChrW(1740) & ChrW(1587) & ChrW(1578) & " " & ChrW(1608) & " " & ChrW(1587) & ChrW(1608) & ChrW(1605)
    Case 30:    txtDay = ChrW(1587) & ChrW(1740) & " " & ChrW(1575) & ChrW(1605)
    Case Else:  txtDay = Horof(val(StandardDay)) & ChrW(1605)
End Select


DayofWeek = J_WEEKDAY(Jdate, 1)


 Select Case mode
    Case 0
            TbH = val(StandardDay) & " " & txtMonth & " " & StandardYear
    Case 1
            TbH = txtDay & " " & txtMonth & " " & txtYear
    Case 2
            TbH = DayofWeek & "¡ " & txtDay & " " & txtMonth & " " & txtYear
End Select
Exit Function

ErrorHandler:
TbH = CVErr(xlErrNum)
End Function



Sub InsertJalaliDate()

Dim Obj1 As New DateClass
Obj1.Initial
ActiveCell.Value = FDate(Obj1.JToday("long"))

End Sub

Function J_TODAY(Optional mode As Integer)

    Application.Volatile True
    Dim x1 As New DateClass
    x1.Initial
    
    If mode = 1 Then
        
        J_TODAY = FDate(x1.JToday("long"))
    Else
        J_TODAY = FDate(x1.JToday("Short"))
    End If

End Function


Function J_WEEKDAY(Jdate As String, Optional mode As Integer)

    Dim x2 As New DateClass
    x2.Initial
    Dim Temp$
    
    If mode = 1 Then Temp$ = "long" Else Temp$ = "short"
        J_WEEKDAY = x2.JWeekDay(x2.NormDate(Jdate), Temp$)
        
    End Function

Function J_NORMDATE(Jdate As String)

    Dim x3 As New DateClass
    x3.Initial
    J_NORMDATE = x3.NormDate(Jdate)
    
End Function


Function J_ADDDAY(Jdate As String, Number As Integer, Optional mode As Integer)

    Dim x4 As New DateClass
    x4.Initial
    
    If mode = 1 Then
        J_ADDDAY = FDate(x4.JAddDay(x4.NormDate(Jdate), Number, "long"))
    Else
        J_ADDDAY = FDate(x4.JAddDay(x4.NormDate(Jdate), Number, "short"))
    End If
    
End Function
Function PadDigits(val, digits)
  PadDigits = Right(String(digits, "0") & val, digits)
End Function
Function J_ADDMONTH(Jdate As String, Number As Integer)
    Dim Year As String, Month As String, day As String
    If Len(Jdate) < 8 Then
        day = Right(Jdate, 2)
        Month = Mid(Jdate, 3, 2)
        Year = Left(Jdate, 2)
    Else
        day = Right(Jdate, 2)
        Month = Mid(Jdate, 6, 2)
        Year = Left(Jdate, 4)
    End If
    Dim nm As Integer
    Dim nd As Integer
    Dim ny As Integer
    nm = CInt(Month)
    nd = CInt(day)
    ny = CInt(Year)
    
    For i = 1 To Number
        If nm = 12 Then
            nm = 1
            ny = ny + 1
        Else
            nm = nm + 1
        End If
    Next
    Dim finalDate As String
    finalDate = Str(ny) + "/" + PadDigits(nm, 2) + "/" + PadDigits(nd, 2)
    J_ADDMONTH = Trim(finalDate)
End Function
Function J_DIFF(Jdate1 As String, JDate2 As String)

    Dim x5 As New DateClass
    x5.Initial
    J_DIFF = x5.JDiff(x5.NormDate(Jdate1), x5.NormDate(JDate2))
    
End Function


Function J_JALALDATE(MDate As String, Optional mode As Integer)

    Dim x6 As New DateClass
    x6.Initial
    
    If IsDate(MDate) Then
        MDate = Year(MDate) & "/" & Month(MDate) & "/" & day(MDate)
    End If
    
    If mode = 1 Then
        
        J_JALALDATE = FDate(x6.JalalDate(x6.NormDate(MDate), "long"))
    Else
        J_JALALDATE = FDate(x6.JalalDate(x6.NormDate(MDate), "Short"))
        
    End If
    
End Function

Function J_SUBDAY(Jdate As String, Number As Integer, Optional mode As Integer)

    Dim x7 As New DateClass
    x7.Initial
    
    If mode = 1 Then
      
        J_SUBDAY = FDate(x7.JSubDay(x7.NormDate(Jdate), Number, "long"))
    Else
        J_SUBDAY = FDate(x7.JSubDay(x7.NormDate(Jdate), Number, "Short"))
    End If
    
End Function


Function J_GREGORIANDATE(MDate As String, Optional mode As Integer)

    Dim x6 As New DateClass
    x6.Initial
    If mode = 1 Then
        
        J_GREGORIANDATE = FDate(GregorianDate(x6.NormDate(MDate), "long"))
    Else
        J_GREGORIANDATE = FDate(GregorianDate(x6.NormDate(MDate), "short"))
    End If
    
End Function





Function CurrencyEn(ByVal MyNumber)
Dim Temp
         Dim Dollars, Cents
         Dim DecimalPlace, Count

         ReDim Place(9) As String
         Place(2) = " Thousand "
         Place(3) = " Million "
         Place(4) = " Billion "
         Place(5) = " Trillion "
         MyNumber = Trim(Str(MyNumber))
         DecimalPlace = InStr(MyNumber, ".")

         If DecimalPlace > 0 Then
              Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Cents = ConvertTens(Temp)

              MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
         End If

         Count = 1
         Do While MyNumber <> ""
                      Temp = ConvertHundreds(Right(MyNumber, 3))
            If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
            If Len(MyNumber) > 3 Then
                 MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
               MyNumber = ""
            End If
            Count = Count + 1
         Loop
         Select Case Dollars
            Case ""
               Dollars = "No Dollars"
            Case "One"
               Dollars = "One Dollar"
            Case Else
               Dollars = Dollars & " Dollars"
         End Select

         Select Case Cents
            Case ""
               Cents = " And No Cents"
            Case "One"
               Cents = " And One Cent"
            Case Else
               Cents = " And " & Cents & " Cents"
         End Select

         CurrencyEn = Dollars & Cents
End Function



Private Function ConvertHundreds(ByVal MyNumber)
Dim Result As String
         If val(MyNumber) = 0 Then Exit Function

         MyNumber = Right("000" & MyNumber, 3)
         If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
         End If

          If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
         Else
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
         End If

         ConvertHundreds = Trim(Result)
End Function



Private Function ConvertTens(ByVal MyTens)
Dim Result As String

         If val(Left(MyTens, 1)) = 1 Then
            Select Case val(MyTens)
               Case 10: Result = "Ten"
               Case 11: Result = "Eleven"
               Case 12: Result = "Twelve"
               Case 13: Result = "Thirteen"
               Case 14: Result = "Fourteen"
               Case 15: Result = "Fifteen"
               Case 16: Result = "Sixteen"
               Case 17: Result = "Seventeen"
               Case 18: Result = "Eighteen"
               Case 19: Result = "Nineteen"
               Case Else
            End Select
         Else
            Select Case val(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select

            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If

         ConvertTens = Result
End Function



Private Function ConvertDigit(ByVal MyDigit)
Select Case val(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select
End Function



Function Jleap(Year As Long) As Long

    Dim tmp As Long
    tmp = Year Mod 33
    If (tmp = 1 Or tmp = 5 Or tmp = 9 Or tmp = 13 Or tmp = 17 Or tmp = 22 Or tmp = 26 Or tmp = 30) Then
        Jleap = 1
    Else
        Jleap = 0
    End If

End Function

Function JDayOfYear(Year As Long, Month As Long, day As Long) As Long
    Dim i As Long, leap As Long
    leap = Jleap(Year)
    For i = 1 To Month - 1
        day = day + JDayTab(leap, i)
    Next i
    JDayOfYear = day
End Function



Function JLeapYears(jyear As Long) As Long

Dim leap As Long, CurrentCycle As Long, Div33 As Long, i As Long
Div33 = Int(jyear / 33)
CurrentCycle = jyear - (Div33 * 33)
leap = Div33 * 8
If CurrentCycle > 0 Then
    i = 1
    Do While i <= CurrentCycle And i <= 18
        leap = leap + 1
        i = i + 4
    Loop
End If
If CurrentCycle > 21 Then
    i = 22
    Do While i <= CurrentCycle And i <= 30
        leap = leap + 1
        i = i + 4
    Loop
End If
JLeapYears = leap

End Function


Function JalaliDays(jyear As Long, jmonth As Long, jday As Long) As Long

Dim TotalDays As Long
Dim leap, tmp As Long
leap = JLeapYears(jyear - 1)
tmp = JDayOfYear(jyear, jmonth, jday)
TotalDays = (jyear - 1) * 365 + leap + tmp
JalaliDays = TotalDays
End Function


Public Function GDayOfYear(Year As Long, Month As Long, day As Long) As Long

Dim i As Long, leap As Long
leap = Gleap(Year)
For i = 1 To Month - 1
    day = day + GDayTab(leap, i)
Next i
GDayOfYear = day

End Function

Public Function FDate(Jdate As String) As String

Dim Year As String, Month As String, day As String
If Len(Jdate) < 8 Then
    day = Right(Jdate, 2)
    Month = Mid(Jdate, 3, 2)
    Year = Left(Jdate, 2)
Else
    day = Right(Jdate, 2)
    Month = Mid(Jdate, 5, 2)
    Year = Left(Jdate, 4)
End If
FDate = Year + "/" + Month + "/" + day

End Function


Public Function Gleap(Year As Long) As Long
If ((Year Mod 4 = 0 And Year Mod 100 <> 0) Or Year Mod 400 = 0) Then
    Gleap = 1
Else
    Gleap = 0
End If
End Function

Public Sub GMonthDay(GYear As Long, GDayOfYear, Month, day)
Dim i As Long, leap As Long
leap = Gleap(GYear)
i = 1
Do While GDayOfYear > GDayTab(leap, i)
    GDayOfYear = GDayOfYear - GDayTab(leap, i)
    i = i + 1
Loop
Month = i
day = GDayOfYear
End Sub

Public Function GregDays(GYear As Long, GMonth As Long, GDay As Long) As Long
Dim Div4 As Long, Div100 As Long, Div400 As Long
Dim TotalDays  As Long, tmp As Long
Div4 = Int((GYear - 1) / 4)
Div100 = Int((GYear - 1) / 100)
Div400 = Int((GYear - 1) / 400)
tmp = GDayOfYear(GYear, GMonth, GDay)
TotalDays = (GYear - 1) * 365 + tmp + Div4 - Div100 + Div400
GregDays = TotalDays
End Function

Public Function GregorianDate(Jdate As String, Optional mode As String) As String
Dim jyear As Long, jmonth As Long, jday As Long
Dim GYear As Long, GMonth As Long, GDay As Long
Dim TotalDays As Long
jyear = Year_(Jdate)
If Len(Jdate) = 6 Then jyear = jyear + 1300
jmonth = Month_(Jdate)
jday = Day_(Jdate)
TotalDays = JalaliDays(jyear, jmonth, jday)
GregorianYMD TotalDays, GYear, GMonth, GDay
GregorianDate = YMD2Str(GYear, GMonth, GDay, mode)
End Function

Public Sub GregorianYMD(TotalDays As Long, GYear As Long, GMonth As Long, GDay As Long)
Dim Div4 As Long, Div100 As Long, Div400 As Long
Dim GDays As Long
TotalDays = TotalDays + GYearOff
GYear = Int(TotalDays / (Solar - 0.25 / 33))
Div4 = Int(GYear / 4)
Div100 = Int(GYear / 100)
Div400 = Int(GYear / 400)
GDays = TotalDays - (365 * GYear) - (Div4 - Div100 + Div400)
GYear = GYear + 1
If GDays = 0 Then
    GYear = GYear - 1
    If Gleap(GYear) Then
        GDays = 366
    Else
        GDays = 365
    End If
ElseIf (GDays = 366 And Gleap(GYear) = 0) Then
    GDays = 1
    GYear = GYear + 1
End If
GMonthDay GYear, GDays, GMonth, GDay
End Sub


Public Function YMD2Str(Year As Long, Month As Long, day As Long, Optional mode As String) As String
Dim Y As String, M As String, d As String
Y = LTrim(Str(Year))
M = LTrim(Str(Month))
If Len(M) <> 2 Then
    M = "0" + M
End If
d = LTrim(Str(day))
If Len(d) <> 2 Then
    d = "0" + d
End If
If mode = "" Then
    mode = "SHORT"
Else
    mode = UCase(mode)
End If
Select Case mode
Case "LONG"
    YMD2Str = Y + M + d
Case Else
    YMD2Str = Right(Y, 2) + M + d
End Select
End Function

Public Sub JalaliYMD(TotalDays As Long, jyear As Long, jmonth As Long, jday As Long)
Dim JDays As Long
Dim leap As Long
TotalDays = TotalDays - GYearOff
jyear = Int(TotalDays / (Solar - 0.25 / 33))
leap = JLeapYears(jyear)
JDays = TotalDays - (365 * jyear + leap)
jyear = jyear + 1
If JDays = 0 Then
    jyear = jyear - 1
    If Jleap(jyear) Then
        JDays = 366
    Else
        JDays = 365
    End If
ElseIf (JDays = 366 And Jleap(jyear) = 0) Then
    JDays = 1
    jyear = jyear + 1
End If
JMonthDay jyear, JDays, jmonth, jday
End Sub

Public Function Day_(Date_ As String) As Long
Day_ = CLng(Right(Date_, 2))
End Function


Public Function Month_(Date_ As String) As Long
Select Case Len(Date_)
Case 8
    Month_ = CLng(Mid(Date_, 5, 2))
Case Else
    Month_ = CLng(Mid(Date_, 3, 2))
End Select
End Function

Public Sub JMonthDay(jyear As Long, JDayOfYear As Long, Month As Long, day As Long)
Dim i As Long, leap As Long
leap = Jleap(jyear)
i = 1
Do While JDayOfYear > JDayTab(leap, i)
    JDayOfYear = JDayOfYear - JDayTab(leap, i)
    i = i + 1
Loop
Month = i
day = JDayOfYear
End Sub

Public Sub DateModuleSetup()
Dim i As Long

JDayTab(0, 0) = 0: JDayTab(1, 0) = 0
For i = 1 To 6
    JDayTab(0, i) = 31
    JDayTab(1, i) = 31
Next i
For i = 7 To 11
    JDayTab(0, i) = 30
    JDayTab(1, i) = 30
Next i
JDayTab(0, 12) = 29: JDayTab(1, 12) = 30

GDayTab(0, 0) = 0: GDayTab(1, 0) = 0
GDayTab(0, 1) = 31: GDayTab(1, 1) = 31
GDayTab(0, 2) = 28: GDayTab(1, 2) = 29
GDayTab(0, 3) = 31: GDayTab(1, 3) = 31
GDayTab(0, 4) = 30: GDayTab(1, 4) = 30
GDayTab(0, 5) = 31: GDayTab(1, 5) = 31
GDayTab(0, 6) = 30: GDayTab(1, 6) = 30
GDayTab(0, 7) = 31: GDayTab(1, 7) = 31
GDayTab(0, 8) = 31: GDayTab(1, 8) = 31
GDayTab(0, 9) = 30: GDayTab(1, 9) = 30
GDayTab(0, 10) = 31: GDayTab(1, 10) = 31
GDayTab(0, 11) = 30: GDayTab(1, 11) = 30
GDayTab(0, 12) = 31: GDayTab(1, 12) = 31
End Sub

Public Function Year_(Date_ As String) As Long
Select Case Len(Date_)
Case 8
    Year_ = CLng(Left(Date_, 4))
Case Else
    Year_ = CLng(Left(Date_, 2))
End Select
End Function

Public Function JalaliDate(GDate As String, Optional mode As String) As String

Dim GYear As Long, GMonth As Long, GDay As Long
Dim jyear As Long, jmonth As Long, jday As Long
Dim TotalDays As Long
GYear = Year_(GDate)
If Len(GDate) = 6 Then GYear = GYear + 1900
GMonth = Month_(GDate)
GDay = Day_(GDate)
TotalDays = GregDays(GYear, GMonth, GDay)
JalaliYMD TotalDays, jyear, jmonth, jday
JalaliDate = YMD2Str(jyear, jmonth, jday, mode)

End Function






