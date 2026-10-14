Attribute VB_Name = "dates"
' Conformance checks for date/time functions and Date arithmetic.
' Expected values follow real VBA 7 behaviour (en-US locale, Windows),
' with examples ported from the VBA language reference (MIT licensed).
Option Explicit

Sub Main()
    TestLiterals
    TestDocExamples
    TestDateSerial
    TestTimeSerial
    TestDateValueTimeValue
    TestParts
    TestWeekdayName
    TestMonthName
    TestIsDate
    TestCDate
    TestArithmetic
    TestCStr
    TestDateAdd
    TestDateDiff
    TestDatePart
    TestNowDateTime
    TestNulls
    TestErrors
End Sub

Private Sub TestLiterals()
    CheckTyped "literal: VarType", VarType(#2/12/1969#), vbDate
    CheckTyped "literal: TypeName", TypeName(#2/12/1969#), "Date"
    Check "literal: long month name", #February 12, 1969# = #2/12/1969#, True
    Check "literal: two-digit year 69", #2/12/69# = #2/12/1969#, True
    Check "literal: 24h time", #4:35:17 PM# = #16:35:17#, True
    Check "literal: serial", CDbl(#2/12/1969#), 25246
    Check "literal: date and time", CDbl(#2/12/1969 12:00:00 PM#), 25246.5
    Check "literal: compare", #1/1/2020# < #1/2/2020#, True
    Dim d As Date
    Check "literal: default Date var", CDbl(d), 0
    d = #3/4/2021#
    Check "literal: assign to Date var", Day(d), 4
End Sub

Private Sub TestDocExamples()
    Dim MyDate, MyTime, MyCheck, MyVar
    MyDate = DateSerial(1969, 2, 12)
    Check "dateserial-function: MyDate", MyDate, #2/12/1969#
    Check "dateserial-function: May 31 1980", DateSerial(1990 - 10, 8 - 2, 1 - 1), #5/31/1980#
    MyDate = DateValue("February 12, 1969")
    Check "datevalue-function: MyDate", MyDate, #2/12/1969#
    MyDate = #February 12, 1969#
    Check "day-function: MyDay", Day(MyDate), 12
    Check "month-function: MyMonth", Month(MyDate), 2
    Check "year-function: MyYear", Year(MyDate), 1969
    Check "weekday-function: MyWeekDay", Weekday(MyDate), 4
    MyTime = #4:35:17 PM#
    Check "hour-function: MyHour", Hour(MyTime), 16
    Check "minute-function: MyMinute", Minute(MyTime), 35
    Check "second-function: MySecond", Second(MyTime), 17
    MyTime = TimeSerial(16, 35, 17)
    Check "timeserial-function: MyTime", MyTime, #4:35:17 PM#
    Check "timeserial-function: 5:45 AM", TimeSerial(12 - 6, -15, 0), #5:45:00 AM#
    MyTime = TimeValue("4:35:17 PM")
    Check "timevalue-function: MyTime", MyTime, #4:35:17 PM#
    Check "dateadd-function: 31-Jan-95", DateAdd("m", 1, "31-Jan-95"), #2/28/1995#
    Check "dateadd-function: 31-Jan-96", DateAdd("m", 1, "31-Jan-96"), #2/29/1996#
    Check "datediff-function: yyyy Dec31-Jan1", DateDiff("yyyy", #12/31/2020#, #1/1/2021#), 1
    MyVar = "04/28/2014"
    MyCheck = IsDate(MyVar)
    Check "isdate-function: 04/28/2014", MyCheck, True
    MyVar = "April 28, 2014"
    Check "isdate-function: April 28, 2014", IsDate(MyVar), True
    MyVar = "13/32/2014"
    Check "isdate-function: 13/32/2014", IsDate(MyVar), False
    MyVar = "04.28.14"
    Check "isdate-function: 04.28.14", IsDate(MyVar), True
    MyVar = "04.28.2014"
    Check "isdate-function: 04.28.2014", IsDate(MyVar), False
End Sub

Private Sub TestDateSerial()
    CheckTyped "dateserial: type", DateSerial(2020, 1, 1), #1/1/2020#
    Check "dateserial: month 13", DateSerial(2020, 13, 1), #1/1/2021#
    Check "dateserial: day 0", DateSerial(2020, 3, 0), #2/29/2020#
    Check "dateserial: day 0 non-leap", DateSerial(2021, 3, 0), #2/28/2021#
    Check "dateserial: month 0", DateSerial(2020, 0, 1), #12/1/2019#
    Check "dateserial: day -1", DateSerial(2020, 1, -1), #12/30/2019#
    Check "dateserial: day 32", DateSerial(2020, 1, 32), #2/1/2020#
    Check "dateserial: day 366", DateSerial(2020, 1, 366), #12/31/2020#
    Check "dateserial: month -1", DateSerial(2020, -1, 1), #11/1/2019#
    Check "dateserial: 2-digit 69", DateSerial(69, 2, 12), #2/12/1969#
    Check "dateserial: 2-digit 29", DateSerial(29, 1, 1), #1/1/2029#
    Check "dateserial: 2-digit 30", DateSerial(30, 1, 1), #1/1/1930#
    Check "dateserial: epoch", CDbl(DateSerial(1899, 12, 30)), 0
End Sub

Private Sub TestTimeSerial()
    CheckTyped "timeserial: type", TimeSerial(12, 0, 0), #12:00:00 PM#
    Check "timeserial: noon", CDbl(TimeSerial(12, 0, 0)), 0.5
    Check "timeserial: 6am", CDbl(TimeSerial(6, 0, 0)), 0.25
    Check "timeserial: 75 minutes", TimeSerial(0, 75, 0), #1:15:00 AM#
    Check "timeserial: 90 seconds", TimeSerial(0, 0, 90), #12:01:30 AM#
    Check "timeserial: minus 15 min", TimeSerial(12, -15, 0), #11:45:00 AM#
    Check "timeserial: 25 hours", Round(CDbl(TimeSerial(25, 0, 0)) * 24, 6), 25
    Check "timeserial: 25 hours Hour", Hour(TimeSerial(25, 0, 0)), 1
    Check "timeserial: midnight", CDbl(TimeSerial(0, 0, 0)), 0
End Sub

Private Sub TestDateValueTimeValue()
    CheckTyped "datevalue: type", DateValue("2/12/1969"), #2/12/1969#
    Check "datevalue: abbreviated month", DateValue("Feb 12, 1969"), #2/12/1969#
    Check "datevalue: drops time", DateValue("2/12/1969 4:35:17 PM"), #2/12/1969#
    Check "datevalue: two-digit year", DateValue("12/30/91"), #12/30/1991#
    Check "datevalue: ISO", DateValue("1969-02-12"), #2/12/1969#
    Check "datevalue: from Date", DateValue(#2/12/1969 4:35:17 PM#), #2/12/1969#
    CheckTyped "timevalue: type", TimeValue("4:35:17 PM"), #4:35:17 PM#
    Check "timevalue: 24h", TimeValue("16:35:17"), #4:35:17 PM#
    Check "timevalue: drops date", TimeValue("2/12/1969 4:35:17 PM"), #4:35:17 PM#
    Check "timevalue: from Date", TimeValue(#2/12/1969 6:00:00 AM#), #6:00:00 AM#
    Check "timevalue: noon serial", CDbl(TimeValue("12:00 PM")), 0.5
    Check "timevalue: midnight", CDbl(TimeValue("12:00:00 AM")), 0
End Sub

Private Sub TestParts()
    Dim dt As Date
    dt = #2/12/1969 4:35:17 PM#
    CheckTyped "day: type", Day(dt), CInt(12)
    CheckTyped "month: type", Month(dt), CInt(2)
    CheckTyped "year: type", Year(dt), CInt(1969)
    CheckTyped "hour: type", Hour(dt), CInt(16)
    CheckTyped "minute: type", Minute(dt), CInt(35)
    CheckTyped "second: type", Second(dt), CInt(17)
    CheckTyped "weekday: type", Weekday(dt), CInt(4)
    Check "day: from string", Day("2/12/1969"), 12
    Check "year: from string", Year("February 12, 1969"), 1969
    Check "hour: from string", Hour("4:35:17 PM"), 16
    Check "day: serial 0", Day(0), 30
    Check "month: serial 0", Month(0), 12
    Check "year: serial 0", Year(0), 1899
    Check "day: serial 2", Day(2), 1
    Check "year: serial 2", Year(2), 1900
    Check "hour: of date-only", Hour(#2/12/1969#), 0
    Check "hour: serial 0.75", Hour(0.75), 18
    Check "weekday: Friday", Weekday(#1/1/2021#), 6
    Check "weekday: vbSunday", Weekday(#1/1/2021#, vbSunday), 6
    Check "weekday: vbMonday", Weekday(#1/1/2021#, vbMonday), 5
    Check "weekday: vbSaturday", Weekday(#1/1/2021#, vbSaturday), 7
    Check "weekday: vbFriday", Weekday(#1/1/2021#, vbFriday), 1
    Check "weekday: system", Weekday(#1/1/2021#, vbUseSystemDayOfWeek), 6
    Check "weekday: Sunday", Weekday(#1/3/2021#), vbSunday
    Check "weekday: Saturday", Weekday(#1/2/2021#), vbSaturday
    Check "weekday: serial 0", Weekday(0), vbSaturday
    Check "dates: negative date day", Day(CDate(-1.5)), 29
    Check "dates: negative date hour", Hour(CDate(-1.5)), 12
End Sub

Private Sub TestWeekdayName()
    CheckTyped "weekdayname-function: 1", WeekdayName(1), "Sunday"
    Check "weekdayname-function: 7", WeekdayName(7), "Saturday"
    Check "weekdayname-function: abbrev", WeekdayName(4, True), "Wed"
    Check "weekdayname-function: not abbrev", WeekdayName(4, False), "Wednesday"
    Check "weekdayname-function: vbMonday", WeekdayName(1, False, vbMonday), "Monday"
    Check "weekdayname-function: vbMonday 7", WeekdayName(7, True, vbMonday), "Sun"
    Check "weekdayname-function: from Weekday", WeekdayName(Weekday(#2/12/1969#)), "Wednesday"
End Sub

Private Sub TestMonthName()
    CheckTyped "monthname-function: 1", MonthName(1), "January"
    Check "monthname-function: 12", MonthName(12), "December"
    Check "monthname-function: abbrev", MonthName(2, True), "Feb"
    Check "monthname-function: not abbrev", MonthName(9, False), "September"
    Check "monthname-function: May abbrev", MonthName(5, True), "May"
End Sub

Private Sub TestIsDate()
    CheckTyped "isdate: literal", IsDate(#1/1/2020#), True
    Check "isdate: Date var", IsDate(DateSerial(2020, 1, 1)), True
    Check "isdate: time string", IsDate("4:35:17 PM"), True
    Check "isdate: ISO string", IsDate("2020-01-31"), True
    Check "isdate: Feb 30", IsDate("2/30/2020"), False
    Check "isdate: Feb 29 leap", IsDate("2/29/2020"), True
    Check "isdate: Feb 29 non-leap", IsDate("2/29/2021"), False
    Check "isdate: word", IsDate("hello"), False
    Check "isdate: empty string", IsDate(""), False
    Check "isdate: Empty", IsDate(Empty), False
    Check "isdate: Null", IsDate(Null), False
    Check "isdate: number", IsDate(42), False
    Check "isdate: bad time", IsDate("25:00"), False
End Sub

Private Sub TestCDate()
    CheckTyped "cdate: string", CDate("2/12/1969"), #2/12/1969#
    Check "cdate: long month", CDate("February 12, 1969"), #2/12/1969#
    Check "cdate: time", CDate("4:35:17 PM"), #4:35:17 PM#
    Check "cdate: date time", CDate("2/12/1969 4:35:17 PM"), #2/12/1969 4:35:17 PM#
    Check "cdate: ISO", CDate("1969-02-12"), #2/12/1969#
    Check "cdate: dd-mmm-yy", CDate("12-Feb-69"), #2/12/1969#
    CheckTyped "cdate: number 2", CDate(2), #1/1/1900#
    Check "cdate: number 0", CDate(0), #12/30/1899#
    Check "cdate: number -1", CDate(-1), #12/29/1899#
    Check "cdate: fractional", CDate(25246.5), #2/12/1969 12:00:00 PM#
    Check "cdate: numeric string", CDbl(CDate("2/12/1969")), 25246
    CheckTyped "cdbl: from date", CDbl(#1/1/1900#), CDbl(2)
    Check "cdbl: epoch", CDbl(#12/30/1899#), 0
    Check "cdbl: 12/31/1899", CDbl(#12/31/1899#), 1
    Check "cdbl: 2020", CDbl(#1/1/2020#), 43831
    Check "clng: date", CLng(#2/12/1969#), 25246
    Check "cdbl: 6pm", CDbl(#6:00:00 PM#), 0.75
End Sub

Private Sub TestArithmetic()
    CheckTyped "arith: date + 1", #1/1/2020# + 1, #1/2/2020#
    CheckTyped "arith: 1 + date", 1 + #1/1/2020#, #1/2/2020#
    CheckTyped "arith: date - 1", #1/1/2020# - 1, #12/31/2019#
    CheckTyped "arith: date - date", #1/2/2020# - #1/1/2020#, CDbl(1)
    Check "arith: leap year span", #3/1/2020# - #2/1/2020#, 29
    Check "arith: date + 0.5", #1/1/2020# + 0.5, #1/1/2020 12:00:00 PM#
    Check "arith: half day diff", #1/1/2020 12:00:00 PM# - #1/1/2020#, 0.5
    Check "arith: negative diff", #1/1/2020# - #1/2/2020#, -1
    Dim d As Date
    d = #12/31/2020#
    d = d + 1
    Check "arith: Date var + 1", d, #1/1/2021#
    Check "arith: Date var year", Year(d), 2021
    d = 2
    Check "arith: assign Double to Date", d, #1/1/1900#
    Check "arith: Int of date", Int(#2/12/1969 4:35:17 PM#), #2/12/1969#
    Check "arith: compare with number", #1/1/1900# = 2, True
End Sub

Private Sub TestCStr()
    CheckTyped "cstr: date", CStr(#2/12/1969#), "2/12/1969"
    Check "cstr: time", CStr(#4:35:17 PM#), "4:35:17 PM"
    Check "cstr: date time", CStr(#2/12/1969 4:35:17 PM#), "2/12/1969 4:35:17 PM"
    Check "cstr: morning", CStr(#9:05:03 AM#), "9:05:03 AM"
    Check "cstr: midnight date", CStr(#1/1/2020 12:00:00 AM#), "1/1/2020"
    Check "cstr: zero date", CStr(CDate(0)), "12:00:00 AM"
    Check "cstr: noon", CStr(CDate(0.5)), "12:00:00 PM"
    Check "cstr: dateserial", CStr(DateSerial(2020, 12, 25)), "12/25/2020"
    Check "cstr: concat", "x" & #1/1/2020#, "x1/1/2020"
    Check "cstr: serial 1", CStr(CDate(1)), "12/31/1899"
    Check "cstr: dateadd", CStr(DateAdd("h", 13, #1/1/2020#)), "1/1/2020 1:00:00 PM"
End Sub

Private Sub TestDateAdd()
    CheckTyped "dateadd: type", DateAdd("d", 1, #12/31/2020#), #1/1/2021#
    Check "dateadd: m Jan31 leap", DateAdd("m", 1, #1/31/2020#), #2/29/2020#
    Check "dateadd: m negative", DateAdd("m", -1, #3/31/2021#), #2/28/2021#
    Check "dateadd: yyyy Feb29", DateAdd("yyyy", 1, #2/29/2020#), #2/28/2021#
    Check "dateadd: yyyy 4", DateAdd("yyyy", 4, #2/29/2020#), #2/29/2024#
    Check "dateadd: q", DateAdd("q", 1, #11/30/2020#), #2/28/2021#
    Check "dateadd: y", DateAdd("y", 10, #1/1/2020#), #1/11/2020#
    Check "dateadd: w", DateAdd("w", 7, #1/1/2021#), #1/8/2021#
    Check "dateadd: ww", DateAdd("ww", 2, #1/1/2021#), #1/15/2021#
    Check "dateadd: h", DateAdd("h", 25, #1/1/2020#), #1/2/2020 1:00:00 AM#
    Check "dateadd: n negative", DateAdd("n", -1, #1/1/2020#), #12/31/2019 11:59:00 PM#
    Check "dateadd: s", DateAdd("s", 90, #1/1/2020#), #1/1/2020 12:01:30 AM#
    Check "dateadd: d negative", DateAdd("d", -1, #1/1/2020#), #12/31/2019#
    Check "dateadd: keeps time", DateAdd("d", 1, #1/1/2020 4:35:17 PM#), #1/2/2020 4:35:17 PM#
    Check "dateadd: m keeps time", DateAdd("m", 1, #1/15/2020 6:00:00 AM#), #2/15/2020 6:00:00 AM#
    Check "dateadd: string date", DateAdd("d", 1, "2/12/1969"), #2/13/1969#
    Check "dateadd: uppercase interval", DateAdd("D", 1, #1/1/2020#), #1/2/2020#
    Check "dateadd: m 12", DateAdd("m", 12, #1/31/2020#), #1/31/2021#
End Sub

Private Sub TestDateDiff()
    CheckTyped "datediff: type", DateDiff("d", #1/1/2020#, #3/1/2020#), CLng(60)
    Check "datediff: yyyy same year", DateDiff("yyyy", #1/1/2020#, #12/31/2020#), 0
    Check "datediff: m boundary", DateDiff("m", #1/31/2020#, #2/1/2020#), 1
    Check "datediff: m within year", DateDiff("m", #1/1/2020#, #12/31/2020#), 11
    Check "datediff: q boundary", DateDiff("q", #3/31/2020#, #4/1/2020#), 1
    Check "datediff: y", DateDiff("y", #1/1/2020#, #3/1/2020#), 60
    Check "datediff: d negative", DateDiff("d", #1/10/2021#, #1/1/2021#), -9
    Check "datediff: d ignores time", DateDiff("d", #1/1/2020 11:00:00 PM#, #1/2/2020 1:00:00 AM#), 1
    Check "datediff: d same day", DateDiff("d", #1/1/2020 1:00:00 AM#, #1/1/2020 11:00:00 PM#), 0
    Check "datediff: h boundary", DateDiff("h", #1/1/2021 11:59:00 PM#, #1/2/2021#), 1
    Check "datediff: h", DateDiff("h", #1/1/2020#, #1/2/2020#), 24
    Check "datediff: n boundary", DateDiff("n", #10:00:59 AM#, #10:01:00 AM#), 1
    Check "datediff: n", DateDiff("n", #10:00:00 AM#, #11:30:00 AM#), 90
    Check "datediff: s", DateDiff("s", #10:00:00 AM#, #10:01:30 AM#), 90
    Check "datediff: s day", DateDiff("s", #1/1/2020#, #1/2/2020#), 86400
    ' "ww" counts Sundays after date1 up to and including date2
    Check "datediff: ww", DateDiff("ww", #1/1/2021#, #1/10/2021#), 2
    Check "datediff: ww date1 Sunday", DateDiff("ww", #1/3/2021#, #1/10/2021#), 1
    Check "datediff: ww vbMonday", DateDiff("ww", #1/1/2021#, #1/10/2021#, vbMonday), 1
    Check "datediff: ww vbSaturday", DateDiff("ww", #1/1/2021#, #1/10/2021#, vbSaturday), 2
    Check "datediff: ww no Sunday", DateDiff("ww", #1/4/2021#, #1/9/2021#), 0
    ' "w" counts whole 7-day periods
    Check "datediff: w 9 days", DateDiff("w", #1/1/2021#, #1/10/2021#), 1
    Check "datediff: w 14 days", DateDiff("w", #1/1/2021#, #1/15/2021#), 2
    Check "datediff: w 6 days", DateDiff("w", #1/1/2021#, #1/7/2021#), 0
    Check "datediff: string args", DateDiff("d", "1/1/2020", "1/31/2020"), 30
    Check "datediff: yyyy negative", DateDiff("yyyy", #1/1/2021#, #12/31/2019#), -2
End Sub

Private Sub TestDatePart()
    Dim dt As Date
    dt = #2/12/1969 4:35:17 PM#
    CheckTyped "datepart: yyyy", DatePart("yyyy", dt), CInt(1969)
    Check "datepart: q", DatePart("q", dt), 1
    Check "datepart: q Nov", DatePart("q", #11/5/2020#), 4
    Check "datepart: m", DatePart("m", dt), 2
    Check "datepart: d", DatePart("d", dt), 12
    Check "datepart: y", DatePart("y", dt), 43
    Check "datepart: y leap Mar1", DatePart("y", #3/1/2020#), 61
    Check "datepart: y Dec31", DatePart("y", #12/31/2021#), 365
    Check "datepart: w", DatePart("w", dt), 4
    Check "datepart: w vbMonday", DatePart("w", dt, vbMonday), 3
    Check "datepart: h", DatePart("h", dt), 16
    Check "datepart: n", DatePart("n", dt), 35
    Check "datepart: s", DatePart("s", dt), 17
    Check "datepart: ww", DatePart("ww", dt), 7
    Check "datepart: ww Jan1", DatePart("ww", #1/1/2021#), 1
    Check "datepart: ww Jan3", DatePart("ww", #1/3/2021#), 2
    Check "datepart: ww Jan3 vbMonday", DatePart("ww", #1/3/2021#, vbMonday), 1
    Check "datepart: ww Dec31", DatePart("ww", #12/31/2021#), 53
    Check "datepart: ww vbFirstJan1", DatePart("ww", #1/1/2021#, vbSunday, vbFirstJan1), 1
    Check "datepart: ww vbFirstFourDays Jan1", DatePart("ww", #1/1/2021#, vbMonday, vbFirstFourDays), 53
    Check "datepart: ww vbFirstFourDays Jan4", DatePart("ww", #1/4/2021#, vbMonday, vbFirstFourDays), 1
    Check "datepart: ww vbFirstFourDays Dec31", DatePart("ww", #12/31/2020#, vbMonday, vbFirstFourDays), 53
    Check "datepart: ww vbFirstFourDays 2020", DatePart("ww", #1/1/2020#, vbMonday, vbFirstFourDays), 1
    Check "datepart: ww vbFirstFourDays Sunday", DatePart("ww", #1/1/2021#, vbSunday, vbFirstFourDays), 53
    Check "datepart: ww vbFirstFullWeek", DatePart("ww", #1/3/2021#, vbSunday, vbFirstFullWeek), 1
    Check "datepart: string date", DatePart("m", "2/12/1969"), 2
End Sub

Private Sub TestNowDateTime()
    Check "now: type", TypeName(Now), "Date"
    Check "date: type", TypeName(Date), "Date"
    Check "time: type", TypeName(Time), "Date"
    Check "date: no time part", CDbl(Date) = Int(CDbl(Date)), True
    Check "time: no date part", CDbl(Time) < 1, True
    Check "timer: type", TypeName(Timer), "Single"
    Check "date$: type", TypeName(Date$), "String"
End Sub

Private Sub TestNulls()
    Check "null: Year", Year(Null), Null
    Check "null: Month", Month(Null), Null
    Check "null: Day", Day(Null), Null
    Check "null: Hour", Hour(Null), Null
    Check "null: Minute", Minute(Null), Null
    Check "null: Second", Second(Null), Null
    Check "null: Weekday", Weekday(Null), Null
    Check "null: DateAdd", DateAdd("d", 1, Null), Null
    Check "null: DateDiff", DateDiff("d", Null, #1/1/2020#), Null
    Check "null: DatePart", DatePart("d", Null), Null
    Check "null: date + Null", #1/1/2020# + Null, Null
End Sub

Private Sub TestErrors()
    Dim v As Variant
    On Error Resume Next
    v = CDate("abc")
    CheckErr "err: CDate bad string", 13
    v = CDate(Null)
    CheckErr "err: CDate Null", 94
    v = DateValue("abc")
    CheckErr "err: DateValue bad string", 13
    v = TimeValue("abc")
    CheckErr "err: TimeValue bad string", 13
    v = Year("abc")
    CheckErr "err: Year bad string", 13
    v = Month("13/32/2014")
    CheckErr "err: Month invalid date", 13
    v = WeekdayName(8)
    CheckErr "err: WeekdayName 8", 5
    v = WeekdayName(0)
    CheckErr "err: WeekdayName 0", 5
    v = MonthName(13)
    CheckErr "err: MonthName 13", 5
    v = MonthName(0)
    CheckErr "err: MonthName 0", 5
    v = DateAdd("x", 1, #1/1/2020#)
    CheckErr "err: DateAdd bad interval", 5
    v = DateDiff("x", #1/1/2020#, #1/2/2020#)
    CheckErr "err: DateDiff bad interval", 5
    v = DatePart("x", #1/1/2020#)
    CheckErr "err: DatePart bad interval", 5
    v = DateAdd("yyyy", -2000, #1/1/2000#)
    CheckErr "err: DateAdd before year 100", 5
    v = DateSerial(2020, 1, 40000)
    CheckErr "err: DateSerial arg out of Integer range", 6
    On Error GoTo 0
End Sub
