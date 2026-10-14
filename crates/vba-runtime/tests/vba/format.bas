Attribute VB_Name = "format"
Option Explicit

' Conformance tests for Format, Format$, FormatNumber, FormatCurrency,
' FormatPercent and FormatDateTime (en-US locale).

Sub Main()
    TestFormatDocExample
    TestNoFormat
    TestNamedNumeric
    TestNamedDateTime
    TestUserNumeric
    TestNumericSections
    TestScientific
    TestUserString
    TestUserDate
    TestUserTime
    TestFormatMisc
    TestFormatErrors
    TestFormatNumber
    TestFormatCurrency
    TestFormatPercent
    TestFormatDateTime
End Sub

' Ported from format-function-visual-basic-for-applications.md
Private Sub TestFormatDocExample()
    Dim MyTime, MyDate, MyStr
    MyTime = #5:04:23 PM#
    MyDate = #1/27/1993#
    MyStr = Format(MyTime, "h:m:s")
    Check "format-function: h:m:s", MyStr, "17:4:23"
    MyStr = Format(MyTime, "hh:mm:ss am/pm")
    Check "format-function: am/pm", MyStr, "05:04:23 pm"
    MyStr = Format(MyTime, "hh:mm:ss AM/PM")
    Check "format-function: AM/PM", MyStr, "05:04:23 PM"
    MyStr = Format(MyDate, "dddd, mmm d yyyy")
    Check "format-function: dddd, mmm d yyyy", MyStr, "Wednesday, Jan 27 1993"
    MyStr = Format(23)
    Check "format-function: Format(23)", MyStr, "23"
    MyStr = Format(5459.4, "##,##0.00")
    Check "format-function: ##,##0.00", MyStr, "5,459.40"
    MyStr = Format(334.9, "###0.00")
    Check "format-function: ###0.00", MyStr, "334.90"
    MyStr = Format(5, "0.00%")
    Check "format-function: 0.00%", MyStr, "500.00%"
    MyStr = Format("HELLO", "<")
    Check "format-function: <", MyStr, "hello"
    MyStr = Format("This is it", ">")
    Check "format-function: >", MyStr, "THIS IS IT"
    ' Section examples from the docs
    Check "format-function: $#,##0;($#,##0) pos", Format(1234, "$#,##0;($#,##0)"), "$1,234"
    Check "format-function: $#,##0;($#,##0) neg", Format(-1234, "$#,##0;($#,##0)"), "($1,234)"
    Check "format-function: $#,##0;;\Z\e\r\o zero", Format(0, "$#,##0;;\Z\e\r\o"), "Zero"
    Check "format-function: $#,##0;;\Z\e\r\o pos", Format(5, "$#,##0;;\Z\e\r\o"), "$5"
    Check "format-function: ##0,, scaling", Format(100000000, "##0,,"), "100"
    Check "format-function: hh:mm:ss AMPM", Format(MyTime, "hh:mm:ss AMPM"), "05:04:23 PM"
End Sub

Private Sub TestNoFormat()
    CheckTyped "format-function: Format(23) type", Format(23), "23"
    Check "format-function: Format(1.5)", Format(1.5), "1.5"
    Check "format-function: Format(-1.5)", Format(-1.5), "-1.5"
    Check "format-function: Format(0.5)", Format(0.5), "0.5"
    Check "format-function: Format(True)", Format(True), "True"
    Check "format-function: Format(False)", Format(False), "False"
    Check "format-function: Format(""abc"")", Format("abc"), "abc"
    Check "format-function: Format(Empty)", Format(Empty), ""
    Check "format-function: Format(Null) is Null", IsNull(Format(Null)), True
    Check "format-function: Format(Null, ""0"") is Null", IsNull(Format(Null, "0")), True
    Check "format-function: Format(1E+20)", Format(1E+20), "1E+20"
    Check "format-function: Format(date)", Format(#1/27/1993#), "1/27/1993"
    Check "format-function: Format(time)", Format(#5:04:23 PM#), "5:04:23 PM"
    Check "format-function: Format(datetime)", Format(#1/27/1993 5:04:23 PM#), "1/27/1993 5:04:23 PM"
    Check "format-function: Format(midnight)", Format(#12:00:00 AM#), "12:00:00 AM"
    Check "format-function: Format(CCur)", Format(CCur(12.5)), "12.5"
    Check "format-function: Format(empty fmt)", Format(1234.5, ""), "1234.5"
    CheckTyped "format-function: Format$", Format$(1.5, "0.0"), "1.5"
End Sub

Private Sub TestNamedNumeric()
    Check "format-function: General Number", Format(1234.5678, "General Number"), "1234.5678"
    Check "format-function: General Number neg", Format(-0.5, "General Number"), "-0.5"
    Check "format-function: Currency", Format(1234.5, "Currency"), "$1,234.50"
    Check "format-function: Currency neg", Format(-1234.5, "Currency"), "($1,234.50)"
    Check "format-function: Currency zero", Format(0, "Currency"), "$0.00"
    Check "format-function: Fixed", Format(1234.5, "Fixed"), "1234.50"
    Check "format-function: Fixed small", Format(0.5, "Fixed"), "0.50"
    Check "format-function: Fixed neg", Format(-0.5, "Fixed"), "-0.50"
    Check "format-function: Fixed round", Format(2.345, "Fixed"), "2.35"
    Check "format-function: Standard", Format(1234.5, "Standard"), "1,234.50"
    Check "format-function: Standard neg", Format(-1234567.891, "Standard"), "-1,234,567.89"
    Check "format-function: Standard small", Format(0.5, "Standard"), "0.50"
    Check "format-function: Percent", Format(0.5, "Percent"), "50.00%"
    Check "format-function: Percent zero", Format(0, "Percent"), "0.00%"
    Check "format-function: Percent big", Format(12.345, "Percent"), "1234.50%"
    Check "format-function: Scientific", Format(1234.5, "Scientific"), "1.23E+03"
    Check "format-function: Scientific small", Format(0.000012345, "Scientific"), "1.23E-05"
    Check "format-function: Scientific neg", Format(-1234.5, "Scientific"), "-1.23E+03"
    Check "format-function: Yes/No 0", Format(0, "Yes/No"), "No"
    Check "format-function: Yes/No 5", Format(5, "Yes/No"), "Yes"
    Check "format-function: Yes/No -1", Format(-1, "Yes/No"), "Yes"
    Check "format-function: True/False 0", Format(0, "True/False"), "False"
    Check "format-function: True/False 1", Format(1, "True/False"), "True"
    Check "format-function: On/Off 0", Format(0, "On/Off"), "Off"
    Check "format-function: On/Off 2", Format(2, "On/Off"), "On"
    Check "format-function: named case-insensitive", Format(1234.5, "standard"), "1,234.50"
End Sub

Private Sub TestNamedDateTime()
    Dim d As Date, t As Date, dt As Date, am As Date
    d = #1/27/1993#
    t = #5:04:23 PM#
    dt = #1/27/1993 5:04:23 PM#
    am = #9:05:07 AM#
    Check "format-function: General Date", Format(dt, "General Date"), "1/27/1993 5:04:23 PM"
    Check "format-function: General Date date only", Format(d, "General Date"), "1/27/1993"
    Check "format-function: General Date time only", Format(t, "General Date"), "5:04:23 PM"
    Check "format-function: General Date number", Format(1.5, "General Date"), "12/31/1899 12:00:00 PM"
    Check "format-function: Long Date", Format(d, "Long Date"), "Wednesday, January 27, 1993"
    Check "format-function: Medium Date", Format(d, "Medium Date"), "27-Jan-93"
    Check "format-function: Short Date", Format(d, "Short Date"), "1/27/1993"
    Check "format-function: Short Date dt", Format(dt, "Short Date"), "1/27/1993"
    Check "format-function: Long Time", Format(t, "Long Time"), "5:04:23 PM"
    Check "format-function: Long Time am", Format(am, "Long Time"), "9:05:07 AM"
    Check "format-function: Medium Time", Format(t, "Medium Time"), "05:04 PM"
    Check "format-function: Medium Time am", Format(am, "Medium Time"), "09:05 AM"
    Check "format-function: Short Time", Format(t, "Short Time"), "17:04"
    Check "format-function: Short Time am", Format(am, "Short Time"), "09:05"
End Sub

Private Sub TestUserNumeric()
    Check "format-function: 0.00", Format(0.5, "0.00"), "0.50"
    Check "format-function: #.##", Format(0.5, "#.##"), ".5"
    Check "format-function: #.## int", Format(12, "#.##"), "12."
    Check "format-function: #.## 12.3", Format(12.3, "#.##"), "12.3"
    Check "format-function: 0.0#", Format(12, "0.0#"), "12.0"
    Check "format-function: 0.0# round", Format(12.345, "0.0#"), "12.35"
    Check "format-function: .00", Format(0.123, ".00"), ".12"
    Check "format-function: 000", Format(5, "000"), "005"
    Check "format-function: 000 neg", Format(-5, "000"), "-005"
    Check "format-function: 00000.0", Format(1234, "00000.0"), "01234.0"
    Check "format-function: 0 extra digits", Format(12345, "00"), "12345"
    Check "format-function: 0 round half up", Format(2.5, "0"), "3"
    Check "format-function: 0 round 3.5", Format(3.5, "0"), "4"
    Check "format-function: 0 round neg", Format(-2.5, "0"), "-3"
    Check "format-function: 0.00 round .125", Format(0.125, "0.00"), "0.13"
    Check "format-function: 0.000 third", Format(1 / 3, "0.000"), "0.333"
    Check "format-function: -0", Format(-0.4, "0"), "-0"
    Check "format-function: # zero", Format(0, "#"), ""
    Check "format-function: #,##0", Format(1234567, "#,##0"), "1,234,567"
    Check "format-function: #,##0 big", Format(1234567890123#, "#,##0"), "1,234,567,890,123"
    Check "format-function: #,##0 small", Format(123, "#,##0"), "123"
    Check "format-function: #,#", Format(1234, "#,#"), "1,234"
    Check "format-function: #,##0.00 neg", Format(-1234.567, "#,##0.00"), "-1,234.57"
    Check "format-function: #,##0, scale", Format(1234567, "#,##0,"), "1,235"
    Check "format-function: 0% ", Format(0.25, "0%"), "25%"
    Check "format-function: 0.0%", Format(0.5, "0.0%"), "50.0%"
    Check "format-function: 0.00% small", Format(0.0525, "0.00%"), "5.25%"
    Check "format-function: $ literal", Format(123, "$0"), "$123"
    Check "format-function: \$ literal", Format(123, "\$0"), "$123"
    Check "format-function: quoted literal", Format(123, """Total: ""0"), "Total: 123"
    Check "format-function: quoted suffix", Format(5, "#0.0 ""units"""), "5.0 units"
    Check "format-function: parens literal", Format(5, "(0)"), "(5)"
    Check "format-function: backslash backslash", Format(5, "0\\"), "5\"
    Check "format-function: numeric string", Format("123", "0.00"), "123.00"
    Check "format-function: non-numeric string", Format("abc", "0.00"), "abc"
    Check "format-function: True numeric", Format(True, "0"), "-1"
    Check "format-function: Currency value", Format(CCur(1234.5678), "0.00"), "1234.57"
    Check "format-function: date as number", Format(#1/1/2000#, "0"), "36526"
    CheckTyped "format-function: result type", Format(1, "0"), "1"
End Sub

Private Sub TestNumericSections()
    Check "format-function: 2 sections neg", Format(-1234.5, "#,##0.00;(#,##0.00)"), "(1,234.50)"
    Check "format-function: 2 sections zero", Format(0, "#,##0.00;(#,##0.00)"), "0.00"
    Check "format-function: 3 sections zero", Format(0, "0.0;(0.0);\z\e\r\o"), "zero"
    Check "format-function: 3 sections pos", Format(1.25, "0.0;(0.0);\z\e\r\o"), "1.3"
    Check "format-function: 3 sections neg", Format(-1.25, "0.0;(0.0);\z\e\r\o"), "(1.3)"
    Check "format-function: 4 sections null", Format(Null, "0;-0;\Z\e\r\o;\N\u\l\l"), "Null"
    Check "format-function: 4 sections pos", Format(7, "0;-0;\Z\e\r\o;\N\u\l\l"), "7"
    Check "format-function: 4 sections neg", Format(-7, "0;-0;\Z\e\r\o;\N\u\l\l"), "-7"
    Check "format-function: empty neg section", Format(-5, "$#,##0;;\Z\e\r\o"), "-$5"
    Check "format-function: neg section literal", Format(-5, "0;\m\i\n\u\s 0"), "minus 5"
End Sub

Private Sub TestScientific()
    Check "format-function: 0.00E+00", Format(12345.678, "0.00E+00"), "1.23E+04"
    Check "format-function: 0.00E-00 small", Format(0.000123, "0.00E-00"), "1.23E-04"
    Check "format-function: 0.00E-00 pos", Format(12345.678, "0.00E-00"), "1.23E04"
    Check "format-function: 0.00e+00", Format(1234, "0.00e+00"), "1.23e+03"
    Check "format-function: 0.0E+0", Format(1500, "0.0E+0"), "1.5E+3"
    Check "format-function: 0E+00 neg", Format(-25000, "0.0E+00"), "-2.5E+04"
End Sub

Private Sub TestUserString()
    Check "format-function: @@@@@", Format("abc", "@@@@@"), "  abc"
    Check "format-function: !@@@@@", Format("abc", "!@@@@@"), "abc  "
    Check "format-function: &&&&&", Format("abc", "&&&&&"), "abc"
    Check "format-function: !&&&&", Format("ab", "!&&&&"), "ab"
    Check "format-function: @@-@@@ full", Format("12345", "@@-@@@"), "12-345"
    Check "format-function: @@-@@@ short", Format("1234", "@@-@@@"), " 1-234"
    Check "format-function: phone", Format("5551234567", "(@@@) @@@-@@@@"), "(555) 123-4567"
    Check "format-function: overflow", Format("abcdef", "@@@"), "abcdef"
    Check "format-function: > with @", Format("Hello", ">@@@@@@@"), "  HELLO"
    Check "format-function: < with @", Format("HeLLo", "<@@@@@"), "hello"
    Check "format-function: number with @", Format(123, "@@@@@"), "  123"
    Check "format-function: string 2 sections str", Format("x", "@;\E\m\p\t\y"), "x"
    Check "format-function: string 2 sections empty", Format("", "@;\E\m\p\t\y"), "Empty"
    Check "format-function: string 2 sections null", Format(Null, "@;\N\u\l\l"), "Null"
End Sub

Private Sub TestUserDate()
    Dim d As Date
    d = #1/27/1993#
    Check "format-function: yyyy-mm-dd", Format(d, "yyyy-mm-dd"), "1993-01-27"
    Check "format-function: m/d/yy", Format(d, "m/d/yy"), "1/27/93"
    Check "format-function: mmmm", Format(d, "mmmm"), "January"
    Check "format-function: mmm", Format(d, "mmm"), "Jan"
    Check "format-function: mm", Format(d, "mm"), "01"
    Check "format-function: m", Format(d, "m"), "1"
    Check "format-function: ddd", Format(d, "ddd"), "Wed"
    Check "format-function: dddd", Format(d, "dddd"), "Wednesday"
    Check "format-function: ddddd", Format(d, "ddddd"), "1/27/1993"
    Check "format-function: dddddd", Format(d, "dddddd"), "Wednesday, January 27, 1993"
    Check "format-function: d", Format(#3/5/2001#, "d"), "5"
    Check "format-function: d/m/yyyy", Format(#3/5/2001#, "d/m/yyyy"), "5/3/2001"
    Check "format-function: dd/mm/yyyy", Format(#3/5/2001#, "dd/mm/yyyy"), "05/03/2001"
    Check "format-function: yy", Format(#3/5/2005#, "yy"), "05"
    Check "format-function: w", Format(d, "w"), "4"
    Check "format-function: w vbMonday", Format(d, "w", vbMonday), "3"
    Check "format-function: ww", Format(d, "ww"), "5"
    Check "format-function: ww jan1", Format(#1/1/2000#, "ww"), "1"
    Check "format-function: ww 54", Format(#12/31/2000#, "ww"), "54"
    Check "format-function: q", Format(d, "q"), "1"
    Check "format-function: q4", Format(#11/2/2010#, "q"), "4"
    Check "format-function: y", Format(d, "y"), "27"
    Check "format-function: y leap", Format(#12/31/2000#, "y"), "366"
    Check "format-function: mmm-yy", Format(d, "mmm-yy"), "Jan-93"
    Check "format-function: escaped literal", Format(d, "\d\a\y\: d"), "day: 27"
    Check "format-function: quoted literal date", Format(d, """Day"" d"), "Day 27"
    Check "format-function: c date", Format(d, "c"), "1/27/1993"
    Check "format-function: c datetime", Format(#1/27/1993 5:04:23 PM#, "c"), "1/27/1993 5:04:23 PM"
    Check "format-function: number as date", Format(36526, "yyyy-mm-dd"), "2000-01-01"
    Check "format-function: date string", Format("1/15/2000", "yyyy"), "2000"
End Sub

Private Sub TestUserTime()
    Dim t As Date, am As Date
    t = #5:04:23 PM#
    am = #9:05:07 AM#
    Check "format-function: h:nn", Format(t, "h:nn"), "17:04"
    Check "format-function: hh:mm minutes", Format(t, "hh:mm"), "17:04"
    Check "format-function: h:mm a/p", Format(t, "h:mm a/p"), "5:04 p"
    Check "format-function: h:mm A/P", Format(t, "h:mm A/P"), "5:04 P"
    Check "format-function: h:n:s", Format(am, "h:n:s"), "9:5:7"
    Check "format-function: hh:nn:ss", Format(am, "hh:nn:ss"), "09:05:07"
    Check "format-function: h AM/PM", Format(am, "h AM/PM"), "9 AM"
    Check "format-function: midnight AM/PM", Format(#12:00:00 AM#, "h AM/PM"), "12 AM"
    Check "format-function: noon AM/PM", Format(#12:30:00 PM#, "h:nn AM/PM"), "12:30 PM"
    Check "format-function: ttttt", Format(t, "ttttt"), "5:04:23 PM"
    Check "format-function: ttttt am", Format(am, "ttttt"), "9:05:07 AM"
    Check "format-function: fraction as time", Format(0.75, "hh:mm"), "18:00"
    Check "format-function: AMPM lower", Format(am, "h:nn ampm"), "9:05 AM"
End Sub

Private Sub TestFormatMisc()
    Check "format-function: FirstDayOfWeek ww", Format(#1/3/1993#, "ww", vbMonday), "1"
    Check "format-function: Format$ string", Format$("abc", ">"), "ABC"
    Check "format-function: Format$ number", Format$(1234.5, "#,##0.0"), "1,234.5"
    CheckTyped "format-function: Format$ no fmt", Format$(42), "42"
End Sub

Private Sub TestFormatErrors()
    Dim s As String
    On Error Resume Next
    s = Format$(Null)
    CheckErr "format-function: Format$(Null)", 94
    s = Format(#1/27/1993#, "w", 8)
    CheckErr "format-function: bad FirstDayOfWeek", 5
    On Error GoTo 0
End Sub

Private Sub TestFormatNumber()
    CheckTyped "formatnumber-function: default", FormatNumber(1234.567), "1,234.57"
    Check "formatnumber-function: 1 digit", FormatNumber(1234.567, 1), "1,234.6"
    Check "formatnumber-function: 0 digits", FormatNumber(1234.567, 0), "1,235"
    Check "formatnumber-function: 4 digits", FormatNumber(1.5, 4), "1.5000"
    Check "formatnumber-function: parens", FormatNumber(-1234.567, 2, , vbTrue), "(1,234.57)"
    Check "formatnumber-function: no parens", FormatNumber(-1234.567, 2, , vbFalse), "-1,234.57"
    Check "formatnumber-function: default neg", FormatNumber(-1234.567), "-1,234.57"
    Check "formatnumber-function: no group", FormatNumber(1234.567, 2, , , vbFalse), "1234.57"
    Check "formatnumber-function: leading digit", FormatNumber(0.5), "0.50"
    Check "formatnumber-function: no leading digit", FormatNumber(0.5, 2, vbFalse), ".50"
    Check "formatnumber-function: no leading digit neg", FormatNumber(-0.5, 2, vbFalse), "-.50"
    Check "formatnumber-function: zero", FormatNumber(0), "0.00"
    Check "formatnumber-function: string arg", FormatNumber("1234.5"), "1,234.50"
    Check "formatnumber-function: round half", FormatNumber(2.5, 0), "3"
    Check "formatnumber-function: round .125", FormatNumber(0.125, 2), "0.13"
    Check "formatnumber-function: big", FormatNumber(1234567890, 0), "1,234,567,890"
    Check "formatnumber-function: vbUseDefault", FormatNumber(1234.5, 2, vbUseDefault, vbUseDefault, vbUseDefault), "1,234.50"
    Check "formatnumber-function: Boolean", FormatNumber(True, 0), "-1"
    On Error Resume Next
    Dim s As String
    s = FormatNumber("abc")
    CheckErr "formatnumber-function: non-numeric", 13
    On Error GoTo 0
End Sub

Private Sub TestFormatCurrency()
    CheckTyped "formatcurrency-function: default", FormatCurrency(1234.5), "$1,234.50"
    Check "formatcurrency-function: neg", FormatCurrency(-1234.5), "($1,234.50)"
    Check "formatcurrency-function: neg no parens", FormatCurrency(-1234.5, , , vbFalse), "-$1,234.50"
    Check "formatcurrency-function: 0 digits", FormatCurrency(1234.5, 0), "$1,235"
    Check "formatcurrency-function: no leading", FormatCurrency(0.5, , vbFalse), "$.50"
    Check "formatcurrency-function: no group", FormatCurrency(1234.5, 2, , , vbFalse), "$1234.50"
    Check "formatcurrency-function: zero", FormatCurrency(0), "$0.00"
    Check "formatcurrency-function: Currency type", FormatCurrency(CCur(99.995)), "$100.00"
    Check "formatcurrency-function: string arg", FormatCurrency("12"), "$12.00"
    On Error Resume Next
    Dim s As String
    s = FormatCurrency("abc")
    CheckErr "formatcurrency-function: non-numeric", 13
    On Error GoTo 0
End Sub

Private Sub TestFormatPercent()
    CheckTyped "formatpercent-function: default", FormatPercent(0.25), "25.00%"
    Check "formatpercent-function: 1 digit", FormatPercent(0.12345, 1), "12.3%"
    Check "formatpercent-function: parens", FormatPercent(-0.5, 0, , vbTrue), "(50%)"
    Check "formatpercent-function: neg", FormatPercent(-0.5, 0), "-50%"
    Check "formatpercent-function: grouped", FormatPercent(12.5), "1,250.00%"
    Check "formatpercent-function: ungrouped", FormatPercent(12.5, 0, , , vbFalse), "1250%"
    Check "formatpercent-function: no leading", FormatPercent(0.001, 2, vbFalse), ".10%"
    Check "formatpercent-function: leading", FormatPercent(0.001, 2), "0.10%"
    Check "formatpercent-function: one", FormatPercent(1), "100.00%"
    Check "formatpercent-function: zero", FormatPercent(0), "0.00%"
End Sub

Private Sub TestFormatDateTime()
    Dim dt As Date
    dt = #1/27/1993 5:04:23 PM#
    CheckTyped "formatdatetime-function: default", FormatDateTime(dt), "1/27/1993 5:04:23 PM"
    Check "formatdatetime-function: vbGeneralDate", FormatDateTime(dt, vbGeneralDate), "1/27/1993 5:04:23 PM"
    Check "formatdatetime-function: vbLongDate", FormatDateTime(dt, vbLongDate), "Wednesday, January 27, 1993"
    Check "formatdatetime-function: vbShortDate", FormatDateTime(dt, vbShortDate), "1/27/1993"
    Check "formatdatetime-function: vbLongTime", FormatDateTime(dt, vbLongTime), "5:04:23 PM"
    Check "formatdatetime-function: vbShortTime", FormatDateTime(dt, vbShortTime), "17:04"
    Check "formatdatetime-function: vbShortTime am", FormatDateTime(#9:05:07 AM#, vbShortTime), "09:05"
    Check "formatdatetime-function: date only", FormatDateTime(#1/27/1993#), "1/27/1993"
    Check "formatdatetime-function: time only", FormatDateTime(#5:04:23 PM#), "5:04:23 PM"
    Check "formatdatetime-function: constants", vbGeneralDate + vbLongDate + vbShortDate + vbLongTime + vbShortTime, 10
    Check "formatdatetime-function: string date", FormatDateTime("1/27/1993", vbLongDate), "Wednesday, January 27, 1993"
    On Error Resume Next
    Dim s As String
    s = FormatDateTime("abc")
    CheckErr "formatdatetime-function: non-date", 13
    s = FormatDateTime(dt, 5)
    CheckErr "formatdatetime-function: bad named format", 5
    On Error GoTo 0
End Sub
