Attribute VB_Name = "conversion"
' Type conversion and inspection: C* conversion functions, Val, Str, Hex,
' Oct, Is* inspection functions, CVErr, VarType, TypeName, IIf, Choose,
' Switch, implicit Variant coercion and literal type suffixes.
' Doc examples come from MicrosoftDocs/VBA-Docs (MIT licensed code samples).
Option Explicit

Sub Main()
    TestDocsConversionFunctions
    TestBankersRounding
    TestConversionOverflow
    TestStringToNumber
    TestBooleanConversions
    TestEmptyNullConversions
    TestCStr
    TestCDate
    TestCDecCLngLng
    TestImplicitCoercion
    TestAssignmentCoercion
    TestLiteralSuffixes
    TestVal
    TestStr
    TestHexOct
    TestIsNumeric
    TestIsDate
    TestIsEmptyIsNull
    TestIsArrayIsObject
    TestErrorValues
    TestIsMissing
    TestVarType
    TestTypeName
    TestIIf
    TestChoose
    TestSwitch
End Sub

' ---------------------------------------------------------------- helpers

Private Function UserFunction() As Variant
    UserFunction = CVErr(32767)
End Function

Private Function CalculateDouble(Number)
    If IsNumeric(Number) Then
        CalculateDouble = Number * 2
    Else
        CalculateDouble = CVErr(2001)
    End If
End Function

Private Function ReturnTwice(Optional A)
    If IsMissing(A) Then
        ReturnTwice = Null
    Else
        ReturnTwice = A * 2
    End If
End Function

Private Function MissingWithDefault(Optional A As Variant = 5) As Boolean
    MissingWithDefault = IsMissing(A)
End Function

Private Function MissingTyped(Optional A As Integer) As Boolean
    MissingTyped = IsMissing(A)
End Function

Private Function TypedDefault(Optional A As Integer) As Integer
    TypedDefault = A
End Function

Private Function CheckIt(TestMe As Integer)
    CheckIt = IIf(TestMe > 1000, "Large", "Small")
End Function

Private Function GetChoice(Ind As Integer)
    GetChoice = Choose(Ind, "Speedy", "United", "Federal")
End Function

Private Function MatchUp(CityName As String)
    MatchUp = Switch(CityName = "London", "English", CityName _
                    = "Rome", "Italian", CityName = "Paris", "French")
End Function

' ---------------------------------------------------------------- tests

Private Sub TestDocsConversionFunctions()
    Dim A, B, MyCheck
    A = 5: B = 5
    MyCheck = CBool(A = B)
    CheckTyped "type-conversion-functions: CBool(A = B)", MyCheck, True
    A = 0
    MyCheck = CBool(A)
    CheckTyped "type-conversion-functions: CBool(0)", MyCheck, False

    Dim MyDouble, MyByte
    MyDouble = 125.5678
    MyByte = CByte(MyDouble)
    CheckTyped "type-conversion-functions: CByte(125.5678)", MyByte, CByte(126)

    Dim MyCurr
    MyDouble = 543.214588
    MyCurr = CCur(MyDouble * 2)
    CheckTyped "type-conversion-functions: CCur(MyDouble * 2)", MyCurr, 1086.4292@

    Dim MyDate, MyShortDate, MyTime, MyShortTime
    MyDate = "February 12, 1969"
    MyShortDate = CDate(MyDate)
    CheckTyped "type-conversion-functions: CDate(""February 12, 1969"")", MyShortDate, #2/12/1969#
    MyTime = "4:35:47 PM"
    MyShortTime = CDate(MyTime)
    CheckTyped "type-conversion-functions: CDate(""4:35:47 PM"")", MyShortTime, #4:35:47 PM#

    MyCurr = CCur(234.456784)
    CheckTyped "type-conversion-functions: CCur(234.456784)", MyCurr, 234.4568@
    MyDouble = CDbl(MyCurr * 8.2 * 0.01)
    CheckTyped "type-conversion-functions: CDbl(MyCurr * 8.2 * 0.01)", MyDouble, 234.4568 * 8.2 * 0.01

    Dim MyDecimal
    MyCurr = 10000000.0587
    MyDecimal = CDec(MyCurr)
    Check "type-conversion-functions: CDec value", MyDecimal, 10000000.0587
    Check "type-conversion-functions: CDec TypeName", TypeName(MyDecimal), "Decimal"

    Dim MyInt
    MyDouble = 2345.5678
    MyInt = CInt(MyDouble)
    CheckTyped "type-conversion-functions: CInt(2345.5678)", MyInt, 2346

    Dim MyVal1, MyVal2, MyLong1, MyLong2
    MyVal1 = 25427.45: MyVal2 = 25427.55
    MyLong1 = CLng(MyVal1)
    MyLong2 = CLng(MyVal2)
    CheckTyped "type-conversion-functions: CLng(25427.45)", MyLong1, 25427&
    CheckTyped "type-conversion-functions: CLng(25427.55)", MyLong2, 25428&

    Dim MyDouble1, MyDouble2, MySingle1, MySingle2
    MyDouble1 = 75.3421115: MyDouble2 = 75.3421555
    MySingle1 = CSng(MyDouble1)
    MySingle2 = CSng(MyDouble2)
    Check "type-conversion-functions: CSng(75.3421115)", CStr(MySingle1), "75.34211"
    Check "type-conversion-functions: CSng(75.3421555)", CStr(MySingle2), "75.34216"
    Check "type-conversion-functions: CSng TypeName", TypeName(MySingle1), "Single"

    Dim MyString
    MyDouble = 437.324
    MyString = CStr(MyDouble)
    CheckTyped "type-conversion-functions: CStr(437.324)", MyString, "437.324"

    Dim MyVar
    MyInt = 4534
    MyVar = CVar(MyInt & "000")
    CheckTyped "type-conversion-functions: CVar(MyInt & ""000"")", MyVar, "4534000"
    CheckTyped "type-conversion-functions: CVar(5)", CVar(5), 5
End Sub

Private Sub TestBankersRounding()
    ' "0.5 rounds to 0, and 1.5 rounds to 2"
    CheckTyped "type-conversion-functions: CInt(0.5)", CInt(0.5), 0
    CheckTyped "type-conversion-functions: CInt(1.5)", CInt(1.5), 2
    CheckTyped "CInt(2.5)", CInt(2.5), 2
    CheckTyped "CInt(3.5)", CInt(3.5), 4
    CheckTyped "CInt(-2.5)", CInt(-2.5), -2
    CheckTyped "CInt(-3.5)", CInt(-3.5), -4
    CheckTyped "CInt(2.51)", CInt(2.51), 3
    CheckTyped "CInt(-2.6)", CInt(-2.6), -3
    CheckTyped "CLng(2.5)", CLng(2.5), 2&
    CheckTyped "CLng(3.5)", CLng(3.5), 4&
    CheckTyped "CLng(-0.5)", CLng(-0.5), 0&
    CheckTyped "CByte(2.5)", CByte(2.5), CByte(2)
    CheckTyped "CByte(3.5)", CByte(3.5), CByte(4)
    CheckTyped "CInt(""2.5"")", CInt("2.5"), 2
    CheckTyped "CLng(""3.5"")", CLng("3.5"), 4&
    CheckTyped "CLng(CCur(2.5))", CLng(CCur(2.5)), 2&
    CheckTyped "CInt(1.5!)", CInt(1.5!), 2
    CheckTyped "CCur(1.23456)", CCur(1.23456), 1.2346@
End Sub

Private Sub TestConversionOverflow()
    Dim v As Variant
    On Error Resume Next
    v = CInt(32767)
    CheckErr "CInt(32767)", 0
    v = CInt(32768)
    CheckErr "CInt(32768)", 6
    v = CInt(32767.5)
    CheckErr "CInt(32767.5)", 6
    v = CInt(-32768)
    CheckErr "CInt(-32768)", 0
    v = CInt(-32769)
    CheckErr "CInt(-32769)", 6
    v = CByte(255)
    CheckErr "CByte(255)", 0
    v = CByte(256)
    CheckErr "CByte(256)", 6
    v = CByte(-1)
    CheckErr "CByte(-1)", 6
    v = CByte(255.5)
    CheckErr "CByte(255.5)", 6
    v = CLng(2147483647)
    CheckErr "CLng(2147483647)", 0
    v = CLng(2147483648#)
    CheckErr "CLng(2147483648#)", 6
    v = CLng(-2147483648#)
    CheckErr "CLng(-2147483648#)", 0
    v = CLng(-2147483649#)
    CheckErr "CLng(-2147483649#)", 6
    v = CSng(1E+300)
    CheckErr "CSng(1E+300)", 6
    v = CCur(1E+20)
    CheckErr "CCur(1E+20)", 6
    v = CInt("40000")
    CheckErr "CInt(""40000"")", 6
    v = CByte(True)
    CheckErr "CByte(True)", 6
    v = CInt(CLng(70000))
    CheckErr "CInt(CLng(70000))", 6
    On Error GoTo 0
End Sub

Private Sub TestStringToNumber()
    CheckTyped "CInt(""&HFF"")", CInt("&HFF"), 255
    CheckTyped "CLng(""&HFF"")", CLng("&HFF"), 255&
    CheckTyped "CDbl(""&H10"")", CDbl("&H10"), 16#
    CheckTyped "CInt(""&O17"")", CInt("&O17"), 15
    CheckTyped "CInt("" 12 "")", CInt(" 12 "), 12
    CheckTyped "CLng("" 12 "")", CLng(" 12 "), 12&
    CheckTyped "CDbl(""1e3"")", CDbl("1e3"), 1000#
    CheckTyped "CLng(""1e3"")", CLng("1e3"), 1000&
    CheckTyped "CDbl(""1.5"")", CDbl("1.5"), 1.5
    CheckTyped "CDbl(""  -3.25  "")", CDbl("  -3.25  "), -3.25
    CheckTyped "CInt(""-7"")", CInt("-7"), -7
    CheckTyped "CDbl(""1,234.5"")", CDbl("1,234.5"), 1234.5
    CheckTyped "CCur(""1.23456"")", CCur("1.23456"), 1.2346@
    CheckTyped "CSng(""0.5"")", CSng("0.5"), 0.5!
    CheckTyped "CDbl("".5"")", CDbl(".5"), 0.5

    Dim v As Variant
    On Error Resume Next
    v = CInt("abc")
    CheckErr "CInt(""abc"")", 13
    v = CDbl("")
    CheckErr "CDbl("""")", 13
    v = CLng("12abc")
    CheckErr "CLng(""12abc"")", 13
    v = CDbl("1.2.3")
    CheckErr "CDbl(""1.2.3"")", 13
    On Error GoTo 0
End Sub

Private Sub TestBooleanConversions()
    CheckTyped "CInt(True)", CInt(True), -1
    CheckTyped "CLng(True)", CLng(True), -1&
    CheckTyped "CLng(False)", CLng(False), 0&
    CheckTyped "CDbl(True)", CDbl(True), -1#
    CheckTyped "CCur(True)", CCur(True), -1@
    CheckTyped "CStr(True)", CStr(True), "True"
    CheckTyped "CStr(False)", CStr(False), "False"
    CheckTyped "True + 0", True + 0, -1
    CheckTyped "True + True", True + True, -2
    CheckTyped "True * 5", True * 5, -5
    CheckTyped "CBool(42)", CBool(42), True
    CheckTyped "CBool(-1)", CBool(-1), True
    CheckTyped "CBool(0.1)", CBool(0.1), True
    CheckTyped "CBool(0#)", CBool(0#), False
    CheckTyped "CBool(""True"")", CBool("True"), True
    CheckTyped "CBool(""false"")", CBool("false"), False
    CheckTyped "CBool(""1"")", CBool("1"), True
    CheckTyped "CBool(""0"")", CBool("0"), False
    CheckTyped "CBool(""-2.5"")", CBool("-2.5"), True
    CheckTyped "CStr(CBool(2))", CStr(CBool(2)), "True"
    CheckTyped "CBool(Empty)", CBool(Empty), False

    Dim v As Variant
    On Error Resume Next
    v = CBool("abc")
    CheckErr "CBool(""abc"")", 13
    v = CBool("yes")
    CheckErr "CBool(""yes"")", 13
    On Error GoTo 0
End Sub

Private Sub TestEmptyNullConversions()
    Dim e As Variant
    CheckTyped "CInt(Empty)", CInt(e), 0
    CheckTyped "CLng(Empty)", CLng(e), 0&
    CheckTyped "CDbl(Empty)", CDbl(e), 0#
    CheckTyped "CStr(Empty)", CStr(e), ""
    CheckTyped "CDate(Empty)", CDate(e), #12:00:00 AM#
    Check "CVar(Null)", IsNull(CVar(Null)), True
    Check "IsEmpty(CVar(Empty))", IsEmpty(CVar(e)), True

    Dim v As Variant
    On Error Resume Next
    v = CInt(Null)
    CheckErr "CInt(Null)", 94
    v = CLng(Null)
    CheckErr "CLng(Null)", 94
    v = CStr(Null)
    CheckErr "CStr(Null)", 94
    v = CBool(Null)
    CheckErr "CBool(Null)", 94
    v = CDbl(Null)
    CheckErr "CDbl(Null)", 94
    v = CDate(Null)
    CheckErr "CDate(Null)", 94
    On Error GoTo 0
End Sub

Private Sub TestCStr()
    CheckTyped "CStr(1.5)", CStr(1.5), "1.5"
    CheckTyped "CStr(-0.5)", CStr(-0.5), "-0.5"
    CheckTyped "CStr(42)", CStr(42), "42"
    CheckTyped "CStr(-42&)", CStr(-42&), "-42"
    CheckTyped "CStr(1E+20)", CStr(1E+20), "1E+20"
    CheckTyped "CStr(1E+15)", CStr(1E+15), "1E+15"
    CheckTyped "CStr(1E+14)", CStr(1E+14), "100000000000000"
    CheckTyped "CStr(123456789012345#)", CStr(123456789012345#), "123456789012345"
    CheckTyped "CStr(1 / 3)", CStr(1 / 3), "0.333333333333333"
    CheckTyped "CStr(CSng(1 / 3))", CStr(CSng(1 / 3)), "0.3333333"
    CheckTyped "CStr(0.0001)", CStr(0.0001), "0.0001"
    CheckTyped "CStr(CCur(1.5))", CStr(CCur(1.5)), "1.5"
    CheckTyped "CStr(CCur(10))", CStr(CCur(10)), "10"
    CheckTyped "CStr(2.5#)", CStr(2.5#), "2.5"
    CheckTyped "CStr(#1/2/2003#)", CStr(#1/2/2003#), "1/2/2003"
    CheckTyped "CStr(#12:00:00 PM#)", CStr(#12:00:00 PM#), "12:00:00 PM"
    CheckTyped "CStr(#1/2/2003 3:04:05 PM#)", CStr(#1/2/2003 3:04:05 PM#), "1/2/2003 3:04:05 PM"
    CheckTyped "CStr(CDate(0))", CStr(CDate(0)), "12:00:00 AM"
    CheckTyped "CStr(CByte(7))", CStr(CByte(7)), "7"
End Sub

Private Sub TestCDate()
    CheckTyped "CDate(0)", CDate(0), #12:00:00 AM#
    CheckTyped "CDate(36526)", CDate(36526), #1/1/2000#
    CheckTyped "CDate(36526.5)", CDate(36526.5), #1/1/2000 12:00:00 PM#
    CheckTyped "CDate(-1)", CDate(-1), #12/29/1899#
    CheckTyped "CDate(""1/2/2003"")", CDate("1/2/2003"), #1/2/2003#
    CheckTyped "CDate(""2003-01-02"")", CDate("2003-01-02"), #1/2/2003#
    CheckTyped "CDbl(#1/1/2000#)", CDbl(#1/1/2000#), 36526#
    CheckTyped "CDbl(#6:00:00 AM#)", CDbl(#6:00:00 AM#), 0.25
    CheckTyped "CLng(#1/1/2000 6:00:00 PM#)", CLng(#1/1/2000 6:00:00 PM#), 36527&
    Dim v As Variant
    On Error Resume Next
    v = CDate("abc")
    CheckErr "CDate(""abc"")", 13
    v = CDate("13/32/2014")
    CheckErr "CDate(""13/32/2014"")", 13
    On Error GoTo 0
End Sub

Private Sub TestCDecCLngLng()
    Check "CDec TypeName", TypeName(CDec(1)), "Decimal"
    Check "CDec(""1.1"") + CDec(""2.2"")", CDec("1.1") + CDec("2.2"), CDec("3.3")
    Check "CDec + Integer TypeName", TypeName(CDec(1) + 1), "Decimal"
    Check "CStr(CDec(""12345678901234567890""))", CStr(CDec("12345678901234567890")), "12345678901234567890"
    Check "CStr(CDec(""0.1"") * 3)", CStr(CDec("0.1") * 3), "0.3"
    Check "CLngLng TypeName", TypeName(CLngLng(1)), "LongLong"
    Check "CLngLng(2.5)", CLngLng(2.5), 2
    Check "CLngLng(3.5)", CLngLng(3.5), 4
    Check "CStr(CLngLng(2147483648#))", CStr(CLngLng(2147483648#)), "2147483648"
End Sub

Private Sub TestImplicitCoercion()
    CheckTyped """1"" + 2", "1" + 2, 3#
    CheckTyped """1"" & 2", "1" & 2, "12"
    CheckTyped """1"" + ""2""", "1" + "2", "12"
    CheckTyped """3"" * ""4""", "3" * "4", 12#
    CheckTyped """10"" / 4", "10" / 4, 2.5
    CheckTyped "1 & 2", 1 & 2, "12"
    CheckTyped "1.5 & """"", 1.5 & "", "1.5"
    CheckTyped "True & """"", True & "", "True"
    Dim e As Variant
    CheckTyped "Empty + 1", e + 1, 1
    CheckTyped "Empty & ""a""", e & "a", "a"
    CheckTyped "Empty * 5", e * 5, 0
    Check "Empty = 0", e = 0, True
    Check "Empty = """"", e = "", True
    Check "Null + 1", IsNull(Null + 1), True
    Check "Null * 0", IsNull(Null * 0), True
    CheckTyped "Null & ""a""", Null & "a", "a"
    Dim s As Variant
    s = "5"
    CheckTyped "Variant ""5"" + 1", s + 1, 6#
    CheckTyped "Variant ""5"" - 1", s - 1, 4#
    CheckTyped "CInt(""5"") + 1", CInt(s) + 1, 6
    CheckTyped "200& * 200", 200& * 200, 40000&
    CheckTyped "Integer + Long", 1 + 1&, 2&
    CheckTyped "Integer + Double", 1 + 1.5, 2.5
    CheckTyped "Integer + Currency", 1 + 1@, 2@
    CheckTyped "Integer / Integer", 4 / 2, 2#
    CheckTyped "Integer \ Integer", 7 \ 2, 3

    Dim v As Variant
    On Error Resume Next
    v = "a" + 1
    CheckErr """a"" + 1", 13
    v = 32767 + 1
    CheckErr "32767 + 1 (Integer overflow)", 6
    v = 200 * 200
    CheckErr "200 * 200 (Integer overflow)", 6
    On Error GoTo 0
End Sub

Private Sub TestAssignmentCoercion()
    Dim i As Integer, l As Long, b As Boolean, s As String
    Dim by As Byte, d As Date, dbl As Double
    i = "12"
    CheckTyped "Integer = ""12""", i, 12
    i = 2.5
    CheckTyped "Integer = 2.5", i, 2
    i = 3.5
    CheckTyped "Integer = 3.5", i, 4
    i = True
    CheckTyped "Integer = True", i, -1
    l = " 7 "
    CheckTyped "Long = "" 7 """, l, 7&
    b = 5
    CheckTyped "Boolean = 5", b, True
    b = "True"
    CheckTyped "Boolean = ""True""", b, True
    b = 0
    CheckTyped "Boolean = 0", b, False
    s = 1.5
    CheckTyped "String = 1.5", s, "1.5"
    s = True
    CheckTyped "String = True", s, "True"
    s = 42
    CheckTyped "String = 42", s, "42"
    d = "1/2/2003"
    CheckTyped "Date = ""1/2/2003""", d, #1/2/2003#
    d = 36526
    CheckTyped "Date = 36526", d, #1/1/2000#
    dbl = "1e3"
    CheckTyped "Double = ""1e3""", dbl, 1000#
    by = 255
    CheckTyped "Byte = 255", by, CByte(255)

    On Error Resume Next
    i = "abc"
    CheckErr "Integer = ""abc""", 13
    i = 40000
    CheckErr "Integer = 40000", 6
    by = 256
    CheckErr "Byte = 256", 6
    by = -1
    CheckErr "Byte = -1", 6
    l = Null
    CheckErr "Long = Null", 94
    s = Null
    CheckErr "String = Null", 94
    b = "abc"
    CheckErr "Boolean = ""abc""", 13
    On Error GoTo 0
End Sub

Private Sub TestLiteralSuffixes()
    Check "TypeName(12%)", TypeName(12%), "Integer"
    Check "TypeName(12&)", TypeName(12&), "Long"
    Check "TypeName(1.5!)", TypeName(1.5!), "Single"
    Check "TypeName(1.5#)", TypeName(1.5#), "Double"
    Check "TypeName(1@)", TypeName(1@), "Currency"
    Check "TypeName(1.5)", TypeName(1.5), "Double"
    Check "TypeName(12)", TypeName(12), "Integer"
    Check "TypeName(32767)", TypeName(32767), "Integer"
    Check "TypeName(32768)", TypeName(32768), "Long"
    Check "TypeName(-32768)", TypeName(-32768), "Long"
    Check "TypeName(2147483647)", TypeName(2147483647), "Long"
    Check "TypeName(2147483648)", TypeName(2147483648#), "Double"
    Check "TypeName(1E3)", TypeName(1E3), "Double"
    Check "TypeName(&HFFFF)", TypeName(&HFFFF), "Integer"
    CheckTyped "&HFFFF", &HFFFF, -1
    Check "TypeName(&HFFFF&)", TypeName(&HFFFF&), "Long"
    CheckTyped "&HFFFF&", &HFFFF&, 65535&
    CheckTyped "&HFF", &HFF, 255
    CheckTyped "&H8000", &H8000, CInt(-32768)
    CheckTyped "&H10000", &H10000, 65536&
    CheckTyped "&HFFFFFFFF", &HFFFFFFFF, -1&
    CheckTyped "&H7FFFFFFF", &H7FFFFFFF, 2147483647
    CheckTyped "oct-function: &O10", &O10, 8
    CheckTyped "&O177777", &O177777, -1
    Check "TypeName(""x"")", TypeName("x"), "String"
    Check "TypeName(#1/1/2000#)", TypeName(#1/1/2000#), "Date"
    Check "TypeName(True)", TypeName(True), "Boolean"
    Check "TypeName(12^)", TypeName(12^), "LongLong"
End Sub

Private Sub TestVal()
    Dim MyValue
    MyValue = Val("2457")
    CheckTyped "val-function: Val(""2457"")", MyValue, 2457#
    MyValue = Val(" 2 45 7")
    CheckTyped "val-function: Val("" 2 45 7"")", MyValue, 2457#
    MyValue = Val("24 and 57")
    CheckTyped "val-function: Val(""24 and 57"")", MyValue, 24#
    Check "val-function: street", Val("    1615 198th Street N.E."), 1615198
    Check "val-function: Val(""&HFFFF"")", Val("&HFFFF"), -1
    Check "val-function: Val(""50%"")", Val("50%"), 50
    Check "Val(""&HFF"")", Val("&HFF"), 255
    Check "Val(""&O17"")", Val("&O17"), 15
    Check "Val(""1e3"")", Val("1e3"), 1000
    Check "Val(""1.5.6"")", Val("1.5.6"), 1.5
    Check "Val("""")", Val(""), 0
    Check "Val(""abc"")", Val("abc"), 0
    Check "Val(""-12.5x"")", Val("-12.5x"), -12.5
    Check "Val("".5"")", Val(".5"), 0.5
    Check "Val(""1,234"")", Val("1,234"), 1
    Check "Val(""$12"")", Val("$12"), 0
    Check "Val(vbTab & ""42"")", Val(vbTab & "42"), 42
    Check "TypeName(Val(""1""))", TypeName(Val("1")), "Double"

    Dim v As Variant
    On Error Resume Next
    v = Val("50.5%")
    CheckErr "val-function: Val(""50.5%"")", 13
    v = Val(Null)
    CheckErr "Val(Null)", 94
    On Error GoTo 0
End Sub

Private Sub TestStr()
    Dim MyString
    MyString = Str(459)
    CheckTyped "str-function: Str(459)", MyString, " 459"
    MyString = Str(-459.65)
    CheckTyped "str-function: Str(-459.65)", MyString, "-459.65"
    MyString = Str(459.001)
    CheckTyped "str-function: Str(459.001)", MyString, " 459.001"
    CheckTyped "Str(0)", Str(0), " 0"
    CheckTyped "Str(1.5)", Str(1.5), " 1.5"
    CheckTyped "Str(-7&)", Str(-7&), "-7"
    Check "Str(Null)", IsNull(Str(Null)), True
End Sub

Private Sub TestHexOct()
    Dim MyHex
    MyHex = Hex(5)
    CheckTyped "hex-function: Hex(5)", MyHex, "5"
    MyHex = Hex(10)
    CheckTyped "hex-function: Hex(10)", MyHex, "A"
    MyHex = Hex(459)
    CheckTyped "hex-function: Hex(459)", MyHex, "1CB"
    CheckTyped "hex-function: Hex(255)", Hex(255), "FF"
    CheckTyped "hex-function: &HFF", &HFF, 255
    CheckTyped "Hex(-1)", Hex(-1), "FFFF"
    CheckTyped "Hex(-1&)", Hex(-1&), "FFFFFFFF"
    CheckTyped "Hex(-32768)", Hex(-32768), "FFFF8000"
    CheckTyped "Hex(-32768%)", Hex(CInt(-32768)), "8000"
    CheckTyped "Hex(65535)", Hex(65535), "FFFF"
    CheckTyped "Hex(65536)", Hex(65536), "10000"
    CheckTyped "Hex(2.5)", Hex(2.5), "2"
    CheckTyped "Hex(3.5)", Hex(3.5), "4"
    CheckTyped "Hex(0)", Hex(0), "0"
    CheckTyped "Hex(""255"")", Hex("255"), "FF"
    Dim e As Variant
    CheckTyped "Hex(Empty)", Hex(e), "0"
    Check "Hex(Null)", IsNull(Hex(Null)), True

    Dim MyOct
    MyOct = Oct(4)
    CheckTyped "oct-function: Oct(4)", MyOct, "4"
    MyOct = Oct(8)
    CheckTyped "oct-function: Oct(8)", MyOct, "10"
    MyOct = Oct(459)
    CheckTyped "oct-function: Oct(459)", MyOct, "713"
    CheckTyped "oct-function: Oct(Empty)", Oct(e), "0"
    Check "oct-function: Oct(Null)", IsNull(Oct(Null)), True
    CheckTyped "Oct(-1)", Oct(-1), "177777"
    CheckTyped "Oct(-1&)", Oct(-1&), "37777777777"
    CheckTyped "Oct(7.5)", Oct(7.5), "10"
End Sub

Private Sub TestIsNumeric()
    Dim MyVar, MyCheck
    MyVar = "53"
    MyCheck = IsNumeric(MyVar)
    CheckTyped "isnumeric-function: ""53""", MyCheck, True
    MyVar = "459.95"
    MyCheck = IsNumeric(MyVar)
    CheckTyped "isnumeric-function: ""459.95""", MyCheck, True
    MyVar = "45 Help"
    MyCheck = IsNumeric(MyVar)
    CheckTyped "isnumeric-function: ""45 Help""", MyCheck, False
    Check "isnumeric-function: date", IsNumeric(#1/1/2000#), False
    Dim e As Variant
    Check "IsNumeric(Empty)", IsNumeric(e), True
    Check "IsNumeric(Null)", IsNumeric(Null), False
    Check "IsNumeric("""")", IsNumeric(""), False
    Check "IsNumeric("" 12 "")", IsNumeric(" 12 "), True
    Check "IsNumeric(""1e3"")", IsNumeric("1e3"), True
    Check "IsNumeric(""&HFF"")", IsNumeric("&HFF"), True
    Check "IsNumeric(""-5"")", IsNumeric("-5"), True
    Check "IsNumeric(""1,234"")", IsNumeric("1,234"), True
    Check "IsNumeric(""1.2.3"")", IsNumeric("1.2.3"), False
    Check "IsNumeric(""abc"")", IsNumeric("abc"), False
    Check "IsNumeric(5)", IsNumeric(5), True
    Check "IsNumeric(1.5)", IsNumeric(1.5), True
    Check "IsNumeric(True)", IsNumeric(True), True
End Sub

Private Sub TestIsDate()
    Dim MyVar, MyCheck
    MyVar = "04/28/2014"
    MyCheck = IsDate(MyVar)
    CheckTyped "isdate-function: ""04/28/2014""", MyCheck, True
    MyVar = "April 28, 2014"
    MyCheck = IsDate(MyVar)
    CheckTyped "isdate-function: ""April 28, 2014""", MyCheck, True
    MyVar = "13/32/2014"
    MyCheck = IsDate(MyVar)
    CheckTyped "isdate-function: ""13/32/2014""", MyCheck, False
    MyVar = "04.28.14"
    MyCheck = IsDate(MyVar)
    CheckTyped "isdate-function: ""04.28.14""", MyCheck, True
    MyVar = "04.28.2014"
    MyCheck = IsDate(MyVar)
    CheckTyped "isdate-function: ""04.28.2014""", MyCheck, False
    Check "IsDate(#1/1/2000#)", IsDate(#1/1/2000#), True
    Check "IsDate(""abc"")", IsDate("abc"), False
    Check "IsDate(""12:30"")", IsDate("12:30"), True
    Check "IsDate(""2/30/2000"")", IsDate("2/30/2000"), False
    Check "IsDate(Null)", IsDate(Null), False
    Check "IsDate(36526)", IsDate(36526), False
End Sub

Private Sub TestIsEmptyIsNull()
    Dim MyVar, MyCheck
    MyCheck = IsEmpty(MyVar)
    CheckTyped "isempty-function: uninitialized", MyCheck, True
    MyVar = Null
    MyCheck = IsEmpty(MyVar)
    CheckTyped "isempty-function: Null", MyCheck, False
    MyVar = Empty
    MyCheck = IsEmpty(MyVar)
    CheckTyped "isempty-function: Empty", MyCheck, True

    Dim MyVar2
    MyCheck = IsNull(MyVar2)
    CheckTyped "isnull-function: uninitialized", MyCheck, False
    MyVar2 = ""
    MyCheck = IsNull(MyVar2)
    CheckTyped "isnull-function: """"", MyCheck, False
    MyVar2 = Null
    MyCheck = IsNull(MyVar2)
    CheckTyped "isnull-function: Null", MyCheck, True

    Dim s As String, n As Long
    Check "IsEmpty(String var)", IsEmpty(s), False
    Check "IsEmpty(Long var)", IsEmpty(n), False
    Check "IsEmpty("""")", IsEmpty(""), False
    Check "IsEmpty(0)", IsEmpty(0), False
    Check "IsNull(Null + 1)", IsNull(Null + 1), True
    Check "IsNull(Null & """")", IsNull(Null & ""), False
    Check "IsNull(Empty)", IsNull(Empty), False
End Sub

Private Sub TestIsArrayIsObject()
    Dim MyArray(1 To 5) As Integer, YourArray, MyCheck
    YourArray = Array(1, 2, 3)
    MyCheck = IsArray(MyArray)
    CheckTyped "isarray-function: MyArray", MyCheck, True
    MyCheck = IsArray(YourArray)
    CheckTyped "isarray-function: YourArray", MyCheck, True
    Dim dyn() As Long
    Check "IsArray(unallocated dynamic)", IsArray(dyn), True
    Check "IsArray(5)", IsArray(5), False
    Check "IsArray(""abc"")", IsArray("abc"), False
    Check "IsArray(Empty)", IsArray(Empty), False

    Dim MyInt As Integer
    Dim YourObject, MyCheck2
    Dim MyObject As Object
    Set YourObject = MyObject
    MyCheck2 = IsObject(YourObject)
    CheckTyped "isobject-function: YourObject", MyCheck2, True
    MyCheck2 = IsObject(MyInt)
    CheckTyped "isobject-function: MyInt", MyCheck2, False
    MyCheck2 = IsObject(Nothing)
    CheckTyped "isobject-function: Nothing", MyCheck2, True
    MyCheck2 = IsObject(Empty)
    CheckTyped "isobject-function: Empty", MyCheck2, False
    MyCheck2 = IsObject(Null)
    CheckTyped "isobject-function: Null", MyCheck2, False
    Dim c As Collection
    Set c = New Collection
    Check "IsObject(Collection)", IsObject(c), True
End Sub

Private Sub TestErrorValues()
    Dim ReturnVal, MyCheck
    ReturnVal = UserFunction()
    MyCheck = IsError(ReturnVal)
    CheckTyped "iserror-function: UserFunction", MyCheck, True
    Check "cverr-function: VarType", VarType(CalculateDouble("345.45robert")), vbError
    Check "cverr-function: IsError", IsError(CalculateDouble("345.45robert")), True
    Check "cverr-function: numeric", CalculateDouble("2"), 4
    Check "cverr-function: explicit CLng", CLng(CVErr(2001)), 2001
    Check "CStr(CVErr(2042))", CStr(CVErr(2042)), "Error 2042"
    Check "TypeName(CVErr(5))", TypeName(CVErr(5)), "Error"
    Check "IsError(5)", IsError(5), False
    Check "IsError(Null)", IsError(Null), False

    Dim l As Long
    On Error Resume Next
    l = CVErr(5)
    CheckErr "cverr-function: implicit conversion to Long", 13
    On Error GoTo 0
End Sub

Private Sub TestIsMissing()
    Dim ReturnValue
    ReturnValue = ReturnTwice()
    Check "ismissing-function: ReturnTwice()", IsNull(ReturnValue), True
    ReturnValue = ReturnTwice(2)
    Check "ismissing-function: ReturnTwice(2)", ReturnValue, 4
    Check "ismissing-function: Variant with default", MissingWithDefault(), False
    Check "ismissing-function: typed Integer", MissingTyped(), False
    Check "Optional Integer default 0", TypedDefault(), 0
End Sub

Private Sub TestVarType()
    Dim IntVar, StrVar, DateVar, ArrayVar
    IntVar = 459: StrVar = "Hello World": DateVar = #2/12/1969#
    ArrayVar = Array("1st Element", "2nd Element")
    Check "vartype-function: IntVar", VarType(IntVar), 2
    Check "vartype-function: DateVar", VarType(DateVar), 7
    Check "vartype-function: StrVar", VarType(StrVar), 8
    Check "vartype-function: ArrayVar", VarType(ArrayVar), 8204
    Dim ia(1 To 3) As Integer
    Check "vartype-function: Integer array", VarType(ia), vbInteger + vbArray
    Check "vartype-function: vbArray constant", vbArray, 8192
    Dim e As Variant
    Check "VarType(Empty)", VarType(e), vbEmpty
    Check "VarType(Null)", VarType(Null), vbNull
    Check "VarType(12&)", VarType(12&), vbLong
    Check "VarType(1.5!)", VarType(1.5!), vbSingle
    Check "VarType(1.5)", VarType(1.5), vbDouble
    Check "VarType(1@)", VarType(1@), vbCurrency
    Check "VarType(Nothing)", VarType(Nothing), vbObject
    Check "VarType(CVErr(1))", VarType(CVErr(1)), vbError
    Check "VarType(True)", VarType(True), vbBoolean
    Check "VarType(CDec(1))", VarType(CDec(1)), vbDecimal
    Check "VarType(CByte(1))", VarType(CByte(1)), vbByte
    Check "VarType(CLngLng(1))", VarType(CLngLng(1)), vbLongLong
    Check "vbLongLong", vbLongLong, 20
    Check "vbDecimal", vbDecimal, 14
    Check "vbByte", vbByte, 17
    Dim sa() As String
    ReDim sa(1)
    Check "VarType(String array)", VarType(sa), vbString + vbArray
End Sub

Private Sub TestTypeName()
    Dim NullVar, MyType, StrVar As String, IntVar As Integer, CurVar As Currency
    Dim ArrayVar(1 To 5) As Integer
    NullVar = Null
    MyType = TypeName(StrVar)
    CheckTyped "typename-function: StrVar", MyType, "String"
    MyType = TypeName(IntVar)
    CheckTyped "typename-function: IntVar", MyType, "Integer"
    MyType = TypeName(CurVar)
    CheckTyped "typename-function: CurVar", MyType, "Currency"
    MyType = TypeName(NullVar)
    CheckTyped "typename-function: NullVar", MyType, "Null"
    MyType = TypeName(ArrayVar)
    CheckTyped "typename-function: ArrayVar", MyType, "Integer()"
    Dim e As Variant
    Check "TypeName(Empty)", TypeName(e), "Empty"
    Check "TypeName(Nothing)", TypeName(Nothing), "Nothing"
    Dim o As Object
    Check "TypeName(unset Object)", TypeName(o), "Nothing"
    Check "TypeName(Array())", TypeName(Array(1, 2)), "Variant()"
    Dim sa() As String
    Check "TypeName(String())", TypeName(sa), "String()"
    Dim c As New Collection
    Check "TypeName(Collection)", TypeName(c), "Collection"
    Check "TypeName(CByte(1))", TypeName(CByte(1)), "Byte"
    Check "TypeName(CDate(1))", TypeName(CDate(1)), "Date"
    Dim d As Double
    Check "TypeName(Double var)", TypeName(d), "Double"
End Sub

Private Sub TestIIf()
    Check "iif-function: CheckIt(1001)", CheckIt(1001), "Large"
    Check "iif-function: CheckIt(5)", CheckIt(5), "Small"
    CheckTyped "IIf typed", IIf(True, 1, "x"), 1
    CheckTyped "IIf(Null, ...)", IIf(Null, "a", "b"), "b"
    CheckTyped "IIf(0, ...)", IIf(0, "a", "b"), "b"
    Dim v As Variant
    On Error Resume Next
    v = IIf(True, 1, 1 / 0)
    CheckErr "iif-function: both parts evaluated", 11
    On Error GoTo 0
End Sub

Private Sub TestChoose()
    Check "choose-function: GetChoice(1)", GetChoice(1), "Speedy"
    Check "choose-function: GetChoice(2)", GetChoice(2), "United"
    Check "choose-function: GetChoice(3)", GetChoice(3), "Federal"
    Check "choose-function: index 3", Choose(3, "one", "two", "three"), "three"
    Check "choose-function: index 0", IsNull(Choose(0, "a", "b")), True
    Check "choose-function: index too large", IsNull(Choose(4, "a", "b", "c")), True
    Check "choose-function: index 2.4", Choose(2.4, "a", "b", "c"), "b"
    CheckTyped "Choose typed", Choose(2, "a", 5&), 5&
End Sub

Private Sub TestSwitch()
    Check "switch-function: London", MatchUp("London"), "English"
    Check "switch-function: Rome", MatchUp("Rome"), "Italian"
    Check "switch-function: Paris", MatchUp("Paris"), "French"
    Check "switch-function: no match", IsNull(MatchUp("Berlin")), True
    Check "switch-function: Null value", IsNull(Switch(True, Null, True, 1)), True
    Check "Switch first true wins", Switch(False, 1, True, 2, True, 3), 2
    Dim v As Variant
    On Error Resume Next
    v = Switch(True)
    CheckErr "switch-function: unpaired", 5
    On Error GoTo 0
End Sub
