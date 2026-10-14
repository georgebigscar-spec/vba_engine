Attribute VB_Name = "strings"
Option Explicit

' Examples ported from the VBA language reference (MIT licensed) plus
' additional checks for VBA 7 behaviour (en-US, Windows-1252 code page).

Type CustomerRecord
    ID As Integer
    Name As String * 10
    Address As String * 30
End Type

Sub Main()
    TestAscChr
    TestAscWChrW
    TestCase
    TestLeftRight
    TestMid
    TestLen
    TestLenUdt
    TestTrim
    TestSpaceString
    TestInStr
    TestInStrRev
    TestStrComp
    TestStrConv
    TestStrReverse
    TestReplace
    TestSplit
    TestJoin
    TestFilter
    TestMidStatement
    TestLSetRSet
    TestConcat
    TestDollarVariants
    TestNullErrors
End Sub

Private Sub TestAscChr()
    Check "asc-function: A", Asc("A"), 65
    Check "asc-function: a", Asc("a"), 97
    Check "asc-function: Apple", Asc("Apple"), 65
    CheckTyped "asc: subtype", Asc("A"), CInt(65)
    CheckTyped "asc: number arg", Asc(5), CInt(53)
    Check "asc: euro cp1252", Asc(ChrW(8364)), 128
    Check "asc: space", Asc(" "), 32

    Check "chr-function: 65", Chr(65), "A"
    Check "chr-function: 97", Chr(97), "a"
    Check "chr-function: 62", Chr(62), ">"
    Check "chr-function: 37", Chr(37), "%"
    Check "chr: 10 is vbLf", Chr(10), vbLf
    Check "chr: 0 is vbNullChar", Chr(0), vbNullChar
    Check "chr: len of Chr(0)", Len(Chr(0)), 1
    Check "chr: 128 is euro", AscW(Chr(128)), 8364
    Check "chr: 233", AscW(Chr(233)), 233
    CheckTyped "chr: subtype", Chr(66), "B"
    Check "chr: roundtrip", Asc(Chr(200)), 200
End Sub

Private Sub TestAscWChrW()
    CheckTyped "ascw: A", AscW("A"), CInt(65)
    Check "ascw: alpha", AscW(ChrW(945)), 945
    Check "ascw: euro", AscW(ChrW(&H20AC)), 8364
    Check "ascw: above 32767 is negative", AscW(ChrW(40000)), -25536
    Check "ascw: FFFF", AscW(ChrW(&HFFFF&)), -1
    Check "chrw: len", Len(ChrW(945)), 1
    Check "chrw: ascii", ChrW(65), "A"
    CheckTyped "chrw$: type", ChrW$(97), "a"
End Sub

Private Sub TestCase()
    Dim UpperCase, LowerCase
    UpperCase = "Hello World 1234"
    LowerCase = LCase(UpperCase)
    Check "lcase-function: LowerCase", LowerCase, "hello world 1234"
    LowerCase = "Hello World 1234"
    UpperCase = UCase(LowerCase)
    Check "ucase-function: UpperCase", UpperCase, "HELLO WORLD 1234"

    Check "ucase: empty string", UCase(""), ""
    Check "ucase: Empty", UCase(Empty), ""
    Check "lcase: Null", LCase(Null), Null
    Check "ucase: Null", UCase(Null), Null
    CheckTyped "ucase: number", UCase(123), "123"
    Check "lcase: True", LCase(True), "true"
    Check "ucase: a-umlaut", UCase(ChrW(228)), ChrW(196)
    Check "lcase: A-umlaut", LCase(ChrW(196)), ChrW(228)
    Check "ucase: punctuation", UCase("a-b_c!"), "A-B_C!"
End Sub

Private Sub TestLeftRight()
    Dim AnyString, MyStr
    AnyString = "Hello World"
    MyStr = Left(AnyString, 1)
    Check "left-function: 1", MyStr, "H"
    MyStr = Left(AnyString, 7)
    Check "left-function: 7", MyStr, "Hello W"
    MyStr = Left(AnyString, 20)
    Check "left-function: 20", MyStr, "Hello World"
    MyStr = Right(AnyString, 1)
    Check "right-function: 1", MyStr, "d"
    MyStr = Right(AnyString, 6)
    Check "right-function: 6", MyStr, " World"
    MyStr = Right(AnyString, 20)
    Check "right-function: 20", MyStr, "Hello World"

    Check "left: 0", Left("abc", 0), ""
    Check "right: 0", Right("abc", 0), ""
    Check "left: Null", Left(Null, 2), Null
    Check "right: Null", Right(Null, 2), Null
    Check "left: number", Left(12345, 2), "12"
    Check "right: number", Right(12345, 2), "45"
    Check "left: Empty", Left(Empty, 3), ""
    Check "left: 2.5 rounds to 2", Left("abcdef", 2.5), "ab"
    Check "left: 3.5 rounds to 4", Left("abcdef", 3.5), "abcd"
    CheckTyped "left: subtype", Left("abc", 1), "a"

    On Error Resume Next
    MyStr = Left("abc", -1)
    CheckErr "left: negative length", 5
    MyStr = Right("abc", -1)
    CheckErr "right: negative length", 5
    MyStr = Left("abc", Null)
    CheckErr "left: Null length", 94
    On Error GoTo 0
End Sub

Private Sub TestMid()
    Dim MyString, FirstWord, LastWord, MidWords
    MyString = "Mid Function Demo"
    FirstWord = Mid(MyString, 1, 3)
    Check "mid-function: FirstWord", FirstWord, "Mid"
    LastWord = Mid(MyString, 14, 4)
    Check "mid-function: LastWord", LastWord, "Demo"
    MidWords = Mid(MyString, 5)
    Check "mid-function: MidWords", MidWords, "Function Demo"

    Check "mid: past end", Mid("abc", 5), ""
    Check "mid: start at len+1", Mid("abc", 4, 2), ""
    Check "mid: length too long", Mid("abc", 2, 100), "bc"
    Check "mid: length 0", Mid("abc", 2, 0), ""
    Check "mid: Null", Mid(Null, 1, 2), Null
    Check "mid: number", Mid(12345, 2, 3), "234"

    Dim r
    On Error Resume Next
    r = Mid("abc", 0)
    CheckErr "mid: start 0", 5
    r = Mid("abc", -1, 1)
    CheckErr "mid: start negative", 5
    r = Mid("abc", 1, -1)
    CheckErr "mid: negative length", 5
    On Error GoTo 0
End Sub

Private Sub TestLen()
    Dim MyInt As Integer, MyCur As Currency
    Dim MyLng As Long, MyDbl As Double, MyByte As Byte, MyBool As Boolean
    Dim MySng As Single, MyDate As Date
    Dim MyString, MyLen
    MyString = "Hello World"
    MyLen = Len(MyInt)
    Check "len-function: Integer", MyLen, 2
    MyLen = Len(MyString)
    Check "len-function: MyString", MyLen, 11
    MyLen = Len(MyCur)
    Check "len-function: Currency", MyLen, 8

    Check "len: Long", Len(MyLng), 4
    Check "len: Double", Len(MyDbl), 8
    Check "len: Byte", Len(MyByte), 1
    Check "len: Boolean", Len(MyBool), 2
    Check "len: Single", Len(MySng), 4
    Check "len: Date", Len(MyDate), 8
    Check "len: empty string", Len(""), 0
    Check "len: Empty", Len(Empty), 0
    Check "len: Null", Len(Null), Null
    CheckTyped "len: subtype", Len("abc"), CLng(3)
    Dim v
    v = 12345
    Check "len: Variant number uses string form", Len(v), 5
    Check "len: vbCrLf", Len(vbCrLf), 2
    Dim s As String
    Check "len: unassigned String", Len(s), 0
End Sub

Private Sub TestLenUdt()
    Dim Customer As CustomerRecord
    Dim MyLen
    MyLen = Len(Customer)
    Check "len-function: Customer", MyLen, 42
    Dim f As String * 10
    Check "len: fixed string", Len(f), 10
End Sub

Private Sub TestTrim()
    Dim MyString, TrimString
    MyString = "  <-Trim->  "
    TrimString = LTrim(MyString)
    Check "ltrim-rtrim-and-trim-functions: LTrim", TrimString, "<-Trim->  "
    TrimString = RTrim(MyString)
    Check "ltrim-rtrim-and-trim-functions: RTrim", TrimString, "  <-Trim->"
    TrimString = LTrim(RTrim(MyString))
    Check "ltrim-rtrim-and-trim-functions: LTrim(RTrim)", TrimString, "<-Trim->"
    TrimString = Trim(MyString)
    Check "ltrim-rtrim-and-trim-functions: Trim", TrimString, "<-Trim->"

    Check "trim: tabs not removed", Trim(vbTab & "a" & vbTab), vbTab & "a" & vbTab
    Check "trim: only spaces", Trim("    "), ""
    Check "trim: inner spaces kept", Trim(" a  b "), "a  b"
    Check "trim: Null", Trim(Null), Null
    Check "ltrim: Null", LTrim(Null), Null
    Check "rtrim: Null", RTrim(Null), Null
    Check "trim: number", Trim(5), "5"
    Check "trim: Empty", Trim(Empty), ""
End Sub

Private Sub TestSpaceString()
    Dim MyString
    MyString = Space(10)
    Check "space-function: Space(10)", MyString, "          "
    MyString = "Hello" & Space(10) & "World"
    Check "space-function: Hello World", MyString, "Hello          World"
    Check "space: 0", Space(0), ""
    CheckTyped "space: subtype", Space(2), "  "

    MyString = String(5, "*")
    Check "string-function: *", MyString, "*****"
    MyString = String(5, 42)
    Check "string-function: 42", MyString, "*****"
    MyString = String(10, "ABC")
    Check "string-function: ABC", MyString, "AAAAAAAAAA"
    Check "string: 0 count", String(0, "x"), ""
    Check "string: code mod 256", String(3, 321), "AAA"
    Check "string: Null char", String(3, Null), Null

    Dim r
    On Error Resume Next
    r = Space(-1)
    CheckErr "space: negative", 5
    r = String(-1, "a")
    CheckErr "string: negative", 5
    r = String(3, "")
    CheckErr "string: empty char", 5
    On Error GoTo 0
End Sub

Private Sub TestInStr()
    Dim SearchString, SearchChar, MyPos
    SearchString = "XXpXXpXXPXXP"
    SearchChar = "P"
    MyPos = InStr(4, SearchString, SearchChar, 1)
    Check "instr-function: text from 4", MyPos, 6
    MyPos = InStr(1, SearchString, SearchChar, 0)
    Check "instr-function: binary from 1", MyPos, 9
    MyPos = InStr(SearchString, SearchChar)
    Check "instr-function: default binary", MyPos, 9
    MyPos = InStr(1, SearchString, "W")
    Check "instr-function: not found", MyPos, 0

    CheckTyped "instr: subtype", InStr("abc", "b"), CLng(2)
    Check "instr: empty string2", InStr("abc", ""), 1
    Check "instr: empty string2 with start", InStr(2, "abc", ""), 2
    Check "instr: empty string1", InStr("", "a"), 0
    Check "instr: start past end", InStr(5, "abc", "a"), 0
    Check "instr: string1 Null", InStr(Null, "a"), Null
    Check "instr: string2 Null", InStr("abc", Null), Null
    Check "instr: vbTextCompare", InStr(1, "ABC", "b", vbTextCompare), 2
    Check "instr: vbBinaryCompare", InStr(1, "ABC", "b", vbBinaryCompare), 0
    Check "instr: module default is binary", InStr("ABC", "b"), 0
    Check "instr: multi-char", InStr("hello world", "o w"), 5
    Check "instr: start skips", InStr(6, "hello world", "o"), 8

    Dim r
    On Error Resume Next
    r = InStr(0, "abc", "a")
    CheckErr "instr: start 0", 5
    r = InStr(Null, "abc", "a")
    CheckErr "instr: start Null", 94
    On Error GoTo 0
End Sub

Private Sub TestInStrRev()
    Check "instrrev: last b", InStrRev("abcabc", "b"), 5
    Check "instrrev: from 4", InStrRev("abcabc", "b", 4), 2
    Check "instrrev-function: match end must be <= start", InStrRev("abc", "c", 2), 0
    Check "instrrev: end at start", InStrRev("abc", "bc", 3), 2
    Check "instrrev-function: not found", InStrRev("abc", "x"), 0
    Check "instrrev-function: start > len", InStrRev("abc", "a", 5), 0
    Check "instrrev-function: stringcheck zero-length", InStrRev("", "a"), 0
    Check "instrrev-function: stringcheck Null", InStrRev(Null, "a"), Null
    Check "instrrev-function: stringmatch Null", InStrRev("abc", Null), Null
    Check "instrrev-function: stringmatch zero-length", InStrRev("abcdef", "", 4), 4
    Check "instrrev: text compare", InStrRev("ABCABC", "b", -1, vbTextCompare), 5
    Check "instrrev: binary compare", InStrRev("ABCABC", "b", -1, vbBinaryCompare), 0
    CheckTyped "instrrev: subtype", InStrRev("abc", "a"), CLng(1)

    Dim r
    On Error Resume Next
    r = InStrRev("abc", "a", 0)
    CheckErr "instrrev: start 0", 5
    r = InStrRev("abc", "a", Null)
    CheckErr "instrrev-function: start Null", 94
    On Error GoTo 0
End Sub

Private Sub TestStrComp()
    Dim MyStr1, MyStr2, MyComp
    MyStr1 = "ABCD": MyStr2 = "abcd"
    MyComp = StrComp(MyStr1, MyStr2, 1)
    Check "strcomp-function: text", MyComp, 0
    MyComp = StrComp(MyStr1, MyStr2, 0)
    Check "strcomp-function: binary", MyComp, -1
    MyComp = StrComp(MyStr2, MyStr1)
    Check "strcomp-function: default", MyComp, 1

    CheckTyped "strcomp: subtype", StrComp("a", "b"), CInt(-1)
    Check "strcomp: equal", StrComp("abc", "abc"), 0
    Check "strcomp: longer greater", StrComp("abc", "ab"), 1
    Check "strcomp: shorter less", StrComp("ab", "abc"), -1
    Check "strcomp: empty strings", StrComp("", ""), 0
    Check "strcomp-function: Null 1", StrComp(Null, "a"), Null
    Check "strcomp-function: Null 2", StrComp("a", Null), Null
    Check "strcomp: binary a vs B", StrComp("a", "B", vbBinaryCompare), 1
    Check "strcomp: text a vs B", StrComp("a", "B", vbTextCompare), -1
    Check "strcomp: binary const", StrComp("A", "a", vbBinaryCompare), -1
    Check "strcomp: text const", StrComp("A", "a", vbTextCompare), 0

    Dim r
    On Error Resume Next
    r = StrComp("a", "b", Null)
    CheckErr "strcomp-function: compare Null", 94
    On Error GoTo 0

    ' Module has no Option Compare: comparisons are binary.
    Check "option compare: = binary", ("a" = "A"), False
    Check "option compare: < binary", ("a" < "B"), False
End Sub

Private Sub TestStrConv()
    Check "strconv: upper", StrConv("Hello World", vbUpperCase), "HELLO WORLD"
    Check "strconv: lower", StrConv("Hello World", vbLowerCase), "hello world"
    Check "strconv: proper", StrConv("hello world", vbProperCase), "Hello World"
    Check "strconv: proper lowers rest", StrConv("HELLO wORLD", vbProperCase), "Hello World"
    Check "strconv: proper apostrophe", StrConv("o'neil", vbProperCase), "O'neil"
    Check "strconv: proper hyphen", StrConv("a-b", vbProperCase), "A-b"
    Check "strconv-function: proper tab separator", StrConv("a" & vbTab & "b", vbProperCase), "A" & vbTab & "B"
    Check "strconv-function: proper lf separator", StrConv("a" & vbLf & "b", vbProperCase), "A" & vbLf & "B"
    Check "strconv: numeric constants", StrConv("abc", 1), "ABC"
    Check "strconv: proper 3", StrConv("abc def", 3), "Abc Def"
    CheckTyped "strconv: subtype", StrConv("x", vbUpperCase), "X"
End Sub

Private Sub TestStrReverse()
    Check "strreverse: abc", StrReverse("abc"), "cba"
    Check "strreverse-function: zero-length", StrReverse(""), ""
    Check "strreverse: palindrome", StrReverse("racecar"), "racecar"
    CheckTyped "strreverse: subtype", StrReverse("ab"), "ba"
    Dim r
    On Error Resume Next
    r = StrReverse(Null)
    CheckErr "strreverse-function: Null", 94
    On Error GoTo 0
End Sub

Private Sub TestReplace()
    Check "replace: all", Replace("aaa", "a", "b"), "bbb"
    Check "replace: non-overlapping", Replace("aaaa", "aa", "b"), "bb"
    Check "replace-function: start returns tail", Replace("abcabc", "b", "X", 3), "caXc"
    Check "replace: count 1", Replace("abcabc", "b", "X", 1, 1), "aXcabc"
    Check "replace-function: count 0", Replace("abcabc", "b", "X", 1, 0), "abcabc"
    Check "replace: text compare", Replace("ABCabc", "b", "x", 1, -1, vbTextCompare), "AxCaxc"
    Check "replace: binary compare", Replace("ABCabc", "b", "x", 1, -1, vbBinaryCompare), "ABCaxc"
    Check "replace-function: find zero-length", Replace("abc", "", "x"), "abc"
    Check "replace-function: replace zero-length", Replace("abcb", "b", ""), "ac"
    Check "replace-function: expression zero-length", Replace("", "a", "b"), ""
    Check "replace-function: start > len", Replace("abc", "b", "X", 5), ""
    Check "replace: longer replacement", Replace("a.b", ".", "--"), "a--b"
    CheckTyped "replace: subtype", Replace("ab", "a", "c"), "cb"

    Dim r
    On Error Resume Next
    r = Replace(Null, "a", "b")
    CheckErr "replace-function: Null expression", 94
    r = Replace("abc", "a", "b", 0)
    CheckErr "replace: start 0", 5
    On Error GoTo 0
End Sub

Private Sub TestSplit()
    Dim strFull As String
    Dim arrSplitStrings1() As String
    Dim arrSplitStrings2() As String
    Dim strSingleString1 As String
    Dim strSingleString2 As String
    Dim strSingleString3 As String
    Dim i As Long
    strFull = "Dow - Fonseca - Graham - Kopke - Noval - Offley - Sandeman - Taylor - Warre"
    arrSplitStrings1 = Split(strFull, "-")
    Check "split-function: arr1 LBound", LBound(arrSplitStrings1), 0
    Check "split-function: arr1 UBound", UBound(arrSplitStrings1), 8
    Check "split-function: arr1(0)", arrSplitStrings1(0), "Dow "
    Check "split-function: arr1(1)", arrSplitStrings1(1), " Fonseca "
    arrSplitStrings2 = Split(strFull, " - ")
    Check "split-function: arr2 UBound", UBound(arrSplitStrings2), 8
    Check "split-function: arr2(0)", arrSplitStrings2(0), "Dow"
    Check "split-function: arr2(1)", arrSplitStrings2(1), "Fonseca"
    strSingleString1 = arrSplitStrings2(3)
    Check "split-function: strSingleString1", strSingleString1, "Kopke"
    strSingleString2 = Split(strFull, " - ")(3)
    Check "split-function: strSingleString2", strSingleString2, "Kopke"
    For i = LBound(arrSplitStrings2, 1) To UBound(arrSplitStrings2, 1)
        If InStr(1, arrSplitStrings2(i), "Kopke", vbTextCompare) > 0 Then
            strSingleString3 = arrSplitStrings2(i)
            Exit For
        End If
    Next i
    Check "split-function: strSingleString3", strSingleString3, "Kopke"

    Dim v
    v = Split("")
    Check "split-function: zero-length UBound", UBound(v), -1
    Check "split-function: zero-length LBound", LBound(v), 0
    v = Split("a b c")
    Check "split: default delimiter", UBound(v), 2
    Check "split: default delimiter item", v(2), "c"
    v = Split("abc", "")
    Check "split-function: zero-length delimiter UBound", UBound(v), 0
    Check "split-function: zero-length delimiter item", v(0), "abc"
    v = Split("a,b,c", ",", 2)
    Check "split: limit UBound", UBound(v), 1
    Check "split: limit rest", v(1), "b,c"
    v = Split("a,,b", ",")
    Check "split: empty middle UBound", UBound(v), 2
    Check "split: empty middle", v(1), ""
    v = Split("aXbxc", "x", -1, vbTextCompare)
    Check "split: text compare", UBound(v), 2
    v = Split("aXbxc", "x", -1, vbBinaryCompare)
    Check "split: binary compare", UBound(v), 1
    Check "split: binary compare item", v(0), "aXb"
    v = Split("a,b,", ",")
    Check "split: trailing delimiter", UBound(v), 2
    Check "split: TypeName", TypeName(Split("a,b", ",")), "String()"
End Sub

Private Sub TestJoin()
    Dim a
    a = Array("a", "b", "c")
    Check "join-function: default delimiter", Join(a), "a b c"
    Check "join-function: zero-length delimiter", Join(a, ""), "abc"
    Check "join: comma", Join(a, ", "), "a, b, c"
    Check "join: numbers", Join(Array(1, 2, 3), "-"), "1-2-3"
    Check "join: roundtrip", Join(Split("x;y;z", ";"), ";"), "x;y;z"
    Check "join: empty split", Join(Split(""), ","), ""
    Check "join: single", Join(Array("only"), ","), "only"
    Dim s(1 To 2) As String
    s(1) = "p": s(2) = "q"
    Check "join: String array 1-based", Join(s, "+"), "p+q"
    CheckTyped "join: subtype", Join(a, ""), "abc"
End Sub

Private Sub TestFilter()
    Dim src, r
    src = Array("apple", "banana", "cherry")
    r = Filter(src, "an")
    Check "filter: include UBound", UBound(r), 0
    Check "filter: include item", r(0), "banana"
    r = Filter(src, "an", False)
    Check "filter: exclude UBound", UBound(r), 1
    Check "filter: exclude 0", r(0), "apple"
    Check "filter: exclude 1", r(1), "cherry"
    r = Filter(src, "AN", True, vbTextCompare)
    Check "filter: text compare", UBound(r), 0
    r = Filter(src, "AN", True, vbBinaryCompare)
    Check "filter-function: no match UBound", UBound(r), -1
    r = Filter(src, "e")
    Check "filter: multiple", UBound(r), 1
    Check "filter: multiple 1", r(1), "cherry"
    Check "filter: LBound", LBound(r), 0
End Sub

Private Sub TestMidStatement()
    Dim MyString
    MyString = "The dog jumps"
    Mid(MyString, 5, 3) = "fox"
    Check "mid-statement: fox", MyString, "The fox jumps"
    Mid(MyString, 5) = "cow"
    Check "mid-statement: cow", MyString, "The cow jumps"
    Mid(MyString, 5) = "cow jumped over"
    Check "mid-statement: cow jumped over", MyString, "The cow jumpe"
    Mid(MyString, 5, 3) = "duck"
    Check "mid-statement: duck", MyString, "The duc jumpe"

    Dim s As String
    s = "abcdef"
    Mid$(s, 2, 2) = "XY"
    Check "mid$ statement", s, "aXYdef"
    s = "abcdef"
    Mid(s, 3, 10) = "Z"
    Check "mid statement: short replacement", s, "abZdef"
    s = "abc"
    Mid(s, 1, 0) = "Q"
    Check "mid statement: length 0", s, "abc"

    On Error Resume Next
    s = "abc"
    Mid(s, 0, 1) = "Q"
    CheckErr "mid statement: start 0", 5
    On Error GoTo 0
End Sub

Private Sub TestLSetRSet()
    Dim MyString
    MyString = "0123456789"
    LSet MyString = "<-Left"
    Check "lset-statement: MyString", MyString, "<-Left    "
    MyString = "0123456789"
    RSet MyString = "Right->"
    Check "rset-statement: MyString", MyString, "   Right->"

    Dim s As String
    s = "abc"
    LSet s = "abcdef"
    Check "lset-statement: truncates", s, "abc"
    s = "abc"
    RSet s = "xyzuvw"
    Check "rset: truncates leftmost", s, "xyz"
    s = "abcde"
    LSet s = ""
    Check "lset: empty source", s, "     "
    s = "abcde"
    RSet s = "Q"
    Check "rset: single char", s, "    Q"
    Dim f As String * 5
    LSet f = "ab"
    Check "lset: fixed string", f, "ab   "
    RSet f = "ab"
    Check "rset: fixed string", f, "   ab"
End Sub

Private Sub TestConcat()
    CheckTyped "concat: string & Null", "a" & Null, "a"
    CheckTyped "concat: Null & string", Null & "a", "a"
    Check "concat: Null & Null", Null & Null, Null
    CheckTyped "concat: string & Empty", "a" & Empty, "a"
    CheckTyped "concat: Empty & Empty", Empty & Empty, ""
    CheckTyped "concat: numbers", 1 & 2, "12"
    CheckTyped "concat: Boolean", True & "", "True"
    CheckTyped "concat: Double", 1.5 & "", "1.5"
    Check "concat: + with Null", "a" + Null, Null
    CheckTyped "concat: + strings", "a" + "b", "ab"
    Dim v
    CheckTyped "concat: uninitialized Variant", v & "x", "x"
    Check "concat: Null & Empty", Null & Empty, ""
End Sub

Private Sub TestDollarVariants()
    CheckTyped "left$", Left$("abc", 2), "ab"
    CheckTyped "right$", Right$("abc", 2), "bc"
    CheckTyped "mid$", Mid$("abcdef", 2, 3), "bcd"
    CheckTyped "mid$ no length", Mid$("abcdef", 4), "def"
    CheckTyped "ucase$", UCase$("abc"), "ABC"
    CheckTyped "lcase$", LCase$("ABC"), "abc"
    CheckTyped "trim$", Trim$("  x  "), "x"
    CheckTyped "ltrim$", LTrim$("  x  "), "x  "
    CheckTyped "rtrim$", RTrim$("  x  "), "  x"
    CheckTyped "space$", Space$(3), "   "
    CheckTyped "string$", String$(3, "x"), "xxx"
    CheckTyped "chr$", Chr$(65), "A"
    CheckTyped "left$ number", Left$(123, 1), "1"
    CheckTyped "left$ Empty", Left$(Empty, 1), ""
End Sub

Private Sub TestNullErrors()
    Dim r
    On Error Resume Next
    r = Left$(Null, 1)
    CheckErr "left$: Null", 94
    r = Right$(Null, 1)
    CheckErr "right$: Null", 94
    r = Mid$(Null, 1)
    CheckErr "mid$: Null", 94
    r = UCase$(Null)
    CheckErr "ucase$: Null", 94
    r = LCase$(Null)
    CheckErr "lcase$: Null", 94
    r = Trim$(Null)
    CheckErr "trim$: Null", 94
    r = Asc(Null)
    CheckErr "asc: Null", 94
    r = Asc("")
    CheckErr "asc: empty string", 5
    r = AscW("")
    CheckErr "ascw: empty string", 5
    r = Chr(256)
    CheckErr "chr: 256", 5
    r = Chr(-1)
    CheckErr "chr: -1", 5
    r = Chr(Null)
    CheckErr "chr: Null", 94
    r = Space(Null)
    CheckErr "space: Null", 94
    r = Left("abc", "x")
    CheckErr "left: type mismatch length", 13
    On Error GoTo 0
End Sub
