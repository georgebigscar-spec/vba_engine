Attribute VB_Name = "operators"
Option Explicit

Sub Main()
    TestDocExamples
    TestArithmeticTypes
    TestOverflow
    TestDivision
    TestIntDivMod
    TestPower
    TestPlusAndConcat
    TestNullEmptyArith
    TestComparisons
    TestVariantComparisons
    TestLogicalTruth
    TestLogicalNull
    TestBitwise
    TestLike
    TestIs
    TestPrecedence
    TestDates
End Sub

Private Sub TestDocExamples()
    Dim A, B, C, D
    A = 10: B = 8: C = 6: D = Null
    ' ampersand-operator
    Check "ampersand-operator: Hello", "Hello" & " World", "Hello World"
    Check "ampersand-operator: Check 123", "Check " & 123 & " Check", "Check 123 Check"
    ' and-operator
    Check "and-operator: A>B And B>C", A > B And B > C, True
    Check "and-operator: B>A And B>C", B > A And B > C, False
    Check "and-operator: A>B And B>D", A > B And B > D, Null
    Check "and-operator: A And B", A And B, 8
    ' backslash-operator
    Check "backslash-operator: 11\4", 11 \ 4, 2
    Check "backslash-operator: 9\3", 9 \ 3, 3
    Check "backslash-operator: 100\3", 100 \ 3, 33
    ' caret-operator
    Check "caret-operator: 2^2", 2 ^ 2, 4
    Check "caret-operator: 3^3^3", 3 ^ 3 ^ 3, 19683
    Check "caret-operator: (-5)^3", (-5) ^ 3, -125
    ' comparison-operators
    Dim Var1, Var2
    Check "comparison-operators: 45<35", (45 < 35), False
    Check "comparison-operators: 45=45", (45 = 45), True
    Check "comparison-operators: 4<>3", (4 <> 3), True
    Check "comparison-operators: ""5"">""4""", ("5" > "4"), True
    Var1 = "5": Var2 = 4
    Check "comparison-operators: Var1>Var2 str/num", (Var1 > Var2), True
    Var1 = 5: Var2 = Empty
    Check "comparison-operators: 5>Empty", (Var1 > Var2), True
    Var1 = 0: Var2 = Empty
    Check "comparison-operators: 0=Empty", (Var1 = Var2), True
    ' eqv-operator
    Check "eqv-operator: A>B Eqv B>C", A > B Eqv B > C, True
    Check "eqv-operator: B>A Eqv B>C", B > A Eqv B > C, False
    Check "eqv-operator: A>B Eqv B>D", A > B Eqv B > D, Null
    Check "eqv-operator: A Eqv B", A Eqv B, -3
    ' forwardslash-operator
    Check "forwardslash-operator: 10/4", 10 / 4, 2.5
    Check "forwardslash-operator: 10/3", 10 / 3, 10# / 3#
    ' imp-operator
    Check "imp-operator: A>B Imp B>C", A > B Imp B > C, True
    Check "imp-operator: A>B Imp C>B", A > B Imp C > B, False
    Check "imp-operator: B>A Imp C>B", B > A Imp C > B, True
    Check "imp-operator: B>A Imp C>D", B > A Imp C > D, True
    Check "imp-operator: C>D Imp B>A", C > D Imp B > A, Null
    Check "imp-operator: B Imp A", B Imp A, -1
    ' like-operator
    Check "like-operator: a*a", "aBBBa" Like "a*a", True
    Check "like-operator: [A-Z]", "F" Like "[A-Z]", True
    Check "like-operator: [!A-Z]", "F" Like "[!A-Z]", False
    Check "like-operator: a#a", "a2a" Like "a#a", True
    Check "like-operator: a[L-P]#[!c-e]", "aM5b" Like "a[L-P]#[!c-e]", True
    Check "like-operator: BAT B?T*", "BAT123khg" Like "B?T*", True
    Check "like-operator: CAT B?T*", "CAT123khg" Like "B?T*", False
    Check "like-operator: a*b", "ab" Like "a*b", True
    Check "like-operator: a*b vs a [*]b", "a*b" Like "a [*]b", False
    Check "like-operator: axxxxxb", "axxxxxb" Like "a [*]b", False
    Check "like-operator: a [[]*", "a [xyz" Like "a [[]*", True
    Dim r
    On Error Resume Next
    r = "a [xyz" Like "a [*"
    CheckErr "like-operator: a [* invalid pattern", 93
    On Error GoTo 0
    ' minus-operator
    Check "minus-operator: 4-2", 4 - 2, 2
    Check "minus-operator: 459.35-334.90", 459.35 - 334.9, 459.35 - 334.9
    Check "minus-operator: 459.35-334.90 approx", Round(459.35 - 334.9, 2), 124.45
    ' mod-operator
    Check "mod-operator: 10 Mod 5", 10 Mod 5, 0
    Check "mod-operator: 10 Mod 3", 10 Mod 3, 1
    Check "mod-operator: 12 Mod 4.3", 12 Mod 4.3, 0
    Check "mod-operator: 12.6 Mod 5", 12.6 Mod 5, 3
    Check "mod-operator: 19 Mod 6.7", 19 Mod 6.7, 5
    ' not-operator
    Check "not-operator: Not(A>B)", Not (A > B), False
    Check "not-operator: Not(B>A)", Not (B > A), True
    Check "not-operator: Not(C>D)", Not (C > D), Null
    Check "not-operator: Not A", Not A, -11
    ' operator (*)
    Check "operator: 2*2", 2 * 2, 4
    Check "operator: 459.35*334.90", Round(459.35 * 334.9, 3), 153836.315
    ' or-operator
    Check "or-operator: A>B Or B>C", A > B Or B > C, True
    Check "or-operator: B>A Or B>C", B > A Or B > C, True
    Check "or-operator: A>B Or B>D", A > B Or B > D, True
    Check "or-operator: B>D Or B>A", B > D Or B > A, Null
    Check "or-operator: A Or B", A Or B, 10
    ' plus-operator
    Check "plus-operator: 2+2", 2 + 2, 4
    Check "plus-operator: 4257.04+98112", 4257.04 + 98112, 102369.04
    Var1 = "34": Var2 = 6
    Check "plus-operator: ""34""+6", Var1 + Var2, 40
    Var1 = "34": Var2 = "6"
    Check "plus-operator: ""34""+""6""", Var1 + Var2, "346"
    ' xor-operator
    Check "xor-operator: A>B Xor B>C", A > B Xor B > C, False
    Check "xor-operator: B>A Xor B>C", B > A Xor B > C, True
    Check "xor-operator: B>A Xor C>B", B > A Xor C > B, False
    Check "xor-operator: B>D Xor A>B", B > D Xor A > B, Null
    Check "xor-operator: A Xor B", A Xor B, 2
End Sub

Private Sub TestArithmeticTypes()
    Dim i As Integer, l As Long, s As Single, d As Double, c As Currency, b As Byte
    i = 2: l = 3: s = 1.5: d = 2.5: c = 4: b = 5
    CheckTyped "plus-operator: Integer+Integer", i + i, CInt(4)
    CheckTyped "plus-operator: Integer+Long", i + l, CLng(5)
    CheckTyped "plus-operator: Byte+Byte", b + b, CByte(10)
    CheckTyped "plus-operator: Byte+Integer", b + i, CInt(7)
    CheckTyped "plus-operator: Integer+Single", i + s, CSng(3.5)
    CheckTyped "plus-operator: Long+Single is Double", l + s, CDbl(4.5)
    CheckTyped "plus-operator: Integer+Double", i + d, CDbl(4.5)
    CheckTyped "plus-operator: Currency+Integer", c + i, CCur(6)
    CheckTyped "plus-operator: Boolean+Boolean", True + True, CInt(-2)
    CheckTyped "minus-operator: Long-Integer", l - i, CLng(1)
    CheckTyped "minus-operator: unary minus Integer", -i, CInt(-2)
    CheckTyped "operator: Integer*Integer", i * i, CInt(4)
    CheckTyped "operator: Integer*Long", i * l, CLng(6)
    CheckTyped "operator: Integer*Double", i * d, CDbl(5)
    CheckTyped "operator: Currency*Integer", c * i, CCur(8)
    CheckTyped "operator: literal 2*2 Integer", 2 * 2, CInt(4)
    CheckTyped "plus-operator: literal 2+2 Integer", 2 + 2, CInt(4)
    CheckTyped "plus-operator: literal 40000+1 Long", 40000 + 1, CLng(40001)
End Sub

Private Sub TestOverflow()
    Dim i As Integer, l As Long, b As Byte, r
    Dim vi, vl, vb
    On Error Resume Next
    i = 32767
    r = i + 1
    CheckErr "plus-operator: Integer 32767+1 overflow", 6
    i = 200
    r = i * i
    CheckErr "operator: Integer 200*200 overflow", 6
    i = -32768
    r = -i
    CheckErr "minus-operator: -(-32768 Integer) overflow", 6
    i = -32768
    r = i - 1
    CheckErr "minus-operator: Integer -32768-1 overflow", 6
    l = 2147483647
    r = l + 1
    CheckErr "plus-operator: Long max+1 overflow", 6
    b = 255
    r = b + CByte(1)
    CheckErr "plus-operator: Byte 255+1 overflow", 6
    b = 0
    r = b - CByte(1)
    CheckErr "minus-operator: Byte 0-1 overflow", 6
    On Error GoTo 0
    ' Variant arithmetic widens instead of overflowing
    vi = CInt(32767)
    CheckTyped "plus-operator: Variant Integer overflow -> Long", vi + CInt(1), CLng(32768)
    vi = CInt(200)
    CheckTyped "operator: Variant Integer*Integer overflow -> Long", vi * vi, CLng(40000)
    vl = CLng(2147483647)
    CheckTyped "plus-operator: Variant Long overflow -> Double", vl + CLng(1), CDbl(2147483648#)
    vb = CByte(255)
    CheckTyped "plus-operator: Variant Byte overflow -> Integer", vb + CByte(1), CInt(256)
End Sub

Private Sub TestDivision()
    Dim i As Integer, l As Long, r
    i = 4: l = 3
    CheckTyped "forwardslash-operator: 10/4 Double", 10 / 4, CDbl(2.5)
    CheckTyped "forwardslash-operator: Integer/Integer Double", i / CInt(2), CDbl(2)
    CheckTyped "forwardslash-operator: Long/Long Double", CLng(9) / l, CDbl(3)
    CheckTyped "forwardslash-operator: Single/Single Single", CSng(3) / CSng(2), CSng(1.5)
    Check "forwardslash-operator: Null/2", Null / 2, Null
    Check "forwardslash-operator: Empty/2", Empty / 2, 0
    On Error Resume Next
    r = 1 / 0
    CheckErr "forwardslash-operator: 1/0", 11
    r = 0 / 0
    CheckErr "forwardslash-operator: 0/0 overflow", 6
    r = 5 / Empty
    CheckErr "forwardslash-operator: 5/Empty", 11
    r = "abc" / 2
    CheckErr "forwardslash-operator: ""abc""/2", 13
    On Error GoTo 0
End Sub

Private Sub TestIntDivMod()
    Dim r
    CheckTyped "backslash-operator: Integer result", 11 \ 4, CInt(2)
    CheckTyped "backslash-operator: Double operands -> Long", 7.2 \ 2, CLng(3)
    CheckTyped "backslash-operator: Long", CLng(100000) \ 3, CLng(33333)
    Check "backslash-operator: 11.5\2 rounds to 12", 11.5 \ 2, 6
    Check "backslash-operator: 10.5\1 banker's", 10.5 \ 1, 10
    Check "backslash-operator: 9.5\1 banker's", 9.5 \ 1, 10
    Check "backslash-operator: 7\2.5 divisor rounds to 2", 7 \ 2.5, 3
    Check "backslash-operator: -7\2 truncates", -7 \ 2, -3
    Check "backslash-operator: 7\-2 truncates", 7 \ -2, -3
    Check "backslash-operator: ""9""\2", "9" \ 2, 4
    Check "backslash-operator: Null\2", Null \ 2, Null
    CheckTyped "mod-operator: Integer result", 10 Mod 3, CInt(1)
    CheckTyped "mod-operator: Double operands -> Long", 10.2 Mod 3, CLng(1)
    Check "mod-operator: 5.5 Mod 2 banker's", 5.5 Mod 2, 0
    Check "mod-operator: 4.5 Mod 3 banker's", 4.5 Mod 3, 1
    Check "mod-operator: 2.5 Mod 3 banker's", 2.5 Mod 3, 2
    Check "mod-operator: -7 Mod 3 sign of dividend", -7 Mod 3, -1
    Check "mod-operator: 7 Mod -3", 7 Mod -3, 1
    Check "mod-operator: -7 Mod -3", -7 Mod -3, -1
    Check "mod-operator: Null Mod 2", Null Mod 2, Null
    Check "mod-operator: Empty Mod 2", Empty Mod 2, 0
    On Error Resume Next
    r = 5 Mod 0
    CheckErr "mod-operator: Mod 0", 11
    r = 5 \ 0
    CheckErr "backslash-operator: \0", 11
    r = 5 Mod 0.4
    CheckErr "mod-operator: Mod 0.4 rounds to 0", 11
    r = 3000000000# \ 1
    CheckErr "backslash-operator: operand beyond Long", 6
    On Error GoTo 0
End Sub

Private Sub TestPower()
    Dim r
    CheckTyped "caret-operator: Integer^Integer Double", CInt(2) ^ CInt(3), CDbl(8)
    Check "caret-operator: 2^-1", 2 ^ -1, 0.5
    Check "caret-operator: 2^-2", 2 ^ -2, 0.25
    Check "caret-operator: 0^0", 0 ^ 0, 1
    Check "caret-operator: 4^0.5", 4 ^ 0.5, 2
    Check "caret-operator: (-2)^2", (-2) ^ 2, 4
    Check "caret-operator: (-2)^3", (-2) ^ 3, -8
    Check "caret-operator: -2^2", -2 ^ 2, -4
    Check "caret-operator: Null^2", Null ^ 2, Null
    Check "caret-operator: Empty^0", Empty ^ 0, 1
    On Error Resume Next
    r = (-8) ^ (1 / 3)
    CheckErr "caret-operator: negative base fractional exponent", 5
    On Error GoTo 0
End Sub

Private Sub TestPlusAndConcat()
    Dim v1, v2, s As String, n As Integer, r
    CheckTyped "plus-operator: String+String", "ab" + "cd", "abcd"
    v1 = "34": v2 = 6
    CheckTyped "plus-operator: Variant str+num Double", v1 + v2, CDbl(40)
    v1 = "1.5": v2 = 1
    Check "plus-operator: ""1.5""+1", v1 + v2, 2.5
    v1 = "abc": v2 = 1
    On Error Resume Next
    r = v1 + v2
    CheckErr "plus-operator: ""abc""+1 type mismatch", 13
    On Error GoTo 0
    v1 = Empty
    CheckTyped "plus-operator: Empty+String", v1 + "abc", "abc"
    CheckTyped "plus-operator: Empty+Long", v1 + CLng(5), CLng(5)
    CheckTyped "plus-operator: Empty+Empty Integer", Empty + Empty, CInt(0)
    Check "plus-operator: String+Null", "abc" + Null, Null
    s = "10": n = 5
    CheckTyped "plus-operator: String var + Integer var", s + n, CDbl(15)
    CheckTyped "minus-operator: ""10""-""4""", "10" - "4", CDbl(6)
    CheckTyped "operator: ""3""*""4""", "3" * "4", CDbl(12)
    ' ampersand
    CheckTyped "ampersand-operator: 1 & 2", 1 & 2, "12"
    CheckTyped "ampersand-operator: 1.5 & """"", 1.5 & "", "1.5"
    CheckTyped "ampersand-operator: True & """"", True & "", "True"
    CheckTyped "ampersand-operator: ""a"" & Null", "a" & Null, "a"
    CheckTyped "ampersand-operator: Null & ""a""", Null & "a", "a"
    Check "ampersand-operator: Null & Null", Null & Null, Null
    CheckTyped "ampersand-operator: Empty & Empty", Empty & Empty, ""
    CheckTyped "ampersand-operator: -1 & 2", -1 & 2, "-12"
    Check "ampersand-operator: precedence 1 + 2 & 3", 1 + 2 & 3, "33"
End Sub

Private Sub TestNullEmptyArith()
    Dim n
    n = Null
    Check "plus-operator: Null+1", n + 1, Null
    Check "minus-operator: 1-Null", 1 - n, Null
    Check "operator: Null*0", n * 0, Null
    Check "minus-operator: -Null", -n, Null
    CheckTyped "plus-operator: Empty+1", Empty + 1, CInt(1)
    CheckTyped "minus-operator: Empty-1", Empty - 1, CInt(-1)
    CheckTyped "operator: Empty*5", Empty * 5, CInt(0)
    CheckTyped "minus-operator: -Empty", -Empty, CInt(0)
    Check "plus-operator: IsNull(Null+Null)", IsNull(n + n), True
End Sub

Private Sub TestComparisons()
    Dim n
    n = Null
    CheckTyped "comparison-operators: result Boolean", 1 < 2, True
    Check "comparison-operators: Null=Null", n = n, Null
    Check "comparison-operators: 1<Null", 1 < n, Null
    Check "comparison-operators: Null<>1", n <> 1, Null
    Check "comparison-operators: ""a""=Null", "a" = n, Null
    Check "comparison-operators: 1=1.0", 1 = 1#, True
    Check "comparison-operators: ""abc""<""abd""", "abc" < "abd", True
    Check "comparison-operators: ""a""<""B"" binary", "a" < "B", False
    Check "comparison-operators: ""abc""=""ABC"" binary", "abc" = "ABC", False
    Check "comparison-operators: ""ab""<""abc""", "ab" < "abc", True
    Check "comparison-operators: """"<""a""", "" < "a", True
    Check "comparison-operators: ""10""<""9"" strings", "10" < "9", True
    Check "comparison-operators: True<False", True < False, True
    Check "comparison-operators: True=-1", True = -1, True
    Check "comparison-operators: Empty=""""", Empty = "", True
    Check "comparison-operators: Empty=0", Empty = 0, True
    Check "comparison-operators: 3>=3", 3 >= 3, True
    Check "comparison-operators: 3<=2", 3 <= 2, False
End Sub

Private Sub TestVariantComparisons()
    Dim v1, v2, s As String, i As Integer, r
    v1 = "1": v2 = 9
    Check "comparison-operators: Variant ""1"">9 (string > number)", v1 > v2, True
    v1 = "10": v2 = "9"
    Check "comparison-operators: Variant ""10""<""9""", v1 < v2, True
    v1 = "10": i = 9
    Check "comparison-operators: Variant ""10"" vs Integer 9 numeric", v1 < i, False
    v1 = "10": i = 10
    Check "comparison-operators: Variant ""10""=Integer 10", v1 = i, True
    s = "10": v2 = 9
    Check "comparison-operators: String ""10"" vs Variant 9 string", s < v2, True
    v1 = 5: v2 = Empty
    Check "comparison-operators: Variant 5 > Empty", v1 > v2, True
    v1 = "": v2 = Empty
    Check "comparison-operators: Variant """" = Empty", v1 = v2, True
    v1 = Empty: v2 = Empty
    Check "comparison-operators: Empty=Empty", v1 = v2, True
    v1 = 1: v2 = 1#
    Check "comparison-operators: Variant Integer=Double", v1 = v2, True
    v1 = "abc": i = 1
    On Error Resume Next
    r = v1 > i
    CheckErr "comparison-operators: ""abc"" vs Integer type mismatch", 13
    On Error GoTo 0
End Sub

Private Sub TestLogicalTruth()
    CheckTyped "and-operator: T And T", True And True, True
    Check "and-operator: T And F", True And False, False
    Check "and-operator: F And F", False And False, False
    Check "or-operator: T Or F", True Or False, True
    Check "or-operator: F Or F", False Or False, False
    Check "xor-operator: T Xor T", True Xor True, False
    Check "xor-operator: T Xor F", True Xor False, True
    Check "eqv-operator: T Eqv T", True Eqv True, True
    Check "eqv-operator: F Eqv F", False Eqv False, True
    Check "eqv-operator: T Eqv F", True Eqv False, False
    Check "imp-operator: T Imp T", True Imp True, True
    Check "imp-operator: T Imp F", True Imp False, False
    Check "imp-operator: F Imp T", False Imp True, True
    Check "imp-operator: F Imp F", False Imp False, True
    CheckTyped "not-operator: Not True", Not True, False
    CheckTyped "not-operator: Not False", Not False, True
End Sub

Private Sub TestLogicalNull()
    Dim n
    n = Null
    Check "and-operator: True And Null", True And n, Null
    Check "and-operator: False And Null", False And n, False
    Check "and-operator: Null And False", n And False, False
    Check "and-operator: Null And Null", n And n, Null
    Check "and-operator: 0 And Null", 0 And n, 0
    Check "or-operator: True Or Null", True Or n, True
    Check "or-operator: Null Or True", n Or True, True
    Check "or-operator: False Or Null", False Or n, Null
    Check "or-operator: Null Or Null", n Or n, Null
    Check "xor-operator: True Xor Null", True Xor n, Null
    Check "xor-operator: False Xor Null", False Xor n, Null
    Check "eqv-operator: True Eqv Null", True Eqv n, Null
    Check "eqv-operator: False Eqv Null", False Eqv n, Null
    Check "imp-operator: True Imp Null", True Imp n, Null
    Check "imp-operator: False Imp Null", False Imp n, True
    Check "imp-operator: Null Imp True", n Imp True, True
    Check "imp-operator: Null Imp False", n Imp False, Null
    Check "imp-operator: Null Imp Null", n Imp n, Null
    Check "not-operator: Not Null", Not n, Null
End Sub

Private Sub TestBitwise()
    Dim r
    CheckTyped "and-operator: Integer And Integer", CInt(12) And CInt(10), CInt(8)
    CheckTyped "or-operator: Integer Or Integer", CInt(12) Or CInt(10), CInt(14)
    CheckTyped "xor-operator: Integer Xor Integer", CInt(12) Xor CInt(10), CInt(6)
    CheckTyped "and-operator: Long And Integer", CLng(65535) And CInt(255), CLng(255)
    CheckTyped "and-operator: Byte And Byte", CByte(12) And CByte(10), CByte(8)
    CheckTyped "not-operator: Not Byte", Not CByte(0), CByte(255)
    CheckTyped "not-operator: Not Integer 0", Not CInt(0), CInt(-1)
    CheckTyped "not-operator: Not Long", Not CLng(0), CLng(-1)
    Check "eqv-operator: 12 Eqv 10", 12 Eqv 10, -7
    Check "imp-operator: 12 Imp 10", 12 Imp 10, -5
    Check "and-operator: -1 And 255", -1 And 255, 255
    Check "and-operator: 2.6 And 7 (rounds)", 2.6 And 7, 3
    Check "or-operator: 2.5 Or 0 (banker's)", 2.5 Or 0, 2
    Check "and-operator: True And 5", True And 5, 5
    Check "or-operator: ""4"" Or 1", "4" Or 1, 5
    On Error Resume Next
    r = "x" And 1
    CheckErr "and-operator: ""x"" And 1 type mismatch", 13
    On Error GoTo 0
End Sub

Private Sub TestLike()
    Dim n, r
    n = Null
    Check "like-operator: case-sensitive binary", "aBc" Like "abc", False
    Check "like-operator: exact", "abc" Like "abc", True
    Check "like-operator: escaped [*]", "a*b" Like "a[*]b", True
    Check "like-operator: escaped [*] no match", "axb" Like "a[*]b", False
    Check "like-operator: escaped [?]", "a?b" Like "a[?]b", True
    Check "like-operator: escaped [#]", "a#b" Like "a[#]b", True
    Check "like-operator: """" Like *", "" Like "*", True
    Check "like-operator: """" Like ?", "" Like "?", False
    Check "like-operator: abc Like """"", "abc" Like "", False
    Check "like-operator: """" Like """"", "" Like "", True
    Check "like-operator: ? exact length", "ab" Like "?", False
    Check "like-operator: ???", "abc" Like "???", True
    Check "like-operator: # digit", "7" Like "#", True
    Check "like-operator: # not letter", "x" Like "#", False
    Check "like-operator: ## two digits", "42" Like "##", True
    Check "like-operator: [a-z] lowercase", "q" Like "[a-z]", True
    Check "like-operator: [a-z] uppercase binary", "Q" Like "[a-z]", False
    Check "like-operator: [!a-z]", "Q" Like "[!a-z]", True
    Check "like-operator: [!a-z] lower", "q" Like "[!a-z]", False
    Check "like-operator: [abc] list", "b" Like "[abc]", True
    Check "like-operator: [abc] miss", "d" Like "[abc]", False
    Check "like-operator: multiple ranges", "5" Like "[a-z0-9]", True
    Check "like-operator: * middle", "hello world" Like "h*d", True
    Check "like-operator: * trailing miss", "hello" Like "h*x", False
    Check "like-operator: *ll*", "hello" Like "*ll*", True
    Check "like-operator: ] outside group", "a]" Like "a]", True
    Check "like-operator: [!] not digit", "a1" Like "a[!0-9]", False
    Check "like-operator: Null string", n Like "a", Null
    Check "like-operator: Null pattern", "a" Like n, Null
    Check "like-operator: number Like", 123 Like "1#3", True
    CheckTyped "like-operator: result Boolean", "a" Like "a", True
    On Error Resume Next
    r = "a" Like "[a"
    CheckErr "like-operator: unclosed bracket", 93
    On Error GoTo 0
End Sub

Private Sub TestIs()
    Dim c1 As Collection, c2 As Collection, c3 As Collection
    Dim o1 As Object, v
    Set c1 = New Collection
    Set c2 = c1
    Set c3 = New Collection
    Check "is-operator: same object", c1 Is c2, True
    Check "is-operator: different objects", c1 Is c3, False
    Check "is-operator: obj Is Nothing", c1 Is Nothing, False
    Check "is-operator: uninit Is Nothing", o1 Is Nothing, True
    Check "is-operator: Nothing Is Nothing", Nothing Is Nothing, True
    Set o1 = c3
    Check "is-operator: Object var Is Collection var", o1 Is c3, True
    Set v = c1
    Check "is-operator: Variant object Is", v Is c2, True
    CheckTyped "is-operator: result Boolean", c1 Is c1, True
    Check "is-operator: Not (x Is Nothing)", Not c1 Is Nothing, True
    Set c2 = Nothing
    Check "is-operator: after Set Nothing", c2 Is Nothing, True
    Check "is-operator: original kept", c1 Is Nothing, False
End Sub

Private Sub TestPrecedence()
    Check "operator-precedence: 2+3*4", 2 + 3 * 4, 14
    Check "operator-precedence: (2+3)*4", (2 + 3) * 4, 20
    Check "operator-precedence: 10-2-3 left assoc", 10 - 2 - 3, 5
    Check "operator-precedence: 8/4/2 left assoc", 8 / 4 / 2, 1
    Check "operator-precedence: 7\2*2 (* before \)", 7 \ 2 * 2, 1
    Check "operator-precedence: 10 Mod 3+1", 10 Mod 3 + 1, 2
    Check "operator-precedence: 10\3 Mod 2 (\ before Mod)", 10 \ 3 Mod 2, 1
    Check "operator-precedence: 2*3^2", 2 * 3 ^ 2, 18
    Check "operator-precedence: -2^2", -2 ^ 2, -4
    Check "operator-precedence: (-2)^2", (-2) ^ 2, 4
    Check "operator-precedence: 2^-2", 2 ^ -2, 0.25
    Check "operator-precedence: -3-(-3)", -3 - -3, 0
    Check "operator-precedence: 1+2 & 3+4", 1 + 2 & 3 + 4, "37"
    Check "operator-precedence: & before =", "a" & "b" = "ab", True
    Check "operator-precedence: comparisons before And", 1 < 2 And 3 < 4, True
    Check "operator-precedence: Not 1 = 2", Not 1 = 2, True
    Check "operator-precedence: Not before And", Not False And False, False
    Check "operator-precedence: And before Or", True Or True And False, True
    Check "operator-precedence: Or before Xor", True Xor True Or True, False
    Check "operator-precedence: 1<2<3", 1 < 2 < 3, True
    Check "operator-precedence: 3>2>1", 3 > 2 > 1, False
    Check "operator-precedence: 5 And 3 = 3", 5 And 3 = 3, 5 And True
    Check "operator-precedence: 2+3 = 5", 2 + 3 = 5, True
End Sub

Private Sub TestDates()
    Dim d As Date, r
    d = #1/1/2020#
    CheckTyped "plus-operator: Date+1", d + 1, #1/2/2020#
    CheckTyped "plus-operator: 1+Date", 1 + d, #1/2/2020#
    CheckTyped "minus-operator: Date-1", d - 1, #12/31/2019#
    CheckTyped "minus-operator: Date-Date Double", #1/2/2020# - d, CDbl(1)
    CheckTyped "plus-operator: Date+0.5", d + 0.5, #1/1/2020 12:00:00 PM#
    CheckTyped "minus-operator: Date-Date fraction", #1/1/2020 6:00:00 AM# - d, CDbl(0.25)
    Check "comparison-operators: Date<Date", d < #1/2/2020#, True
    Check "comparison-operators: Date=number", d = 43831, True
    Check "plus-operator: Date+Null", d + Null, Null
    CheckTyped "ampersand-operator: Date & """"", #1/2/2020# & "", "1/2/2020"
End Sub
