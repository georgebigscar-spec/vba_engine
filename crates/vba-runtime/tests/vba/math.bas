Attribute VB_Name = "math"
' Math and financial functions: Abs, Atn, Cos, Sin, Tan, Exp, Log, Sqr,
' Sgn, Int, Fix, Round, Rnd/Randomize, Partition, the financial functions
' and the "Derived math functions" formulas.
Option Explicit

Sub Main()
    TestAbs
    TestTrig
    TestExpLog
    TestSqr
    TestSgn
    TestIntFix
    TestRound
    TestRnd
    TestPartition
    TestDerived
    TestDepreciation
    TestAnnuity
    TestCashFlows
End Sub

' True when a and b agree to a relative tolerance (absolute near zero).
Private Function Near(ByVal a As Variant, ByVal b As Variant, Optional ByVal tol As Double = 0.000000001) As Boolean
    Dim scale_ As Double
    scale_ = Abs(b)
    If scale_ < 1 Then scale_ = 1
    Near = (Abs(a - b) <= tol * scale_)
End Function

Private Sub TestAbs()
    Dim MyNumber
    MyNumber = Abs(50.3)
    Check "abs-function: Abs(50.3)", MyNumber, 50.3
    MyNumber = Abs(-50.3)
    Check "abs-function: Abs(-50.3)", MyNumber, 50.3
    Check "abs-function: Abs(-1)", Abs(-1), 1
    Check "abs-function: Abs(1)", Abs(1), 1

    CheckTyped "abs: Integer keeps type", Abs(CInt(-5)), CInt(5)
    CheckTyped "abs: Long keeps type", Abs(-5&), 5&
    CheckTyped "abs: Double keeps type", Abs(-2.5), 2.5
    CheckTyped "abs: Single keeps type", Abs(CSng(-2.5)), CSng(2.5)
    CheckTyped "abs: Currency keeps type", Abs(CCur(-1.5)), CCur(1.5)
    CheckTyped "abs: Decimal keeps type", Abs(CDec(-1.5)), CDec(1.5)
    CheckTyped "abs: String gives Double", Abs("-3"), 3#
    Check "abs-function: Null", Abs(Null), Null
    Check "abs-function: uninitialized gives zero", Abs(Empty), 0
    Check "abs: zero", Abs(0), 0

    On Error Resume Next
    Dim r
    r = Abs(CInt(-32768))
    CheckErr "abs: Integer min overflows", 6
    r = Abs(CLng(-2147483647) - 1)
    CheckErr "abs: Long min overflows", 6
    r = Abs("abc")
    CheckErr "abs: non-numeric string", 13
    On Error GoTo 0
End Sub

Private Sub TestTrig()
    Dim MyAngle, MySecant, MyCosecant, MyCotangent
    MyAngle = 1.3
    MySecant = 1 / Cos(MyAngle)
    Check "cos-function: secant of 1.3", Near(MySecant, 3.73833412707544), True
    MyCosecant = 1 / Sin(MyAngle)
    Check "sin-function: cosecant of 1.3", Near(MyCosecant, 1.0378200456748), True
    MyCotangent = 1 / Tan(MyAngle)
    Check "tan-function: cotangent of 1.3", Near(MyCotangent, 0.277615646541125), True

    Check "atn-function: 4*Atn(1) = pi", CStr(4 * Atn(1)), "3.14159265358979"
    Check "atn: Atn(0)", Atn(0), 0
    CheckTyped "atn: returns Double", Atn(CInt(1)), Atn(1#)
    Check "atn: large argument near pi/2", Near(Atn(1E+300), 2 * Atn(1)), True
    Check "atn: negative", Atn(-1), -Atn(1)
    Check "cos: Cos(0)", Cos(0), 1
    CheckTyped "cos: returns Double", Cos(CInt(0)), 1#
    Check "cos: Cos(pi) = -1", Near(Cos(4 * Atn(1)), -1), True
    Check "sin: Sin(0)", Sin(0), 0
    Check "sin: Sin(pi/2) = 1", Near(Sin(2 * Atn(1)), 1), True
    Check "tan: Tan(0)", Tan(0), 0
    Check "tan: Tan(pi/4) = 1", Near(Tan(Atn(1)), 1), True
    Check "sin: string argument", Sin("0"), 0

    On Error Resume Next
    Dim r
    r = Cos("abc")
    CheckErr "cos: non-numeric string", 13
    ' Sin/Cos/Tan/Atn/Exp/Log/Sqr take a Double argument: Null is error 94.
    r = Sin(Null)
    CheckErr "sin: Null", 94
    r = Cos(Null)
    CheckErr "cos: Null", 94
    r = Tan(Null)
    CheckErr "tan: Null", 94
    r = Atn(Null)
    CheckErr "atn: Null", 94
    r = Exp(Null)
    CheckErr "exp: Null", 94
    r = Log(Null)
    CheckErr "log: Null", 94
    r = Sqr(Null)
    CheckErr "sqr: Null", 94
    On Error GoTo 0
End Sub

Private Function Log10(X)
    Log10 = Log(X) / Log(10#)
End Function

Private Sub TestExpLog()
    Dim MyAngle, MyHSin, MyLog
    MyAngle = 1.3
    MyHSin = (Exp(MyAngle) - Exp(-1 * MyAngle)) / 2
    Check "exp-function: hyperbolic sine of 1.3", Near(MyHSin, 1.69838243729262), True
    MyLog = Log(MyAngle + Sqr(MyAngle * MyAngle + 1))
    Check "log-function: inverse hyperbolic sine of 1.3", Near(MyLog, 1.0784510589549), True
    Check "log-function: Log10(1000)", Near(Log10(1000), 3), True
    Check "log-function: Log10(10)", Log10(10), 1

    Check "exp: Exp(0)", Exp(0), 1
    Check "exp: Exp(1)", CStr(Exp(1)), "2.71828182845905"
    CheckTyped "exp: returns Double", Exp(CInt(0)), 1#
    Check "exp: Exp(709) finite", Exp(709) > 1E+307, True
    Check "exp: Exp(-1000) underflows to 0", Exp(-1000), 0
    Check "log: Log(1)", Log(1), 0
    Check "log: Log(Exp(1))", Near(Log(Exp(1)), 1), True
    Check "log: Log(Exp(2.5))", Near(Log(Exp(2.5)), 2.5), True
    CheckTyped "log: returns Double", Log(CInt(1)), 0#

    On Error Resume Next
    Dim r
    r = Log(0)
    CheckErr "log: Log(0)", 5
    r = Log(-1)
    CheckErr "log: Log(-1)", 5
    r = Exp(710)
    CheckErr "exp-function: Exp(710) overflow", 6
    r = Log("abc")
    CheckErr "log: non-numeric string", 13
    On Error GoTo 0
End Sub

Private Sub TestSqr()
    Dim MySqr
    MySqr = Sqr(4)
    Check "sqr-function: Sqr(4)", MySqr, 2
    MySqr = Sqr(23)
    Check "sqr-function: Sqr(23)", CStr(MySqr), "4.79583152331272"
    MySqr = Sqr(0)
    Check "sqr-function: Sqr(0)", MySqr, 0
    On Error Resume Next
    MySqr = Sqr(-4)
    CheckErr "sqr-function: Sqr(-4)", 5
    MySqr = Sqr(-1)
    CheckErr "sqr: Sqr(-1)", 5
    MySqr = Sqr("x")
    CheckErr "sqr: non-numeric string", 13
    On Error GoTo 0

    CheckTyped "sqr: returns Double", Sqr(CInt(9)), 3#
    Check "sqr: string argument", Sqr("16"), 4
    Check "sqr: Sqr(2)^2", Near(Sqr(2) ^ 2, 2), True
End Sub

Private Sub TestSgn()
    Dim MyVar1, MyVar2, MyVar3, MySign
    MyVar1 = 12: MyVar2 = -2.4: MyVar3 = 0
    MySign = Sgn(MyVar1)
    Check "sgn-function: Sgn(12)", MySign, 1
    MySign = Sgn(MyVar2)
    Check "sgn-function: Sgn(-2.4)", MySign, -1
    MySign = Sgn(MyVar3)
    Check "sgn-function: Sgn(0)", MySign, 0

    CheckTyped "sgn: returns Integer (Double arg)", Sgn(-2.4), CInt(-1)
    CheckTyped "sgn: returns Integer (Long arg)", Sgn(100000), CInt(1)
    CheckTyped "sgn: returns Integer (Currency arg)", Sgn(CCur(-3)), CInt(-1)
    Check "sgn: tiny positive", Sgn(1E-300), 1
    Check "sgn: string argument", Sgn("-5"), -1
    Check "sgn: Empty", Sgn(Empty), 0
    On Error Resume Next
    Dim r
    r = Sgn("abc")
    CheckErr "sgn: non-numeric string", 13
    On Error GoTo 0
End Sub

Private Sub TestIntFix()
    Dim MyNumber
    MyNumber = Int(99.8)
    Check "int-fix-functions: Int(99.8)", MyNumber, 99
    MyNumber = Fix(99.2)
    Check "int-fix-functions: Fix(99.2)", MyNumber, 99
    MyNumber = Int(-99.8)
    Check "int-fix-functions: Int(-99.8)", MyNumber, -100
    MyNumber = Fix(-99.8)
    Check "int-fix-functions: Fix(-99.8)", MyNumber, -99
    MyNumber = Int(-99.2)
    Check "int-fix-functions: Int(-99.2)", MyNumber, -100
    MyNumber = Fix(-99.2)
    Check "int-fix-functions: Fix(-99.2)", MyNumber, -99
    Check "int-fix-functions: Int(-8.4)", Int(-8.4), -9
    Check "int-fix-functions: Fix(-8.4)", Fix(-8.4), -8
    Check "int-fix-functions: Fix = Sgn*Int(Abs)", Fix(-7.6), Sgn(-7.6) * Int(Abs(-7.6))

    CheckTyped "int: Double keeps type", Int(2.7), 2#
    CheckTyped "int: Integer keeps type", Int(CInt(5)), CInt(5)
    CheckTyped "int: Long keeps type", Int(70000), 70000
    CheckTyped "int: Single keeps type", Int(CSng(-1.5)), CSng(-2)
    CheckTyped "int: Currency keeps type", Int(CCur(-1.5)), CCur(-2)
    CheckTyped "int: Decimal keeps type", Int(CDec(-2.5)), CDec(-3)
    CheckTyped "fix: Currency keeps type", Fix(CCur(-1.5)), CCur(-1)
    CheckTyped "int: String gives Double", Int("3.7"), 3#
    CheckTyped "fix: String gives Double", Fix("-3.7"), -3#
    CheckTyped "int: Date keeps type", Int(#1/2/2000 6:00:00 PM#), #1/2/2000#
    Check "int-fix-functions: Int(Null)", Int(Null), Null
    Check "int-fix-functions: Fix(Null)", Fix(Null), Null
    Check "int: Empty", Int(Empty), 0
    Check "int: large double", Int(12345678901.5), 12345678901#
    Check "int: -0.5", Int(-0.5), -1
    Check "fix: -0.5", Fix(-0.5), 0

    On Error Resume Next
    Dim r
    r = Int("abc")
    CheckErr "int: non-numeric string", 13
    r = Fix("abc")
    CheckErr "fix: non-numeric string", 13
    On Error GoTo 0
End Sub

Private Sub TestRound()
    Check "round-function: Round(0.12335,4)", Round(0.12335, 4), 0.1234
    Check "round-function: Round(0.12345,4)", Round(0.12345, 4), 0.1234
    Check "round-function: Round(0.12355,4)", Round(0.12355, 4), 0.1236
    Check "round-function: Round(0.12365,4)", Round(0.12365, 4), 0.1236
    Check "round-function: Round(0.00005,4)", Round(0.00005, 4), 0

    ' Banker's rounding to an integer.
    Check "round: Round(0.5)", Round(0.5), 0
    Check "round: Round(1.5)", Round(1.5), 2
    Check "round: Round(2.5)", Round(2.5), 2
    Check "round: Round(3.5)", Round(3.5), 4
    Check "round: Round(-2.5)", Round(-2.5), -2
    Check "round: Round(-3.5)", Round(-3.5), -4
    Check "round: Round(2.6)", Round(2.6), 3
    Check "round: Round(-2.4)", Round(-2.4), -2
    Check "round: Round(2.567, 2)", Round(2.567, 2), 2.57
    Check "round: Round(12.345, 1)", Round(12.345, 1), 12.3
    Check "round: Round(0.25, 1)", Round(0.25, 1), 0.2
    Check "round: Round(0.35, 1)", Round(0.35, 1), 0.4
    Check "round: zero places explicit", Round(7.5, 0), 8

    CheckTyped "round: Double keeps type", Round(2.5), 2#
    CheckTyped "round: Integer keeps type", Round(CInt(7)), CInt(7)
    CheckTyped "round: Long keeps type", Round(70000), 70000
    CheckTyped "round: Single keeps type", Round(CSng(1.5)), CSng(2)
    CheckTyped "round: Currency keeps type", Round(CCur(2.5)), CCur(2)
    CheckTyped "round: Currency places", Round(CCur(1.2345), 2), CCur(1.23)
    CheckTyped "round: Decimal keeps type", Round(CDec(2.5)), CDec(2)
    CheckTyped "round: String gives Double", Round("2.5"), 2#
    Check "round: Null", Round(Null), Null
    Check "round: Empty", Round(Empty), 0

    On Error Resume Next
    Dim r
    r = Round(1.5, -1)
    CheckErr "round: negative places", 5
    r = Round("abc")
    CheckErr "round: non-numeric string", 13
    On Error GoTo 0
End Sub

Private Sub TestRnd()
    Dim i As Integer, x As Single, okRange As Boolean, okDie As Boolean
    Dim MyValue As Integer
    CheckTyped "rnd-function: returns Single", TypeName(Rnd), "Single"
    okRange = True: okDie = True
    For i = 1 To 200
        x = Rnd
        If x < 0 Or x >= 1 Then okRange = False
        MyValue = Int((6 * Rnd) + 1)
        If MyValue < 1 Or MyValue > 6 Then okDie = False
    Next i
    Check "rnd-function: 0 <= Rnd < 1", okRange, True
    Check "rnd-function: die value 1..6", okDie, True

    ' Rnd(0) repeats the most recent number.
    x = Rnd
    Check "rnd-function: Rnd(0) = last number", Rnd(0), x
    Check "rnd-function: Rnd(0) again", Rnd(0), x
    ' Rnd(positive) gives the next number, same as Rnd.
    Check "rnd-function: Rnd(1) in range", (Rnd(1) >= 0) And (Rnd(1) < 1), True

    ' Negative argument: same number every time for the same seed.
    Dim n1 As Single, n2 As Single, n3 As Single
    n1 = Rnd(-5)
    n2 = Rnd(-5)
    Check "rnd-function: Rnd(-5) repeatable", n1, n2
    n3 = Rnd(-7)
    Check "rnd-function: Rnd(-5) <> Rnd(-7)", n1 <> n3, True
    Check "rnd-function: Rnd(-5) in range", (n1 >= 0) And (n1 < 1), True

    ' Rnd(-1) followed by Randomize n repeats the sequence.
    Dim a1 As Single, a2 As Single, a3 As Single
    Dim b1 As Single, b2 As Single, b3 As Single
    Rnd -1
    Randomize 42
    a1 = Rnd: a2 = Rnd: a3 = Rnd
    Rnd -1
    Randomize 42
    b1 = Rnd: b2 = Rnd: b3 = Rnd
    Check "randomize-statement: repeated sequence 1", b1, a1
    Check "randomize-statement: repeated sequence 2", b2, a2
    Check "randomize-statement: repeated sequence 3", b3, a3
    Check "randomize-statement: sequence varies", (a1 <> a2) Or (a2 <> a3), True

    ' After Randomize with no argument Rnd still lies in [0, 1).
    Randomize
    x = Rnd
    Check "randomize-statement: Randomize then Rnd in range", (x >= 0) And (x < 1), True
End Sub

Private Sub TestPartition()
    ' Ranges are padded to the width of Stop + 1 plus one character.
    Check "partition-function: before first (20,199,10)", Partition(5, 20, 199, 10), "    :  19"
    Check "partition-function: first range (20,199,10)", Partition(20, 20, 199, 10), "  20:  29"
    Check "partition-function: last range (20,199,10)", Partition(195, 20, 199, 10), " 190: 199"
    Check "partition-function: after last (20,199,10)", Partition(200, 20, 199, 10), " 200:    "
    Check "partition-function: before first (100,1010,20)", Partition(50, 100, 1010, 20), "     :   99"
    Check "partition-function: first range (100,1010,20)", Partition(100, 100, 1010, 20), "  100:  119"
    Check "partition-function: last range (100,1010,20)", Partition(1005, 100, 1010, 20), " 1000: 1010"
    Check "partition-function: after last (100,1010,20)", Partition(2000, 100, 1010, 20), " 1011:     "
    Check "partition-function: interval 1", Partition(100, 0, 1000, 1), "  100:  100"
    Check "partition-function: 0..500 by 50", Partition(267, 0, 500, 50), " 250: 299"
    Check "partition-function: first of 0..500 by 50", Partition(0, 0, 500, 50), "   0:  49"
    Check "partition-function: Null number", Partition(Null, 0, 500, 50), Null
    CheckTyped "partition-function: returns String", Partition(267, 0, 500, 50), " 250: 299"

    On Error Resume Next
    Dim r
    r = Partition(5, -1, 100, 10)
    CheckErr "partition-function: start < 0", 5
    r = Partition(5, 10, 10, 1)
    CheckErr "partition-function: stop = start", 5
    r = Partition(5, 0, 100, 0)
    CheckErr "partition-function: interval < 1", 5
    On Error GoTo 0
End Sub

Private Sub TestDerived()
    Dim X As Double, Pi As Double
    Pi = 4 * Atn(1)
    X = 1.3
    Check "derived-math-functions: Sec", Near(1 / Cos(X), 3.73833412707544), True
    Check "derived-math-functions: Cosec", Near(1 / Sin(X), 1.0378200456748), True
    Check "derived-math-functions: Cotan", Near(1 / Tan(X), 0.277615646541125), True
    X = 0.5
    Check "derived-math-functions: Arcsin(0.5)", Near(Atn(X / Sqr(-X * X + 1)), Pi / 6), True
    Check "derived-math-functions: Arccos(0.5)", Near(Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1), Pi / 3), True
    X = 2
    Check "derived-math-functions: Arcsec(2)", Near(Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1)), 2.42786827464503), True
    X = 1
    Check "derived-math-functions: Arccotan(1)", Near(Atn(X) + 2 * Atn(1), 3 * Pi / 4), True
    X = 1.3
    Check "derived-math-functions: HSin", Near((Exp(X) - Exp(-X)) / 2, 1.69838243729262), True
    X = 1
    Check "derived-math-functions: HCos(1)", Near((Exp(X) + Exp(-X)) / 2, 1.54308063481524), True
    X = 0.5
    Check "derived-math-functions: HTan(0.5)", Near((Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X)), 0.46211715726001), True
    Check "derived-math-functions: HSec(0.5)", Near(2 / (Exp(X) + Exp(-X)), 1 / ((Exp(X) + Exp(-X)) / 2)), True
    X = 1.3
    Check "derived-math-functions: HArcsin", Near(Log(X + Sqr(X * X + 1)), 1.0784510589549), True
    X = 2
    Check "derived-math-functions: HArccos(2)", Near(Log(X + Sqr(X * X - 1)), 1.31695789692482), True
    X = 0.5
    Check "derived-math-functions: HArctan(0.5)", Near(Log((1 + X) / (1 - X)) / 2, 0.549306144334055), True
    X = 3
    Check "derived-math-functions: HArccotan(3)", Near(Log((X + 1) / (X - 1)) / 2, 0.346573590279973), True
    Check "derived-math-functions: LogN(8, 2)", Near(Log(8) / Log(2), 3), True
End Sub

Private Sub TestDepreciation()
    CheckTyped "sln-function: returns Double", SLN(2400, 300, 10), 210#
    Check "sln-function: SLN(30000,7500,10)", SLN(30000, 7500, 10), 2250
    Check "syd-function: SYD(30000,7500,10,1)", Near(SYD(30000, 7500, 10, 1), 4090.90909090909), True
    Check "syd-function: SYD(30000,7500,10,10)", Near(SYD(30000, 7500, 10, 10), 409.090909090909), True
    CheckTyped "syd-function: returns Double", TypeName(SYD(30000, 7500, 10, 1)), "Double"

    Check "ddb-function: year 1", Near(DDB(2400, 300, 10, 1), 480), True
    Check "ddb-function: year 2", Near(DDB(2400, 300, 10, 2), 384), True
    Check "ddb-function: factor 1.5", Near(DDB(2400, 300, 10, 1, 1.5), 360), True
    Check "ddb-function: monthly", Near(DDB(2400, 300, 120, 1), 40), True
    Check "ddb-function: daily", Near(DDB(2400, 300, 3650, 1), 1.31506849315068), True
    Check "ddb-function: year 10 limited by salvage", Near(DDB(2400, 300, 10, 10), 22.1225472, 0.0000001), True
    CheckTyped "ddb-function: returns Double", TypeName(DDB(2400, 300, 10, 1)), "Double"
End Sub

Private Sub TestAnnuity()
    Check "pmt-function: PMT(0.08/12,10,10000)", Near(Pmt(0.08 / 12, 10, 10000), -1037.03208935916), True
    Check "pmt-function: 30-year mortgage", Near(Pmt(0.1 / 12, 360, 100000), -877.571570088799), True
    Check "pmt-function: zero rate", Near(Pmt(0, 10, 1000), -100), True
    CheckTyped "pmt-function: returns Double", TypeName(Pmt(0.01, 10, 1000)), "Double"

    Check "fv-function: FV(0.06/12,10,-200,-500,1)", Near(FV(0.06 / 12, 10, -200, -500, 1), 2581.40337406014), True
    Check "fv-function: FV(0.1,5,-100)", Near(FV(0.1, 5, -100), 610.51), True
    Check "fv-function: zero rate", Near(FV(0, 5, -100, -1000), 1500), True

    Check "pv-function: PV(0.08/12,240,500)", Near(PV(0.08 / 12, 240, 500), -59777.1458511878), True
    Check "pv-function: PV(0.05,10,-1000)", Near(PV(0.05, 10, -1000), 7721.73492918482), True

    Check "nper-function: NPer(0.01,-100,-1000,10000,1)", Near(NPer(0.12 / 12, -100, -1000, 10000, 1), 59.6738656742946), True
    Check "nper-function: NPer(0.01,-100,1000)", Near(NPer(0.01, -100, 1000), 10.5886444594232), True

    Check "rate-function: Rate(48,-200,8000)", Near(Rate(48, -200, 8000), 0.00770147248820189, 0.000001), True
    Check "rate-function: Rate(10,-100,700)", Near(Rate(10, -100, 700), 0.0707282083662951, 0.000001), True
    Check "rate-function: Rate with guess", Near(Rate(48, -200, 8000, 0, 0, 0.1), 0.00770147248820189, 0.000001), True

    Check "ipmt-function: IPmt(0.1/12,1,36,8000)", Near(IPmt(0.1 / 12, 1, 36, 8000), -66.6666666666667), True
    Check "ipmt-function: IPmt(0.1/12,3,36,8000)", Near(IPmt(0.1 / 12, 3, 36, 8000), -63.4621895664547), True
    Check "ppmt-function: PPmt(0.1/12,1,24,2000)", Near(PPmt(0.1 / 12, 1, 24, 2000), -75.6231860083666), True
    Check "ppmt-function: IPmt + PPmt = Pmt", Near(IPmt(0.1 / 12, 5, 24, 2000) + PPmt(0.1 / 12, 5, 24, 2000), Pmt(0.1 / 12, 24, 2000)), True

    On Error Resume Next
    Dim r
    r = IPmt(0.1 / 12, 0, 36, 8000)
    CheckErr "ipmt-function: period 0", 5
    r = PPmt(0.1 / 12, 37, 36, 8000)
    CheckErr "ppmt-function: period > nper", 5
    On Error GoTo 0
End Sub

Private Sub TestCashFlows()
    Dim Guess, RetRate, NetPVal, LoanAPR, InvAPR
    Static Values(5) As Double
    Guess = 0.1
    Values(0) = -70000
    Values(1) = 22000: Values(2) = 25000
    Values(3) = 28000: Values(4) = 31000
    RetRate = IRR(Values(), Guess) * 100
    Check "irr-function: example", Near(RetRate, 17.7435884422527, 0.000001), True
    Check "irr-function: Format of example", Format(RetRate, "#0.00"), "17.74"
    Check "irr-function: default guess", Near(IRR(Values()), 0.177435884422527, 0.000001), True
    CheckTyped "irr-function: returns Double", TypeName(IRR(Values())), "Double"

    RetRate = 0.0625
    NetPVal = NPV(RetRate, Values())
    Check "npv-function: example", Near(NetPVal, 19312.5702095352), True
    Check "npv-function: Format of example", Format(NetPVal, "###,##0.00"), "19,312.57"

    LoanAPR = 0.1
    InvAPR = 0.12
    RetRate = MIRR(Values(), LoanAPR, InvAPR)
    Check "mirr-function: example (6 elements)", Near(RetRate, 0.148014600794883), True
    Check "mirr-function: Format of example", Format(Abs(RetRate) * 100, "#0.00"), "14.80"

    Dim V2(5) As Double
    V2(0) = -120000: V2(1) = 39000: V2(2) = 30000
    V2(3) = 21000: V2(4) = 37000: V2(5) = 46000
    Check "mirr-function: six cash flows", Near(MIRR(V2(), 0.1, 0.12), 0.126094130365905), True
    V2(0) = -70000: V2(1) = 12000: V2(2) = 15000
    V2(3) = 18000: V2(4) = 21000: V2(5) = 26000
    Check "irr-function: six cash flows", Near(IRR(V2()), 0.0866309480365316, 0.000001), True

    Dim V3(3) As Double
    V3(0) = -10000: V3(1) = 3000: V3(2) = 4200: V3(3) = 6800
    Check "npv-function: four cash flows", Near(NPV(0.1, V3()), 1188.44341233522), True

    On Error Resume Next
    Dim r, AllPos(2) As Double
    AllPos(0) = 1: AllPos(1) = 2: AllPos(2) = 3
    r = IRR(AllPos())
    CheckErr "irr-function: no negative cash flow", 5
    On Error GoTo 0
End Sub
