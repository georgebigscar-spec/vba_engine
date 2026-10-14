Attribute VB_Name = "control_flow"
' Statements, procedures and arrays. Expected values follow the VBA language
' reference (if-then-else, select-case, doloop, fornext, on-error, ...)
' and real VBA 7 behaviour.
Option Explicit

Private Type Person
    Name As String
    Age As Integer
    Scores(1 To 3) As Long
End Type

Private Type EmployeeRecord
    ID As Integer
    Name As String * 20
    Phone As Long
    HireDate As Date
End Type

Private Enum SecurityLevel
    IllegalEntry = -1
    SecurityLevel1 = 0
    SecurityLevel2 = 1
End Enum

Private Enum Shade
    shRed
    shGreen = 5
    shBlue
End Enum

Private Enum InterfaceColors
    icMistyRose = &HE1E4FF
    icSlateGray = &H908070
End Enum

Const MyVar = 459
Private Const MyInt As Integer = 5
Const MyStr = "Hello", MyDouble As Double = 3.4567
Private Const DerivedConst = MyVar * 2

Private gTrace As String
Private gCount As Long

Sub Main()
    TestIfThenElse
    TestSingleLineIf
    TestSelectCase
    TestDoLoop
    TestWhileWend
    TestForNext
    TestForEach
    TestExit
    TestGoTo
    TestOnErrorResumeNext
    TestOnErrorGoTo
    TestResume
    TestErrObject
    TestErrorPropagation
    TestResumeWithoutError
    TestErrorStatement
    TestWith
    TestStatic
    TestConst
    TestDimReDim
    TestLBoundUBound
    TestErase
    TestArrayFunction
    TestCall
    TestByRefByVal
    TestOptional
    TestParamArray
    TestNamedArgs
    TestRecursion
    TestFunctionReturn
    TestType
    TestEnum
    TestLineSyntax
    TestTypeSuffixDim
    TestGoSub
End Sub

' ---------------------------------------------------------------- helpers

Private Sub Inc(x As Long)
    x = x + 1
End Sub

Private Sub IncByVal(ByVal x As Long)
    x = x + 1
End Sub

Private Sub IncVariant(v)
    v = v + 1
End Sub

Private Function KeepTotal(Number)
    ' Only the variable Accumulate preserves its value between calls.
    Static Accumulate
    Accumulate = Accumulate + Number
    KeepTotal = Accumulate
End Function

Private Static Function MyFunction(Arg1, Arg2, Arg3)
    ' All local variables preserve value between function calls.
    Dim Accumulate, Half, Calls
    Calls = Calls + 1
    Accumulate = Arg1 + Arg2 + Arg3
    Half = Accumulate / 2
    MyFunction = Half & "/" & Calls
End Function

Private Static Sub StaticSubCounter(ByRef result As Long)
    Dim n As Long
    n = n + 1
    result = n
End Sub

Private Function NonStaticCounter() As Long
    Dim n As Long
    n = n + 1
    NonStaticCounter = n
End Function

Private Function StaticArr() As Long
    Static a(2) As Long
    a(1) = a(1) + 10
    StaticArr = a(1)
End Function

Private Function ReturnTwice(Optional A)
    If IsMissing(A) Then
        ReturnTwice = Null
    Else
        ReturnTwice = A * 2
    End If
End Function

Private Function MyFunc(MyStr As String, Optional MyArg1 As _
    Integer = 5, Optional MyArg2 = "Dolly") As String
    MyFunc = MyStr & "|" & MyArg1 & "|" & MyArg2
End Function

Private Function OptTyped(Optional n As Long) As String
    OptTyped = n & ":" & IsMissing(n)
End Function

Private Function OptVariantDefault(Optional v As Variant = 7) As String
    OptVariantDefault = v & ":" & IsMissing(v)
End Function

Private Function CalcSum(ByVal FirstArg As Integer, ParamArray OtherArgs()) As Long
    Dim i As Long, s As Long
    s = FirstArg
    For i = LBound(OtherArgs) To UBound(OtherArgs)
        s = s + OtherArgs(i)
    Next i
    CalcSum = s
End Function

Private Function PCount(ParamArray args()) As String
    PCount = LBound(args) & ".." & UBound(args)
End Function

Private Function PFirst(ParamArray args()) As Variant
    PFirst = args(0)
End Function

Private Sub PInc(ParamArray args())
    args(0) = args(0) + 1
End Sub

Private Function Sub2(a As Long, b As Long) As Long
    Sub2 = a - b
End Function

Private Function Factorial(ByVal n As Long) As Double
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Private Function Fib(ByVal n As Long) As Long
    If n < 2 Then Fib = n Else Fib = Fib(n - 1) + Fib(n - 2)
End Function

Private Function CalculateSquareRoot(NumberArg As Double) As Double
    If NumberArg < 0 Then
        Exit Function
    Else
        CalculateSquareRoot = Sqr(NumberArg)
    End If
End Function

Private Function NoAssignVariant()
End Function

Private Function NoAssignLong() As Long
End Function

Private Function NoAssignString() As String
End Function

Private Function NoAssignBool() As Boolean
End Function

Private Function AccumulateInName() As Long
    Dim i As Long
    AccumulateInName = 1
    For i = 1 To 4
        AccumulateInName = AccumulateInName * 2
    Next i
End Function

Private Function ReturnsArray() As Variant
    ReturnsArray = Array(1, 2, 3)
End Function

Private Function ReturnsLongArray() As Variant
    Dim a(1 To 2) As Long
    a(1) = 5: a(2) = 6
    ReturnsLongArray = a
End Function

Private Sub SetTrace(ByVal s As String)
    gTrace = gTrace & s
End Sub

Private Function DivideNoHandler(ByVal a As Long, ByVal b As Long) As Long
    gTrace = gTrace & "d"
    DivideNoHandler = a \ b
    gTrace = gTrace & "D"
End Function

Private Sub RaiseInHandler()
    Dim x As Long
    On Error GoTo H
    x = 1 \ 0
    Exit Sub
H:
    Err.Raise 1234, "Inner", "from handler"
End Sub

Private Function HandledInside() As Long
    Dim x As Long
    On Error GoTo H
    x = 1 \ 0
    HandledInside = 1
    Exit Function
H:
    HandledInside = Err.Number
End Function

Private Function ResumeNextInside() As Long
    Dim x As Long
    On Error Resume Next
    x = 1 \ 0
    ResumeNextInside = Err.Number
    Exit Function
End Function

Private Sub BadResume()
    Resume Next
End Sub

Private Sub RaiseCustom()
    Err.Raise vbObjectError + 513, "MyProj.MyObject", "No ""bob"" allowed in your name"
End Sub

Private Function TwoLevelInner() As Long
    TwoLevelInner = CLng("abc")
End Function

Private Function TwoLevelMiddle() As Long
    gTrace = gTrace & "m"
    TwoLevelMiddle = TwoLevelInner()
    gTrace = gTrace & "M"
End Function

' ---------------------------------------------------------------- tests

Private Sub TestIfThenElse()
    Dim Number As Long, Digits As Long, s As String
    Number = 53
    If Number < 10 Then
        Digits = 1
    ElseIf Number < 100 Then
        Digits = 2
    Else
        Digits = 3
    End If
    Check "if-thenelse-statement: Digits", Digits, 2

    Number = 5
    If Number < 10 Then
        Digits = 1
    ElseIf Number < 100 Then
        Digits = 2
    Else
        Digits = 3
    End If
    Check "if-thenelse-statement: Digits 1", Digits, 1

    Number = 500
    If Number < 10 Then
        Digits = 1
    ElseIf Number < 100 Then
        Digits = 2
    Else
        Digits = 3
    End If
    Check "if-thenelse-statement: Digits 3", Digits, 3

    ' Null condition is treated as False
    Dim n As Variant
    n = Null
    If n Then s = "T" Else s = "F"
    Check "if-thenelse-statement: Null condition", s, "F"
    If n = 1 Then
        s = "T"
    Else
        s = "F"
    End If
    Check "if-thenelse-statement: Null comparison", s, "F"

    ' Nonzero numbers are True
    If -1.5 Then s = "T" Else s = "F"
    Check "if-thenelse-statement: nonzero true", s, "T"
    If "True" Then s = "T" Else s = "F"
    Check "if-thenelse-statement: string True", s, "T"

    ' No short circuit evaluation: both operands are evaluated
    Dim z As Long
    On Error Resume Next
    s = "none"
    If False And (1 \ z = 1) Then s = "T" Else s = "F"
    CheckErr "if-thenelse-statement: no short circuit", 11
    On Error GoTo 0

    ' Nested ElseIf with no matching branch and no Else
    s = "x"
    If 1 > 2 Then
        s = "a"
    ElseIf 2 > 3 Then
        s = "b"
    End If
    Check "if-thenelse-statement: no branch", s, "x"
End Sub

Private Sub TestSingleLineIf()
    Dim a As Long, b As Long, c As Long
    If 1 = 1 Then a = 1: b = 2 Else c = 3
    Check "if-thenelse-statement: single line a", a, 1
    Check "if-thenelse-statement: single line b", b, 2
    Check "if-thenelse-statement: single line c", c, 0
    a = 0: b = 0: c = 0
    If 1 = 2 Then a = 1: b = 2 Else c = 3: a = 9
    Check "if-thenelse-statement: single line else c", c, 3
    Check "if-thenelse-statement: single line else a", a, 9
    Check "if-thenelse-statement: single line else b", b, 0
    If a = 9 Then If b = 0 Then c = 42
    Check "if-thenelse-statement: nested single line", c, 42
End Sub

Private Function Classify(v As Variant) As String
    Select Case v
        Case 1 To 4, 7 To 9, 11, 13, Is > 100
            Classify = "A"
        Case Is < 0
            Classify = "neg"
        Case Else
            Classify = "else"
    End Select
End Function

Private Function ClassifyStr(v As Variant) As String
    Select Case v
        Case "everything", "nuts" To "soup"
            ClassifyStr = "B"
        Case Else
            ClassifyStr = "else"
    End Select
End Function

Private Sub TestSelectCase()
    Dim Number, s As String
    Number = 8
    Select Case Number
    Case 1 To 5
        s = "Between 1 and 5"
    Case 6, 7, 8
        s = "Between 6 and 8"
    Case 9 To 10
        s = "Greater than 8"
    Case Else
        s = "Not between 1 and 10"
    End Select
    Check "select-case-statement: Number 8", s, "Between 6 and 8"

    Check "select-case-statement: 3", Classify(3), "A"
    Check "select-case-statement: 4.5", Classify(4.5), "else"
    Check "select-case-statement: 13", Classify(13), "A"
    Check "select-case-statement: 12", Classify(12), "else"
    Check "select-case-statement: Is > 100", Classify(101), "A"
    Check "select-case-statement: 100", Classify(100), "else"
    Check "select-case-statement: everything", ClassifyStr("everything"), "B"
    Check "select-case-statement: pear in nuts To soup", ClassifyStr("pear"), "B"
    Check "select-case-statement: soup", ClassifyStr("soup"), "B"
    Check "select-case-statement: soupy", ClassifyStr("soupy"), "else"
    Check "select-case-statement: binary compare", ClassifyStr("Pear"), "else"
    Check "select-case-statement: Is < 0", Classify(-3), "neg"
    Check "select-case-statement: Null goes to Else", Classify(Null), "else"
    Check "select-case-statement: Null goes to Else (strings)", ClassifyStr(Null), "else"

    ' Only first matching clause runs
    Select Case 5
        Case 1 To 10: s = "first"
        Case 5: s = "second"
    End Select
    Check "select-case-statement: first match", s, "first"

    ' No match and no Case Else: nothing happens
    s = "unchanged"
    Select Case 99
        Case 1: s = "one"
    End Select
    Check "select-case-statement: no match", s, "unchanged"

    ' Select Case True idiom
    Dim score As Long
    score = 75
    Select Case True
        Case score >= 90: s = "A"
        Case score >= 70: s = "C"
        Case Else: s = "F"
    End Select
    Check "select-case-statement: Select Case True", s, "C"

    ' Is with other operators
    Select Case "b"
        Case Is >= "c": s = "ge c"
        Case Is <> "a": s = "not a"
    End Select
    Check "select-case-statement: Is <>", s, "not a"

    ' Test expression evaluated once
    gCount = 0
    Select Case NextCount()
        Case 5: s = "5"
        Case 4: s = "4"
        Case 1: s = "1"
    End Select
    Check "select-case-statement: evaluated once", s, "1"
    Check "select-case-statement: evaluated once count", gCount, 1
End Sub

Private Function NextCount() As Long
    gCount = gCount + 1
    NextCount = gCount
End Function

Private Sub TestDoLoop()
    Dim Check1 As Boolean, Counter As Long, Total As Long
    Check1 = True: Counter = 0: Total = 0
    Do
        Do While Counter < 20
            Counter = Counter + 1
            If Counter Mod 10 = 0 Then
                Check1 = False  ' stands in for the MsgBox answer "No"
                If Not Check1 Then Exit Do
            End If
        Loop
        Total = Total + Counter
        Counter = 0
    Loop Until Check1 = False
    Check "doloop-statement: Total", Total, 10

    Dim n As Long
    n = 0
    Do While False
        n = n + 1
    Loop
    Check "doloop-statement: Do While pre-test", n, 0
    n = 0
    Do
        n = n + 1
    Loop While False
    Check "doloop-statement: Loop While post-test", n, 1
    n = 0
    Do Until True
        n = n + 1
    Loop
    Check "doloop-statement: Do Until pre-test", n, 0
    n = 0
    Do
        n = n + 1
    Loop Until True
    Check "doloop-statement: Loop Until post-test", n, 1
    n = 0
    Do Until n >= 5
        n = n + 1
    Loop
    Check "doloop-statement: Do Until count", n, 5
    n = 0
    Do
        n = n + 1
        If n = 7 Then Exit Do
    Loop
    Check "doloop-statement: Exit Do", n, 7
    ' Null condition in Do While is treated as False
    Dim v As Variant
    v = Null
    n = 0
    Do While v
        n = n + 1
        Exit Do
    Loop
    Check "doloop-statement: Do While Null", n, 0
End Sub

Private Sub TestWhileWend()
    Dim Counter
    Counter = 0
    While Counter < 20
        Counter = Counter + 1
    Wend
    Check "whilewend-statement: Counter", Counter, 20
    Dim i As Long, j As Long, n As Long
    i = 0
    While i < 3
        j = 0
        While j < 4
            n = n + 1
            j = j + 1
        Wend
        i = i + 1
    Wend
    Check "whilewend-statement: nested", n, 12
End Sub

Private Sub TestForNext()
    Dim Words, Chars, MyString
    For Words = 10 To 1 Step -1
        For Chars = 0 To 9
            MyString = MyString & Chars
        Next Chars
        MyString = MyString & " "
    Next Words
    Check "fornext-statement: MyString length", Len(MyString), 110
    Check "fornext-statement: MyString start", Left(MyString, 11), "0123456789 "

    Dim i As Long, n As Long, s As String
    For i = 1 To 10 Step 2
        n = n + 1
    Next i
    Check "fornext-statement: Step 2 count", n, 5
    Check "fornext-statement: Step 2 var after", i, 11

    s = ""
    For i = 10 To 1 Step -3
        s = s & i & ","
    Next
    Check "fornext-statement: Step -3 values", s, "10,7,4,1,"
    Check "fornext-statement: Step -3 var after", i, -2

    n = 0
    For i = 5 To 1
        n = n + 1
    Next i
    Check "fornext-statement: start > end count", n, 0
    Check "fornext-statement: start > end var", i, 5

    Dim d As Double
    n = 0
    For d = 0 To 1 Step 0.25
        n = n + 1
    Next d
    Check "fornext-statement: Step 0.25 count", n, 5
    Check "fornext-statement: Step 0.25 var after", d, 1.25

    ' Bounds are evaluated once
    Dim lim As Long
    lim = 3: n = 0
    For i = 1 To lim
        lim = 10
        n = n + 1
    Next i
    Check "fornext-statement: modify end inside", n, 3

    Dim stp As Long
    stp = 1: s = ""
    For i = 1 To 5 Step stp
        stp = 2
        s = s & i
    Next i
    Check "fornext-statement: modify step inside", s, "12345"

    ' Changing the counter inside affects iteration
    n = 0
    For i = 1 To 10
        i = i + 1
        n = n + 1
    Next i
    Check "fornext-statement: modify counter count", n, 5
    Check "fornext-statement: modify counter after", i, 11

    ' Variant counter takes the type of the bounds
    Dim v As Variant
    For v = 1 To 2
    Next v
    CheckTyped "fornext-statement: Variant counter Integer", v, CInt(3)

    Dim j As Long

    ' Exit For leaves only the innermost loop
    n = 0
    For i = 1 To 3
        For j = 1 To 10
            If j = 2 Then Exit For
            n = n + 1
        Next j
    Next i
    Check "fornext-statement: Exit For inner only", n, 3
    Check "fornext-statement: outer var after Exit For", i, 4
    Check "fornext-statement: inner var at Exit For", j, 2

    ' Integer counter overflow at the end
    Dim k As Integer
    On Error Resume Next
    For k = 32766 To 32767
    Next k
    CheckErr "fornext-statement: Integer counter overflow", 6
    On Error GoTo 0
End Sub

Private Sub TestForEach()
    Dim a As Variant, x As Variant, total As Long
    a = Array(1, 2, 3, 4)
    For Each x In a
        total = total + x
    Next x
    Check "for-eachnext-statement: array sum", total, 10

    ' Element variable is a copy
    For Each x In a
        x = x * 100
    Next
    Check "for-eachnext-statement: array unchanged", a(0), 1

    ' Multi-dimensional arrays enumerate first index fastest
    Dim m(1, 1) As Long, s As String
    m(0, 0) = 1: m(1, 0) = 2: m(0, 1) = 3: m(1, 1) = 4
    For Each x In m
        s = s & x
    Next
    Check "for-eachnext-statement: 2D order", s, "1234"

    ' Collection
    Dim col As New Collection
    col.Add "Hello"
    col.Add "World"
    col.Add "!"
    s = ""
    For Each x In col
        s = s & x
    Next
    Check "for-eachnext-statement: collection", s, "HelloWorld!"

    Dim Found As Boolean, n As Long
    Found = False
    For Each x In col
        n = n + 1
        If x = "World" Then
            Found = True
            Exit For
        End If
    Next
    Check "for-eachnext-statement: Found", Found, True
    Check "for-eachnext-statement: Exit For count", n, 2

    ' Empty collection: body not run
    Dim empty1 As New Collection
    n = 0
    For Each x In empty1
        n = n + 1
    Next
    Check "for-eachnext-statement: empty collection", n, 0

    ' Typed array
    Dim strs(2) As String
    strs(0) = "a": strs(1) = "b": strs(2) = "c"
    s = ""
    For Each x In strs
        s = s & x
    Next
    Check "for-eachnext-statement: string array", s, "abc"

    ' Nested For Each
    s = ""
    Dim y As Variant
    For Each x In Array("a", "b")
        For Each y In Array(1, 2)
            s = s & x & y
        Next y
    Next x
    Check "for-eachnext-statement: nested", s, "a1a2b1b2"
End Sub

Private Function ExitDemo(ByVal target As Long) As String
    Dim i As Long, k As Long
    ExitDemo = "end"
    Do
        k = k + 1
        For i = 1 To 100
            Select Case i
            Case target: Exit For
            Case 29: Exit Do
            Case 54: ExitDemo = "sub": Exit Function
            End Select
        Next i
        If k = 2 Then ExitDemo = "for" & i: Exit Do
    Loop
    If ExitDemo = "end" Then ExitDemo = "do" & i
End Function

Private Sub TestExit()
    Check "exit-statement: Exit For", ExitDemo(7), "for7"
    Check "exit-statement: Exit Do", ExitDemo(40), "do29"
    Dim d As Double
    d = 4
    Check "function-statement: CalculateSquareRoot", CalculateSquareRoot(d), 2
    d = -4
    Check "function-statement: Exit Function returns default", CalculateSquareRoot(d), 0
    gTrace = ""
    ExitSubDemo
    Check "exit-statement: Exit Sub", gTrace, "a"
End Sub

Private Sub ExitSubDemo()
    gTrace = gTrace & "a"
    Exit Sub
    gTrace = gTrace & "b"
End Sub

Private Sub TestGoTo()
    Dim n As Long, s As String
    n = 0
Again:
    n = n + 1
    If n < 5 Then GoTo Again
    Check "goto-statement: backward", n, 5

    GoTo Skip
    s = "not skipped"
Skip:
    Check "goto-statement: forward", s, ""
End Sub

Private Sub TestGoSub()
    Dim Num
    Num = 10
    If Num > 0 Then GoSub MyRoutine
    Check "gosubreturn-statement: Num", Num, 5
    GoSub MyRoutine
    Check "gosubreturn-statement: twice", Num, 2.5

    ' Return without GoSub is error 3
    On Error Resume Next
    Return
    CheckErr "gosubreturn-statement: Return without GoSub", 3
    On Error GoTo 0
    Exit Sub
MyRoutine:
    Num = Num / 2
    Return
End Sub

Private Sub TestOnErrorResumeNext()
    Dim x As Long, z As Long
    On Error Resume Next
    x = 5
    x = 1 \ z
    Check "on-error-statement: Resume Next keeps value", x, 5
    Check "on-error-statement: Err.Number 11", Err.Number, 11
    Check "on-error-statement: Err.Description", Err.Description, "Division by zero"
    Err.Clear
    Check "on-error-statement: Err.Clear", Err.Number, 0
    Check "on-error-statement: Err.Description cleared", Err.Description, ""

    ' Err is not cleared by a successful statement
    x = 1 \ z
    x = 2
    Check "on-error-statement: Err persists", Err.Number, 11

    ' On Error GoTo 0 clears Err
    On Error GoTo 0
    Check "on-error-statement: GoTo 0 clears Err", Err.Number, 0

    ' Error in a callee without handler skips the rest of the callee
    On Error Resume Next
    gTrace = ""
    x = DivideNoHandler(1, 0)
    gTrace = gTrace & "c"
    Check "on-error-statement: callee aborted", gTrace, "dc"
    CheckErr "on-error-statement: callee error number", 11

    ' Through two levels
    gTrace = ""
    x = TwoLevelMiddle()
    Check "on-error-statement: two levels trace", gTrace, "m"
    CheckErr "on-error-statement: two levels error", 13

    ' A procedure that handled its own error returns normally
    Err.Clear
    Check "on-error-statement: handled inside", HandledInside(), 11
    Check "on-error-statement: Err cleared after Exit Function", Err.Number, 0
    Check "on-error-statement: resume next inside", ResumeNextInside(), 11
    Check "on-error-statement: Err cleared after Resume Next callee exit", Err.Number, 0

    ' Error raised inside a handler propagates to the caller
    RaiseInHandler
    Check "on-error-statement: error in handler number", Err.Number, 1234
    Check "on-error-statement: error in handler source", Err.Source, "Inner"
    Err.Clear

    Dim arr(2) As Long
    x = arr(5)
    CheckErr "on-error-statement: subscript 9", 9
    Dim i As Integer
    i = 40000
    CheckErr "on-error-statement: overflow 6", 6
    x = CLng("zz")
    CheckErr "on-error-statement: type mismatch 13", 13
    Dim v As Variant
    v = Null
    x = v
    CheckErr "on-error-statement: invalid use of Null 94", 94
    On Error GoTo 0
End Sub

Private Function GoToHandler(ByVal z As Long) As String
    Dim x As Long
    On Error GoTo ErrorHandler
    x = 10 \ z
    GoToHandler = "ok" & x
    Exit Function
ErrorHandler:
    GoToHandler = "err" & Err.Number
End Function

Private Function NestedCallHandler() As String
    Dim x As Long
    On Error GoTo H
    gTrace = ""
    x = DivideNoHandler(1, 0)
    gTrace = gTrace & "after"
    NestedCallHandler = gTrace
    Exit Function
H:
    gTrace = gTrace & "H" & Err.Number
    Resume Next
End Function

Private Function HandlerWithResumeNext() As String
    Dim s As String, z As Long, x As Long
    On Error GoTo H
    s = "a"
    x = 1 \ z
    s = s & "b"
    HandlerWithResumeNext = s & Err.Number
    Exit Function
H:
    s = s & "h"
    Resume Next
End Function

Private Sub TestOnErrorGoTo()
    Check "on-error-statement: GoTo label ok", GoToHandler(2), "ok5"
    Check "on-error-statement: GoTo label err", GoToHandler(0), "err11"
    Check "on-error-statement: callee error to caller handler", NestedCallHandler(), "dH11after"
    Check "on-error-statement: Resume Next clears Err", HandlerWithResumeNext(), "ahb0"
    Dim e As Long
    e = vbObjectError + 1052
    Check "on-error-statement: vbObjectError + 1052", e, -2147220452
End Sub

Private Function ResumeRetry() As String
    Dim z As Long, x As Long, tries As Long
    On Error GoTo H
    x = 12 \ z
    ResumeRetry = x & "/" & tries
    Exit Function
H:
    tries = tries + 1
    z = 3
    Resume
End Function

Private Function ResumeLabel() As String
    Dim z As Long, x As Long, s As String
    On Error GoTo H
    s = "1"
    x = 12 \ z
    s = s & "2"
Continue:
    s = s & "3"
    ResumeLabel = s & ":" & Err.Number
    Exit Function
H:
    s = s & "h"
    Resume Continue
End Function

Private Sub TestResume()
    Check "resume-statement: Resume retries", ResumeRetry(), "4/1"
    Check "resume-statement: Resume label", ResumeLabel(), "1h3:0"
End Sub

Private Sub TestErrObject()
    On Error Resume Next
    RaiseCustom
    Check "raise-method: Number", Err.Number, vbObjectError + 513
    Check "raise-method: Number value", Err.Number, -2147220991
    Check "raise-method: Source", Err.Source, "MyProj.MyObject"
    Check "raise-method: Description", Err.Description, "No ""bob"" allowed in your name"
    Err.Clear
    Check "raise-method: Source cleared", Err.Source, ""

    Err.Raise 1000
    Check "raise-method: user number", Err.Number, 1000
    Check "raise-method: default description", Err.Description, "Application-defined or object-defined error"
    Err.Clear

    Err.Raise 11
    Check "raise-method: 11 description", Err.Description, "Division by zero"
    Err.Clear
    Err.Raise 9
    Check "raise-method: 9 description", Err.Description, "Subscript out of range"
    Err.Clear
    Err.Raise 13
    Check "raise-method: 13 description", Err.Description, "Type mismatch"
    Err.Clear
    Err.Raise 5
    Check "raise-method: 5 description", Err.Description, "Invalid procedure call or argument"
    Err.Clear
    Err.Raise 6
    Check "raise-method: 6 description", Err.Description, "Overflow"
    Err.Clear
    Err.Raise 94
    Check "raise-method: 94 description", Err.Description, "Invalid use of Null"
    Err.Clear

    Err.Raise 0
    CheckErr "raise-method: Raise 0 is error 5", 5

    Err.Number = 42
    Check "number-property: set Number", Err.Number, 42
    Err.Clear

    Err.Raise 70, , "custom text"
    Check "raise-method: custom description", Err.Description, "custom text"
    CheckErr "raise-method: 70", 70

    ' Err is the default property Number
    Err.Raise 7
    Check "number-property: Err default", Err, 7
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub TestErrorPropagation()
    Dim s As String
    On Error GoTo H
    gTrace = ""
    s = CStr(TwoLevelMiddle())
    gTrace = gTrace & "X"
After:
    Check "on-error-statement: propagate to caller handler", gTrace, "mH13"
    Exit Sub
H:
    gTrace = gTrace & "H" & Err.Number
    Resume After
End Sub

Private Sub TestResumeWithoutError()
    On Error Resume Next
    BadResume
    CheckErr "resume-without-error-error-20", 20
    On Error GoTo 0
End Sub

Private Sub TestErrorStatement()
    On Error Resume Next
    Error 11
    Check "error-statement: Error 11 number", Err.Number, 11
    Check "error-statement: Error 11 description", Err.Description, "Division by zero"
    Err.Clear
    Error 1000
    CheckErr "error-statement: Error 1000", 1000
    On Error GoTo 0
End Sub


Private Sub TestWith()
    Dim col As New Collection
    With col
        .Add "a"
        .Add "b", "kb"
        Check "with-statement: Count", .Count, 2
        Check "with-statement: Item", .Item("kb"), "b"
    End With
    Dim p As Person
    With p
        .Name = "Ann"
        .Age = 30
        .Scores(2) = 7
    End With
    Check "with-statement: UDT Name", p.Name, "Ann"
    Check "with-statement: UDT Age", p.Age, 30
    Check "with-statement: UDT array member", p.Scores(2), 7

    ' With expression is evaluated once
    Dim outer As New Collection, inner As New Collection
    inner.Add 5
    outer.Add inner
    With outer(1)
        .Add 6
        Check "with-statement: nested object count", .Count, 2
    End With
    Check "with-statement: object modified", inner.Count, 2
End Sub

Private Sub TestStatic()
    Check "static-statement: KeepTotal 1", KeepTotal(1), 1
    Check "static-statement: KeepTotal 2", KeepTotal(2), 3
    Check "static-statement: KeepTotal 3", KeepTotal(10), 13
    Check "static-statement: Static Function 1", MyFunction(1, 2, 3), "3/1"
    Check "static-statement: Static Function 2", MyFunction(2, 2, 2), "3/2"
    Dim r As Long
    StaticSubCounter r
    StaticSubCounter r
    StaticSubCounter r
    Check "static-statement: Static Sub", r, 3
    Check "static-statement: non static 1", NonStaticCounter(), 1
    Check "static-statement: non static 2", NonStaticCounter(), 1
    Check "static-statement: static array 1", StaticArr(), 10
    Check "static-statement: static array 2", StaticArr(), 20
End Sub

Private Sub TestConst()
    Check "const-statement: MyVar", MyVar, 459
    CheckTyped "const-statement: MyVar type", MyVar, CInt(459)
    CheckTyped "const-statement: MyInt", MyInt, CInt(5)
    Check "const-statement: MyStr", MyStr, "Hello"
    CheckTyped "const-statement: MyDouble", MyDouble, CDbl(3.4567)
    Check "const-statement: derived", DerivedConst, 918
    Const LocalLong = 5&
    CheckTyped "const-statement: Long literal", LocalLong, CLng(5)
    Const LocalDbl = 1.5
    CheckTyped "const-statement: Double literal", LocalDbl, CDbl(1.5)
    Const BigConst = 100000
    CheckTyped "const-statement: large literal is Long", BigConst, CLng(100000)
    Const TypedStr As String = "x" & "y"
    CheckTyped "const-statement: typed String", TypedStr, "xy"
End Sub

Private Sub TestDimReDim()
    Dim MyArray() As Integer, I As Long
    ReDim MyArray(5)
    For I = 1 To 5
        MyArray(I) = I
    Next I
    ReDim MyArray(10)
    Check "redim-statement: ReDim clears", MyArray(3), 0
    For I = 1 To 10
        MyArray(I) = I
    Next I
    ReDim Preserve MyArray(15)
    Check "redim-statement: Preserve keeps", MyArray(10), 10
    Check "redim-statement: Preserve new element", MyArray(15), 0
    Check "redim-statement: Preserve UBound", UBound(MyArray), 15
    ReDim Preserve MyArray(3)
    Check "redim-statement: Preserve shrink UBound", UBound(MyArray), 3
    Check "redim-statement: Preserve shrink value", MyArray(3), 3

    Dim X() As Long
    ReDim X(10, 10, 10)
    X(2, 3, 4) = 99
    ReDim Preserve X(10, 10, 15)
    Check "redim-statement: Preserve last dim", X(2, 3, 4), 99
    Check "redim-statement: Preserve last dim UBound", UBound(X, 3), 15
    On Error Resume Next
    ReDim Preserve X(12, 10, 15)
    CheckErr "redim-statement: Preserve first dim error 9", 9
    On Error GoTo 0

    Dim b() As String
    ReDim b(2 To 5)
    Check "redim-statement: LBound 2", LBound(b), 2
    Check "redim-statement: UBound 5", UBound(b), 5
    Check "redim-statement: String default", b(3), ""

    Dim v As Variant
    ReDim v(3)
    Check "redim-statement: Variant ReDim", UBound(v), 3
    Check "redim-statement: Variant elements Empty", IsEmpty(v(0)), True
    Check "redim-statement: TypeName", TypeName(v), "Variant()"

    Dim la() As Long
    Check "dim-statement: TypeName Long()", TypeName(la), "Long()"
    Check "dim-statement: IsArray", IsArray(la), True

    ' Array assignment copies
    Dim c1() As Long, c2() As Long
    ReDim c1(1)
    c1(0) = 1
    c2 = c1
    c2(0) = 99
    Check "dim-statement: array copy", c1(0), 1
    Check "dim-statement: array copied value", c2(0), 99

    ' Subscript out of range
    Dim f(1 To 3) As Long
    On Error Resume Next
    f(0) = 1
    CheckErr "dim-statement: below LBound", 9
    f(4) = 1
    CheckErr "dim-statement: above UBound", 9
    On Error GoTo 0

    ' Multi-dimensional
    Dim g(1 To 3, 1 To 4) As Long, r As Long, cc As Long, total As Long
    For r = 1 To 3
        For cc = 1 To 4
            g(r, cc) = r * 10 + cc
        Next cc
    Next r
    Check "dim-statement: 2D element", g(2, 3), 23
    For r = 1 To 3
        For cc = 1 To 4
            total = total + g(r, cc)
        Next cc
    Next r
    Check "dim-statement: 2D sum", total, 270

    ' Default values
    Dim di(2) As Integer, dd(2) As Double, db(2) As Boolean, dv(2) As Variant
    Check "dim-statement: Integer default", di(1), 0
    Check "dim-statement: Boolean default", db(1), False
    Check "dim-statement: Variant default Empty", IsEmpty(dv(1)), True
    CheckTyped "dim-statement: Double default type", dd(1), CDbl(0)

    ' Function returning arrays
    Dim ra As Variant, rl() As Long
    ra = ReturnsArray()
    Check "function-statement: returns Variant array", ra(2), 3
    rl = ReturnsLongArray()
    Check "function-statement: returns Long()", rl(2), 6
    Check "function-statement: returns Long() LBound", LBound(rl), 1
End Sub

Private Sub TestLBoundUBound()
    Dim Lower, Upper
    Dim MyArray(1 To 10, 5 To 15, 10 To 20)
    Dim AnyArray(10)
    Lower = LBound(MyArray, 1)
    Check "lbound-function: dim 1", Lower, 1
    Lower = LBound(MyArray, 3)
    Check "lbound-function: dim 3", Lower, 10
    Lower = LBound(AnyArray)
    Check "lbound-function: AnyArray", Lower, 0
    Check "lbound-function: dim 2", LBound(MyArray, 2), 5
    Upper = UBound(MyArray, 1)
    Check "ubound-function: dim 1", Upper, 10
    Upper = UBound(MyArray, 3)
    Check "ubound-function: dim 3", Upper, 20
    Upper = UBound(AnyArray)
    Check "ubound-function: AnyArray", Upper, 10
    CheckTyped "ubound-function: returns Long", UBound(AnyArray), CLng(10)

    Dim dyn() As Long, v As Variant
    On Error Resume Next
    Upper = UBound(MyArray, 4)
    CheckErr "ubound-function: dimension too large", 9
    Upper = UBound(MyArray, 0)
    CheckErr "ubound-function: dimension 0", 9
    Upper = UBound(dyn)
    CheckErr "ubound-function: unallocated", 9
    v = 5
    Upper = UBound(v)
    CheckErr "ubound-function: not an array", 13
    On Error GoTo 0
End Sub

Private Sub TestErase()
    Dim NumArray(10) As Integer
    Dim StrVarArray(10) As String
    Dim VarArray(10) As Variant
    Dim DynamicArray() As Integer
    ReDim DynamicArray(10)
    NumArray(3) = 7: StrVarArray(3) = "x": VarArray(3) = "y": DynamicArray(3) = 9
    Erase NumArray
    Check "erase-statement: Integer element 0", NumArray(3), 0
    Check "erase-statement: fixed array keeps bounds", UBound(NumArray), 10
    Erase StrVarArray
    Check "erase-statement: String element empty", StrVarArray(3), ""
    Erase VarArray
    Check "erase-statement: Variant element Empty", IsEmpty(VarArray(3)), True
    Erase DynamicArray
    Dim u As Long
    On Error Resume Next
    u = UBound(DynamicArray)
    CheckErr "erase-statement: dynamic array freed", 9
    On Error GoTo 0
    Dim m(1, 1) As Long
    m(1, 1) = 5
    Erase m
    Check "erase-statement: 2D element 0", m(1, 1), 0
End Sub

Private Sub TestArrayFunction()
    Dim A As Variant, B As Long, i As Long
    A = Array(10, 20, 30)
    B = A(2)
    Check "array-function: B", B, 30
    ReDim Preserve A(4)
    A(4) = 40
    Check "array-function: A(4)", A(4), 40
    Check "array-function: A(0) after Preserve", A(0), 10
    Check "array-function: A(3) Empty", IsEmpty(A(3)), True
    Check "array-function: UBound", UBound(A), 4

    Dim MyWeek, MyDay
    MyWeek = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    MyDay = MyWeek(2)
    Check "array-function: MyWeek(2) base 0", MyDay, "Wed"
    MyDay = MyWeek(4)
    Check "array-function: MyWeek(4) base 0", MyDay, "Fri"
    Check "array-function: LBound", LBound(MyWeek), 0
    Check "array-function: TypeName", TypeName(MyWeek), "Variant()"
    Dim e As Variant
    e = Array()
    Check "array-function: empty UBound", UBound(e), -1
    Check "array-function: empty LBound", LBound(e), 0
    CheckTyped "array-function: element subtype", Array(1, "a", 2.5)(0), CInt(1)
    CheckTyped "array-function: element subtype Double", Array(1, "a", 2.5)(2), CDbl(2.5)
    ' Nested arrays (jagged)
    Dim j As Variant
    j = Array(Array(1, 2), Array(3, 4, 5))
    Check "array-function: jagged", j(1)(2), 5
    Check "array-function: jagged UBound", UBound(j(1)), 2
End Sub

Private Sub TestCall()
    gTrace = ""
    Call SetTrace("a")
    SetTrace "b"
    Call SetTrace("c"): SetTrace ("d")
    Check "call-statement: trace", gTrace, "abcd"
    ' Calling a function as a statement discards the result
    gCount = 0
    Call NextCount
    NextCount
    Check "call-statement: function as statement", gCount, 2
End Sub

Private Sub TestByRefByVal()
    Dim x As Long
    x = 1
    Inc x
    Check "function-statement: ByRef default", x, 2
    Call Inc(x)
    Check "function-statement: Call ByRef", x, 3
    Inc (x)
    Check "function-statement: parenthesised arg is ByVal", x, 3
    x = 3
    Call Inc((x))
    Check "function-statement: double parens ByVal", x, 3
    x = 3
    IncByVal x
    Check "function-statement: ByVal", x, 3
    Dim v As Variant
    v = 10
    IncVariant v
    Check "function-statement: ByRef Variant", v, 11
    ' Array element passed ByRef
    Dim arr(2) As Long
    arr(1) = 5
    Inc arr(1)
    Check "function-statement: ByRef array element", arr(1), 6
    ' UDT member passed ByRef
    Dim p As Person
    Dim age As Long
    p.Scores(1) = 1
    Inc p.Scores(1)
    Check "function-statement: ByRef UDT member", p.Scores(1), 2
    ' Expression argument
    x = 3
    Inc x + 0
    Check "function-statement: expression arg", x, 3
    ' Type mismatch for ByRef with different type
    Check "function-statement: ByRef same variable", Sub2(x, x), 0
End Sub

Private Sub TestOptional()
    Check "ismissing-function: ReturnTwice()", IsNull(ReturnTwice()), True
    Check "ismissing-function: ReturnTwice(2)", ReturnTwice(2), 4
    Check "function-statement: MyFunc all", MyFunc("Hello", 2, "World"), "Hello|2|World"
    Check "function-statement: MyFunc second omitted", MyFunc("Test", , 5), "Test|5|5"
    Check "function-statement: MyFunc named", MyFunc(MyStr:="Hello ", MyArg1:=7), "Hello |7|Dolly"
    Check "function-statement: MyFunc defaults", MyFunc("x"), "x|5|Dolly"
    Check "ismissing-function: typed Optional", OptTyped(), "0:False"
    Check "ismissing-function: typed Optional given", OptTyped(3), "3:False"
    Check "ismissing-function: Variant with default", OptVariantDefault(), "7:False"
    Check "ismissing-function: Variant with default given", OptVariantDefault(1), "1:False"
End Sub

Private Sub TestParamArray()
    Check "function-statement: CalcSum", CalcSum(4, 3, 2, 1), 10
    Check "function-statement: CalcSum one arg", CalcSum(4), 4
    Check "paramarray: bounds", PCount(1, 2, 3), "0..2"
    Check "paramarray: no args", PCount(), "0..-1"
    Check "paramarray: first", PFirst("z", 1), "z"
    Dim x As Long
    x = 5
    PInc x
    Check "paramarray: elements are ByRef", x, 6
End Sub

Private Sub TestNamedArgs()
    Check "function-statement: named args reordered", Sub2(b:=1, a:=10), 9
    Check "function-statement: mixed positional named", Sub2(10, b:=4), 6
    Check "function-statement: MyFunc named third", MyFunc("a", MyArg2:="z"), "a|5|z"
End Sub

Private Sub TestRecursion()
    Check "function-statement: Factorial 10", Factorial(10), 3628800
    Check "function-statement: Factorial 1", Factorial(1), 1
    Check "function-statement: Fib 15", Fib(15), 610
End Sub

Private Sub TestFunctionReturn()
    Check "function-statement: unassigned Variant Empty", IsEmpty(NoAssignVariant()), True
    CheckTyped "function-statement: unassigned Long", NoAssignLong(), CLng(0)
    CheckTyped "function-statement: unassigned String", NoAssignString(), ""
    CheckTyped "function-statement: unassigned Boolean", NoAssignBool(), False
    Check "function-statement: name used as variable", AccumulateInName(), 16
    CheckTyped "function-statement: typed return converts", CalculateSquareRoot(9#), CDbl(3)
End Sub

Private Sub TestType()
    Dim MyRecord As EmployeeRecord
    MyRecord.ID = 12003
    Check "type-statement: ID", MyRecord.ID, 12003
    CheckTyped "type-statement: ID Integer", MyRecord.ID, CInt(12003)
    Check "type-statement: fixed string length", Len(MyRecord.Name), 20
    MyRecord.Name = "Bob"
    Check "type-statement: fixed string padded", MyRecord.Name, "Bob" & Space(17)
    Check "type-statement: Long default", MyRecord.Phone, 0
    Check "type-statement: Date default", MyRecord.HireDate, CDate(0)

    Dim p1 As Person, p2 As Person
    p1.Name = "Ann": p1.Age = 30: p1.Scores(3) = 9
    p2 = p1
    p2.Name = "Bob"
    p2.Scores(3) = 1
    Check "type-statement: copy semantics name", p1.Name, "Ann"
    Check "type-statement: copy semantics array", p1.Scores(3), 9
    Check "type-statement: copy name", p2.Name, "Bob"
    Check "type-statement: copy age", p2.Age, 30
    Check "type-statement: member array bounds", LBound(p1.Scores) & "-" & UBound(p1.Scores), "1-3"

    Dim people(1 To 2) As Person
    people(2).Name = "Cy"
    people(2).Scores(1) = 4
    Check "type-statement: array of UDT", people(2).Name, "Cy"
    Check "type-statement: array of UDT member array", people(2).Scores(1), 4
    Check "type-statement: array of UDT default", people(1).Name, ""

    On Error Resume Next
    p1.Age = 70000
    CheckErr "type-statement: Integer member overflow", 6
    On Error GoTo 0
End Sub

Private Sub TestEnum()
    Check "enum-statement: IllegalEntry", IllegalEntry, -1
    Check "enum-statement: SecurityLevel1", SecurityLevel1, 0
    Check "enum-statement: SecurityLevel2", SecurityLevel2, 1
    Check "enum-statement: icMistyRose", icMistyRose, &HE1E4FF
    Check "enum-statement: icSlateGray", icSlateGray, 9470064
    Check "enum-statement: implicit first", shRed, 0
    Check "enum-statement: explicit", shGreen, 5
    Check "enum-statement: implicit after explicit", shBlue, 6
    CheckTyped "enum-statement: type Long", shBlue, CLng(6)
    Check "enum-statement: qualified", Shade.shGreen, 5
    Dim s As Shade
    s = shBlue
    Check "enum-statement: variable", s, 6
    Check "enum-statement: default 0", NoAssignLong() = shRed, True
End Sub

Private Sub TestLineSyntax()
    Dim a As Long, b As Long, s As String
    a = 1 + _
        2 + _
        3
    Check "line continuation: sum", a, 6
    s = "abc" & _
        "def"
    Check "line continuation: string", s, "abcdef"
    a = 1: b = 2: s = "x"
    Check "multiple statements: a", a, 1
    Check "multiple statements: b", b, 2
    Check "multiple statements: s", s, "x"
    a = 0
Lbl: a = a + 1
    If a < 4 Then GoTo Lbl
    Check "labels: label on same line", a, 4
    ' Comment with Rem
    Rem this is a comment
    a = 5 ' trailing comment
    Check "rem-statement: comment", a, 5
    ' Dim with multiple variables and types
    Dim c As Integer, d, e As String
    CheckTyped "dim-statement: Integer", c, CInt(0)
    Check "dim-statement: untyped is Variant Empty", IsEmpty(d), True
    CheckTyped "dim-statement: String", e, ""
End Sub


Private Sub TestTypeSuffixDim()
    ' Type-declaration characters
    Dim f%
    Dim h$
    Dim k#
    Dim n@
    CheckTyped "dim-statement: %", f, CInt(0)
    CheckTyped "dim-statement: $", h, ""
    CheckTyped "dim-statement: #", k, CDbl(0)
    CheckTyped "dim-statement: @", n, CCur(0)
End Sub
