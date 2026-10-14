use std::rc::Rc;

use vba_runtime::{CaptureHost, DynObject, Engine, EngineError, Value};

fn run(src: &str) -> Vec<String> {
    let host = CaptureHost::new();
    let mut engine = Engine::new(host.clone());
    engine.load_source(src, "Module1").unwrap();
    engine.run("Main", vec![]).unwrap();
    host.lines()
}

fn run_err(src: &str) -> EngineError {
    let mut engine = Engine::new(CaptureHost::new());
    engine.load_source(src, "Module1").unwrap();
    engine.run("Main", vec![]).unwrap_err()
}

#[test]
fn recursion_and_functions() {
    let out = run(r#"
Function Fact(ByVal n As Long) As Long
    If n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)
End Function

Sub Main()
    Debug.Print Fact(10)
End Sub
"#);
    assert_eq!(out, vec!["3628800"]);
}

#[test]
fn byref_is_default() {
    let out = run(r#"
Sub Bump(x As Long, ByVal y As Long)
    x = x + 1
    y = y + 1
End Sub

Sub Main()
    Dim a As Long, b As Long
    Bump a, b
    Call Bump(a, b)
    Debug.Print a, b
End Sub
"#);
    assert_eq!(out, vec!["2 0"]);
}

#[test]
fn loops_select_and_strings() {
    let out = run(r#"
Sub Main()
    Dim i As Integer, s As String
    For i = 10 To 1 Step -3
        s = s & i & ";"
    Next
    Debug.Print s
    i = 0
    Do
        i = i + 1
        If i = 3 Then Exit Do
    Loop While True
    Select Case i
        Case 1, 2: Debug.Print "small"
        Case 3 To 5: Debug.Print "medium"
        Case Else: Debug.Print "large"
    End Select
    Debug.Print UCase(Left("hello world", 5)) & "|" & Mid$("abcdef", 2, 3) & "|" & Len("xyz")
    Debug.Print Join(Split("a,b,c", ","), "-"), InStr("hello", "ll")
End Sub
"#);
    assert_eq!(out, vec!["10;7;4;1;", "medium", "HELLO|bcd|3", "a-b-c 3"]);
}

#[test]
fn arrays_and_collections() {
    let out = run(r#"
Sub Main()
    Dim a(1 To 3) As Long, i As Long, total As Long
    For i = 1 To 3: a(i) = i * i: Next
    ReDim Preserve a(1 To 4)
    a(4) = 100
    For Each v In a
        total = total + v
    Next
    Debug.Print total, UBound(a), LBound(a)

    Dim c As New Collection
    c.Add "x", "first"
    c.Add "y"
    Debug.Print c.Count, c("first"), c(2)
End Sub
"#);
    assert_eq!(out, vec!["114 4 1", "2 x y"]);
}

#[test]
fn on_error_resume_next_and_goto() {
    let out = run(r#"
Sub Main()
    On Error Resume Next
    x = 1 / 0
    Debug.Print "after", Err.Number
    On Error GoTo Handler
    Err.Raise 1234, , "custom"
    Debug.Print "not reached"
    Exit Sub
Handler:
    Debug.Print "handled", Err.Number, Err.Description
    Resume Next
End Sub
"#);
    assert_eq!(out, vec!["after 11", "handled 1234 custom", "not reached"]);
}

#[test]
fn unhandled_error_reports_line() {
    let e = run_err("Sub Main()\n  Dim a(2)\n  a(5) = 1\nEnd Sub\n");
    match e {
        EngineError::Runtime(e) => {
            assert_eq!(e.number, 9);
            assert_eq!(e.line, Some(3));
        }
        other => panic!("{other}"),
    }
}

#[test]
fn option_explicit() {
    let e = run_err("Option Explicit\nSub Main()\n  y = 1\nEnd Sub\n");
    assert!(e.to_string().contains("Variable not defined"));
}

#[test]
fn host_objects() {
    let mut host = CaptureHost::new();
    let sheet = Rc::new(
        DynObject::new("Worksheet")
            .with_prop("Name", "Sheet1")
            .with_prop("Value", 0),
    );
    host.globals
        .push(("ActiveSheet".into(), Value::from(sheet.clone())));
    let mut engine = Engine::new(host.clone());
    engine
        .load_source(
            r#"
Sub Main()
    With ActiveSheet
        .Name = "Report"
        .Value = .Value + 42
    End With
    Debug.Print ActiveSheet.Name, TypeName(ActiveSheet)
End Sub
"#,
            "Module1",
        )
        .unwrap();
    engine.run("Main", vec![]).unwrap();
    assert_eq!(host.lines(), vec!["Report Worksheet"]);
    assert_eq!(sheet.get_prop("Value").unwrap().to_string(), "42");
}

#[test]
fn function_called_from_host_with_args() {
    let mut engine = Engine::new(CaptureHost::new());
    engine
        .load_source("Function Add(a, b)\n  Add = a + b\nEnd Function\n", "M")
        .unwrap();
    let v = engine
        .run("Add", vec![Value::Long(2), Value::Double(0.5)])
        .unwrap();
    assert_eq!(v.to_string(), "2.5");
}
