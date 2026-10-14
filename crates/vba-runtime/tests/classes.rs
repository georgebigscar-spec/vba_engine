use vba_runtime::{CaptureHost, Engine, EngineError};

const ACCOUNT_CLS: &str = r#"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Option Explicit

Public Owner As String
Private mBalance As Currency
Private mHistory As Collection

Private Sub Class_Initialize()
    Set mHistory = New Collection
    Owner = "nobody"
End Sub

Public Property Get Balance() As Currency
    Balance = mBalance
End Property

Public Property Let Balance(ByVal value As Currency)
    If value < 0 Then Err.Raise 5, "Account", "Negative balance"
    mBalance = value
    mHistory.Add value
End Property

Public Property Get Changes() As Long
    Changes = mHistory.Count
End Property

Public Sub Deposit(ByVal amount As Currency)
    Balance = Balance + amount
End Sub

Public Function Describe() As String
    Describe = Me.Owner & ": " & Format(Me.Balance, "0.00")
End Function

Private Sub Secret()
End Sub
"#;

fn run_with(modules: &[(&str, &str)]) -> Result<Vec<String>, EngineError> {
    let host = CaptureHost::new();
    let mut engine = Engine::new(host.clone());
    for (name, src) in modules {
        engine.load_source(src, name)?;
    }
    engine.run("Main", vec![])?;
    Ok(host.lines())
}

#[test]
fn class_properties_methods_and_me() {
    let main = r#"
Sub Main()
    Dim a As Account
    Set a = New Account
    Debug.Print a.Owner, TypeName(a), TypeOf a Is Account
    a.Owner = "Ann"
    a.Balance = 10
    a.Deposit 5.5
    Debug.Print a.Balance, a.Changes, a.Describe()

    Dim b As New Account
    Debug.Print b.Balance, a Is b
End Sub
"#;
    let out = run_with(&[("Account", ACCOUNT_CLS), ("Module1", main)]).unwrap();
    assert_eq!(
        out,
        vec!["nobody Account True", "15.5 2 Ann: 15.50", "0 False"]
    );
}

#[test]
fn class_errors_and_private_members() {
    let main = "Sub Main()\n  Dim a As New Account\n  On Error Resume Next\n  a.Balance = -1\n  Debug.Print Err.Number, Err.Description\n  Err.Clear\n  a.Secret\n  Debug.Print Err.Number\nEnd Sub\n";
    let out = run_with(&[("Account", ACCOUNT_CLS), ("Module1", main)]).unwrap();
    assert_eq!(out, vec!["5 Negative balance", "438"]);
}

#[test]
fn user_defined_types() {
    let src = r#"
Type Point
    X As Long
    Y As Long
End Type

Type Shape
    Name As String
    Corners(1 To 2) As Point
End Type

Sub Move(p As Point, ByVal dx As Long)
    p.X = p.X + dx
End Sub

Sub Main()
    Dim a As Point, b As Point
    a.X = 1: a.Y = 2
    b = a
    b.X = 100
    Move a, 10
    Debug.Print a.X, a.Y, b.X, TypeName(a)

    Dim s As Shape
    s.Name = "box"
    s.Corners(2).Y = 7
    With s.Corners(1)
        .X = 3
    End With
    Debug.Print s.Name, s.Corners(1).X, s.Corners(2).Y

    Dim pts(1 To 3) As Point
    pts(2).X = 42
    Debug.Print pts(2).X, pts(3).X
End Sub
"#;
    let out = run_with(&[("Module1", src)]).unwrap();
    assert_eq!(out, vec!["11 2 100 Point", "box 3 7", "42 0"]);
}

#[test]
fn enums_and_param_array() {
    let src = r#"
Option Explicit

Enum Color
    Red
    Green = 10
    Blue
End Enum

Function Sum(ParamArray xs() As Variant) As Double
    Dim x As Variant
    For Each x In xs
        Sum = Sum + x
    Next
End Function

Sub Main()
    Dim c As Color
    c = Blue
    Debug.Print Red, Color.Green, c, TypeName(c)
    Debug.Print Sum(1, 2, 3.5), Sum()
End Sub
"#;
    let out = run_with(&[("Module1", src)]).unwrap();
    assert_eq!(out, vec!["0 10 11 Long", "6.5 0"]);
}
