//! Class and module features: Class_Terminate, Implements, events, default
//! members, custom enumerators, `As New`, module-level properties,
//! `Module.Member` and `Private` isolation.

use vba_runtime::{CaptureHost, Engine, EngineError};

fn class(name: &str, body: &str) -> String {
    format!("VERSION 1.0 CLASS\nBEGIN\nEND\nAttribute VB_Name = \"{name}\"\n{body}")
}

fn run_with(modules: &[(&str, String)]) -> Result<Vec<String>, EngineError> {
    let host = CaptureHost::new();
    let mut engine = Engine::new(host.clone());
    for (name, src) in modules {
        engine.load_source(src, name)?;
    }
    engine.run("Main", vec![])?;
    Ok(host.lines())
}

fn run(modules: &[(&str, String)]) -> Vec<String> {
    run_with(modules).unwrap_or_else(|e| panic!("{e}"))
}

#[test]
fn class_terminate_runs_when_last_reference_goes() {
    let tracer = class(
        "Tracer",
        r#"
Public Name As String
Private Sub Class_Terminate()
    Debug.Print "bye " & Name
End Sub
"#,
    );
    let out = run(&[
        ("Tracer", tracer),
        (
            "Module1",
            r#"
Sub Scoped()
    Dim t As New Tracer
    t.Name = "local"
    Debug.Print "in scope"
End Sub

Sub Main()
    Dim a As Tracer, b As Tracer
    Set a = New Tracer
    a.Name = "first"
    Set b = a
    Set a = Nothing
    Debug.Print "still held"
    Set b = Nothing
    Debug.Print "after"
    Scoped
    Debug.Print "back"
    Set a = New Tracer
    a.Name = "replaced"
    Set a = New Tracer
    a.Name = "last"
End Sub
"#
            .into(),
        ),
    ]);
    assert_eq!(
        out,
        [
            "still held",
            "bye first",
            "after",
            "in scope",
            "bye local",
            "back",
            "bye replaced",
            "bye last"
        ]
    );
}

#[test]
fn implements_dispatches_to_interface_members() {
    let ishape = class(
        "IShape",
        r#"
Public Function Area() As Double
End Function
Public Property Get Name() As String
End Property
"#,
    );
    let square = class(
        "Square",
        r#"
Implements IShape
Public Side As Double
Private Function IShape_Area() As Double
    IShape_Area = Side * Side
End Function
Private Property Get IShape_Name() As String
    IShape_Name = "square"
End Property
"#,
    );
    let circle = class(
        "Circle",
        r#"
Implements IShape
Public R As Double
Private Function IShape_Area() As Double
    IShape_Area = 3 * R * R
End Function
Private Property Get IShape_Name() As String
    IShape_Name = "circle"
End Property
"#,
    );
    let out = run(&[
        ("IShape", ishape),
        ("Square", square),
        ("Circle", circle),
        (
            "Module1",
            r#"
Sub Main()
    Dim shapes As New Collection, s As IShape, sq As New Square, c As New Circle
    sq.Side = 2
    c.R = 1
    shapes.Add sq
    shapes.Add c
    For Each s In shapes
        Debug.Print s.Name; s.Area; TypeOf s Is IShape; TypeName(s)
    Next
End Sub
"#
            .into(),
        ),
    ]);
    assert_eq!(out, ["square 4 TrueSquare", "circle 3 TrueCircle"]);
}

#[test]
fn events_reach_with_events_handlers() {
    let account = class(
        "Account",
        r#"
Public Event Changed(ByVal balance As Double)
Public Event Withdrawing(ByVal amount As Double, Cancel As Boolean)
Private mBalance As Double
Public Sub Deposit(ByVal amount As Double)
    mBalance = mBalance + amount
    RaiseEvent Changed(mBalance)
End Sub
Public Sub Withdraw(ByVal amount As Double)
    Dim cancel As Boolean
    RaiseEvent Withdrawing(amount, cancel)
    If cancel Then
        Debug.Print "withdrawal cancelled"
        Exit Sub
    End If
    mBalance = mBalance - amount
    RaiseEvent Changed(mBalance)
End Sub
"#,
    );
    let audit = class(
        "Audit",
        r#"
Private WithEvents acc As Account
Public Sub Watch(a As Account)
    Set acc = a
End Sub
Public Sub StopWatching()
    Set acc = Nothing
End Sub
Private Sub acc_Changed(ByVal balance As Double)
    Debug.Print "balance now"; balance
End Sub
Private Sub acc_Withdrawing(ByVal amount As Double, Cancel As Boolean)
    If amount > 100 Then Cancel = True
End Sub
"#,
    );
    let out = run(&[
        ("Account", account),
        ("Audit", audit),
        (
            "Module1",
            r#"
Sub Main()
    Dim a As New Account, log As New Audit
    a.Deposit 50
    log.Watch a
    a.Deposit 70
    a.Withdraw 500
    a.Withdraw 20
    log.StopWatching
    a.Deposit 1
    Debug.Print "done"
End Sub
"#
            .into(),
        ),
    ]);
    assert_eq!(
        out,
        [
            "balance now 120 ",
            "withdrawal cancelled",
            "balance now 100 ",
            "done"
        ]
    );
}

#[test]
fn default_member_and_custom_enumerator() {
    let list = class(
        "NameList",
        r#"
Private mItems As New Collection
Public Sub Add(ByVal s As String)
    mItems.Add s
End Sub
Public Property Get Item(ByVal i As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = mItems(i)
End Property
Public Property Get Count() As Long
    Count = mItems.Count
End Property
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mItems.[_NewEnum]
End Function
"#,
    );
    let money = class(
        "Money",
        r#"
Private mAmount As Double
Public Property Get Amount() As Double
Attribute Amount.VB_UserMemId = 0
    Amount = mAmount
End Property
Public Property Let Amount(ByVal v As Double)
    mAmount = v
End Property
"#,
    );
    let out = run(&[
        ("NameList", list),
        ("Money", money),
        (
            "Module1",
            r#"
Sub Main()
    Dim names As New NameList, n As Variant, m As New Money, total As Double
    names.Add "Ann"
    names.Add "Bob"
    Debug.Print names(2); names.Count
    For Each n In names
        Debug.Print n;
    Next
    Debug.Print
    m = 40
    total = m + 2
    Debug.Print total; CStr(m); m > 10
    If m Then Debug.Print "truthy"
End Sub
"#
            .into(),
        ),
    ]);
    assert_eq!(out, ["Bob 2 ", "AnnBob", " 42 40True", "truthy"]);
}

#[test]
fn as_new_creates_on_first_use() {
    let counter = class(
        "Counter",
        r#"
Public N As Long
Private Sub Class_Initialize()
    Debug.Print "created"
End Sub
"#,
    );
    let out = run(&[
        ("Counter", counter),
        (
            "Module1",
            r#"
Sub Main()
    Dim c As New Counter
    Debug.Print "declared"
    c.N = 5
    Debug.Print c.N
    Set c = Nothing
    Debug.Print c Is Nothing
    Debug.Print c.N
End Sub
"#
            .into(),
        ),
    ]);
    assert_eq!(
        out,
        ["declared", "created", " 5 ", "created", "False", " 0 "]
    );
}

#[test]
fn module_properties_qualified_names_and_private_scope() {
    let config = r#"
Private mLevel As Long
Public Total As Long
Private Secret As String

Public Property Get Level() As Long
    Level = mLevel
End Property

Public Property Let Level(ByVal v As Long)
    If v > 10 Then v = 10
    mLevel = v
End Property

Public Function Describe() As String
    Describe = "level " & mLevel
End Function

Private Function Helper() As String
    Helper = "config helper"
End Function

Public Function UseHelper() As String
    UseHelper = Helper()
End Function
"#;
    let main = r#"
Private Function Helper() As String
    Helper = "main helper"
End Function

Sub Main()
    Level = 42
    Debug.Print Level; Config.Level; Config.Describe
    Config.Level = 3
    Config.Total = 7
    Total = Total + 1
    Debug.Print Level; Total; Config.Total
    Debug.Print Helper(); "|"; UseHelper()
    On Error Resume Next
    Debug.Print Config.Secret
    Debug.Print Err.Number <> 0
End Sub
"#;
    let out = run(&[("Config", config.into()), ("Module1", main.into())]);
    assert_eq!(
        out,
        [
            " 10  10 level 10",
            " 3  8  8 ",
            "main helper|config helper",
            "True"
        ]
    );
}

#[test]
fn terminate_on_error_unwind_nesting_and_not_on_end() {
    let node = class(
        "Node",
        r#"
Public Name As String
Public Child As Node
Private Sub Class_Terminate()
    Debug.Print "free " & Name
End Sub
"#,
    );
    let main = r#"
Sub Fails()
    Dim n As New Node
    n.Name = "temp"
    Err.Raise 5
End Sub

Sub Main()
    On Error Resume Next
    Fails
    Debug.Print "handled"; Err.Number
    On Error GoTo 0
    Dim root As New Node
    root.Name = "root"
    Set root.Child = New Node
    root.Child.Name = "child"
    Set root = Nothing
    Debug.Print "tree gone"
    Dim keep As New Node
    keep.Name = "kept"
    End
End Sub
"#;
    let out = run(&[("Node", node), ("Module1", main.into())]);
    assert_eq!(
        out,
        [
            "free temp",
            "handled 5 ",
            "free root",
            "free child",
            "tree gone"
        ]
    );
}

#[test]
fn raise_event_requires_declaration_and_private_procs_stay_private() {
    let src = class(
        "Src",
        "Public Sub Fire()\n    RaiseEvent Missing\nEnd Sub\n",
    );
    let err = run_with(&[
        ("Src", src),
        (
            "Module1",
            "Sub Main()\n Dim s As New Src\n s.Fire\nEnd Sub\n".into(),
        ),
    ])
    .unwrap_err();
    assert!(err.to_string().contains("Event not found"), "{err}");

    let err = run_with(&[
        ("Lib", "Private Sub Hidden()\nEnd Sub\n".into()),
        ("Module1", "Sub Main()\n Lib.Hidden\nEnd Sub\n".into()),
    ])
    .unwrap_err();
    assert!(err.to_string().contains("Hidden"), "{err}");
}
