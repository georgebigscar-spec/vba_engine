//! File I/O through the host, fixed-length strings, Erase, LSet/RSet and
//! the Mid statement.

use vba_runtime::{CaptureHost, Engine, EngineError, StdHost};

fn run_on(host: &CaptureHost, src: &str) -> Result<Vec<String>, EngineError> {
    let mut engine = Engine::new(host.clone());
    engine.load_source(src, "Module1")?;
    engine.run("Main", vec![])?;
    Ok(host.lines())
}

fn run(src: &str) -> Vec<String> {
    run_on(&CaptureHost::new(), src).unwrap_or_else(|e| panic!("{e}"))
}

#[test]
fn write_then_read_text_files() {
    let host = CaptureHost::new();
    let out = run_on(
        &host,
        r#"
Sub Main()
    Dim f As Integer, line As String, n As Long
    f = FreeFile
    Open "data/out.txt" For Output As #f
    Print #f, "Name"; Tab(10); "Qty"
    Print #f, "apple", 3
    Print #f, "half";
    Print #f, " line"
    Close #f

    Open "data/out.txt" For Append As #1
    Print #1, "appended"
    Close

    Open "data/out.txt" For Input As #2
    Do Until EOF(2)
        Line Input #2, line
        n = n + 1
        Debug.Print n; "[" & line & "]"
    Loop
    Debug.Print LOF(2); FreeFile
    Close #2
End Sub
"#,
    )
    .unwrap();
    assert_eq!(
        out,
        [
            " 1 [Name     Qty]",
            " 2 [apple          3 ]",
            " 3 [half line]",
            " 4 [appended]",
            " 54  1 "
        ]
    );
    assert!(host
        .file_text("data/out.txt")
        .unwrap()
        .ends_with("appended\r\n"));
}

#[test]
fn write_and_input_round_trip_typed_fields() {
    let host = CaptureHost::new();
    let out = run_on(
        &host,
        r#"
Sub Main()
    Dim name As String, qty As Long, price As Double, ok As Boolean, d As Date
    Open "rows.csv" For Output As #1
    Write #1, "Widget, large", 3, 2.5, True, DateSerial(2024, 3, 1)
    Write #1, "Bolt", 10, 0.1, False, #1/15/2024#
    Close #1
    Open "rows.csv" For Input As #1
    Do While Not EOF(1)
        Input #1, name, qty, price, ok, d
        Debug.Print name; "|"; qty; "|"; price; "|"; ok; "|"; Format(d, "yyyy-mm-dd")
    Loop
    Close #1
End Sub
"#,
    )
    .unwrap();
    assert_eq!(
        host.file_text("rows.csv").unwrap(),
        "\"Widget, large\",3,2.5,#TRUE#,#2024-03-01#\r\n\"Bolt\",10,0.1,#FALSE#,#2024-01-15#\r\n"
    );
    assert_eq!(
        out,
        [
            "Widget, large| 3 | 2.5 |True|2024-03-01",
            "Bolt| 10 | 0.1 |False|2024-01-15"
        ]
    );
}

#[test]
fn dir_kill_and_file_errors() {
    let host = CaptureHost::new();
    for name in ["in/a.txt", "in/b.txt", "in/c.csv"] {
        host.files.borrow_mut().insert(name.into(), b"x".to_vec());
    }
    let out = run_on(
        &host,
        r#"
Sub Main()
    Dim f As String
    f = Dir("in/*.txt")
    Do While f <> ""
        Debug.Print f;
        f = Dir()
    Loop
    Debug.Print
    Debug.Print Dir("in/c.csv"); "|"; Dir("in/missing.txt"); "|"; FileLen("in/a.txt")
    FileCopy "in/a.txt", "in/copy.txt"
    Kill "in/*.txt"
    Debug.Print Dir("in/*")
    On Error Resume Next
    Open "nope.txt" For Input As #1
    Debug.Print Err.Number
    Err.Clear
    Open "in/c.csv" For Input As #1
    Open "in/c.csv" For Input As #1
    Debug.Print Err.Number
    Err.Clear
    Print #1, "x"
    Debug.Print Err.Number
    Err.Clear
    Dim s As String
    Line Input #1, s
    Line Input #1, s
    Debug.Print Err.Number
End Sub
"#,
    )
    .unwrap();
    assert_eq!(
        out,
        [
            "a.txtb.txt",
            "c.csv|| 1 ",
            "c.csv",
            " 53 ",
            " 55 ",
            " 54 ",
            " 62 "
        ]
    );
}

#[test]
fn files_are_denied_unless_the_host_allows_them() {
    // A host that does not override the file methods.
    struct Quiet;
    impl vba_runtime::Host for Quiet {
        fn debug_print(&mut self, _: &str) {}
    }
    let mut engine = Engine::new(Quiet);
    engine
        .load_source(
            "Sub Main()\n Open \"x.txt\" For Output As #1\nEnd Sub\n",
            "M",
        )
        .unwrap();
    let err = engine.run("Main", vec![]).unwrap_err();
    assert!(err.to_string().contains("Path/File access error"), "{err}");

    // StdHost uses the real file system.
    let dir = std::env::temp_dir().join(format!("vba-files-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("std.txt");
    let src = format!(
        "Sub Main()\n Open \"{p}\" For Output As #1\n Print #1, \"hi\"\n Close #1\nEnd Sub\n",
        p = path.display()
    );
    let mut engine = Engine::new(StdHost);
    engine.load_source(&src, "M").unwrap();
    engine.run("Main", vec![]).unwrap();
    assert_eq!(std::fs::read_to_string(&path).unwrap(), "hi\r\n");
    std::fs::remove_dir_all(&dir).ok();
}

#[test]
fn unclosed_files_are_flushed_at_the_end() {
    let host = CaptureHost::new();
    run_on(
        &host,
        "Sub Main()\n Open \"log.txt\" For Output As #3\n Print #3, \"kept\"\nEnd Sub\n",
    )
    .unwrap();
    assert_eq!(host.file_text("log.txt").unwrap(), "kept\r\n");
}

#[test]
fn fixed_strings_erase_lset_rset_and_mid() {
    let out = run(r#"
Type Rec
    Code As String * 4
    Qty As Long
End Type

Sub Main()
    Dim code As String * 5, r As Rec, s As String
    code = "AB"
    Debug.Print "[" & code & "]"; Len(code)
    code = "ABCDEFG"
    Debug.Print "[" & code & "]"
    r.Code = "X1"
    Debug.Print "[" & r.Code & "]"

    s = "12345678"
    LSet s = "ab"
    Debug.Print "[" & s & "]"
    RSet s = "ab"
    Debug.Print "[" & s & "]"

    s = "Hello World"
    Mid(s, 7) = "Earth!!!"
    Debug.Print s
    Mid(s, 1, 2) = "JJJJ"
    Debug.Print s

    Dim fixed(1 To 3) As Long, dyn() As String, recs(1 To 2) As Rec
    fixed(2) = 7
    ReDim dyn(5)
    dyn(1) = "x"
    recs(1).Qty = 9
    Erase fixed, dyn, recs
    Debug.Print UBound(fixed); fixed(2); recs(1).Qty
    On Error Resume Next
    Debug.Print UBound(dyn)
    Debug.Print Err.Number
End Sub
"#);
    assert_eq!(
        out,
        [
            "[AB   ] 5 ",
            "[ABCDE]",
            "[X1  ]",
            "[ab      ]",
            "[      ab]",
            "Hello Earth",
            "JJllo Earth",
            " 3  0  0 ",
            " 9 "
        ]
    );
}
