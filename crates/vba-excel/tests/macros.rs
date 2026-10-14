//! Typical Excel macros run against the in-memory object model.

use vba_excel::ExcelHost;
use vba_runtime::{Engine, Value};

fn run(host: &ExcelHost, src: &str) -> Vec<String> {
    let mut engine = Engine::new(host.clone());
    engine.load_source(src, "Module1").expect("parse");
    if let Err(e) = engine.run("Main", vec![]) {
        panic!("{e}\noutput: {:?}", host.lines());
    }
    host.lines()
}

fn cell(host: &ExcelHost, addr: &str) -> String {
    host.cell("Sheet1", addr)
        .map(|v| v.to_string())
        .unwrap_or_default()
}

#[test]
fn last_row_and_loop_over_rows() {
    let host = ExcelHost::new();
    for (i, (name, qty)) in [("apple", 3.0), ("pear", 5.0), ("plum", 2.0)]
        .iter()
        .enumerate()
    {
        host.set_cell("Sheet1", &format!("A{}", i + 2), Value::str(*name));
        host.set_cell("Sheet1", &format!("B{}", i + 2), Value::Double(*qty));
    }
    host.set_cell("Sheet1", "A1", Value::str("Item"));
    let out = run(
        &host,
        r#"
Sub Main()
    Dim lastRow As Long, i As Long, total As Double
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print "last"; lastRow
    For i = 2 To lastRow
        Cells(i, 3).Value = Cells(i, 2).Value * 10
        total = total + Cells(i, "B")
    Next i
    Range("C1") = "Total"
    Cells(lastRow + 1, 2).Value = total
    Debug.Print ActiveSheet.UsedRange.Address
    Debug.Print Range("A1").End(xlDown).Address(False, False)
    Debug.Print Range("A1").CurrentRegion.Address
End Sub
"#,
    );
    assert_eq!(out, ["last 4 ", "$A$1:$C$5", "A4", "$A$1:$C$5"]);
    assert_eq!(cell(&host, "C3"), "50");
    assert_eq!(cell(&host, "B5"), "10");
}

#[test]
fn value_arrays_and_for_each() {
    let host = ExcelHost::new();
    let out = run(
        &host,
        r#"
Sub Main()
    Dim data As Variant, c As Range, s As Double, r As Long
    Range("A1:C1").Value = Array("x", "y", "z")
    Range("A2:B3").Value = 7
    data = Range("A1:C3").Value
    Debug.Print LBound(data, 1); UBound(data, 1); UBound(data, 2); data(1, 3); data(3, 2)
    For Each c In Range("A2:B3")
        s = s + c.Value
    Next
    Debug.Print s; Range("A2:B3").Count; Range("A2:B3").Rows.Count
    Dim out(1 To 2, 1 To 1)
    out(1, 1) = "p": out(2, 1) = "q"
    Range("E1").Resize(2, 1).Value = out
    For r = 1 To Range("A1:C3").Rows.Count
        Debug.Print Range("A1:C3").Cells(r, 1).Address(0, 0);
    Next
    Debug.Print
    Debug.Print Range("B2").Offset(1, 1).Address; IsEmpty(Range("C3").Value)
End Sub
"#,
    );
    assert_eq!(out, [" 1  3  3 z 7 ", " 28  4  2 ", "A1A2A3", "$C$3True"]);
    assert_eq!(cell(&host, "E2"), "q");
}

#[test]
fn sheets_books_and_formatting() {
    let host = ExcelHost::new();
    let out = run(
        &host,
        r#"
Sub Main()
    Dim ws As Worksheet, sh As Object
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Report"
    ws.Range("A1").Value = "Header"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Interior.Color = RGB(255, 255, 0)
    ws.Range("B1").Value = 0.5
    ws.Range("B1").NumberFormat = "0%"
    Debug.Print ws.Range("B1").Text; ws.Range("A1").Font.Bold; Worksheets.Count; ws.Index
    For Each sh In ThisWorkbook.Worksheets
        Debug.Print sh.Name;
    Next
    Debug.Print
    Worksheets("Sheet1").Activate
    Debug.Print ActiveSheet.Name; Sheets("Report").Range("A1").Value
    With Worksheets("Report")
        .Cells(2, 1).Value = .Cells(1, 1).Value & "!"
    End With
    Debug.Print Worksheets("Report").Range("A2")
    Debug.Print Application.Name; TypeName(ws); TypeName(ws.Range("A1"))
End Sub
"#,
    );
    assert_eq!(
        out,
        [
            "50%True 2  2 ",
            "Sheet1Report",
            "Sheet1Header",
            "Header!",
            "Microsoft ExcelWorksheetRange"
        ]
    );
}

#[test]
fn worksheet_functions() {
    let host = ExcelHost::new();
    let out = run(
        &host,
        r#"
Sub Main()
    Dim i As Long, wf As Object
    For i = 1 To 5
        Cells(i, 1).Value = i * 2
        Cells(i, 2).Value = Choose(i, "a", "b", "a", "c", "a")
    Next
    Set wf = Application.WorksheetFunction
    Debug.Print wf.Sum(Range("A1:A5")); wf.Average(Range("A:A")); wf.Max(Range("A1:A5"), 100); wf.Min(Columns(1))
    Debug.Print wf.CountIf(Range("B1:B5"), "a"); wf.SumIf(Range("B1:B5"), "a", Range("A1:A5")); wf.CountIf(Range("A1:A5"), ">4")
    Debug.Print wf.CountA(Range("A1:B10")); wf.CountBlank(Range("A1:B10")); wf.Count(Range("A1:B5"))
    Debug.Print wf.VLookup(6, Range("A1:B5"), 2, False); wf.Match("c", Range("B1:B5"), 0); wf.Index(Range("A1:B5"), 4, 2)
    Debug.Print wf.Round(2.345, 2); WorksheetFunction.Sum(1, 2, Array(3, 4))
    On Error Resume Next
    Debug.Print wf.VLookup(7, Range("A1:B5"), 2, False)
    Debug.Print Err.Number
End Sub
"#,
    );
    assert_eq!(
        out,
        [
            " 30  6  100  2 ",
            " 3  18  3 ",
            " 10  10  5 ",
            "a 4 c",
            " 2.35  10 ",
            " 1004 "
        ]
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_round_trip() {
    let dir = std::env::temp_dir().join(format!("vba-excel-test-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("out.xlsx");
    let host = ExcelHost::new();
    let src = format!(
        r#"
Sub Main()
    Range("A1").Value = "Name"
    Range("A1").Font.Bold = True
    Range("A2").Value = "Total"
    Range("B2").Value = 42.5
    Range("B3").Formula = "=B2*2"
    Range("C2").Value = DateSerial(2024, 3, 1)
    Range("D2").Value = True
    Worksheets.Add(After:=Worksheets(1)).Name = "Data"
    Worksheets("Data").Range("C5").Value = "far"
    ActiveWorkbook.SaveAs "{}"
End Sub
"#,
        path.display()
    );
    run(&host, &src);

    let loaded = ExcelHost::open(&path).unwrap();
    let book = loaded.active_workbook();
    assert_eq!(book.sheets.borrow().len(), 2);
    assert_eq!(loaded.cell("Sheet1", "A2").unwrap().to_string(), "Total");
    assert_eq!(loaded.cell("Sheet1", "B2").unwrap().to_string(), "42.5");
    assert!(matches!(
        loaded.cell("Sheet1", "C2").unwrap(),
        Value::Date(_)
    ));
    assert_eq!(loaded.cell("Sheet1", "D2").unwrap().to_string(), "True");
    assert_eq!(loaded.cell("Data", "C5").unwrap().to_string(), "far");
    let sheet = book.sheet("Sheet1").unwrap();
    assert_eq!(sheet.data(3, 2).formula.as_deref(), Some("=B2*2"));
    let out = run(
        &loaded,
        "Sub Main()\n Debug.Print Workbooks(1).Name; Worksheets(2).Range(\"C5\")\nEnd Sub",
    );
    assert_eq!(out, ["out.xlsxfar"]);
    std::fs::remove_dir_all(&dir).ok();
}
