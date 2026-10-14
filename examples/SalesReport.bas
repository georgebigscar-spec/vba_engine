Attribute VB_Name = "SalesReport"
Option Explicit

' Builds a sales table on Sheet1 and a per-region summary on a new sheet.
' Run with:  vba run examples/SalesReport.bas --excel --save report.xlsx
' or on an existing file:  vba run examples/SalesReport.bas --workbook in.xlsx --save out.xlsx

Sub Main()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    If IsEmpty(ws.Range("A1").Value) Then FillSample ws

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Debug.Print "Rows of data:"; lastRow - 1

    ' Amount = Qty * Price, read and written as arrays like real macros do.
    Dim data As Variant, amounts() As Variant, i As Long
    data = ws.Range("A2:D" & lastRow).Value
    ReDim amounts(1 To UBound(data, 1), 1 To 1)
    For i = 1 To UBound(data, 1)
        amounts(i, 1) = data(i, 3) * data(i, 4)
    Next
    ws.Range("E1").Value = "Amount"
    ws.Range("E2").Resize(UBound(amounts, 1), 1).Value = amounts
    ws.Range("A1:E1").Font.Bold = True

    Dim summary As Worksheet
    Set summary = Worksheets.Add(After:=ws)
    summary.Name = "Summary"
    summary.Range("A1:C1").Value = Array("Region", "Total", "Share")
    summary.Range("A1:C1").Font.Bold = True
    summary.Range("A1:C1").Interior.Color = RGB(221, 235, 247)

    Dim regions As New Collection, r As Variant, row As Long, grand As Double
    On Error Resume Next
    For i = 2 To lastRow
        regions.Add ws.Cells(i, 2).Value, CStr(ws.Cells(i, 2).Value)
    Next
    On Error GoTo 0

    grand = WorksheetFunction.Sum(ws.Range("E2:E" & lastRow))
    row = 2
    For Each r In regions
        summary.Cells(row, 1).Value = r
        summary.Cells(row, 2).Value = WorksheetFunction.SumIf(ws.Range("B2:B" & lastRow), r, ws.Range("E2:E" & lastRow))
        summary.Cells(row, 3).Value = summary.Cells(row, 2).Value / grand
        summary.Cells(row, 3).NumberFormat = "0.0%"
        Debug.Print r, summary.Cells(row, 2).Value, summary.Cells(row, 3).Text
        row = row + 1
    Next
    summary.Cells(row, 1).Value = "Total"
    summary.Cells(row, 2).Value = grand
    summary.Cells(row, 1).Resize(1, 2).Font.Bold = True
    Debug.Print "Total", grand
End Sub

Private Sub FillSample(ws As Worksheet)
    ws.Range("A1:D1").Value = Array("Date", "Region", "Qty", "Price")
    Dim rows As Variant, i As Long
    rows = Array( _
        Array(#1/5/2024#, "North", 10, 2.5), _
        Array(#1/6/2024#, "South", 4, 10), _
        Array(#1/9/2024#, "North", 7, 3), _
        Array(#1/12/2024#, "East", 12, 1.25), _
        Array(#1/15/2024#, "South", 3, 9.5))
    For i = 0 To UBound(rows)
        ws.Cells(i + 2, 1).Resize(1, 4).Value = rows(i)
    Next
    ws.Columns(1).NumberFormat = "yyyy-mm-dd"
End Sub
