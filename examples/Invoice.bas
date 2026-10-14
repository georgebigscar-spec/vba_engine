Attribute VB_Name = "Invoice"
Option Explicit

' Demo module: run with `vba run examples/Invoice.bas`
' or debug with `vba debug examples/Invoice.bas`.

Private Const TAX_RATE As Double = 0.2

Function LineTotal(ByVal qty As Long, ByVal price As Currency) As Currency
    LineTotal = qty * price
End Function

Function ApplyTax(ByVal amount As Currency) As Currency
    ApplyTax = amount * (1 + TAX_RATE)
End Function

Sub Main()
    Dim items As New Collection
    Dim item As Variant, net As Currency, i As Long

    items.Add Array("Paper", 10, 4.5)
    items.Add Array("Toner", 2, 39.99)
    items.Add Array("Stapler", 1, 12)

    For Each item In items
        i = i + 1
        net = net + LineTotal(item(1), item(2))
        Debug.Print i & ". " & item(0) & vbTab & Format(LineTotal(item(1), item(2)), "#,##0.00")
    Next item

    Debug.Print "Net:   " & Format(net, "#,##0.00")
    Debug.Print "Gross: " & Format(ApplyTax(net), "#,##0.00")

    On Error GoTo Failed
    Debug.Print "Average: " & net / (i - 3)
    Exit Sub
Failed:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub
