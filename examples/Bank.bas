Attribute VB_Name = "Bank"
Option Explicit

' Run: vba run examples/Bank.bas examples/Account.cls

Public Enum TxKind
    Deposit = 1
    Withdrawal
End Enum

Private Type Tx
    Kind As TxKind
    Amount As Currency
End Type

Sub Main()
    Dim acc As New Account
    acc.Owner = "Ann"

    Dim txs(1 To 3) As Tx, i As Long
    txs(1).Kind = Deposit: txs(1).Amount = 100
    txs(2).Kind = Withdrawal: txs(2).Amount = 30
    With txs(3)
        .Kind = Withdrawal
        .Amount = 500
    End With

    On Error GoTo Rejected
    For i = 1 To 3
        acc.Post txs(i).Kind, txs(i).Amount
        Debug.Print "after #" & i & ": " & acc.Balance
    Next i
    Exit Sub
Rejected:
    Debug.Print "tx #" & i & " rejected: " & Err.Description
    Debug.Print acc.Owner & " has " & acc.Balance & " after " & acc.Operations & " operations"
End Sub
