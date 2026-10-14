VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Option Explicit

Public Owner As String
Private mBalance As Currency
Private mLog As Collection

Private Sub Class_Initialize()
    Set mLog = New Collection
End Sub

Public Property Get Balance() As Currency
    Balance = mBalance
End Property

Public Property Let Balance(ByVal value As Currency)
    If value < 0 Then Err.Raise vbObjectError + 1, "Account", "Insufficient funds"
    mBalance = value
End Property

Public Sub Post(ByVal kind As TxKind, ByVal amount As Currency)
    Select Case kind
        Case Deposit: Balance = Balance + amount
        Case Withdrawal: Balance = Balance - amount
    End Select
    mLog.Add kind
End Sub

Public Property Get Operations() As Long
    Operations = mLog.Count
End Property
