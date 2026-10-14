Attribute VB_Name = "smoke"
Sub Main()
    Check "mid", Mid("Mid Function Demo", 1, 3), "Mid"
    CheckTyped "clng", CLng("12"), 12&
    On Error Resume Next
    Dim x: x = 1 / 0
    CheckErr "div0", 11
End Sub
