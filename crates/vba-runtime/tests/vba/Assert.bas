Attribute VB_Name = "Assert"
' Assertion helpers shared by every conformance file in this folder.
' A failed check prints a line starting with "FAIL"; the Rust harness
' (tests/conformance.rs) collects those lines.
Option Explicit

Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = "<" & TypeName(v) & ">"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsArray(v) Then
        Show = "<array>"
    ElseIf VarType(v) = vbString Then
        Show = """" & v & """"
    Else
        Show = CStr(v) & " (" & TypeName(v) & ")"
    End If
End Function

Private Function Same(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        Same = IsNull(a) And IsNull(b)
    ElseIf IsEmpty(a) Or IsEmpty(b) Then
        Same = IsEmpty(a) And IsEmpty(b)
    ElseIf VarType(a) = vbString Or VarType(b) = vbString Then
        Same = (VarType(a) = VarType(b)) And (StrComp(a, b, vbBinaryCompare) = 0)
    Else
        Same = (a = b)
    End If
End Function

' Value check: strings compare exactly, numbers by value.
Public Sub Check(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant)
    If Not Same(actual, expected) Then
        Debug.Print "FAIL " & label & ": got " & Show(actual) & ", expected " & Show(expected)
    End If
End Sub

' Value and Variant subtype check.
Public Sub CheckTyped(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant)
    Check label, actual, expected
    If TypeName(actual) <> TypeName(expected) Then
        Debug.Print "FAIL " & label & ": type " & TypeName(actual) & ", expected " & TypeName(expected)
    End If
End Sub

' Checks that Err holds the given number, then clears it.
Public Sub CheckErr(ByVal label As String, ByVal expected As Long)
    If Err.Number <> expected Then
        Debug.Print "FAIL " & label & ": error " & Err.Number & " (" & Err.Description & "), expected " & expected
    End If
    Err.Clear
End Sub
