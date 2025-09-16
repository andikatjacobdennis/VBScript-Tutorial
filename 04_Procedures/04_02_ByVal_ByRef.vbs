' =======================================
'   VBScript Example: ByVal and ByRef
' =======================================

' ---------------------------------------
' Example 1: ByVal (pass by value, multiple arguments)
' ---------------------------------------
Sub ChangeByVal(ByVal x, ByVal y)
    x = x + 10
    y = y + 20
End Sub

Dim a, b
a = 5
b = 15
Call ChangeByVal(a, b)

MsgBox "After ChangeByVal call:" & vbCrLf & _
       "a = " & a & vbCrLf & "b = " & b, vbInformation, "ByVal Example"


' ---------------------------------------
' Example 2: ByRef (pass by reference, multiple arguments)
' ---------------------------------------
Sub ChangeByRef(ByRef x, ByRef y)
    x = x + 10
    y = y + 20
End Sub

Dim c, d
c = 5
d = 15
Call ChangeByRef(c, d)

MsgBox "After ChangeByRef call:" & vbCrLf & _
       "c = " & c & vbCrLf & "d = " & d, vbInformation, "ByRef Example"

' =======================================