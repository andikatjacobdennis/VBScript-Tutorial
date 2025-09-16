' =======================================
'   VBScript Example: Procedures
' =======================================

' ---------------------------------------
' Example 1: Sub Procedure (no return value)
' ---------------------------------------
Sub ShowMessage(text)
    MsgBox text, vbInformation, "Sub Procedure Example"
End Sub

' Call the Sub procedure
ShowMessage "Hello from a Sub procedure!"


' ---------------------------------------
' Example 2: Function Procedure (with return value)
' ---------------------------------------
Function AddNumbers(a, b)
    AddNumbers = a + b   ' return result
End Function

' Use the Function procedure
Dim num1, num2, result
num1 = 10
num2 = 20
result = AddNumbers(num1, num2)

MsgBox "The sum of " & num1 & " and " & num2 & " is " & result, vbInformation, "Function Procedure Example"


' ---------------------------------------
' Example 3: Function for Boolean Logic
' ---------------------------------------
Function IsEven(n)
    If n Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

Dim checkNum
checkNum = 7

If IsEven(checkNum) Then
    MsgBox checkNum & " is even.", vbInformation, "Boolean Function Example"
Else
    MsgBox checkNum & " is odd.", vbInformation, "Boolean Function Example"
End If

' =======================================
