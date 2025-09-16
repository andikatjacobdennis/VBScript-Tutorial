' =======================================
'   VBScript Example: Input
' =======================================

' ---------------------------------------
' InputBox Example
' ---------------------------------------
Dim userName
userName = InputBox("Please enter your name:", "Input Example")

If userName <> "" Then
    MsgBox "Hello, " & userName & "!", vbInformation, "Greeting"
Else
    MsgBox "No name entered.", vbExclamation, "Greeting"
End If

' ---------------------------------------
' Input Numeric Value Example
' ---------------------------------------
Dim age
age = InputBox("Please enter your age:", "Numeric Input Example")

If IsNumeric(age) Then
    MsgBox "You are " & age & " years old.", vbInformation, "Age Input"
Else
    MsgBox "Invalid input. Please enter a number.", vbExclamation, "Age Input"
End If

' ---------------------------------------
' Confirm Input Example
' ---------------------------------------
Dim response
response = MsgBox("Do you want to continue?", vbYesNo + vbQuestion, "Confirm Input")

If response = vbYes Then
    MsgBox "You chose Yes.", vbInformation, "Confirmation"
Else
    MsgBox "You chose No.", vbInformation, "Confirmation"
End If

' =======================================
