' =======================================
'   VBScript Example: Condition
' =======================================

' ---------------------------------------
' Simple If...Then
' ---------------------------------------
Dim num, message
num = 5
message = "Simple If...Then:" & vbCrLf

If num > 0 Then
    message = message & "Number is positive." & vbCrLf
End If

MsgBox message, vbInformation, "If...Then Example"


' ---------------------------------------
' If...Then...Else
' ---------------------------------------
num = -3
message = "If...Then...Else:" & vbCrLf

If num >= 0 Then
    message = message & "Number is non-negative." & vbCrLf
Else
    message = message & "Number is negative." & vbCrLf
End If

MsgBox message, vbInformation, "If...Then...Else Example"


' ---------------------------------------
' If...Then...ElseIf
' ---------------------------------------
num = 0
message = "If...Then...ElseIf:" & vbCrLf

If num > 0 Then
    message = message & "Number is positive." & vbCrLf
ElseIf num < 0 Then
    message = message & "Number is negative." & vbCrLf
Else
    message = message & "Number is zero." & vbCrLf
End If

MsgBox message, vbInformation, "If...Then...ElseIf Example"


' ---------------------------------------
' Select Case
' ---------------------------------------
Dim grade
grade = "B"
message = "Select Case:" & vbCrLf

Select Case grade
    Case "A"
        message = message & "Excellent!" & vbCrLf
    Case "B"
        message = message & "Good job!" & vbCrLf
    Case "C"
        message = message & "You passed." & vbCrLf
    Case Else
        message = message & "Invalid grade." & vbCrLf
End Select

MsgBox message, vbInformation, "Select Case Example"

' =======================================
