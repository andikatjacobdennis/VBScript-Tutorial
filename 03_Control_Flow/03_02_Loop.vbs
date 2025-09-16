' =======================================
'   VBScript Example: Loops
' =======================================

' ---------------------------------------
' For...Next Loop
' ---------------------------------------
Dim i, message
message = "For...Next Loop (1 to 10):" & vbCrLf

For i = 1 To 10
    message = message & i & vbCrLf
Next

MsgBox message, vbInformation, "For...Next Example"


' ---------------------------------------
' While...Wend Loop
' ---------------------------------------
i = 1
message = "While...Wend Loop (1 to 10):" & vbCrLf

While i <= 10
    message = message & i & vbCrLf
    i = i + 1
Wend

MsgBox message, vbInformation, "While...Wend Example"


' ---------------------------------------
' Do...Loop Until
' ---------------------------------------
i = 1
message = "Do...Loop Until (1 to 10):" & vbCrLf

Do
    message = message & i & vbCrLf
    i = i + 1
Loop Until i > 10

MsgBox message, vbInformation, "Do...Loop Until Example"


' ---------------------------------------
' Do While...Loop
' ---------------------------------------
i = 1
message = "Do While...Loop (1 to 10):" & vbCrLf

Do While i <= 10
    message = message & i & vbCrLf
    i = i + 1
Loop

MsgBox message, vbInformation, "Do While...Loop Example"

' =======================================