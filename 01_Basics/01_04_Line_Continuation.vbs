' =======================================
'   VBScript Example: Line Continuation
' =======================================

' Declare Variables
Dim num1, num2, result

' Declare Constants
Const MESSAGE_TITLE = "Addition Result"

' Assign Values to Variables
num1 = 10
num2 = 20

' Perform Addition
result = num1 + num2

' Display Result in Message Box
MsgBox "The sum of " & num1 & " and " & num2 & " is " & result, vbInformation, MESSAGE_TITLE

' Demonstrate Line Continuation
MsgBox "The sum of " & num1 & _
       " and " & num2 & _
       " is " & result, vbInformation, MESSAGE_TITLE

' =======================================