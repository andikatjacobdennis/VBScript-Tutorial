' =======================================
'   VBScript Example: Option Explicit & Error Handling
'   Demonstrates how undeclared variables behave
' =======================================

' Enforce variable declaration (helps catch typos like num3 instead of num2)
Option Explicit

' Ignore runtime errors (script continues execution even if an error occurs)
On Error Resume Next

' Declare Variables
Dim num1, num2, result

' Declare Constants
Const MESSAGE_TITLE = "Addition Result"

' Assign Values to Variables
num1 = 10
num3 = 20   ' <-- num3 is NOT declared (will cause an error if Option Explicit is on)

' Perform Addition
result = num1 + num3

' Display Result in Message Box
MsgBox "The sum of " & num1 & " and " & num3 & " is " & result, vbInformation, MESSAGE_TITLE

' =======================================