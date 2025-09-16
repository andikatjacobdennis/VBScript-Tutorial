' =======================================
'   VBScript Example: Error Handling
' =======================================

' ---------------------------------------
' This section demonstrates how to handle runtime errors in VBScript
' using "On Error Resume Next" and "Err" object
' ---------------------------------------

' Enable error handling: VBScript will continue running even if an error occurs
On Error Resume Next

Dim result, divisor

' Example 1: Division by zero
divisor = 0
result = 10 / divisor   ' This would normally cause a runtime error

' Check if an error occurred
If Err.Number <> 0 Then
    ' Display the error number and description
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation, "Error Handling"
    ' Clear the error so we can continue
    Err.Clear
End If

' Example 2: Invalid type conversion
Dim strValue
strValue = "abc"
result = CInt(strValue)  ' This will cause a type mismatch error

If Err.Number <> 0 Then
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbExclamation, "Error Handling"
    Err.Clear
End If

' Disable error handling (optional, restores default behavior)
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. "On Error Resume Next" allows the script to continue after an error.
' 2. "Err.Number" gives the numeric code of the error.
' 3. "Err.Description" provides a textual description of the error.
' 4. Always use "Err.Clear" after handling an error to reset the error object.
' ---------------------------------------