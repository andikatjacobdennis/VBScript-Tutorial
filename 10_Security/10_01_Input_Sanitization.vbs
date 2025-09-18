' =======================================
'   VBScript Example: Input Sanitization
' =======================================

' ---------------------------------------
' This section demonstrates how to sanitize
' and validate user input in VBScript
' ---------------------------------------

' Function to sanitize input by trimming spaces and removing unsafe characters
Function SanitizeInput(userInput)
    Dim cleanInput
    ' Trim leading and trailing spaces
    cleanInput = Trim(userInput)

    ' Replace dangerous characters with safe equivalents
    cleanInput = Replace(cleanInput, ";", "")
    cleanInput = Replace(cleanInput, "--", "")
    cleanInput = Replace(cleanInput, "'", "")
    cleanInput = Replace(cleanInput, """", "")

    SanitizeInput = cleanInput
End Function

' Function to check if input is numeric
Function IsNumericInput(userInput)
    If IsNumeric(userInput) Then
        IsNumericInput = True
    Else
        IsNumericInput = False
    End If
End Function

' Prompt user for input
Dim rawInput, safeInput
rawInput = InputBox("Enter a number:", "Input Sanitization Example")

' Sanitize the input
safeInput = SanitizeInput(rawInput)

' Validate numeric input
If IsNumericInput(safeInput) Then
    MsgBox "Valid input: " & safeInput, vbInformation, "Sanitization Success"
Else
    MsgBox "Invalid input detected!" & vbCrLf & _
           "Original: " & rawInput & vbCrLf & _
           "Sanitized: " & safeInput, vbCritical, "Sanitization Failed"
End If

' ---------------------------------------
' Notes:
' 1. Always sanitize inputs from users.
' 2. Remove or escape unsafe characters to prevent misuse.
' 3. Validate inputs (e.g., numeric check) to ensure correctness.
' 4. Combine sanitization + validation for safer scripts.
' ---------------------------------------
