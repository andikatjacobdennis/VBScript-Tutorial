' =======================================
'   VBScript Example: Built-in Functions
' =======================================

' Declare a variable to hold messages
Dim message

' ---------------------------------------
' String Functions
' ---------------------------------------
' Demonstrates common string manipulation functions
message = "String Functions:" & vbCrLf

' Len: returns the length of a string
message = message & "Len(""Hello"") = " & Len("Hello") & vbCrLf

' LCase: converts string to lowercase
message = message & "LCase(""WELCOME"") = " & LCase("WELCOME") & vbCrLf

' UCase: converts string to uppercase
message = message & "UCase(""welcome"") = " & UCase("welcome") & vbCrLf

' Left: returns a specified number of characters from the left of a string
message = message & "Left(""VBScript"", 2) = " & Left("VBScript", 2) & vbCrLf

' Right: returns a specified number of characters from the right of a string
message = message & "Right(""VBScript"", 6) = " & Right("VBScript", 6) & vbCrLf

' Mid: returns a substring starting from a specific position
message = message & "Mid(""VBScript"", 3, 4) = " & Mid("VBScript", 3, 4) & vbCrLf

' Display the string functions in a message box
MsgBox message, vbInformation, "String Functions"


' ---------------------------------------
' Numeric Functions
' ---------------------------------------
' Demonstrates common numeric functions
message = "Numeric Functions:" & vbCrLf

' Abs: returns the absolute value of a number
message = message & "Abs(-15) = " & Abs(-15) & vbCrLf

' Sqr: returns the square root of a number
message = message & "Sqr(25) = " & Sqr(25) & vbCrLf

' Rnd: returns a random number between 0 and 1
message = message & "Rnd() = " & Rnd() & vbCrLf

' Round: rounds a number to a specified number of decimal places
message = message & "Round(3.14159, 2) = " & Round(3.14159, 2) & vbCrLf

' Display the numeric functions in a message box
MsgBox message, vbInformation, "Numeric Functions"


' ---------------------------------------
' Date/Time Functions
' ---------------------------------------
' Demonstrates date and time functions
message = "Date/Time Functions:" & vbCrLf

' Date: returns the current system date
message = message & "Date() = " & Date() & vbCrLf

' Time: returns the current system time
message = message & "Time() = " & Time() & vbCrLf

' Now: returns the current date and time
message = message & "Now() = " & Now() & vbCrLf

' Year, Month, Day: extract year, month, and day from a date
message = message & "Year(Now()) = " & Year(Now()) & vbCrLf
message = message & "Month(Now()) = " & Month(Now()) & vbCrLf
message = message & "Day(Now()) = " & Day(Now()) & vbCrLf

' Display the date/time functions in a message box
MsgBox message, vbInformation, "Date/Time Functions"


' ---------------------------------------
' Conversion Functions
' ---------------------------------------
' Demonstrates converting between data types
message = "Conversion Functions:" & vbCrLf

' CInt: converts a string to an integer
message = message & "CInt(""123"") = " & CInt("123") & vbCrLf

' CDbl: converts a string to a double (decimal) number
message = message & "CDbl(""45.67"") = " & CDbl("45.67") & vbCrLf

' CStr: converts a number to a string
message = message & "CStr(100) = " & CStr(100) & vbCrLf

' Display the conversion functions in a message box
MsgBox message, vbInformation, "Conversion Functions"

' =======================================

' Hex: converts a number to its hexadecimal representation
x = 123
MsgBox "y value after converting to Hex -" & Hex(x), vbInformation, "Hex Function"

' ======================================

' FormatNumber: formats a number to a specified number of decimal places
Dim num1 : num1 = -645.998651
MsgBox FormatNumber(num1, 3), vbInformation, "Format Number"      ' Output: -645.999

' =======================================

' Rnd with a seed: returns a pseudo-random number based on the given seed
Dim num2 : num2 = -645.998651
MsgBox "Rnd Result of num is : " & Rnd(num2), vbInformation, "Rnd Function" ' Example output: 0.5130115
' =======================================

' Regular expression for email validation
Dim emailPattern, testEmail, regEx, isMatch
emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
testEmail = "abc@example.com"
Set regEx = New RegExp
regEx.Pattern = emailPattern
regEx.IgnoreCase = True
regEx.Global = False
isMatch = regEx.Test(testEmail)
If isMatch Then
    MsgBox "The email address '" & testEmail & "' is valid.", vbInformation, "Email Validation"
Else
    MsgBox "The email address '" & testEmail & "' is NOT valid.", vbExclamation, "Email Validation"
End If
Set regEx = Nothing
' =======================================