' =======================================
'   VBScript Example: Simple Unit Tests
' =======================================
'
' Filename: 14_01_Simple_Unit_Tests.vbs
'
' Demonstrates:
' - Writing simple test functions in VBScript
' - Creating an assertion helper
' - Running test cases and reporting results
' - Using MsgBox for summary
'
' Notes:
' - VBScript has no built-in testing framework
' - We simulate "unit tests" using functions and assertions
' - For larger projects, consider structured logging of results
' =======================================

Option Explicit

Dim totalTests, passedTests, failedTests
totalTests = 0
passedTests = 0
failedTests = 0

' ---------------------------------------
' Assertion helper
' ---------------------------------------
Sub AssertEquals(testName, expected, actual)
    totalTests = totalTests + 1
    If expected = actual Then
        passedTests = passedTests + 1
        WScript.Echo "[PASS] " & testName
    Else
        failedTests = failedTests + 1
        WScript.Echo "[FAIL] " & testName & vbCrLf & _
                     "   Expected: " & expected & vbCrLf & _
                     "   Actual: " & actual
    End If
End Sub

' ---------------------------------------
' Example functions to test
' ---------------------------------------
Function Add(a, b)
    Add = a + b
End Function

Function IsEven(n)
    If n Mod 2 = 0 Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

' ---------------------------------------
' Test cases
' ---------------------------------------
Sub RunTests()
    ' Test Add()
    AssertEquals "Add(2, 3) should be 5", 5, Add(2, 3)
    AssertEquals "Add(-1, 1) should be 0", 0, Add(-1, 1)

    ' Test IsEven()
    AssertEquals "IsEven(4) should be True", True, IsEven(4)
    AssertEquals "IsEven(5) should be False", False, IsEven(5)
End Sub

' Run the tests
RunTests()

' ---------------------------------------
' Report summary
' ---------------------------------------
Dim summary
summary = "Unit Test Results:" & vbCrLf & _
          "Total: " & totalTests & vbCrLf & _
          "Passed: " & passedTests & vbCrLf & _
          "Failed: " & failedTests

MsgBox summary, vbInformation, "Test Summary"

' ---------------------------------------
' Notes:
' 1. Use WScript.Echo for console output (works if run with cscript.exe).
' 2. MsgBox is used here for a final summary.
' 3. You can extend this by writing results to a log file.
' 4. This script illustrates the concept of "unit tests" in plain VBScript.
' ---------------------------------------
