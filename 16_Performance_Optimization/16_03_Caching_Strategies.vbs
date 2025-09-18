' =======================================
'   VBScript Example: Caching Strategies
' =======================================
'
' Filename: 16_03_Caching_Strategies.vbs
'
' Demonstrates:
' - Implementing simple caching in VBScript
' - Reducing repeated calculations and object access
' - Comparing performance with and without caching
'
' Notes:
' - Caching can significantly improve performance in loops or repeated operations
' - Examples use in-memory dictionaries for storing computed results
' =======================================

Option Explicit

Dim startTime, endTime, elapsed, i

' ---------- Example 1: Without caching ----------
Function ExpensiveCalc(x)
    ' Simulate expensive calculation
    WScript.Sleep 1  ' 1 ms delay
    ExpensiveCalc = x * x
End Function

Dim result
startTime = Timer()
For i = 1 To 50
    result = ExpensiveCalc(i)
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 1 (no caching) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Caching Demo"

' ---------- Example 2: With caching using Dictionary ----------
Dim cache
Set cache = CreateObject("Scripting.Dictionary")

Function ExpensiveCalcCached(x)
    If cache.Exists(x) Then
        ExpensiveCalcCached = cache(x)
    Else
        ' Simulate expensive calculation
        WScript.Sleep 1  ' 1 ms delay
        cache(x) = x * x
        ExpensiveCalcCached = cache(x)
    End If
End Function

startTime = Timer()
For i = 1 To 50
    result = ExpensiveCalcCached(i)
Next
' Repeat same calls to demonstrate caching effect
For i = 1 To 50
    result = ExpensiveCalcCached(i)
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 2 (with caching) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Caching Demo"

' ---------- Example 3: Caching file reads ----------
Dim fso, filePath, fileContent
Set fso = CreateObject("Scripting.FileSystemObject")
filePath = "testfile.txt"

' Create a test file if it doesn't exist
If Not fso.FileExists(filePath) Then
    Dim f
    Set f = fso.CreateTextFile(filePath, True)
    f.WriteLine "Line 1"
    f.WriteLine "Line 2"
    f.WriteLine "Line 3"
    f.Close
End If

' Without caching: read file multiple times
startTime = Timer()
For i = 1 To 100
    Dim f, txt
    Set f = fso.OpenTextFile(filePath, 1)
    txt = f.ReadAll
    f.Close
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 3 (file read without caching) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Caching Demo"

' With caching: read once and reuse
Dim fileCache
fileCache = ""
startTime = Timer()
If fileCache = "" Then
    Dim f
    Set f = fso.OpenTextFile(filePath, 1)
    fileCache = f.ReadAll
    f.Close
End If

' Simulate multiple accesses
Dim j
For j = 1 To 100
    txt = fileCache
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 3 (file read with caching) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Caching Demo"

' Cleanup
Set cache = Nothing
Set fso = Nothing

' ---------------------------------------
' Notes:
' 1. Use Dictionary objects for caching computed values or repeated lookups.
' 2. File reads, database queries, and expensive calculations benefit most from caching.
' 3. Always ensure cached data is valid or refreshed when necessary.
' 4. Profiling is important to measure the benefit of caching in real scenarios.
' ---------------------------------------
