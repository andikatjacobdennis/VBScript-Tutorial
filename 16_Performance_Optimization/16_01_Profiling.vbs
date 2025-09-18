' =======================================
'   VBScript Example: Profiling & Timing
' =======================================
'
' Filename: 16_01_Profiling.vbs
'
' Demonstrates:
' - Measuring execution time for code sections
' - Using Timer() function for profiling
' - Comparing performance of different approaches
' - Reporting results
'
' Notes:
' - Timer() returns the number of seconds since midnight with fractional seconds
' - For long-running scripts, store timestamps at multiple points
' - This simple profiler can help identify slow sections in VBScript
' =======================================

Option Explicit

Dim startTime, endTime, elapsed

' ---------- Example 1: Profiling a loop ----------
startTime = Timer()

Dim i, sum
sum = 0
For i = 1 To 1000000
    sum = sum + i
Next

endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 1: Sum loop executed in " & FormatNumber(elapsed, 4) & " seconds." & vbCrLf & _
       "Result: " & sum, vbInformation, "Profiling - Example 1"

' ---------- Example 2: Profiling array operations ----------
Dim arr(), j
ReDim arr(1 To 1000000)
startTime = Timer()

For j = 1 To 1000000
    arr(j) = j * 2
Next

endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 2: Array population executed in " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Profiling - Example 2"

' ---------- Example 3: Comparing string concatenation methods ----------
Dim s, k
Const ITERATIONS = 50000

' Method A: Using simple concatenation
s = ""
startTime = Timer()
For k = 1 To ITERATIONS
    s = s & "x"
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "String concatenation (Method A) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Profiling - Example 3A"

' Method B: Using array join (faster for large strings)
Dim parts()
ReDim parts(1 To ITERATIONS)
For k = 1 To ITERATIONS
    parts(k) = "x"
Next
startTime = Timer()
s = Join(parts, "")
endTime = Timer()
elapsed = endTime - startTime
MsgBox "String concatenation (Method B using Join) took " & FormatNumber(elapsed, 4) & " seconds.", vbInformation, "Profiling - Example 3B"

' ---------------------------------------
' Notes:
' 1. Use Timer() for simple profiling; for very long-running scripts spanning midnight, handle Timer rollover.
' 2. Profiling helps identify bottlenecks such as loops, string operations, or file I/O.
' 3. Consider breaking large tasks into smaller functions for easier measurement and optimization.
' 4. Array.Join is usually faster than repeated string concatenation for large datasets.
' ---------------------------------------
