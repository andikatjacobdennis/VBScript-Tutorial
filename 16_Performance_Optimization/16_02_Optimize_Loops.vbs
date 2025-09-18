' =======================================
'   VBScript Example: Loop Optimization
' =======================================
'
' Filename: 16_02_Optimize_Loops.vbs
'
' Demonstrates:
' - Techniques to optimize loops in VBScript
' - Using arrays, precomputed values, and avoiding repetitive operations
' - Profiling loop performance with Timer()
'
' Notes:
' - Loops are a common source of performance bottlenecks
' - Avoid unnecessary recalculations or object access inside loops
' =======================================

Option Explicit

Dim startTime, endTime, elapsed, i

' ---------- Example 1: Looping with repeated property access ----------
Dim dict, key
Set dict = CreateObject("Scripting.Dictionary")
For i = 1 To 1000
    dict.Add "Key" & i, i
Next

' Inefficient loop: accessing Count each iteration
startTime = Timer()
For i = 0 To dict.Count - 1
    key = dict.Keys()(i)
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 1 (inefficient Count access): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' Optimized loop: store Count in variable
Dim nCount
nCount = dict.Count
startTime = Timer()
For i = 0 To nCount - 1
    key = dict.Keys()(i)
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 1 (optimized Count stored): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' ---------- Example 2: Avoid repeated object creation ----------
Dim s, ITERATIONS
ITERATIONS = 50000

' Inefficient: concatenating string repeatedly
s = ""
startTime = Timer()
For i = 1 To ITERATIONS
    s = s & "x"
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 2 (repeated concatenation): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' Optimized: using array and Join
Dim arr(), k
ReDim arr(1 To ITERATIONS)
For k = 1 To ITERATIONS
    arr(k) = "x"
Next
startTime = Timer()
s = Join(arr, "")
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 2 (array + Join): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' ---------- Example 3: Minimize expensive function calls ----------
Function ExpensiveCalc(x)
    ExpensiveCalc = Sqr(x) * Log(x + 1)
End Function

Dim result
startTime = Timer()
For i = 1 To 10000
    result = ExpensiveCalc(i)  ' calling function each iteration
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 3 (direct function call each iteration): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' Optimized: precompute if possible
Dim precomputed(1 To 10000)
For i = 1 To 10000
    precomputed(i) = ExpensiveCalc(i)
Next
startTime = Timer()
For i = 1 To 10000
    result = precomputed(i)  ' use precomputed value
Next
endTime = Timer()
elapsed = endTime - startTime
MsgBox "Example 3 (using precomputed array): " & FormatNumber(elapsed, 5) & " seconds.", vbInformation, "Loop Optimization"

' ---------------------------------------
' Notes:
' 1. Store loop bounds, object properties, or expensive calculations outside the loop if possible.
' 2. Avoid repetitive object creation inside loops.
' 3. Using arrays and Join can drastically improve string concatenation performance.
' 4. Profiling helps identify which loops benefit most from optimization.
' ---------------------------------------
