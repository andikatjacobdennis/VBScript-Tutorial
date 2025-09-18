' =======================================
'   VBScript Example: Event Viewer Access
' =======================================
'
' Demonstrates:
' - Reading Windows Event Viewer logs (Application/System/Security)
' - Prompting user for which log to read
' - Sanitizing input
' - Displaying recent entries in a MsgBox
'
' Notes:
' - Requires appropriate permissions to read certain logs (e.g., Security).
' - WMI is used here for portability across Windows versions.
' =======================================

Option Explicit
On Error Resume Next

' ---------------------------------------
' Utility: Sanitize input (log name)
' ---------------------------------------
Function SanitizeLogName(logName)
    Dim s
    s = Trim(LCase(logName))
    Select Case s
        Case "application", "system", "security"
            SanitizeLogName = UCase(Left(s,1)) & Mid(s,2)  ' Capitalize
        Case Else
            SanitizeLogName = "Application"  ' default fallback
    End Select
End Function

' ---------------------------------------
' Prompt user for which log to view
' ---------------------------------------
Dim rawInput, logName
rawInput = InputBox("Enter which Event Log to view (Application, System, Security):", _
                    "Event Viewer Example", "Application")
logName = SanitizeLogName(rawInput)

If logName = "" Then
    MsgBox "No log selected. Exiting.", vbExclamation, "Aborted"
    WScript.Quit
End If

' ---------------------------------------
' Connect to WMI and query events
' ---------------------------------------
Dim objWMI, query, colEvents, evt, output, count
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

If Err.Number <> 0 Then
    MsgBox "Unable to connect to WMI." & vbCrLf & _
           "Err#: " & Err.Number & " Desc: " & Err.Description, _
           vbCritical, "WMI Error"
    WScript.Quit
End If

' Build query: get latest 5 events
query = "SELECT * FROM Win32_NTLogEvent WHERE Logfile='" & logName & "' ORDER BY TimeGenerated DESC"
Set colEvents = objWMI.ExecQuery(query)

If Err.Number <> 0 Then
    MsgBox "Error querying event log." & vbCrLf & _
           "Err#: " & Err.Number & " Desc: " & Err.Description, _
           vbCritical, "Query Error"
    WScript.Quit
End If

' ---------------------------------------
' Display results (up to 5 entries)
' ---------------------------------------
output = "Latest events from " & logName & " log:" & vbCrLf & String(40, "-") & vbCrLf
count = 0

For Each evt In colEvents
    count = count + 1
    output = output & "Time: " & evt.TimeGenerated & vbCrLf & _
                      "Source: " & evt.SourceName & vbCrLf & _
                      "Type: " & evt.Type & vbCrLf & _
                      "Event ID: " & evt.EventCode & vbCrLf & _
                      "Message: " & evt.Message & vbCrLf & _
                      String(40, "-") & vbCrLf
    If count >= 5 Then Exit For
Next

If count = 0 Then
    MsgBox "No events found in the " & logName & " log.", vbInformation, "No Results"
Else
    MsgBox output, vbInformation, logName & " Events"
End If

' ---------------------------------------
' Clean up
' ---------------------------------------
Set colEvents = Nothing
Set objWMI = Nothing
On Error GoTo 0

' ---------------------------------------
' Notes / Guidance:
' 1. Only Application and System logs are accessible by normal users.
'    Security log requires admin privileges.
' 2. Queries can filter by EventCode, SourceName, or TimeGenerated.
' 3. MsgBox truncates long messages; for serious use, write to a file.
' 4. Always sanitize user input to avoid WMI query injection.
' ---------------------------------------
