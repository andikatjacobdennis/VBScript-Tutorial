' =======================================
'   VBScript Example: Remote Scripting
' =======================================
'
' Filename: 13_03_Remote_Scripting.vbs
'
' Demonstrates:
' - Executing WMI queries against a remote computer
' - Using authentication (username/password) in connection string
' - Error handling for network or permission failures
' - Retrieving remote system information (OS, processes)
'
' Notes:
' - Requires administrative rights on the remote computer
' - WMI service must be running on the remote computer
' - Firewall must allow WMI/DCOM traffic
' - User account must have rights to perform WMI queries remotely
' =======================================

Option Explicit
On Error Resume Next

Dim strComputer, strUser, strPassword, wmiService, colOS, objOS, colProcesses, objProcess

' Ask user for remote computer name
strComputer = InputBox("Enter remote computer name or IP address:", "Remote WMI", ".")
If Trim(strComputer) = "" Then strComputer = "."

' Optionally ask for credentials (leave blank for current user)
strUser = InputBox("Enter username (DOMAIN\User or User):", "Remote WMI Credentials", "")
strPassword = InputBox("Enter password (leave blank to use current user):", "Remote WMI Credentials", "")

Dim locator
Set locator = CreateObject("WbemScripting.SWbemLocator")

If strUser = "" Then
    ' Connect with current credentials
    Set wmiService = locator.ConnectServer(strComputer, "root\cimv2")
Else
    ' Connect with explicit credentials
    Set wmiService = locator.ConnectServer(strComputer, "root\cimv2", strUser, strPassword)
End If

If Err.Number <> 0 Or wmiService Is Nothing Then
    MsgBox "Failed to connect to WMI on " & strComputer & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Remote WMI Error"
    WScript.Quit 1
End If

' Set security impersonation level
wmiService.Security_.ImpersonationLevel = 3   ' impersonate

' Example 1: Query OS info
Set colOS = wmiService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
If Err.Number = 0 Then
    For Each objOS In colOS
        MsgBox "Remote Computer: " & strComputer & vbCrLf & _
               "OS: " & objOS.Caption & vbCrLf & _
               "Version: " & objOS.Version & vbCrLf & _
               "Last Bootup: " & objOS.LastBootUpTime, vbInformation, "Remote OS Info"
    Next
Else
    MsgBox "Error querying remote OS info: " & Err.Description, vbExclamation, "Remote Query Error"
    Err.Clear
End If

' Example 2: List top 5 running processes
Set colProcesses = wmiService.ExecQuery("SELECT * FROM Win32_Process")
If Err.Number = 0 Then
    Dim count, procList
    count = 0
    procList = "Top 5 Processes on " & strComputer & ":" & vbCrLf
    For Each objProcess In colProcesses
        procList = procList & objProcess.Name & " (PID: " & objProcess.ProcessId & ")" & vbCrLf
        count = count + 1
        If count >= 5 Then Exit For
    Next
    MsgBox procList, vbInformation, "Remote Processes"
Else
    MsgBox "Error querying remote processes: " & Err.Description, vbExclamation, "Remote Query Error"
    Err.Clear
End If

' Cleanup
Set colOS = Nothing
Set colProcesses = Nothing
Set wmiService = Nothing
Set locator = Nothing
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. Run this script with appropriate permissions.
' 2. Remote machine must have WMI/DCOM enabled and firewall opened.
' 3. To test locally, enter "." as the computer name.
' 4. For secure environments, prefer Kerberos authentication and proper domain accounts.
' ---------------------------------------
