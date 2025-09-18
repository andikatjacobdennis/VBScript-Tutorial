' =======================================
'   VBScript Example: WMI Basics
' =======================================
'
' Filename: 13_01_WMI.vbs
'
' Demonstrates:
' - Connecting to Windows Management Instrumentation (WMI)
' - Querying system information (OS, processes, services)
' - Executing WMI queries with error handling
' - Iterating results and displaying properties
' - Using impersonation level "impersonate" for queries
'
' Notes:
' - WMI is a powerful automation and management interface.
' - Requires Windows Management Instrumentation service to be running.
' - For remote queries, you can connect to \\ComputerName\root\cimv2
'   (requires permissions and firewall exceptions).
' - Always validate input and use least privilege when accessing remote machines.
' =======================================

Option Explicit
On Error Resume Next

Dim wmiService, colOS, objOS, colProcesses, objProcess, colServices, objService

' Connect to WMI service on local machine
Set wmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

If Err.Number <> 0 Or wmiService Is Nothing Then
    MsgBox "Failed to connect to WMI service. Ensure WMI is running.", vbCritical, "WMI Error"
    WScript.Quit 1
End If

' Example 1: Query Operating System Information
Set colOS = wmiService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
If Err.Number = 0 Then
    For Each objOS In colOS
        MsgBox "Operating System: " & objOS.Caption & vbCrLf & _
               "Version: " & objOS.Version & vbCrLf & _
               "Build Number: " & objOS.BuildNumber & vbCrLf & _
               "Registered User: " & objOS.RegisteredUser & vbCrLf & _
               "Serial Number: " & objOS.SerialNumber, vbInformation, "WMI - OS Info"
    Next
    Set colOS = Nothing
Else
    MsgBox "Error executing OS query: " & Err.Description, vbExclamation, "WMI Query Error"
    Err.Clear
End If

' Example 2: List Top 5 Running Processes
Set colProcesses = wmiService.ExecQuery("SELECT * FROM Win32_Process")
If Err.Number = 0 Then
    Dim count
    count = 0
    Dim procList
    procList = "Running Processes:" & vbCrLf
    For Each objProcess In colProcesses
        procList = procList & objProcess.Name & " (PID: " & objProcess.ProcessId & ")" & vbCrLf
        count = count + 1
        If count >= 5 Then Exit For
    Next
    MsgBox procList, vbInformation, "WMI - Processes"
    Set colProcesses = Nothing
Else
    MsgBox "Error executing Process query: " & Err.Description, vbExclamation, "WMI Query Error"
    Err.Clear
End If

' Example 3: List first 5 Services and their states
Set colServices = wmiService.ExecQuery("SELECT * FROM Win32_Service")
If Err.Number = 0 Then
    Dim svcList
    svcList = "Services:" & vbCrLf
    count = 0
    For Each objService In colServices
        svcList = svcList & objService.Name & " - State: " & objService.State & vbCrLf
        count = count + 1
        If count >= 5 Then Exit For
    Next
    MsgBox svcList, vbInformation, "WMI - Services"
    Set colServices = Nothing
Else
    MsgBox "Error executing Service query: " & Err.Description, vbExclamation, "WMI Query Error"
    Err.Clear
End If

' Cleanup
Set wmiService = Nothing
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. Use "SELECT * FROM Win32_ClassName" to query WMI classes.
' 2. Common classes: Win32_OperatingSystem, Win32_Process, Win32_Service,
'    Win32_ComputerSystem, Win32_LogicalDisk, Win32_NetworkAdapter.
' 3. WMI is extensible — many applications (e.g., antivirus, SQL Server) add providers.
' 4. Use wbemtest.exe or PowerShell Get-WmiObject / Get-CimInstance to explore WMI.
' 5. For automation, you can combine WMI queries with scripts to monitor and manage systems.
' ---------------------------------------
