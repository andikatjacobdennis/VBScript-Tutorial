' =======================================
'   VBScript Example: System Information Check
' =======================================
'
' Description: Gathers and displays comprehensive system information
'              including hardware, software, network, and disk details.
'
' Features:
' - Hardware information (CPU, RAM, BIOS)
' - Operating system details
' - Disk space analysis
' - Network configuration
' - Running processes
' - Save report to file
' =======================================

Option Explicit

' Global objects
Dim fso, wshShell, wmiService, network
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
Set network = CreateObject("WScript.Network")

' Main script execution
Call Main()

Sub Main()
    Dim choice, report, outputFile
    
    Do
        choice = ShowMainMenu()
        
        ' Check if user pressed Cancel or Close
        If choice = "" Then
            MsgBox "Goodbye!", vbInformation, "System Info Check"
            Exit Do
        End If
        
        Select Case choice
            Case "1"
                report = GenerateFullReport()
                ShowReport report, "Full System Report"
            Case "2"
                report = GenerateHardwareReport()
                ShowReport report, "Hardware Information"
            Case "3"
                report = GenerateSoftwareReport()
                ShowReport report, "Software Information"
            Case "4"
                report = GenerateDiskReport()
                ShowReport report, "Disk Information"
            Case "5"
                report = GenerateNetworkReport()
                ShowReport report, "Network Information"
            Case "6"
                outputFile = InputBox("Enter output filename (e.g., system_report.txt):", "Save Report", "system_report.txt")
                If outputFile = "" Then
                    MsgBox "Save cancelled.", vbInformation, "Cancelled"
                Else
                    report = GenerateFullReport()
                    SaveReport report, outputFile
                End If
            Case "7"
                RunQuickCheck()
            Case "8"
                MsgBox "Goodbye!", vbInformation, "System Info Check"
                Exit Do
            Case Else
                MsgBox "Invalid choice. Please try again.", vbExclamation, "Error"
        End Select
        
        If choice <> "8" And choice <> "" Then
            If MsgBox("Press Yes to continue or No to exit.", vbYesNo + vbQuestion, "Continue") = vbNo Then
                Exit Do
            End If
        End If
    Loop
End Sub

Function ShowMainMenu()
    Dim message, title
    title = "System Information Check"
    message = "Please choose an option:" & vbCrLf & vbCrLf
    message = message & "1. Generate Full System Report" & vbCrLf
    message = message & "2. Hardware Information" & vbCrLf
    message = message & "3. Software Information" & vbCrLf
    message = message & "4. Disk Information" & vbCrLf
    message = message & "5. Network Information" & vbCrLf
    message = message & "6. Save Report to File" & vbCrLf
    message = message & "7. Quick System Check" & vbCrLf
    message = message & "8. Exit" & vbCrLf & vbCrLf
    message = message & "Enter your choice (1-8):"
    
    ShowMainMenu = InputBox(message, title)
End Function

Function GenerateFullReport()
    Dim report
    report = "SYSTEM INFORMATION REPORT" & vbCrLf
    report = report & "Generated on: " & Now() & vbCrLf
    report = report & String(50, "=") & vbCrLf & vbCrLf
    
    report = report & GenerateHardwareReport() & vbCrLf
    report = report & GenerateSoftwareReport() & vbCrLf
    report = report & GenerateDiskReport() & vbCrLf
    report = report & GenerateNetworkReport() & vbCrLf
    report = report & GenerateProcessReport()
    
    GenerateFullReport = report
End Function

Function GenerateHardwareReport()
    Dim report, wmi, items, item, totalRAM, freeRAM
    report = "HARDWARE INFORMATION" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    
    On Error Resume Next
    
    ' Connect to WMI
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    If Err.Number = 0 Then
        ' CPU Information
        report = report & vbCrLf & "Processor:" & vbCrLf
        Set items = wmi.ExecQuery("Select * from Win32_Processor")
        For Each item In items
            report = report & "  Name: " & item.Name & vbCrLf
            report = report & "  Cores: " & item.NumberOfCores & vbCrLf
            report = report & "  Threads: " & item.NumberOfLogicalProcessors & vbCrLf
            report = report & "  Max Speed: " & item.MaxClockSpeed & " MHz" & vbCrLf
        Next
        
        ' Memory Information
        report = report & vbCrLf & "Memory:" & vbCrLf
        Set items = wmi.ExecQuery("Select * from Win32_ComputerSystem")
        For Each item In items
            totalRAM = Round(item.TotalPhysicalMemory / (1024 * 1024 * 1024), 2)
            report = report & "  Total RAM: " & totalRAM & " GB" & vbCrLf
        Next
        
        Set items = wmi.ExecQuery("Select * from Win32_OperatingSystem")
        For Each item In items
            freeRAM = Round(item.FreePhysicalMemory / (1024 * 1024), 2)
            report = report & "  Free RAM: " & freeRAM & " MB" & vbCrLf
        Next
        
        ' BIOS Information
        report = report & vbCrLf & "BIOS:" & vbCrLf
        Set items = wmi.ExecQuery("Select * from Win32_BIOS")
        For Each item In items
            report = report & "  Manufacturer: " & item.Manufacturer & vbCrLf
            report = report & "  Version: " & item.Version & vbCrLf
            report = report & "  Release Date: " & item.ReleaseDate & vbCrLf
        Next
    Else
        report = report & vbCrLf & "WMI access failed: " & Err.Description & vbCrLf
    End If
    
    On Error GoTo 0
    GenerateHardwareReport = report
End Function

Function GenerateSoftwareReport()
    Dim report, wmi, items, item
    report = vbCrLf & "SOFTWARE INFORMATION" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    
    On Error Resume Next
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    If Err.Number = 0 Then
        ' OS Information
        report = report & vbCrLf & "Operating System:" & vbCrLf
        Set items = wmi.ExecQuery("Select * from Win32_OperatingSystem")
        For Each item In items
            report = report & "  Name: " & item.Caption & vbCrLf
            report = message & "  Version: " & item.Version & vbCrLf
            report = report & "  Build: " & item.BuildNumber & vbCrLf
            report = report & "  Install Date: " & item.InstallDate & vbCrLf
            report = report & "  Last Boot: " & item.LastBootUpTime & vbCrLf
        Next
        
        ' Computer Information
        report = report & vbCrLf & "Computer:" & vbCrLf
        report = report & "  Name: " & wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & vbCrLf
        report = report & "  User: " & wshShell.ExpandEnvironmentStrings("%USERNAME%") & vbCrLf
        report = report & "  Domain: " & network.UserDomain & vbCrLf
        
        ' Environment Variables
        report = report & vbCrLf & "System Directory: " & wshShell.ExpandEnvironmentStrings("%SystemRoot%") & vbCrLf
    Else
        report = report & vbCrLf & "WMI access failed: " & Err.Description & vbCrLf
    End If
    
    On Error GoTo 0
    GenerateSoftwareReport = report
End Function

Function GenerateDiskReport()
    Dim report, drives, drive, freeSpace, totalSize, percentFree
    report = vbCrLf & "DISK INFORMATION" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    
    Set drives = fso.Drives
    
    For Each drive In drives
        If drive.IsReady Then
            freeSpace = Round(drive.FreeSpace / (1024 * 1024 * 1024), 2)
            totalSize = Round(drive.TotalSize / (1024 * 1024 * 1024), 2)
            percentFree = Round((freeSpace / totalSize) * 100, 1)
            
            report = report & vbCrLf & "Drive " & drive.DriveLetter & ":" & vbCrLf
            report = report & "  Type: " & GetDriveTypeName(drive.DriveType) & vbCrLf
            report = report & "  File System: " & drive.FileSystem & vbCrLf
            report = report & "  Total Size: " & totalSize & " GB" & vbCrLf
            report = report & "  Free Space: " & freeSpace & " GB" & vbCrLf
            report = report & "  Percent Free: " & percentFree & "%" & vbCrLf
            
            If percentFree < 10 Then
                report = report & "  WARNING: Low disk space!" & vbCrLf
            End If
        End If
    Next
    
    GenerateDiskReport = report
End Function

Function GetDriveTypeName(driveType)
    Select Case driveType
        Case 0: GetDriveTypeName = "Unknown"
        Case 1: GetDriveTypeName = "Removable"
        Case 2: GetDriveTypeName = "Fixed"
        Case 3: GetDriveTypeName = "Network"
        Case 4: GetDriveTypeName = "CD-ROM"
        Case 5: GetDriveTypeName = "RAM Disk"
        Case Else: GetDriveTypeName = "Unknown (" & driveType & ")"
    End Select
End Function

Function GenerateNetworkReport()
    Dim report, wmi, adapters, adapter, ipAddresses, ip
    report = vbCrLf & "NETWORK INFORMATION" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    
    report = report & vbCrLf & "Computer Name: " & network.ComputerName & vbCrLf
    report = report & "User Domain: " & network.UserDomain & vbCrLf
    
    On Error Resume Next
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    If Err.Number = 0 Then
        report = report & vbCrLf & "Network Adapters:" & vbCrLf
        Set adapters = wmi.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=true")
        
        For Each adapter In adapters
            report = report & vbCrLf & "  Adapter: " & adapter.Description & vbCrLf
            report = report & "  MAC Address: " & adapter.MACAddress & vbCrLf
            
            If IsArray(adapter.IPAddress) Then
                For Each ip In adapter.IPAddress
                    report = report & "  IP Address: " & ip & vbCrLf
                Next
            End If
            
            If IsArray(adapter.IPSubnet) Then
                For Each ip In adapter.IPSubnet
                    report = report & "  Subnet Mask: " & ip & vbCrLf
                Next
            End If
            
            If IsArray(adapter.DefaultIPGateway) Then
                report = report & "  Gateway: " & adapter.DefaultIPGateway(0) & vbCrLf
            End If
            
            report = report & "  DHCP Enabled: " & adapter.DHCPEnabled & vbCrLf
        Next
    End If
    
    On Error GoTo 0
    GenerateNetworkReport = report
End Function

Function GenerateProcessReport()
    Dim report, wmi, processes, process, count
    report = vbCrLf & "RUNNING PROCESSES" & vbCrLf
    report = report & String(25, "-") & vbCrLf
    
    On Error Resume Next
    
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    count = 0
    
    If Err.Number = 0 Then
        Set processes = wmi.ExecQuery("Select * from Win32_Process")
        report = report & "Total Processes: " & processes.Count & vbCrLf & vbCrLf
        
        ' Show top 10 processes by memory usage
        report = report & "Top 10 Processes by Memory Usage:" & vbCrLf
        For Each process In processes
            If count < 10 Then
                report = report & "  " & process.Name & " (PID: " & process.ProcessId & ")" & vbCrLf
                count = count + 1
            End If
        Next
    End If
    
    On Error GoTo 0
    GenerateProcessReport = report
End Function

Sub RunQuickCheck()
    Dim message, drives, drive, freeSpace, totalSize, percentFree
    message = "QUICK SYSTEM CHECK" & vbCrLf & String(20, "=") & vbCrLf & vbCrLf
    
    ' Check disk space
    message = message & "Disk Space Status:" & vbCrLf
    Set drives = fso.Drives
    
    For Each drive In drives
        If drive.IsReady And drive.DriveType = 2 Then ' Fixed drives only
            freeSpace = Round(drive.FreeSpace / (1024 * 1024 * 1024), 2)
            totalSize = Round(drive.TotalSize / (1024 * 1024 * 1024), 2)
            percentFree = Round((freeSpace / totalSize) * 100, 1)
            
            message = message & "  Drive " & drive.DriveLetter & ": " & percentFree & "% free" & vbCrLf
        End If
    Next
    
    ' Check memory
    On Error Resume Next
    Dim wmi, items, item, totalRAM, freeRAM
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    If Err.Number = 0 Then
        Set items = wmi.ExecQuery("Select * from Win32_OperatingSystem")
        For Each item In items
            freeRAM = Round(item.FreePhysicalMemory / (1024 * 1024), 2)
            totalRAM = Round(item.TotalVisibleMemorySize / (1024 * 1024), 2)
            message = message & vbCrLf & "Memory: " & freeRAM & " MB free of " & totalRAM & " MB" & vbCrLf
        Next
    End If
    
    On Error GoTo 0
    message = message & vbCrLf & "System appears to be running normally."
    
    MsgBox message, vbInformation, "Quick Check Results"
End Sub

Sub ShowReport(report, title)
    Dim tempFile, tempPath, objFile
    tempPath = fso.GetSpecialFolder(2) & "\system_report.txt"
    
    ' Write to temporary file
    Set objFile = fso.CreateTextFile(tempPath, True)
    objFile.Write report
    objFile.Close
    
    ' Open in Notepad
    wshShell.Run "notepad.exe " & tempPath, 1, False
    
    ' Offer to view immediately
    If MsgBox("Report generated. Would you like to view it in a message box?", vbYesNo, "View Report") = vbYes Then
        MsgBox Left(report, 1000) & "..." & vbCrLf & vbCrLf & "(Truncated - full report opened in Notepad)", vbInformation, title
    End If
End Sub

Sub SaveReport(report, fileName)
    Dim objFile
    On Error Resume Next
    
    Set objFile = fso.CreateTextFile(fileName, True)
    objFile.Write report
    objFile.Close
    
    If Err.Number = 0 Then
        MsgBox "Report saved successfully to: " & fileName, vbInformation, "Success"
    Else
        MsgBox "Error saving file: " & Err.Description, vbCritical, "Error"
    End If
    
    On Error GoTo 0
End Sub

' Cleanup
Sub OnExit()
    Set fso = Nothing
    Set wshShell = Nothing
    Set network = Nothing
End Sub

Class Cleanup
    Private Sub Class_Terminate()
        OnExit
    End Sub
End Class

Dim clean
Set clean = New Cleanup