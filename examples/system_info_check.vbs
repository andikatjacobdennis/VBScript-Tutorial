' =======================================
'   VBScript Example: System Information Check (Open in Notepad)
' =======================================

On Error Resume Next  ' Continue execution even if errors occur

' ---------------------------------------
' Create FileSystemObject to save temporary report
' ---------------------------------------
Dim fso, tempFilePath
Set fso = CreateObject("Scripting.FileSystemObject")

' Use Windows Temp folder for temporary file
tempFilePath = fso.GetSpecialFolder(2) & "\SystemReport.txt"  ' 2 = TemporaryFolder

Dim reportFile
Set reportFile = fso.CreateTextFile(tempFilePath, True)

' ---------------------------------------
' Function: WriteReport
' Writes text to the report file
' ---------------------------------------
Sub WriteReport(text)
    reportFile.WriteLine text
End Sub

' ---------------------------------------
' Create WMI Object
' ---------------------------------------
Dim objWMI, colItems, objItem
Set objWMI = GetObject("winmgmts:\\.\root\CIMV2")

' ---------------------------------------
' Hardware Information
' ---------------------------------------
WriteReport "========== Hardware Information =========="

Set colItems = objWMI.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
    WriteReport "CPU Name: " & objItem.Name
    WriteReport "Number of Cores: " & objItem.NumberOfCores
    WriteReport "Max Clock Speed: " & objItem.MaxClockSpeed & " MHz"
Next

Set colItems = objWMI.ExecQuery("Select * from Win32_ComputerSystem")
For Each objItem in colItems
    WriteReport "Total Physical Memory: " & FormatNumber(objItem.TotalPhysicalMemory / 1024 / 1024, 0) & " MB"
Next

Set colItems = objWMI.ExecQuery("Select * from Win32_BIOS")
For Each objItem in colItems
    WriteReport "BIOS Version: " & objItem.SMBIOSBIOSVersion
    WriteReport "BIOS Manufacturer: " & objItem.Manufacturer
Next

' ---------------------------------------
' Operating System Information
' ---------------------------------------
WriteReport vbCrLf & "========== Operating System =========="
Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem")
For Each objItem in colItems
    WriteReport "OS Name: " & objItem.Caption
    WriteReport "OS Version: " & objItem.Version
    WriteReport "OS Architecture: " & objItem.OSArchitecture
    WriteReport "Last Boot Up Time: " & objItem.LastBootUpTime
Next

' ---------------------------------------
' Disk Information
' ---------------------------------------
WriteReport vbCrLf & "========== Disk Drives =========="
Set colItems = objWMI.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3")
For Each objItem in colItems
    WriteReport "Drive: " & objItem.DeviceID
    WriteReport "File System: " & objItem.FileSystem
    WriteReport "Total Size: " & FormatNumber(objItem.Size / 1024 / 1024 / 1024, 2) & " GB"
    WriteReport "Free Space: " & FormatNumber(objItem.FreeSpace / 1024 / 1024 / 1024, 2) & " GB"
Next

' ---------------------------------------
' Network Information
' ---------------------------------------
WriteReport vbCrLf & "========== Network Configuration =========="
Set colItems = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=True")
For Each objItem in colItems
    WriteReport "Description: " & objItem.Description
    If Not IsNull(objItem.IPAddress) Then
        WriteReport "IP Address: " & Join(objItem.IPAddress, ", ")
    End If
    WriteReport "MAC Address: " & objItem.MACAddress
Next

' ---------------------------------------
' Running Processes
' ---------------------------------------
WriteReport vbCrLf & "========== Running Processes =========="
Set colItems = objWMI.ExecQuery("Select * from Win32_Process")
For Each objItem in colItems
    WriteReport "Process: " & objItem.Name & " | PID: " & objItem.ProcessId
Next

' ---------------------------------------
' Finalize Report and Open in Notepad
' ---------------------------------------
reportFile.Close
Set reportFile = Nothing

' Open the temporary file in Notepad
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "notepad.exe " & tempFilePath, 1, False  ' 1 = normal window, False = don't wait
Set shell = Nothing

On Error GoTo 0
