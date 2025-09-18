' =======================================
'   VBScript Example: Registry
' =======================================

' ---------------------------------------
' Reading a Registry Value
' ---------------------------------------
Dim WshShell, regValue
Set WshShell = CreateObject("WScript.Shell")

' Example: Read Windows version from registry
On Error Resume Next
regValue = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
If Err.Number = 0 Then
    MsgBox "Windows Version: " & regValue, vbInformation, "Registry Read"
Else
    MsgBox "Failed to read registry value.", vbExclamation, "Registry Read"
End If
On Error GoTo 0

' ---------------------------------------
' Writing a Registry Value
' ---------------------------------------
' Example: Add a test string value (HKEY_CURRENT_USER)
On Error Resume Next
WshShell.RegWrite "HKCU\Software\VBScriptExample\TestValue", "Hello Registry", "REG_SZ"
If Err.Number = 0 Then
    MsgBox "Registry value written successfully.", vbInformation, "Registry Write"
Else
    MsgBox "Failed to write registry value.", vbExclamation, "Registry Write"
End If
On Error GoTo 0

' ---------------------------------------
' Deleting a Registry Value
' ---------------------------------------
' Example: Delete the test value
On Error Resume Next
WshShell.RegDelete "HKCU\Software\VBScriptExample\TestValue"
If Err.Number = 0 Then
    MsgBox "Registry value deleted successfully.", vbInformation, "Registry Delete"
Else
    MsgBox "Failed to delete registry value.", vbExclamation, "Registry Delete"
End If
On Error GoTo 0

' =======================================
