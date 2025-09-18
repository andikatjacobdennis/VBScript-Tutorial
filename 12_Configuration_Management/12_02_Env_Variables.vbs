' =======================================
'   VBScript Example: Environment Variables
' =======================================
'
' Filename: 12_02_Env_Variables.vbs
'
' Demonstrates:
' - Reading environment variables (Process, User, System)
' - Setting process-level variables (temporary)
' - Setting user-level variables (persistent via HKCU\Environment)
' - Attempting system-level changes (requires admin) using setx /M
' - Deleting user-level variables (registry) and process-level removal
' - Listing variables in each scope
' - Sanitization, error handling and safety notes
'
' Notes:
' - Changing Process environment only affects the current process and children.
' - Writing to HKCU\Environment makes a persistent user-level change but may require
'   a logoff/login or notifying other processes to pick up the change.
' - setx utility can set user or system variables; /M requires elevated privileges.
' - Modifying HKLM for system variables requires administrative rights; this script
'   prefers setx for system writes (if available).
' - This is intended as a tutorial; in production, be careful with destructive ops.
' =======================================

Option Explicit
On Error Resume Next

Dim wshShell
Set wshShell = CreateObject("WScript.Shell")

' ---------- Configuration ----------
Const SCOPE_PROCESS = "Process"
Const SCOPE_USER    = "User"
Const SCOPE_SYSTEM  = "System"

' ---------- Utilities: sanitization ----------
Function SanitizeName(n)
    ' Environment variable names should be non-empty, not contain = and trimmed
    If IsNull(n) Then
        SanitizeName = ""
        Exit Function
    End If
    n = Trim(CStr(n))
    If n = "" Then
        SanitizeName = ""
        Exit Function
    End If
    ' Disallow equals and control chars
    n = Replace(n, "=", "")
    Dim i, ch, out
    out = ""
    For i = 1 To Len(n)
        ch = Asc(Mid(n, i, 1))
        If ch >= 32 And ch <> 61 Then
            out = out & Chr(ch)
        End If
    Next
    SanitizeName = out
End Function

Function SanitizeValue(v)
    If IsNull(v) Then
        SanitizeValue = ""
        Exit Function
    End If
    v = CStr(v)
    ' Trim trailing/leading whitespace but preserve internal formatting
    v = Trim(v)
    ' Remove embedded nulls and control chars except tab/newline if you want
    Dim i, ch, out
    out = ""
    For i = 1 To Len(v)
        ch = Asc(Mid(v, i, 1))
        If ch >= 9 Then
            out = out & Chr(ch)
        End If
    Next
    SanitizeValue = out
End Function

' ---------- Core operations ----------

' GetEnv(scope, name) -> returns value or empty string if not found
Function GetEnv(scope, name)
    Dim sName, env, val
    sName = SanitizeName(name)
    If sName = "" Then
        GetEnv = ""
        Exit Function
    End If
    Select Case scope
        Case SCOPE_PROCESS
            Set env = wshShell.Environment("Process")
        Case SCOPE_USER
            Set env = wshShell.Environment("User")
        Case SCOPE_SYSTEM
            Set env = wshShell.Environment("System")
        Case Else
            Set env = wshShell.Environment("Process")
    End Select
    On Error Resume Next
    val = env.Item(sName)
    If Err.Number <> 0 Then
        Err.Clear
        val = ""
    End If
    GetEnv = val
End Function

' SetEnvProcess(name, value) - sets env var only for this process (and children)
Function SetEnvProcess(name, value)
    Dim sName, sValue, env
    sName = SanitizeName(name)
    sValue = SanitizeValue(value)
    If sName = "" Then
        SetEnvProcess = False
        Exit Function
    End If
    Set env = wshShell.Environment("Process")
    On Error Resume Next
    env.Item(sName) = sValue
    If Err.Number <> 0 Then
        Err.Clear
        SetEnvProcess = False
    Else
        SetEnvProcess = True
    End If
End Function

' Remove a process-level variable
Function DeleteEnvProcess(name)
    Dim sName, env
    sName = SanitizeName(name)
    If sName = "" Then
        DeleteEnvProcess = False
        Exit Function
    End If
    Set env = wshShell.Environment("Process")
    On Error Resume Next
    env.Remove sName
    If Err.Number <> 0 Then
        Err.Clear
        DeleteEnvProcess = False
    Else
        DeleteEnvProcess = True
    End If
End Function

' Set user-level persistent variable by writing to HKCU\Environment
Function SetEnvUser(name, value)
    Dim sName, sValue, keyPath
    sName = SanitizeName(name)
    sValue = SanitizeValue(value)
    If sName = "" Then
        SetEnvUser = False
        Exit Function
    End If
    keyPath = "HKCU\Environment\" & sName
    On Error Resume Next
    wshShell.RegWrite keyPath, sValue, "REG_SZ"
    If Err.Number <> 0 Then
        Err.Clear
        SetEnvUser = False
    Else
        SetEnvUser = True
    End If
End Function

' Delete user-level variable from HKCU\Environment
Function DeleteEnvUser(name)
    Dim sName, keyPath
    sName = SanitizeName(name)
    If sName = "" Then
        DeleteEnvUser = False
        Exit Function
    End If
    keyPath = "HKCU\Environment\" & sName
    On Error Resume Next
    wshShell.RegDelete keyPath
    If Err.Number <> 0 Then
        ' If the key doesn't exist RegDelete raises; treat non-existence as success
        Err.Clear
        DeleteEnvUser = True
    Else
        DeleteEnvUser = True
    End If
End Function

' Set system-level variable using setx /M (requires administrative privileges)
' Falls back to attempting registry write to HKLM (not recommended here).
Function SetEnvSystem(name, value)
    Dim sName, sValue, cmd, exitCode
    sName = SanitizeName(name)
    sValue = SanitizeValue(value)
    If sName = "" Then
        SetEnvSystem = False
        Exit Function
    End If

    ' Use setx if available - /M sets machine environment (requires elevation)
    cmd = "cmd /c setx """ & sName & """ """ & sValue & """ /M"
    On Error Resume Next
    exitCode = wshShell.Run(cmd, 0, True) ' hide window, wait
    If Err.Number <> 0 Then
        Err.Clear
        SetEnvSystem = False
        Exit Function
    End If
    ' setx returns 0 on success (but this may vary); assume success
    SetEnvSystem = (exitCode = 0)
End Function

' ListEnv(scope) -> returns array of "NAME=VALUE" entries (0-based; empty array if none)
Function ListEnv(scope)
    Dim env, a(), i
    ReDim a(-1)
    Select Case scope
        Case SCOPE_PROCESS
            Set env = wshShell.Environment("Process")
        Case SCOPE_USER
            Set env = wshShell.Environment("User")
        Case SCOPE_SYSTEM
            Set env = wshShell.Environment("System")
        Case Else
            Set env = wshShell.Environment("Process")
    End Select

    On Error Resume Next
    Dim e
    i = -1
    For Each e In env
        i = i + 1
        ReDim Preserve a(i)
        a(i) = e
    Next
    ListEnv = a
End Function

' ---------- Helper: Pretty-print a list ----------
Function ArrayToText(arr)
    Dim i, out
    out = ""
    If IsEmpty(arr) Then
        ArrayToText = ""
        Exit Function
    End If
    If UBound(arr) < 0 Then
        ArrayToText = ""
        Exit Function
    End If
    For i = 0 To UBound(arr)
        out = out & arr(i) & vbCrLf
    Next
    ArrayToText = out
End Function

' ---------- Example usage (interactive demo) ----------
Dim scopeChoice, varName, varValue, curValue, ok, items

scopeChoice = InputBox("Choose scope: Process, User, or System (default: Process):", "Env Variables Demo", "Process")
If Trim(scopeChoice) = "" Then scopeChoice = SCOPE_PROCESS
scopeChoice = Trim(UCase(scopeChoice))
If scopeChoice = "PROCESS" Then scopeChoice = SCOPE_PROCESS
If scopeChoice = "USER" Then scopeChoice = SCOPE_USER
If scopeChoice = "SYSTEM" Then scopeChoice = SCOPE_SYSTEM
If scopeChoice <> SCOPE_PROCESS And scopeChoice <> SCOPE_USER And scopeChoice <> SCOPE_SYSTEM Then
    MsgBox "Invalid scope specified. Using Process.", vbExclamation, "Notice"
    scopeChoice = SCOPE_PROCESS
End If

' List current variables in chosen scope
items = ListEnv(scopeChoice)
MsgBox "Current environment variables in scope: " & scopeChoice & vbCrLf & vbCrLf & ArrayToText(items), vbInformation, "ListEnv"

' Ask for variable name to inspect / modify
varName = InputBox("Enter the environment variable name to inspect / modify:", "Env Variable Name")
varName = SanitizeName(varName)
If varName = "" Then
    MsgBox "No variable name provided. Exiting demo.", vbExclamation, "Aborted"
    WScript.Quit
End If

curValue = GetEnv(scopeChoice, varName)
If curValue = "" Then
    MsgBox "Variable '" & varName & "' not set in " & scopeChoice & " scope.", vbInformation, "Inspect"
Else
    MsgBox "Current value of '" & varName & "' in " & scopeChoice & ": " & vbCrLf & curValue, vbInformation, "Inspect"
End If

' Ask user what to do: set / delete / none
Dim action
action = InputBox("Action? (set / delete / none) [set]:", "Action", "set")
action = LCase(Trim(action))
If action = "" Then action = "set"

If action = "set" Then
    varValue = InputBox("Enter new value for '" & varName & "':", "Set Value", curValue)
    varValue = SanitizeValue(varValue)
    If scopeChoice = SCOPE_PROCESS Then
        ok = SetEnvProcess(varName, varValue)
        If ok Then
            MsgBox "Process-level variable set: " & varName & "=" & varValue, vbInformation, "Success"
        Else
            MsgBox "Failed to set process-level variable.", vbCritical, "Failure"
        End If
    ElseIf scopeChoice = SCOPE_USER Then
        ok = SetEnvUser(varName, varValue)
        If ok Then
            MsgBox "User-level variable written to HKCU\Environment. Note: other processes may need logoff/login or a WM_SETTINGCHANGE broadcast to pick it up.", vbInformation, "Success"
        Else
            MsgBox "Failed to write user-level variable (HKCU).", vbCritical, "Failure"
        End If
    ElseIf scopeChoice = SCOPE_SYSTEM Then
        ok = SetEnvSystem(varName, varValue)
        If ok Then
            MsgBox "Attempted to set system-level variable using setx /M." & vbCrLf & _
                   "If elevated, this succeeded; otherwise it likely failed. A reboot or WM_SETTINGCHANGE may be required for other processes to see it.", vbInformation, "System Set Attempted"
        Else
            MsgBox "Failed to set system-level variable (requires administrative privileges).", vbCritical, "Failure"
        End If
    End If
ElseIf action = "delete" Then
    If scopeChoice = SCOPE_PROCESS Then
        ok = DeleteEnvProcess(varName)
        If ok Then
            MsgBox "Process-level variable removed (current process).", vbInformation, "Deleted"
        Else
            MsgBox "Failed to remove process-level variable.", vbCritical, "Failure"
        End If
    ElseIf scopeChoice = SCOPE_USER Then
        ok = DeleteEnvUser(varName)
        If ok Then
            MsgBox "User-level variable removed from HKCU\Environment (if it existed).", vbInformation, "Deleted"
        Else
            MsgBox "Failed to delete user-level variable.", vbCritical, "Failure"
        End If
    ElseIf scopeChoice = SCOPE_SYSTEM Then
        ' Attempt to remove using setx (set to empty not supported); advise manual removal or admin registry edit
        MsgBox "Removing system environment variables is not implemented in this demo. Use an elevated registry editor or appropriate administration tools. Be careful.", vbExclamation, "Not Implemented"
    End If
Else
    MsgBox "No action taken.", vbInformation, "Done"
End If

' Final listing (show that changes may not be visible to existing processes)
items = ListEnv(scopeChoice)
MsgBox "Final listing for scope: " & scopeChoice & vbCrLf & vbCrLf & ArrayToText(items), vbInformation, "Final List"

' Clean up
Set wshShell = Nothing
On Error GoTo 0

' ---------------------------------------
' Final notes:
' 1. Process scope: temporary and safe for child processes you spawn.
' 2. User scope: persistent via HKCU and safe for single-user changes.
' 3. System scope: changes affect all users and require admin — use with extreme caution.
' 4. Changes to persistent environment variables may require a logoff/login, a reboot,
'    or broadcasting WM_SETTINGCHANGE for other processes to detect them.
' 5. Always sanitize environment names/values and avoid overwriting important variables (PATH, COMSPEC, etc.)
' ---------------------------------------
