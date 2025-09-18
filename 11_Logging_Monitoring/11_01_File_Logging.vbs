' =======================================
'   VBScript Example: File Logging
' =======================================
'
' Demonstrates:
' - Appending sanitized log entries to a file
' - Timestamps for each entry
' - Basic log rotation based on file size (keeps N backups)
' - Error handling and safe path construction
'
' Notes:
' - This is a simple file-logging example for VBScript.
' - For heavy-duty logging use an external logger or service.
' =======================================

Option Explicit
On Error Resume Next

' ---------- Configuration ----------
Const LOG_DIR = "."                    ' directory to store logs (use absolute path in production)
Const LOG_NAME = "application.log"     ' base log filename
Const MAX_LOG_BYTES = 1024 * 50        ' rotate when file > 50 KB (adjust as needed)
Const MAX_BACKUPS = 3                  ' number of rotated backups to keep
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

' ---------- Utilities ----------

' Return current timestamp in ISO-ish format: yyyy-mm-dd hh:nn:ss
Function NowStamp()
    Dim d, yyyy, mm, dd, hh, nn, ss
    d = Now
    yyyy = Year(d)
    mm = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d), 2)
    hh = Right("0" & Hour(d), 2)
    nn = Right("0" & Minute(d), 2)
    ss = Right("0" & Second(d), 2)
    NowStamp = yyyy & "-" & mm & "-" & dd & " " & hh & ":" & nn & ":" & ss
End Function

' Basic sanitization for log messages (remove CR/LF to keep single-line entries and strip dangerous chars)
Function SanitizeLogMessage(msg)
    Dim s
    s = msg
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    ' Remove nulls and control chars < 32 (except tab)
    Dim i, ch, out
    out = ""
    For i = 1 To Len(s)
        ch = Asc(Mid(s, i, 1))
        If ch >= 32 Or ch = 9 Then
            out = out & Chr(ch)
        Else
            out = out & " "
        End If
    Next
    ' Trim and collapse excessive spaces
    out = Trim(out)
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    ' Avoid newlines that might be injected
    out = Replace(out, vbCrLf, " ")
    out = Replace(out, vbCr, " ")
    out = Replace(out, vbLf, " ")
    SanitizeLogMessage = out
End Function

' Sanitize file name to remove path traversal and invalid chars
Function SanitizeFileName(name)
    Dim s
    s = Trim(name)
    s = Replace(s, "..", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, ":", "")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, """", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    SanitizeFileName = s
End Function

' ---------- Log rotation ----------
' Rotate log files: logfile -> logfile.1, logfile.1 -> logfile.2, ... up to MAX_BACKUPS
Sub RotateLogs(fullPath, baseName)
    Dim fso, i, src, dst, backupPath
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Move older backups up (highest first)
    For i = MAX_BACKUPS - 1 To 1 Step -1
        src = fullPath & "." & i
        dst = fullPath & "." & (i + 1)
        If fso.FileExists(src) Then
            ' If destination exists, delete it first
            If fso.FileExists(dst) Then
                On Error Resume Next
                fso.DeleteFile dst, True
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
            On Error Resume Next
            fso.MoveFile src, dst
            If Err.Number <> 0 Then
                ' Non-fatal: continue
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next

    ' Finally rename current log to .1
    If fso.FileExists(fullPath) Then
        backupPath = fullPath & ".1"
        If fso.FileExists(backupPath) Then
            On Error Resume Next
            fso.DeleteFile backupPath, True
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
        On Error Resume Next
        fso.MoveFile fullPath, backupPath
        If Err.Number <> 0 Then
            ' Could not rotate — non-fatal, clear error
            Err.Clear
        End If
        On Error GoTo 0
    End If

    Set fso = Nothing
End Sub

' ---------- Core: Write a log entry ----------
Sub WriteLog(level, message)
    Dim fso, safeLogDir, safeLogName, fullPath, fileObj, fileSize

    ' Sanitize inputs
    safeLogDir = Trim(LOG_DIR)
    If safeLogDir = "" Then safeLogDir = "."
    safeLogName = SanitizeFileName(LOG_NAME)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure directory exists (attempt to create if missing)
    If Not fso.FolderExists(safeLogDir) Then
        On Error Resume Next
        fso.CreateFolder safeLogDir
        If Err.Number <> 0 Then
            ' could not create folder; fallback to current dir
            Err.Clear
            safeLogDir = "."
        End If
        On Error GoTo 0
    End If

    fullPath = fso.GetAbsolutePathName(safeLogDir) & "\" & safeLogName

    ' Check file size and rotate if necessary
    If fso.FileExists(fullPath) Then
        On Error Resume Next
        fileSize = fso.GetFile(fullPath).Size
        If Err.Number <> 0 Then
            fileSize = 0
            Err.Clear
        End If
        On Error GoTo 0
        If fileSize >= MAX_LOG_BYTES Then
            RotateLogs fullPath, safeLogName
        End If
    End If

    ' Open for append and write sanitized single-line entry
    On Error Resume Next
    Set fileObj = fso.OpenTextFile(fullPath, ForAppending, True)
    If Err.Number <> 0 Then
        ' Failed to open file - inform user (non-fatal)
        MsgBox "Unable to open log file for appending." & vbCrLf & _
               "Err#: " & Err.Number & " Desc: " & Err.Description, vbExclamation, "Logging Error"
        Err.Clear
        On Error GoTo 0
        Set fileObj = Nothing
        Set fso = Nothing
        Exit Sub
    End If
    On Error GoTo 0

    Dim entry, safeMessage
    safeMessage = SanitizeLogMessage(message)
    entry = NowStamp() & " [" & UCase(level) & "] " & safeMessage

    On Error Resume Next
    fileObj.WriteLine entry
    If Err.Number <> 0 Then
        ' write failed; clear and notify (non-fatal)
        MsgBox "Failed to write to log file." & vbCrLf & _
               "Err#: " & Err.Number & " Desc: " & Err.Description, vbExclamation, "Logging Error"
        Err.Clear
    End If
    On Error GoTo 0

    fileObj.Close
    Set fileObj = Nothing
    Set fso = Nothing
End Sub

' ---------- Example usage ----------
' Demonstrate logging different levels and a simulated error

WriteLog "info", "Application started."
WriteLog "debug", "User clicked 'Process' button."
WriteLog "warning", "Configuration value missing; using default."

' Simulated operation with error handling
On Error Resume Next
Dim x, y, z
x = 10
y = 0
z = x / y  ' will raise error
If Err.Number <> 0 Then
    WriteLog "error", "Runtime error during division: Err# " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WriteLog "info", "Division result: " & z
End If
On Error GoTo 0

WriteLog "info", "Application finished."

' ---------------------------------------
' Notes / Guidance:
' 1. Keep log messages concise and single-line for easier parsing.
' 2. Sanitize any user-supplied or external data before logging it.
' 3. Rotate logs to prevent unbounded growth; this example rotates by size.
' 4. For concurrent writes in multi-process scenarios consider a proper service or use locking.
' 5. Use absolute paths and secure directories in production.
' ---------------------------------------
