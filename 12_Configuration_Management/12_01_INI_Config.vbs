' =======================================
'   VBScript Example: INI Configuration
' =======================================
'
' Filename: 12_01_INI_Config.vbs
'
' Demonstrates:
' - Reading and writing simple INI-style config files
' - Handling sections ([section]) and key=value pairs
' - Safe writes using temp file + atomic rename
' - Basic sanitization of filenames/keys/values
' - Error handling and simple backup before overwriting
'
' Notes:
' - This is a lightweight INI handler intended for simple use.
' - It does not implement every INI corner-case (escapes, multiline values).
' - For complex needs consider a dedicated parser/library.
' =======================================

Option Explicit
On Error Resume Next

' ---------- Configuration ----------
Const BACKUP_EXT = ".bak"
Const TEMP_EXT = ".tmp"
Const COMMENT_CHARS = ";#"    ' characters that start comments

' ---------- Utilities ----------

' Simple sanitization for file names (prevent path traversal and invalid chars)
Function SanitizeFileName(name)
    Dim s
    s = Trim(name)
    If s = "" Then
        SanitizeFileName = ""
        Exit Function
    End If
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

' Normalize section name (trim, remove brackets if present)
Function NormalizeSection(sec)
    Dim s
    s = Trim(sec)
    If Left(s,1) = "[" And Right(s,1) = "]" Then
        s = Mid(s, 2, Len(s) - 2)
    End If
    NormalizeSection = Trim(s)
End Function

' Normalize key name (trim)
Function NormalizeKey(k)
    NormalizeKey = Trim(k)
End Function

' Normalize value (trim trailing/leading spaces but keep internal formatting)
Function NormalizeValue(v)
    NormalizeValue = v
End Function

' Return True if a line is a comment or blank
Function IsCommentOrBlank(line)
    Dim t, i, ch
    t = Trim(line)
    If t = "" Then
        IsCommentOrBlank = True
        Exit Function
    End If
    ' check if starts with comment char
    ch = Left(t,1)
    If InStr(COMMENT_CHARS, ch) > 0 Then
        IsCommentOrBlank = True
    Else
        IsCommentOrBlank = False
    End If
End Function

' ---------- Core INI operations ----------
' ReadValue(filePath, section, key, defaultValue)
' WriteValue(filePath, section, key, value)
' DeleteKey(filePath, section, key)
' DeleteSection(filePath, section)
' ListSections(filePath) -> returns array of section names (or empty array)
' ListKeys(filePath, section) -> returns array of keys in section

' Internal: load file lines into a VB array
Function LoadFileLines(path)
    Dim fso, file, lines(), idx
    ReDim lines(-1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        LoadFileLines = lines
        Set fso = Nothing
        Exit Function
    End If
    On Error Resume Next
    Set file = fso.OpenTextFile(path, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        ReDim lines(-1)
        LoadFileLines = lines
        Set fso = Nothing
        Exit Function
    End If
    idx = -1
    Do While Not file.AtEndOfStream
        idx = idx + 1
        ReDim Preserve lines(idx)
        lines(idx) = file.ReadLine
    Loop
    file.Close
    Set file = Nothing
    Set fso = Nothing
    LoadFileLines = lines
End Function

' Internal: write array of lines to file safely (temp -> backup -> rename)
Function SafeWriteFile(path, lines)
    Dim fso, tempPath, bakPath, file, i, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = path & TEMP_EXT

    ' write to temp
    On Error Resume Next
    Set file = fso.CreateTextFile(tempPath, True)
    If Err.Number <> 0 Then
        SafeWriteFile = False
        Err.Clear
        Set fso = Nothing
        Exit Function
    End If
    For i = 0 To UBound(lines)
        file.WriteLine lines(i)
    Next
    file.Close
    Set file = Nothing

    ' make backup if target exists
    If fso.FileExists(path) Then
        bakPath = path & BACKUP_EXT
        ' rotate previous backup (overwrite)
        On Error Resume Next
        If fso.FileExists(bakPath) Then
            fso.DeleteFile bakPath, True
            If Err.Number <> 0 Then Err.Clear
        End If
        On Error GoTo 0
        On Error Resume Next
        fso.MoveFile path, bakPath
        If Err.Number <> 0 Then
            ' If move fails, try delete and continue
            Err.Clear
            On Error Resume Next
            fso.DeleteFile path, True
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    End If

    ' move temp into place
    On Error Resume Next
    fso.MoveFile tempPath, path
    If Err.Number <> 0 Then
        ' if move fails, try copy+delete
        Err.Clear
        On Error Resume Next
        fso.CopyFile tempPath, path, True
        If Err.Number = 0 Then
            fso.DeleteFile tempPath, True
        End If
        If Err.Number <> 0 Then
            SafeWriteFile = False
            Err.Clear
            Set fso = Nothing
            Exit Function
        End If
    End If
    SafeWriteFile = True
    Set fso = Nothing
End Function

' Read a value from INI. Returns defaultValue if not found.
Function ReadValue(path, section, key, defaultValue)
    Dim lines(), i, curSection, kName, vPart, sepPos
    ReadValue = defaultValue
    section = NormalizeSection(section)
    key = NormalizeKey(key)

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then Exit Function

    curSection = ""
    For i = 0 To UBound(lines)
        If IsCommentOrBlank(lines(i)) Then
            ' skip
        Else
            Dim trimmed
            trimmed = Trim(lines(i))
            If Left(trimmed,1) = "[" And Right(trimmed,1) = "]" Then
                curSection = NormalizeSection(trimmed)
            ElseIf curSection = section Then
                sepPos = InStr(trimmed, "=")
                If sepPos > 0 Then
                    kName = Trim(Left(trimmed, sepPos - 1))
                    If LCase(kName) = LCase(key) Then
                        vPart = Mid(trimmed, sepPos + 1)
                        ReadValue = NormalizeValue(Trim(vPart))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

' Write or update a key in a section. Returns True on success.
Function WriteValue(path, section, key, value)
    Dim lines(), i, curSection, foundSection, foundKey, out(), idx
    section = NormalizeSection(section)
    key = NormalizeKey(key)
    value = NormalizeValue(value)

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then
        ' file doesn't exist or empty -> create minimal structure
        ReDim lines(0)
        lines(0) = "[" & section & "]"
        ReDim Preserve lines(1)
        lines(1) = key & "=" & value
        SafeWriteFile path, lines
        WriteValue = True
        Exit Function
    End If

    curSection = ""
    foundSection = False
    foundKey = False
    idx = -1
    ReDim out(-1)

    For i = 0 To UBound(lines)
        idx = idx + 1
        ReDim Preserve out(idx)
        Dim ln
        ln = lines(i)
        If IsCommentOrBlank(ln) Then
            out(idx) = ln
        Else
            Dim t
            t = Trim(ln)
            If Left(t,1) = "[" And Right(t,1) = "]" Then
                ' when we move into a new section, if we passed the desired section and didn't find the key,
                ' append the key before emitting the new section header
                If foundSection And Not foundKey Then
                    idx = idx + 1
                    ReDim Preserve out(idx)
                    out(idx) = key & "=" & value
                    foundKey = True
                End If
                curSection = NormalizeSection(t)
                out(idx) = ln
                If LCase(curSection) = LCase(section) Then
                    foundSection = True
                End If
            ElseIf foundSection Then
                Dim pos
                pos = InStr(ln, "=")
                If pos > 0 Then
                    Dim existingKey
                    existingKey = Trim(Left(ln, pos - 1))
                    If LCase(existingKey) = LCase(key) Then
                        ' replace line with new value
                        out(idx) = key & "=" & value
                        foundKey = True
                    Else
                        out(idx) = ln
                    End If
                Else
                    out(idx) = ln
                End If
            Else
                out(idx) = ln
            End If
        End If
    Next

    ' If we finished and foundSection is false, append section + key
    If Not foundSection Then
        idx = idx + 1
        ReDim Preserve out(idx)
        out(idx) = ""                 ' blank line for readability
        idx = idx + 1
        ReDim Preserve out(idx)
        out(idx) = "[" & section & "]"
        idx = idx + 1
        ReDim Preserve out(idx)
        out(idx) = key & "=" & value
        foundKey = True
    ElseIf foundSection And Not foundKey Then
        ' we found section but not key and didn't insert earlier (e.g., section at end)
        idx = idx + 1
        ReDim Preserve out(idx)
        out(idx) = key & "=" & value
        foundKey = True
    End If

    WriteValue = SafeWriteFile(path, out)
End Function

' Delete a key from a section. Returns True if changed (or True if key didn't exist but operation OK)
Function DeleteKey(path, section, key)
    Dim lines(), i, out(), changed, curSection, pos
    section = NormalizeSection(section)
    key = NormalizeKey(key)

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then
        DeleteKey = True
        Exit Function
    End If

    curSection = ""
    changed = False
    ReDim out(-1)
    For i = 0 To UBound(lines)
        Dim ln
        ln = lines(i)
        If IsCommentOrBlank(ln) Then
            ReDim Preserve out(Ubound(out) + 1)
            out(UBound(out)) = ln
        Else
            Dim t
            t = Trim(ln)
            If Left(t,1) = "[" And Right(t,1) = "]" Then
                curSection = NormalizeSection(t)
                ReDim Preserve out(Ubound(out) + 1)
                out(UBound(out)) = ln
            ElseIf LCase(curSection) = LCase(section) Then
                pos = InStr(ln, "=")
                If pos > 0 Then
                    Dim existingKey
                    existingKey = Trim(Left(ln, pos - 1))
                    If LCase(existingKey) = LCase(key) Then
                        ' skip this line (delete)
                        changed = True
                    Else
                        ReDim Preserve out(Ubound(out) + 1)
                        out(UBound(out)) = ln
                    End If
                Else
                    ReDim Preserve out(Ubound(out) + 1)
                    out(UBound(out)) = ln
                End If
            Else
                ReDim Preserve out(Ubound(out) + 1)
                out(UBound(out)) = ln
            End If
        End If
    Next

    If changed Then
        DeleteKey = SafeWriteFile(path, out)
    Else
        ' nothing changed, treat as success
        DeleteKey = True
    End If
End Function

' Delete entire section. Returns True on success (even if section didn't exist)
Function DeleteSection(path, section)
    Dim lines(), i, out(), changed, curSection
    section = NormalizeSection(section)

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then
        DeleteSection = True
        Exit Function
    End If

    curSection = ""
    changed = False
    ReDim out(-1)
    For i = 0 To UBound(lines)
        Dim ln
        ln = lines(i)
        If IsCommentOrBlank(ln) Then
            ReDim Preserve out(Ubound(out) + 1)
            out(UBound(out)) = ln
        Else
            Dim t
            t = Trim(ln)
            If Left(t,1) = "[" And Right(t,1) = "]" Then
                curSection = NormalizeSection(t)
                If LCase(curSection) = LCase(section) Then
                    ' skip this section: consume lines until next section header
                    changed = True
                    ' advance i until next section header or end
                    Do While i < UBound(lines)
                        i = i + 1
                        If Left(Trim(lines(i)),1) = "[" And Right(Trim(lines(i)),1) = "]" Then
                            ' back up one so the outer loop will process this header
                            i = i - 1
                            Exit Do
                        End If
                    Loop
                Else
                    ReDim Preserve out(Ubound(out) + 1)
                    out(UBound(out)) = ln
                End If
            Else
                ReDim Preserve out(Ubound(out) + 1)
                out(UBound(out)) = ln
            End If
        End If
    Next

    If changed Then
        DeleteSection = SafeWriteFile(path, out)
    Else
        DeleteSection = True
    End If
End Function

' ListSections: returns a 0-based array of section names (may be empty)
Function ListSections(path)
    Dim lines(), i, sections(), cur, seen
    ReDim sections(-1)
    Set seen = CreateObject("Scripting.Dictionary")

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then
        ListSections = sections
        Exit Function
    End If

    For i = 0 To UBound(lines)
        If Not IsCommentOrBlank(lines(i)) Then
            Dim t
            t = Trim(lines(i))
            If Left(t,1) = "[" And Right(t,1) = "]" Then
                cur = NormalizeSection(t)
                If Not seen.Exists(LCase(cur)) Then
                    seen.Add LCase(cur), cur
                    ReDim Preserve sections(UBound(sections) + 1)
                    sections(UBound(sections)) = cur
                End If
            End If
        End If
    Next

    Set seen = Nothing
    ListSections = sections
End Function

' ListKeys: returns array of key names in section (may be empty)
Function ListKeys(path, section)
    Dim lines(), i, keys(), curSection, pos, kName
    ReDim keys(-1)
    section = NormalizeSection(section)

    lines = LoadFileLines(path)
    If UBound(lines) < 0 Then
        ListKeys = keys
        Exit Function
    End If

    curSection = ""
    For i = 0 To UBound(lines)
        If Not IsCommentOrBlank(lines(i)) Then
            Dim t
            t = Trim(lines(i))
            If Left(t,1) = "[" And Right(t,1) = "]" Then
                curSection = NormalizeSection(t)
            ElseIf LCase(curSection) = LCase(section) Then
                pos = InStr(t, "=")
                If pos > 0 Then
                    kName = Trim(Left(t, pos - 1))
                    ReDim Preserve keys(UBound(keys) + 1)
                    keys(UBound(keys)) = kName
                End If
            End If
        End If
    Next

    ListKeys = keys
End Function

' ---------- Example usage ----------
' NOTE: adjust the path or use a relative path in the same folder as script
Dim cfgPath, val

cfgPath = SanitizeFileName("example_config.ini")
If cfgPath = "" Then
    MsgBox "Invalid config filename. Exiting.", vbExclamation, "Error"
    WScript.Quit
End If

' If file doesn't exist, create a sample one
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(cfgPath) Then
    Dim sampleLines()
    ReDim sampleLines(3)
    sampleLines(0) = "; sample INI file created by 12_01_INI_Config.vbs"
    sampleLines(1) = "[General]"
    sampleLines(2) = "AppName=VBScriptDemo"
    sampleLines(3) = "Version=1.0"
    SafeWriteFile cfgPath, sampleLines
End If

' Read an existing value (with default)
val = ReadValue(cfgPath, "General", "AppName", "UnknownApp")
MsgBox "AppName from INI: " & val, vbInformation, "ReadValue"

' Write/update a value
If WriteValue(cfgPath, "General", "LastRun", Now) Then
    MsgBox "Wrote LastRun to INI.", vbInformation, "WriteValue"
Else
    MsgBox "Failed to write to INI.", vbCritical, "WriteValue Error"
End If

' List sections
Dim secs(), i
secs = ListSections(cfgPath)
If UBound(secs) >= 0 Then
    Dim sList
    sList = ""
    For i = 0 To UBound(secs)
        sList = sList & secs(i) & vbCrLf
    Next
    MsgBox "Sections found:" & vbCrLf & sList, vbInformation, "ListSections"
Else
    MsgBox "No sections found.", vbInformation, "ListSections"
End If

' List keys in General
Dim keys(), kList
keys = ListKeys(cfgPath, "General")
kList = ""
If UBound(keys) >= 0 Then
    For i = 0 To UBound(keys)
        kList = kList & keys(i) & vbCrLf
    Next
    MsgBox "Keys in [General]:" & vbCrLf & kList, vbInformation, "ListKeys"
Else
    MsgBox "No keys in [General].", vbInformation, "ListKeys"
End If

' Delete a key (example)
If DeleteKey(cfgPath, "General", "Version") Then
    MsgBox "Deleted key 'Version' (if it existed).", vbInformation, "DeleteKey"
End If

' Delete a section (example) -- commented out by default
' If DeleteSection(cfgPath, "Obsolete") Then
'    MsgBox "Deleted section 'Obsolete' (if it existed).", vbInformation, "DeleteSection"
' End If

Set fso = Nothing
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. This implementation preserves comments and ordering where possible.
' 2. Writes are done to a temp file and then moved into place for safety.
' 3. A .bak copy of the previous file is created when overwriting.
' 4. Keep INI files small and simple; they are not transactional or robustly concurrent.
' 5. For secure or concurrent configuration storage use registry, database, or dedicated config stores.
' ---------------------------------------
