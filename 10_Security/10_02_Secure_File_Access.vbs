' =======================================
'   VBScript Example: Secure File Access
' =======================================

' ---------------------------------------
' This section demonstrates how to safely
' handle file read/write operations in VBScript
' with error handling and input sanitization
' ---------------------------------------

On Error Resume Next   ' Enable error handling

' Function: Sanitize file path to avoid unsafe characters
Function SanitizePath(filePath)
    Dim cleanPath
    cleanPath = Trim(filePath)

    ' Remove characters not allowed in file names/paths
    cleanPath = Replace(cleanPath, ":", "")
    cleanPath = Replace(cleanPath, "*", "")
    cleanPath = Replace(cleanPath, "?", "")
    cleanPath = Replace(cleanPath, """", "")
    cleanPath = Replace(cleanPath, "<", "")
    cleanPath = Replace(cleanPath, ">", "")
    cleanPath = Replace(cleanPath, "|", "")

    SanitizePath = cleanPath
End Function

' Prompt user for filename
Dim rawFileName, safeFileName, fso, file
rawFileName = InputBox("Enter the file name to create (without extension):", _
                       "Secure File Access Example")

safeFileName = SanitizePath(rawFileName)

' If user cancels input, exit
If safeFileName = "" Then
    MsgBox "No file name provided. Exiting.", vbExclamation, "Aborted"
    WScript.Quit
End If

' Create FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Define safe full path (in current folder, with .txt extension)
Dim fullPath
fullPath = fso.GetAbsolutePathName(".") & "\" & safeFileName & ".txt"

' Try creating and writing to the file
Set file = fso.CreateTextFile(fullPath, True)

If Err.Number <> 0 Then
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "File Access Error"
    Err.Clear
Else
    ' Write safe content
    file.WriteLine "This is a secure file created with VBScript."
    file.WriteLine "Filename provided (sanitized): " & safeFileName
    file.Close

    MsgBox "File successfully created at:" & vbCrLf & fullPath, _
           vbInformation, "File Created"
End If

' Clean up
Set file = Nothing
Set fso = Nothing

' Disable error handling
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. Always sanitize file paths to avoid unsafe characters.
' 2. Use FileSystemObject for controlled access.
' 3. Always handle errors to prevent crashes.
' 4. Write only safe, validated data to files.
' ---------------------------------------
