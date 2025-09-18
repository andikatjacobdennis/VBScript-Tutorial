' =======================================
'   VBScript Example: Find and Replace in Files
' =======================================
'
' Description: This script searches for text in files and replaces it.
'              It demonstrates file reading, writing, and string manipulation.
'
' Usage: Customize the variables below or modify to accept command-line arguments.
' =======================================

' Enable explicit variable declaration
Option Explicit

' Create FileSystemObject
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' =======================================
' CONFIGURATION - Customize these values
' =======================================
Dim folderPath, filePattern, findText, replaceText, includeSubfolders

' Folder to search in (use current directory if empty)
folderPath = "C:\ExampleFiles\"  ' Change this to your target folder
filePattern = "*.txt"           ' File pattern to search (e.g., "*.txt", "*.log", "*.*")
findText = "oldtext"             ' Text to find
replaceText = "newtext"          ' Text to replace with
includeSubfolders = True         ' Set to False to search only main folder

' =======================================
' MAIN SCRIPT
' =======================================

Dim message, response, fileCount, replaceCount

' Verify folder exists
If folderPath = "" Then folderPath = "."
If Not fso.FolderExists(folderPath) Then
    MsgBox "Error: Folder not found: " & folderPath, vbCritical, "Error"
    WScript.Quit 1
End If

' Confirm operation
message = "Find and Replace Operation:" & vbCrLf & vbCrLf
message = message & "Folder: " & folderPath & vbCrLf
message = message & "Files: " & filePattern & vbCrLf
message = message & "Find: """ & findText & """" & vbCrLf
message = message & "Replace with: """ & replaceText & """" & vbCrLf
message = message & "Include subfolders: " & includeSubfolders & vbCrLf & vbCrLf
message = message & "Do you want to continue?"

response = MsgBox(message, vbYesNo + vbQuestion, "Find and Replace Confirmation")

If response = vbNo Then
    MsgBox "Operation cancelled.", vbInformation, "Cancelled"
    WScript.Quit
End If

' Perform find and replace
PerformFindReplace folderPath, filePattern, findText, replaceText, includeSubfolders

' Display results
message = "Operation completed!" & vbCrLf & vbCrLf
message = message & "Files processed: " & fileCount & vbCrLf
message = message & "Replacements made: " & replaceCount
MsgBox message, vbInformation, "Results"

WScript.Quit

' =======================================
' FUNCTIONS AND SUBROUTINES
' =======================================

Sub PerformFindReplace(folderPath, filePattern, findText, replaceText, includeSubfolders)
    Dim folder, file, subfolder, content, newContent
    Set folder = fso.GetFolder(folderPath)
    
    ' Process files in current folder
    For Each file In folder.Files
        If FileMatchesPattern(file.Name, filePattern) Then
            ProcessFile file.Path, findText, replaceText
        End If
    Next
    
    ' Process subfolders if requested
    If includeSubfolders Then
        For Each subfolder In folder.SubFolders
            PerformFindReplace subfolder.Path, filePattern, findText, replaceText, True
        Next
    End If
End Sub

Function FileMatchesPattern(fileName, pattern)
    If pattern = "*.*" Then
        FileMatchesPattern = True
    Else
        FileMatchesPattern = (LCase(fso.GetExtensionName(fileName)) = LCase(Mid(pattern, 2)))
    End If
End Function

Sub ProcessFile(filePath, findText, replaceText)
    Dim file, content, newContent, tempFilePath
    On Error Resume Next
    
    ' Read file content
    Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    content = file.ReadAll
    file.Close
    
    ' Check if text exists in file
    If InStr(1, content, findText, vbTextCompare) > 0 Then
        ' Create backup (optional - uncomment next line to enable)
        ' fso.CopyFile filePath, filePath & ".bak", True
        
        ' Perform replacement
        newContent = Replace(content, findText, replaceText, 1, -1, vbTextCompare)
        
        ' Write new content to temporary file
        tempFilePath = filePath & ".tmp"
        Set file = fso.CreateTextFile(tempFilePath, True)
        file.Write newContent
        file.Close
        
        ' Replace original file with temporary file
        fso.DeleteFile filePath
        fso.MoveFile tempFilePath, filePath
        
        fileCount = fileCount + 1
        replaceCount = replaceCount + (Len(newContent) - Len(content)) / Len(replaceText)
    End If
    
    On Error GoTo 0
End Sub

' =======================================
' COMMAND-LINE VERSION (Alternative approach)
' Uncomment and modify if you want command-line usage
' =======================================
'
' If WScript.Arguments.Count >= 3 Then
'     folderPath = WScript.Arguments(0)
'     findText = WScript.Arguments(1)
'     replaceText = WScript.Arguments(2)
'     
'     If WScript.Arguments.Count >= 4 Then
'         filePattern = WScript.Arguments(3)
'     End If
'     
'     If WScript.Arguments.Count >= 5 Then
'         includeSubfolders = CBool(WScript.Arguments(4))
'     End If
' End If
