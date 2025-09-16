' =======================================
'   VBScript Example: Folder & File Operations
' =======================================

Dim fso, folderPath, newFolderPath, filePath, file, content
Set fso = CreateObject("Scripting.FileSystemObject")

' ---------------------------------------
' 1. Create a Folder
' ---------------------------------------
folderPath = "C:\ExampleFolder"
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder folderPath
    MsgBox "Folder created successfully: " & folderPath, vbInformation, "Folder Creation"
End If

' ---------------------------------------
' 2. Rename the Folder
' ---------------------------------------
newFolderPath = "C:\RenamedFolder"
If fso.FolderExists(folderPath) Then
    fso.MoveFolder folderPath, newFolderPath
    MsgBox "Folder renamed successfully to: " & newFolderPath, vbInformation, "Folder Rename"
End If

' ---------------------------------------
' 3. List Files in Folder (if any)
' ---------------------------------------
Dim fileItem, fileList
fileList = "Files in " & newFolderPath & ":" & vbCrLf
If fso.FolderExists(newFolderPath) Then
    For Each fileItem In fso.GetFolder(newFolderPath).Files
        fileList = fileList & fileItem.Name & vbCrLf
    Next
    MsgBox fileList, vbInformation, "Folder Contents"
End If

' ---------------------------------------
' 4. Create and Write to a File
' ---------------------------------------
filePath = newFolderPath & "\example.txt"
If Not fso.FileExists(filePath) Then
    Set file = fso.CreateTextFile(filePath, True)
    file.WriteLine "Hello, this is a VBScript file example!"
    file.Close
    MsgBox "File created successfully: " & filePath, vbInformation, "File Creation"
End If

' ---------------------------------------
' 5. Read from the File
' ---------------------------------------
If fso.FileExists(filePath) Then
    Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    content = file.ReadAll
    file.Close
    MsgBox "File contents:" & vbCrLf & content, vbInformation, "Read File"
End If

' ---------------------------------------
' 6. Append Text to the File
' ---------------------------------------
Set file = fso.OpenTextFile(filePath, 8, True) ' 8 = ForAppending
file.WriteLine "Appending a new line to the file."
file.Close
MsgBox "Text appended successfully.", vbInformation, "Append File"

' ---------------------------------------
' 7. Delete the File
' ---------------------------------------
If fso.FileExists(filePath) Then
    fso.DeleteFile(filePath)
    MsgBox "File deleted successfully: " & filePath, vbInformation, "Delete File"
End If

' ---------------------------------------
' 8. Delete the Folder
' ---------------------------------------
If fso.FolderExists(newFolderPath) Then
    fso.DeleteFolder newFolderPath
    MsgBox "Folder deleted successfully: " & newFolderPath, vbInformation, "Delete Folder"
End If

' =======================================
