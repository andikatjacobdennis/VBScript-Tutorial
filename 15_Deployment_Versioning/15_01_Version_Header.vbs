' =======================================
'   VBScript Example: Version Header
' =======================================
'
' Filename: 15_01_Version_Header.vbs
'
' Demonstrates:
' - Embedding version information in scripts
' - Displaying version info dynamically
' - Using constants for version management
' - Outputting metadata (author, description)
'
' Notes:
' - Keeping a version header helps with deployment and maintenance
' - Version can follow Semantic Versioning (Major.Minor.Patch)
' =======================================

Option Explicit

' ---------------------------------------
' Version Header Metadata
' ---------------------------------------
Const SCRIPT_NAME    = "15_01_Version_Header.vbs"
Const SCRIPT_VERSION = "1.0.0"
Const SCRIPT_AUTHOR  = "Your Name"
Const SCRIPT_DESC    = "Demonstrates embedding version information in VBScript."

' ---------------------------------------
' Helper: Show version info
' ---------------------------------------
Sub ShowVersionInfo()
    Dim info
    info = "Script: " & SCRIPT_NAME & vbCrLf & _
           "Version: " & SCRIPT_VERSION & vbCrLf & _
           "Author: " & SCRIPT_AUTHOR & vbCrLf & _
           "Description: " & SCRIPT_DESC & vbCrLf & _
           "Last Run: " & Now
    MsgBox info, vbInformation, "Version Header"
End Sub

' ---------------------------------------
' Main Program
' ---------------------------------------
ShowVersionInfo()

' Example functionality (dummy)
MsgBox "Hello! This is an example script with versioning.", vbOKOnly, "Demo"

' ---------------------------------------
' Notes:
' 1. Update SCRIPT_VERSION with every meaningful change.
' 2. Semantic versioning: 
'    - MAJOR: incompatible changes
'    - MINOR: new features, backward-compatible
'    - PATCH: bug fixes, small updates
' 3. Include author and description for team collaboration.
' ---------------------------------------
```
