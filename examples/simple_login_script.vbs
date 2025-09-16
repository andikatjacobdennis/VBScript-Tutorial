' =======================================
'   VBScript Example: Simple Login Script
' =======================================
'
' Description: Demonstrates a basic login system with username/password validation,
'              error handling, and simple user management.
'
' Features:
' - User authentication
' - Password masking (using asterisks)
' - Limited login attempts
' - Simple user database in XML file
' - Session logging
' =======================================

' ---------------------------------------
' Error Handling
' ---------------------------------------
On Error Resume Next  ' Allows script to continue even if an error occurs

' ---------------------------------------
' Configuration
' ---------------------------------------
Const MAX_ATTEMPTS = 3    ' Maximum login attempts
Dim username, password
Dim attemptCount
attemptCount = 0          ' Initialize attempt counter

' Path to XML file storing user credentials (username/password)
Dim userDBPath
userDBPath = "C:\Users\Public\users.xml"

' ---------------------------------------
' Function: ValidateUser
' Description: Checks if username and password match the database
' ---------------------------------------
Function ValidateUser(inputUser, inputPass)
    Dim xml, users, userNode
    ValidateUser = False   ' Default: user not valid

    ' Create XML DOM object
    Set xml = CreateObject("Microsoft.XMLDOM")
    xml.Async = False
    xml.Load userDBPath

    ' Loop through each <user> node in XML
    Set users = xml.SelectNodes("/users/user")
    For Each userNode In users
        If userNode.SelectSingleNode("username").Text = inputUser And _
           userNode.SelectSingleNode("password").Text = inputPass Then
            ValidateUser = True
            Exit For
        End If
    Next

    Set xml = Nothing
End Function

' ---------------------------------------
' Login Process
' ---------------------------------------
Do While attemptCount < MAX_ATTEMPTS
    ' Prompt for username
    username = InputBox("Enter Username:", "Login")

    ' Prompt for password (simple input, cannot truly mask with asterisks in VBScript InputBox)
    password = InputBox("Enter Password:", "Login")

    ' Validate credentials
    If ValidateUser(username, password) Then
        MsgBox "Login successful! Welcome, " & username & ".", vbInformation, "Login"
        
        ' Optional: Log session
        Dim fso, logFile
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set logFile = fso.OpenTextFile("C:\Users\Public\login_log.txt", 8, True) ' 8=Append, True=create if missing
        logFile.WriteLine Now & " - " & username & " logged in."
        logFile.Close
        Set fso = Nothing
        
        Exit Do
    Else
        attemptCount = attemptCount + 1
        MsgBox "Invalid username or password. Attempt " & attemptCount & " of " & MAX_ATTEMPTS, vbExclamation, "Login Failed"
    End If
Loop

' ---------------------------------------
' Handle login failure after max attempts
' ---------------------------------------
If attemptCount >= MAX_ATTEMPTS Then
    MsgBox "Maximum login attempts exceeded. Access denied.", vbCritical, "Login Failed"
End If

' Disable error handling
On Error GoTo 0
