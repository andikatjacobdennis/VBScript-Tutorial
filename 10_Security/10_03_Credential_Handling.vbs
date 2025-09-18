' =======================================
'   VBScript Example: Credential Handling
' =======================================

' ---------------------------------------
' This section demonstrates how to securely
' handle user credentials in VBScript.
'
' Note:
' - VBScript cannot provide modern encryption.
' - Instead, we use hashing (via Scripting library)
'   or basic obfuscation for demonstration.
' - Never store raw passwords in plain text.
' ---------------------------------------

On Error Resume Next   ' Enable error handling

' Function: Simple hashing using Microsoft Scriptlet library
Function GetHash(inputText)
    Dim objHash, byteArray, i, hashVal
    Set objHash = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

    ' Convert string to byte array
    ReDim byteArray(LenB(inputText) - 1)
    For i = 1 To LenB(inputText)
        byteArray(i - 1) = AscB(MidB(inputText, i, 1))
    Next

    ' Compute hash
    hashVal = ""
    objHash.ComputeHash_2(byteArray)
    For i = 1 To UBound(objHash.Hash) 
        hashVal = hashVal & LCase(Right("0" & Hex(objHash.Hash(i)), 2))
    Next

    GetHash = hashVal
    Set objHash = Nothing
End Function

' Prompt user for username and password
Dim username, password, passwordHash
username = InputBox("Enter your username:", "Credential Handling Example")
password = InputBox("Enter your password:", "Credential Handling Example")

' Sanitize username (avoid unsafe characters)
username = Replace(username, " ", "_")
username = Replace(username, ";", "")
username = Replace(username, "'", "")
username = Replace(username, """", "")

' Hash the password (never store plain text)
passwordHash = GetHash(password)

If Err.Number <> 0 Then
    MsgBox "Error while hashing password." & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical, "Hashing Error"
    Err.Clear
Else
    ' Simulate saving credentials (username + hashed password)
    MsgBox "Username: " & username & vbCrLf & _
           "Password Hash: " & passwordHash, _
           vbInformation, "Credentials Processed"
End If

' Disable error handling
On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. Never store plaintext passwords.
' 2. Use hashing to store only password digests.
' 3. Always sanitize usernames before use.
' 4. VBScript has limited cryptography — for
'    real security, use modern languages.
' ---------------------------------------
