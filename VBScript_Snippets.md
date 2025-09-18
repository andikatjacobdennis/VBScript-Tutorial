# Snippets

## 1. The Absolute Basics (The "Hello World" and Structure)

This is the foundation. You must be able to write this in your sleep.

```vbscript
' Always use Option Explicit to force variable declaration
Option Explicit

' Declare variables with Dim
Dim myMessage

' Assign a value
myMessage = "Hello, World!"

' Output to the user (GUI)
MsgBox myMessage

' Output to the console (for automation)
WScript.Echo myMessage
```

**What to remember:** `Option Explicit`, `Dim`, `MsgBox` (for popups), `WScript.Echo` (for command line).

---

## 2. User Interaction (Getting Input)

```vbscript
Dim userName
' InputBox displays a prompt and stores the result
userName = InputBox("Please enter your name:", "User Input")

If userName <> "" Then ' Check if the user didn't click Cancel
    MsgBox "Hello, " & userName & "!", vbInformation, "Greeting"
Else
    MsgBox "You did not enter a name.", vbExclamation, "Notice"
End If
```

**What to remember:** `InputBox("Prompt", "Title")`, always check if the result is empty (`""`).

---

## 3. Conditional Logic (Making Decisions)

```vbscript
Dim number
number = InputBox("Enter a number:")

' Basic If...Then...ElseIf...Else structure
If IsNumeric(number) Then
    If number > 10 Then
        MsgBox "The number is large."
    ElseIf number > 0 Then
        MsgBox "The number is small."
    Else
        MsgBox "The number is zero or negative."
    End If
Else
    MsgBox "That's not a valid number!"
End If
```

**What to remember:** `If...Then...ElseIf...Else...End If`, `IsNumeric()` to check if input is a number.

---

## 4. Loops (Doing Things Repeatedly)

a) For...Next Loop (When you know how many times to run)

```vbscript
Dim i
For i = 1 To 5
    WScript.Echo "Iteration number: " & i
Next
' Output: 1, 2, 3, 4, 5
```

b) For Each...Next Loop (Go through each item in a collection)

```vbscript
Dim fruit, fruitList
fruitList = Array("Apple", "Banana", "Orange") ' Create an array

For Each fruit In fruitList
    WScript.Echo fruit
Next
' Output: Apple, Banana, Orange
```

c) Do While...Loop (Run while a condition is true)

```vbscript
Dim count
count = 1
Do While count <= 3
    WScript.Echo "Count is: " & count
    count = count + 1 ' INCREMENT THE COUNTER! Forgetting this is a common mistake.
Loop
```

**What to remember:**

* `For i = 1 To 5 ... Next`
* `For Each item In collection ... Next`
* `Do While condition ... Loop` (Don't forget to change the condition inside the loop!)

---

### 5. Working with Files (CRITICAL for Admin Tasks)

This is a very common interview topic. Remember the object name: `FileSystemObject`.

```vbscript
' Create the main filesystem object
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if a file exists
Dim filePath
filePath = "C:\test.txt"
If fso.FileExists(filePath) Then
    MsgBox "The file exists."
Else
    MsgBox "The file does not exist."
End If

' Read all text from a file
Dim fileContent
If fso.FileExists(filePath) Then
    Dim file
    Set file = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    fileContent = file.ReadAll
    file.Close
    MsgBox fileContent
End If

' Write text to a file (OVERWRITES existing content!)
Dim newText
newText = "This is new content."
Set file = fso.OpenTextFile(filePath, 2, True) ' 2 = ForWriting, True = Create if needed
file.Write newText
file.Close
```

**What to remember:**

* `Set fso = CreateObject("Scripting.FileSystemObject")`
* `fso.FileExists(path)`
* `fso.OpenTextFile(path, Mode, Create)` Modes: `1`=Read, `2`=Write, `8`=Append.

---

### 6. Error Handling (Making Your Scripts Robust)

Showing you know error handling is a huge plus. It separates beginners from pros.

```vbscript
On Error Resume Next ' Tells VBScript to continue on error

' A risky operation, like accessing a non-existent drive
Dim riskyFile
Set riskyFile = fso.OpenTextFile("Z:\non_existent_file.txt", 1)

' Check if the previous operation caused an error
If Err.Number <> 0 Then
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Err.Clear ' Clear the error object
End If

On Error Goto 0 ' Turn off custom error handling
```

**What to remember:**

* `On Error Resume Next` (Starts error trapping)
* `If Err.Number <> 0 Then` (Check for an error)
* `Err.Clear` (Reset the error object)

---

### 7. Running System Commands

This is how you leverage other programs from your script.

```vbscript
Dim wshell
Set wshell = CreateObject("WScript.Shell")

' Run a command and wait for it to finish (0 = hidden window)
wshell.Run "ping 8.8.8.8", 0, True

' Or just run a program
wshell.Run "notepad.exe"
```

**What to remember:** `Set wshell = CreateObject("WScript.Shell")` and `wshell.Run "command"`.
