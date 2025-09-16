' =======================================
'   VBScript Example: MsgBox
' =======================================

' ---------------------------------------
'   Demonstrates MsgBox
' ---------------------------------------
' 
' ---------------------------------------
'   Button Constants
' ---------------------------------------
'   0  vbOKOnly           OK button
'   1  vbOKCancel         OK and Cancel buttons
'   2  vbAbortRetryIgnore Abort, Retry, and Ignore buttons
'   3  vbYesNoCancel      Yes, No, and Cancel buttons
'   4  vbYesNo            Yes and No buttons
'   5  vbRetryCancel      Retry and Cancel buttons
'
' ---------------------------------------
'   Icon Constants
' ---------------------------------------
'   16 vbCritical         Critical message icon
'   32 vbQuestion         Question mark icon
'   48 vbExclamation      Exclamation point icon
'   64 vbInformation      Information message icon
'
' ---------------------------------------
'   Return Values
' ---------------------------------------
'   1  vbOK
'   2  vbCancel
'   3  vbAbort
'   4  vbRetry
'   5  vbIgnore
'   6  vbYes
'   7  vbNo
'
' =======================================

Dim message, answer
message = "Do you want to continue?"

' Show MsgBox with OK + Cancel buttons and Exclamation icon
answer = MsgBox(message, vbOKCancel + vbExclamation, "Hello !!!")

If answer = vbOK Then
    MsgBox "OK"
ElseIf answer = vbCancel Then
    MsgBox "Cancel"
Else
    MsgBox "Unknown"
End If

' =======================================




