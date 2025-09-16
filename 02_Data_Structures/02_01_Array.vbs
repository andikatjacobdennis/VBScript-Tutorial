' =======================================
'   VBScript Example: Array
' =======================================

option Explicit

Dim fruits, i, fruitList

' Declare and initialize an array
fruits = Array("Apple", "Banana", "Cherry", "Date")
fruitList = "Fruits in the array:" & vbCrLf

' Loop through the array and build a string
For i = LBound(fruits) To UBound(fruits)
    fruitList = fruitList & "- " & fruits(i) & vbCrLf
Next

' Display the list of fruits in a message box
MsgBox fruitList, vbInformation, "Array Example"
' =======================================