' =======================================
'   VBScript Example: COM Automation
' =======================================
'
' Filename: 13_02_COM_Automation.vbs
'
' Demonstrates:
' - Creating and using COM automation objects
' - Example 1: Automating Microsoft Excel
' - Example 2: Automating Internet Explorer (legacy example)
' - Handling errors when COM objects are not available
'
' Notes:
' - COM Automation allows VBScript to control applications that expose an automation interface.
' - Requires the target application to be installed (e.g., Excel, IE).
' - Internet Explorer is deprecated but shown here for historical/educational purposes.
' =======================================

Option Explicit
On Error Resume Next

' ---------- Example 1: Excel Automation ----------
Dim xlApp, xlBook, xlSheet

Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Or xlApp Is Nothing Then
    MsgBox "Excel is not installed or not available for automation.", vbCritical, "Excel COM Error"
    Err.Clear
Else
    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    xlSheet.Cells(1,1).Value = "Hello from VBScript!"
    xlSheet.Cells(2,1).Value = "The time is:"
    xlSheet.Cells(2,2).Value = Now

    xlSheet.Range("A1:B2").Columns.AutoFit

    MsgBox "Excel workbook created. Please check Excel window.", vbInformation, "Excel Automation"

    ' Cleanup
    Set xlSheet = Nothing
    Set xlBook = Nothing
    ' Leave Excel running so user can see it
    ' xlApp.Quit  ' Uncomment to close Excel automatically
    Set xlApp = Nothing
End If

' ---------- Example 2: Internet Explorer Automation (legacy) ----------
Dim ieApp
Set ieApp = CreateObject("InternetExplorer.Application")
If Err.Number <> 0 Or ieApp Is Nothing Then
    MsgBox "Internet Explorer is not available on this system.", vbExclamation, "IE COM Error"
    Err.Clear
Else
    ieApp.Visible = True
    ieApp.Navigate "https://www.example.com"

    ' Wait until page loads (simple loop)
    Do While ieApp.Busy Or ieApp.ReadyState <> 4
        WScript.Sleep 500
    Loop

    MsgBox "Internet Explorer navigated to: " & ieApp.LocationURL, vbInformation, "IE Automation"

    ' Cleanup
    ' ieApp.Quit   ' Uncomment to close IE automatically
    Set ieApp = Nothing
End If

On Error GoTo 0

' ---------------------------------------
' Notes:
' 1. You can automate many COM-enabled apps: Word, Outlook, PowerPoint, Access, etc.
' 2. Use CreateObject("ProgID") where ProgID is the programmatic identifier (e.g., "Word.Application").
' 3. Always check for errors — not all systems have the same applications installed.
' 4. Internet Explorer is deprecated; for modern browser automation, use dedicated tools like Selenium.
' ---------------------------------------
