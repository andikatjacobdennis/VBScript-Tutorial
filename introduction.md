# Introduction to VBScript

## What is VBScript?

VBScript (Visual Basic Scripting Edition) is a lightweight, interpreted scripting language developed by Microsoft. It is a subset of the Visual Basic for Applications (VBA) language and is designed for simplicity and ease of use. While its use in web pages (as a client-side script) has been completely deprecated, it remains a powerful tool for:

*   **Windows Administration:** Automating system administration tasks.
*   **Task Automation:** Writing logon scripts, automating software installations, file management, etc.
*   **Testing:** It was historically used in Quick Test Professional (QTP/UFT) for test automation.

VBScript is known for its simple syntax, making it an excellent language for beginners to learn fundamental programming concepts.

## Key Applications of VBScript

Despite being a legacy technology, VBScript is still useful in specific Windows-centric environments:

1.  **Windows System Administration:** The primary modern use case. System administrators use it to automate repetitive tasks like user management, event log monitoring, and configuring system settings.
2.  **Login Scripts:** In Active Directory environments, VBScript can be used to map network drives, install printers, or launch applications when a user logs in.
3.  **File and Folder Management:** Automating complex file operations (copying, moving, renaming, deleting) based on specific conditions.
4.  **Interacting with Windows Objects:** It can automate other applications like Microsoft Office (Excel, Word, Outlook) or interact with the Windows Registry and Windows Management Instrumentation (WMI).
5.  **Simple GUI Applications:** Creating basic input/output dialog boxes for user interaction, as demonstrated in many tutorials in this repo.

## The Windows Script Host (WSH): `wscript` vs `cscript`

VBScript files (`.vbs`) are executed by the **Windows Script Host (WSH)**, a built-in administration tool in Windows. WSH provides two hosts for running scripts:

| Feature | **WScript** (Windows-based) | **CScript** (Command-based) |
| :--- | :--- | :--- |
| **Executable** | `wscript.exe` | `cscript.exe` |
| **Interface** | Graphical (GUI). Displays output in message boxes (MsgBox). | Console (CLI). Displays output in the command prompt window. |
| **Best For** | Scripts that require user interaction via pop-up windows. | Scripts that run automated, unattended tasks and log output to the console. |
| **Default Host** | Yes. Double-clicking a `.vbs` file uses `wscript`. | No. Must be explicitly called from the command line. |

### How to Use Them

*   **By default,** when you double-click a `.vbs` file, it runs with `wscript.exe`.
*   To run a script from the command prompt with `cscript`, use:
    ```batch
    cscript //nologo C:\Path\To\Your\Script.vbs
    ```
    The `//nologo` flag suppresses the Microsoft banner text, giving you cleaner output.

*   You can change the default host. To set `cscript` as the default, run Command Prompt as Administrator and execute:
    ```batch
    cscript //H:CScript //S
    ```
    To switch back to `wscript`, use:
    ```batch
    cscript //H:WScript //S
    ```

## Helpful Tips for Getting Started

1.  **Use a Good Editor:** While Notepad works, an editor with syntax highlighting like **VS Code**, **Notepad++**, or **VBSEdit** will make your life much easier by color-coding keywords and helping spot errors.

2.  **Start Small:** Begin with the basics. Understand variables (`Dim`), message boxes (`MsgBox`), and input boxes (`InputBox`) before moving on to loops and file operations.

3.  **Error Handling is Crucial:** Always include basic error handling (`On Error Resume Next` and checking `Err.Number`) in your scripts, especially when dealing with files or the registry. This prevents the script from crashing and provides helpful feedback. See `15_Error.vbs`.

4.  **Test in a Safe Environment:** Be extremely careful with scripts that delete files, modify the registry, or change system settings. Always test them in a virtual machine or a non-critical environment first.

5.  **Comment Your Code:** Use the apostrophe (`'`) to add comments. This explains what your code does, both for others and for your future self. All the tutorial files are well-commented for this reason.

6.  **`Option Explicit` is Your Friend:** Add `Option Explicit` at the top of your scripts. This forces you to declare all variables with `Dim`, which helps catch typos (e.g., `usarName` vs `userName`) that would otherwise cause confusing errors. See `05_Option_Explicit.vbs`.

7.  **The Object Browser is Key:** To unlock VBScript's full potential, you need to use objects like `FileSystemObject` or `WScript.Shell`. Use the Microsoft documentation or the Object Browser in tools like VBSEdit to discover their methods and properties.

8.  **Embrace the Command Line:** Learn to run your scripts using `cscript` in the command prompt. It's essential for scheduling tasks with **Task Scheduler** and for logging output.

## The Future of VBScript

Microsoft has announced the deprecation of VBScript as a feature of the Windows operating system. It will be available as a feature on demand (FOD) before being completely removed in a future Windows release. While this means it's not a technology for new green-field projects, the vast amount of legacy automation still in use makes understanding it a valuable skill for IT professionals.

---

**Proceed to the [Tutorial Files](./README.md#-tutorial-files) to begin your VBScript journey!**
