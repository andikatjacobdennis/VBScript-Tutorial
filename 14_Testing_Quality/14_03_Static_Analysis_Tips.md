# Static Analysis Tips

**Filename:** 14_03_Static_Analysis_Tips.md  
**Category:** Testing & Quality – Ensuring Code Reliability  

---

## Purpose

Static analysis is the process of examining code **without executing it**.  
In VBScript and other scripting languages, static analysis helps catch issues **before runtime**, improving code reliability, readability, and security.  

---

## What to Look For

### 🧩 Syntax Issues
- Check for **missing `End If`, `Next`, `Loop`** statements.  
- Look for **unbalanced parentheses, quotes, or string concatenations**.  
- Verify **Option Explicit** is used to enforce variable declarations.  

### ⚙️ Code Quality
- Ensure variables are **declared with `Dim`** and not reused for unrelated values.  
- Avoid **duplicate code blocks** – extract them into functions.  
- Remove **unused variables or functions**.  

### 🛡️ Security
- Flag **hardcoded passwords, usernames, or connection strings**.  
- Identify places where **user input is directly used** (e.g., in file paths or WMI queries) without sanitization.  
- Look for **dangerous functions** (e.g., `Execute`, `Eval`) that could enable code injection.  

### 🧹 Maintainability
- Ensure functions are **short and focused** on one task.  
- Use **consistent indentation and spacing**.  
- Check that **comments match the actual behavior** of the code.  

---

## Tools & Techniques

While VBScript does not have modern static analysis tools like Python or JavaScript, you can still apply these techniques:

- **Manual Review**  
  Go line by line, looking for the issues above.  

- **Windows Script Host Error Checking**  
  Running `cscript.exe script.vbs //X` enables step-by-step debugging.  

- **VBScript Linters (Community Tools)**  
  Some community editors (e.g., Notepad++, VS Code with plugins) provide basic linting or highlighting for VBScript.  

- **Search & Grep**  
  Use text search to quickly find unsafe patterns (e.g., `Execute`, `On Error Resume Next`).  

---

## Best Practices for Static Checks

- Always start scripts with `Option Explicit`.  
- Review **all error handling blocks** to ensure they don’t mask problems.  
- Verify **MsgBox and logging messages** are accurate and helpful.  
- Check for **hardcoded file paths**; use configuration or environment variables instead.  
- Standardize **naming conventions** across all scripts.  

---

## Example Static Review Findings

❌ **Problematic Code:**  
```vbscript
On Error Resume Next
result = 10 / 0
````

⚠️ **Static Analysis Note:**

* Using `On Error Resume Next` without error checks hides division by zero.

✅ **Improved Code:**

```vbscript
On Error Resume Next
result = 10 / 0
If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description
    Err.Clear
End If
```

---

## Conclusion

Static analysis is about **prevention, not reaction**.
By systematically checking scripts for errors, bad practices, and security flaws, you can make VBScript code **safer, cleaner, and more reliable** — before it ever runs.
