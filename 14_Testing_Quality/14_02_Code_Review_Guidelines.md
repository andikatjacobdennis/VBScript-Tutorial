# Code Review Guidelines

**Filename:** 14_02_Code_Review_Guidelines.md  
**Category:** Testing & Quality – Ensuring Code Reliability  

---

## Purpose

Code reviews help ensure that VBScript (and other code) is:

- **Correct** – behaves as intended and produces the right output  
- **Readable** – easy for others to understand and maintain  
- **Secure** – avoids unsafe practices that could lead to vulnerabilities  
- **Consistent** – follows shared style and best practices  
- **Reliable** – handles errors gracefully and avoids breaking in production  

---

## Checklist for Reviewing VBScript Code

### ✅ Functionality
- Does the code meet the stated requirements?  
- Are edge cases handled (e.g., empty inputs, null values, errors)?  
- Has the code been manually tested (if unit tests are not available)?  

### 🛡️ Security
- Are inputs sanitized and validated before use?  
- Are credentials or sensitive values stored securely (never hardcoded)?  
- Does the code avoid unsafe operations like unrestricted file writes or registry edits?  

### 🧩 Readability & Style
- Is the code organized into logical sections with comments?  
- Are variables named meaningfully (not just `x`, `y`, `tmp`)?  
- Are magic numbers/strings avoided (use constants where appropriate)?  

### ⚙️ Error Handling
- Is `On Error Resume Next` used carefully (not hiding bugs)?  
- Are error messages clear and helpful for debugging?  
- Is `Err.Clear` used properly after handling an error?  

### 📄 Documentation
- Are header comments included (purpose, filename, usage)?  
- Are tricky parts of the code explained inline with comments?  
- Are example inputs/outputs or usage notes provided?  

---

## Best Practices

- **Small Commits:** Keep code changes small to make reviews manageable.  
- **Ask Questions:** If something is unclear, ask the author to explain.  
- **Be Respectful:** Reviews should be constructive, not critical.  
- **Consistency Over Preference:** Enforce team conventions, not personal opinions.  
- **Automate Where Possible:** Use linting or static analysis tools for consistency.  

---

## Example Review Feedback

❌ **Bad Feedback:**  
> "This code is wrong. Rewrite it."  

✅ **Good Feedback:**  
> "The function works, but it doesn’t handle empty input. Consider adding a check before processing the value."  

---

## Conclusion

Code reviews are not about blame — they are about **collaboration and improvement**.  
The goal is to make the code **simpler, safer, and easier to maintain**.  
