# VBScript Tutorial 📜

A comprehensive, well-structured collection of VBScript examples and tutorials. This repository is designed to take you from absolute beginner to proficient in Windows scripting, covering everything from basic message boxes to file system operations and registry access.

## 📖 Introduction

VBScript (Visual Basic Scripting Edition) is a lightweight, interpreted scripting language from Microsoft, perfect for automating tasks in a Windows environment. While a legacy technology, it remains a powerful tool for system administrators and developers for task automation, logon scripts, and simple GUI applications.

**New to VBScript?** Start with the **[Introduction Guide](./introduction.md)** to learn about its applications, the Windows Script Host (WSH), and helpful tips for getting started.

## 📁 Repository Structure

The tutorials are organized into logical folders with consistent numbering for easy navigation.

```text
vbs-tutorial/
│
├── 📂 01_Basics/                 # Fundamental Concepts
│   ├── 01_01_My_First_Vb_Script.vbs
│   ├── 01_02_MsgBox.vbs
│   ├── 01_03_Operations.vbs
│   ├── 01_04_Line_Continuation.vbs
│   └── 01_05_Option_Explicit.vbs
│
├── 📂 02_Data_Structures/        # Working with Data
│   └── 02_01_Array.vbs
│
├── 📂 03_Control_Flow/           # Program Logic & Loops
│   ├── 03_01_Condition.vbs
│   └── 03_02_Loop.vbs
│
├── 📂 04_Procedures/             # Functions & Subroutines
│   ├── 04_01_Procedures.vbs
│   └── 04_02_ByVal_ByRef.vbs
│
├── 📂 05_Built_in_Functions/     # Core Language Functions
│   └── 05_01_Built_in_Functions.vbs
│
├── 📂 06_User_Interaction/       # Getting Input from Users
│   └── 06_01_Input.vbs
│
├── 📂 07_File_System_Operations/ # Interacting with Files & Folders
│   └── 07_01_Folder_File.vbs
│
├── 📂 08_System_Integration/     # Advanced Windows Features
│   └── 08_01_Registery.vbs
│
├── 📂 09_Error_Handling/         # Making Scripts Robust
│   └── 09_01_Error.vbs
│
├── 📂 examples/                  # Practical example scripts
│   ├── simple_login_script.vbs
│   ├── find_and_replace.vbs
│   └── system_info_check.vbs
│
├── 📂 resources/                 # Helpful references
│   └── useful_links.md
│
├── 📄 introduction.md
├── 📄 README.md
└── 📄 LICENSE
```

## 🚀 Getting Started

### Prerequisites

- A Windows operating system (Windows 7, 10, or 11)
- Windows Script Host (WSH), which is included by default in Windows

### How to Run the Scripts

You can execute any VBScript (`.vbs`) file in two ways:

1. **Double-Click (Graphical):**
   Simply double-click the file. It will run using `wscript.exe` and display output in message boxes.

2. **Command Line (Console):**
   Open Command Prompt and use `cscript` for cleaner output, which is ideal for automation.

   ```cmd
   cscript //nologo "01_Basics\01_01_My_First_Vb_Script.vbs"
   ```

   The `//nologo` flag suppresses the copyright banner.

## 🧭 Learning Path

We recommend following the tutorials in this order:

1. **Basics:** Understand variables, message boxes, and basic operations.
   - `01_01_My_First_Vb_Script.vbs` → `01_02_MsgBox.vbs` → `01_03_Operations.vbs`

2. **Foundations:** Learn best practices and how to structure code.
   - `01_04_Line_Continuation.vbs` → `01_05_Option_Explicit.vbs`

3. **Data & Logic:** Work with arrays, conditional statements, and loops.
   - `02_01_Array.vbs` → `03_01_Condition.vbs` → `03_02_Loop.vbs`

4. **Modularity:** Organize your code into reusable procedures and functions.
   - `04_01_Procedures.vbs` → `04_02_ByVal_ByRef.vbs`

5. **Power Tools:** Use built-in functions for strings, dates, and conversions.
   - `05_01_Built_in_Functions.vbs`

6. **Interaction & IO:** Get user input and work with the file system.
   - `06_01_Input.vbs` → `07_01_Folder_File.vbs`

7. **Advanced Topics:** Integrate with Windows and handle errors gracefully.
   - `08_01_Registery.vbs` → `09_01_Error.vbs`

## ⚠️ Important Notes

- **Test in a Safe Environment:** Be cautious with scripts that modify files (`07_01_Folder_File.vbs`) or the Windows Registry (`08_01_Registery.vbs`). Always test them in a non-critical environment first.
- **Administrative Rights:** Some operations, especially writing to protected areas of the registry or file system, may require running the script as an Administrator.
- **Deprecation Notice:** VBScript is a deprecated technology. While it's invaluable for maintaining legacy systems and learning core concepts, consider PowerShell or Python for new automation projects.

## 🤝 Contributing

Contributions are welcome! If you have a useful example script, an improvement to an existing tutorial, or a correction, please feel free to:

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/amazing-example`).
3. Commit your changes (`git commit -m 'Add some amazing example'`).
4. Push to the branch (`git push origin feature/amazing-example`).
5. Open a Pull Request.

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔗 Resources

- [Windows Script Host Documentation (Microsoft)](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1et7k7c(v=vs.84))
- [VBScript Language Reference (Microsoft)](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1wf56tt(v=vs.84))

---

**Happy Scripting!** If you find this tutorial helpful, please give it a ⭐!
