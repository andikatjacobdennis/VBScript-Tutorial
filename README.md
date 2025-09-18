# VBScript Tutorial

A comprehensive, well-structured collection of VBScript examples and tutorials.
This repository is designed to take you from **absolute beginner** to **proficient Windows scripter**, covering everything from **basic message boxes** to **system integration, security, logging, testing, and deployment**.

---

## Introduction

VBScript (Visual Basic Scripting Edition) is a lightweight, interpreted scripting language developed by Microsoft. It is ideal for **automating tasks in Windows environments**, such as:

* File and folder manipulation
* Registry access
* Logon/startup scripts
* System integration via WSH, WMI, and COM

While VBScript is now considered **legacy technology**, it remains a valuable skill for:

* Maintaining and modernizing legacy systems
* Understanding Windows automation fundamentals
* Learning structured scripting before moving to PowerShell or Python

**New to VBScript?** Start with the **[Introduction Guide](./Introduction.md)** to learn about its applications, the Windows Script Host (WSH), and helpful tips for getting started.

---

## Repository Structure

The tutorials are organized into **logical modules** with consistent numbering for easy navigation. Each module addresses either **language fundamentals** or **architectural concerns**.

```text
VBScript-Tutorial/
│
├── 📂 01_Basics/                  # Fundamental Concepts
├── 📂 02_Data_Structures/         # Working with Data
├── 📂 03_Control_Flow/            # Program Logic & Loops
├── 📂 04_Procedures/              # Functions & Subroutines
├── 📂 05_Built_in_Functions/      # Core Language Functions
├── 📂 06_User_Interaction/        # Getting Input from Users
├── 📂 07_File_System_Operations/  # Interacting with Files & Folders
├── 📂 08_System_Integration/      # Advanced Windows Features
├── 📂 09_Error_Handling/          # Making Scripts Robust
├── 📂 10_Security/                # Security Best Practices
├── 📂 11_Logging_Monitoring/      # Observability & Diagnostics
├── 📂 12_Configuration_Management/# Externalizing Settings
├── 📂 13_InterProcess_Communication/ # Automation & Integration
├── 📂 14_Testing_Quality/         # Ensuring Code Reliability
├── 📂 15_Deployment_Versioning/   # Packaging & Version Control
├── 📂 16_Performance_Optimization/# Optimizing Execution
├── 📂 Examples/                   # Practical Example Scripts
├── 📄 Introduction.md
├── 📄 README.md
├── 📄 LICENSE
└── 📄 VBScript_Snippets.md
```

---

## Module Overview

| Module                            | Purpose                                                                                                                     |
| --------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| **01–09 Core Language**           | Syntax, data structures, flow, procedures, built-in functions, user input, file system, system integration, error handling. |
| **10 Security**                   | Safe input handling, least-privilege scripting, secure file & registry access.                                              |
| **11 Logging & Monitoring**       | Script observability: log files, event viewer integration, troubleshooting.                                                 |
| **12 Configuration Management**   | Externalizing settings: INI files, environment variables, registry keys.                                                    |
| **13 Interprocess Communication** | WMI, COM automation, remote scripting, integration with other processes.                                                    |
| **14 Testing & Quality**          | Unit testing, static analysis, coding guidelines, best practices.                                                           |
| **15 Deployment & Versioning**    | Script version headers, distribution strategies, CI/CD integration.                                                         |
| **16 Performance Optimization**   | Profiling, efficient loops, caching, minimizing system calls.                                                               |

---

## Getting Started

### Prerequisites

* Windows operating system (7, 10, or 11)
* Windows Script Host (WSH), included by default

### Running Scripts

You can run any `.vbs` file in two ways:

1. **Double-Click (Graphical):**
   Runs with `wscript.exe` and shows message boxes.

2. **Command Line (Console):**
   Recommended for automation:

   ```cmd
   cscript //nologo "01_Basics\01_01_My_First_VBScript.vbs"
   ```

   The `//nologo` flag hides the copyright banner.

---

## Learning Path

Follow this order for a smooth progression:

1. **Basics → Data Structures → Control Flow**
2. **Procedures → Built-in Functions → User Interaction**
3. **File System → System Integration → Error Handling**
4. **Security → Logging → Configuration Management**
5. **Interprocess Communication → Testing & Quality**
6. **Deployment & Versioning → Performance Optimization**

This path mirrors how an architect designs software: start with fundamentals, then add cross-cutting concerns like **security, logging, configuration, testing, deployment, and performance**.

---

## Important Notes

* **Test Safely:** Always test scripts that modify files or the registry in a non-production environment first.
* **Administrative Rights:** Certain scripts (e.g., registry edits) require Administrator privileges.
* **Deprecation Notice:** VBScript is deprecated. For modern automation, prefer **PowerShell or Python**—but understanding VBScript is invaluable for maintaining legacy systems.

---

## Contributing

We welcome contributions! You can add:

* New example scripts
* Improvements to tutorials
* Fixes or best practices

### Workflow

1. Fork the repo
2. Create a feature branch (`git checkout -b feature/amazing-example`)
3. Commit changes (`git commit -m "Add amazing example"`)
4. Push to your branch (`git push origin feature/amazing-example`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

---

## Resources

* [Windows Script Host Documentation (Microsoft)](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1et7k7c%28v=vs.84%29)
* [VBScript Language Reference (Microsoft)](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1wf56tt%28v=vs.84%29)

---

**Happy Scripting!**
If you find this tutorial helpful, please give it a ⭐
