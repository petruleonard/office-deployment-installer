# Office Deployment Tool Installer

This script automates Microsoft Office installation using the Office Deployment Tool (ODT).  
It allows you to select the product, language, and excluded apps, then generates a native XML configuration and runs the installer.  
A console-based progress bar is displayed during setup.exe download.

---

## How to run

Open **PowerShell** (Run as Administrator) and execute:

    irm https://raw.githubusercontent.com/petruleonard/office-deployment-installer/refs/heads/main/odt.ps1 iex

This will automatically download and run the installer script.

---

## Notes
- Requires Windows with PowerShell 5.1 or newer.
- Must be executed with Administrator privileges.
- Internet connection is required during installation.
