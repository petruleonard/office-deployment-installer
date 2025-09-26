# Office Deployment Tool Installer

This PowerShell script automates the installation of Microsoft Office using the **Office Deployment Tool (ODT)**.  
It generates a configuration XML based on user selections (product, language, excluded apps) and runs the installer.  

✅ Features:  
- Runs with administrator check  
- Interactive product and language selection  
- Option to exclude specific Office apps  
- Generates native configuration XML  
- Downloads the latest ODT setup with a console progress bar  
- Cleans up temporary files after installation  

⚠️ Note: The script must be run as Administrator.  

## How to run

Open **PowerShell** (Run as Administrator) and execute:

   [irm https://raw.githubusercontent.com/petruleonard/office-deployment-installer/refs/heads/main/odt.ps1 | iex

This will automatically download and run the installer script.

---

## Notes
- Requires Windows with PowerShell 5.1 or newer.
- Must be executed with Administrator privileges.
- Internet connection is required during installation.


