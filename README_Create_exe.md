📦 How to Convert a Python Script to a .EXE File (Windows)
This guide walks you through converting a Python script into a standalone Windows executable (.exe) using PyInstaller.

✅ Requirements

Python installed (recommended: Python 3.10+)
Script must be working before compilation
If automating Excel, Microsoft Excel must be installed

🔧 Step-by-step Guide
Open Command Prompt or Terminal
Install PyInstaller (first time only):

pip install pyinstaller


Navigate to the folder with your .py file:
cd "C:\Path\To\Your\Script\Folder"


(Optional) Clean previous builds:
rmdir /s /q build
rmdir /s /q dist
del your_script.spec

Create the .exe file:
pyinstaller --onefile your_script.py
✅ The .exe will be located in the dist/ folder.

🖥️ Hide Console Window (Optional)
To prevent the terminal from appearing when running the .exe:
pyinstaller --onefile --noconsole your_script.py


💡 Tips
Use raw strings in Python for paths: r"C:\folder\file.xlsx"
Always use full paths for input/output files when compiling
Excel automation requires Excel to be installed on the user's machine


📤 Sharing the .EXE
You can send only the .exe to other users
Users do not need Python installed
Ensure they have all required input files (e.g., templates)

