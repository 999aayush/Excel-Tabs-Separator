# Excel-Tabs-Separator
Splits Excel workbooks into individual files per tab while preserving full formatting, formulas, and layouts. Includes a Windows GUI EXE for native support of all formats (.xlsx, .xls, .xlsb, .xlsm) and a Cross-Platform Python Script for standard .xlsx/.xlsm support on macOS/Linux.

# Universal Excel Tab Separator

A dual-mode utility to split Excel workbooks into individual files per tab. 

This repository includes two versions of the tool:
1.  **Windows GUI (Pro)**: Uses native Excel automation to preserve **100%** of formatting and supports legacy formats (`.xls`, `.xlsb`).
2.  **Command Line (Lite)**: A lightweight, cross-platform script for standard `.xlsx` files.

---

## 🚀 Features & Compatibility

| Feature | 🖥️ GUI Version (`script-gui-exe.pyw`) | 💻 CLI Version (`script.py`) |
| :--- | :--- | :--- |
| **Interface** | Dark Mode GUI | Command Line / Terminal |
| **OS Support** | **Windows Only** | Windows, macOS, Linux |
| **Engine** | Microsoft Excel (via `win32com`) | Python (`openpyxl`) |
| **Supported Files** | `.xlsx`, `.xls`, `.xlsb`, `.xlsm` | `.xlsx`, `.xlsm` |
| **Formatting** | **Perfect** (Native Excel Copy) | Basic (Data & Styles) |

---

## 🛠️ Option 1: GUI Version (Windows Only)
*Best for: Users who need to preserve complex formatting, print layouts, or work with older .xls/.xlsb files.*

### How to Run
**Method A: Using the Executable (Recommended)**
1.  Download `Excel-Tabs-Separator.exe` from the Releases section (or the folder provided).
2.  Double-click to launch.
3.  Drag and drop your Excel file or browse to select it.

**Method B: Running from Source**
If you want to run the raw Python script (`script-gui-exe.pyw`), you **cannot** double-click it directly due to environment handling logic.
1.  Install requirements: `pip install pywin32`
2.  Run the included `.bat` file:
    ```bash
    run.bat
    ```
    *(Or launch it manually via CMD: `pythonw script-gui-exe.pyw`)*

**Note for Developers:**
This script (`script-gui-exe.pyw`) is optimized for compilation using **Auto Py to Exe** (PyInstaller). It contains logic to handle path freezing and environment switching automatically when compiled.

---

## 💻 Option 2: Command Line Version (Cross-Platform)
*Best for: Mac/Linux users, or quick batch processing of standard .xlsx files.*

### Requirements
* Python 3.x
* Library: `pip install openpyxl`

### How to Run
1.  Open your terminal or command prompt.
2.  Run the script:
    ```bash
    python script.py
    ```
3.  **Drag and Drop:** You can drag your Excel file directly onto the terminal window when prompted.

⚠️ **Important Note on Drag & Drop (Windows):**
Windows prevents drag-and-drop operations between programs with different permission levels.
* **DO NOT** run Command Prompt as Administrator if you intend to drag and drop files.
* Use a standard (non-admin) CMD window.

---

## 📝 License
[MIT](LICENSE)
