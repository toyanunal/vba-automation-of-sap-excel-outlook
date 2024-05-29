# Integrate SAP, Excel, and Outlook using Excel VBA

## 0. SAP_GUI.bat
**Why Use It:** Easily log into SAP by executing `SAP_GUI.bat`.

**How to Use:**
- Run the following code in Command Prompt:
    ```sh
    Call Shell("cmd.exe /s /c " & "C:\Users\" & Environ$("UserName") & "\Desktop\SAP_GUI.bat", vbNormalFocus)
    ```
- Double-click the `SAP_GUI.bat` icon at its location (for our case, it is `C:\Users\<username>\Desktop`).

## 1. VBA_to_SAP.txt
**Why Use It:** Manipulate SAP using Excel VBA codes.

**How to Use:**
- Create a Macro-Enabled Excel file (`.xlsm`).
- Paste the code inside a module in that Excel file.

## 2. Delete_Empty_Cells.txt
**Why Use It:** Delete empty cells, rows, and columns in an Excel workbook.

**How to Use:**
- Create a Macro-Enabled Excel file (`.xlsm`).
- Paste the code inside a module in that Excel file.

## 3. Send_Email.txt
**Why Use It:** Send emails via Outlook using Excel VBA codes.

**How to Use:**
- Create a Macro-Enabled Excel file (`.xlsm`).
- Paste the code inside a module in that Excel file.

## 4. Workbook_Open_Event.txt
**Why Use It:** Automatically trigger Excel VBA codes when the Excel file is opened.

**How to Use:**
- Create a Macro-Enabled Excel file (`.xlsm`).
- Paste the code inside the `ThisWorkbook` object in that Excel file.
- Schedule the opening of that Excel file using Windows Task Scheduler.
