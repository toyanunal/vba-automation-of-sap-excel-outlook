## Integrate SAP, Excel and Outlook using Excel VBA 


0. SAP_GUI.bat => WHY TO USE: easily log into SAP by executing "SAP_GUI.bat"
    
    HOW TO USE:
    * run the code below on Command Prompt
    
      `Call Shell("cmd.exe /s /c " & "C:\Users\" & Environ$("UserName") & "\Desktop\SAP_GUI.bat", vbNormalFocus)`
    * double-click "SAP_GUI.bat" icon on its location (for our case, it is C:\Users\<username>\Desktop)
    


1. VBA_to_SAP.txt => WHY TO USE: manipulate SAP using Excel VBA codes
    
    HOW TO USE:
    * create a Macro-Enabled Excel file (.xlsm)
    * paste this code inside a module in that Excel file


2. Delete_Empty_Cells.txt => WHY TO USE: delete empty cells, rows and columns in Excel workbook
    
    HOW TO USE:
    * create a Macro-Enabled Excel file (.xlsm)
    * paste this code inside a module in that Excel file



3. Send_Email.txt => WHY TO USE: send e-mails via Outlook using Excel VBA codes
    
    HOW TO USE:
    * create a Macro-Enabled Excel file (.xlsm)
    * paste this code inside a module in that Excel file


4. Workbook_Open_Event.txt => WHY TO USE: automatically trigger Excel VBA codes when its Excel file is opened   
    
    HOW TO USE:
    * create a Macro-Enabled Excel file (.xlsm)
    * paste this code inside ThisWorkbook object in that Excel file
    * schedule opening of that Excel file using Windows Task Scheduler
