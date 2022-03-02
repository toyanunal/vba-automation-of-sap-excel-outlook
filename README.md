### Integrate SAP, Excel and Outlook using Excel VBA 


0. SAP_GUI.bat => WHY TO USE: easily log into SAP by executing "SAP_GUI.bat"
    
    HOW TO USE:
    * run the code below on Command Prompt
    
      `Call Shell("cmd.exe /s /c " & "C:\Users\" & Environ$("UserName") & "\Desktop\SAP_GUI.bat", vbNormalFocus)`
    * double-click "SAP_GUI.bat" icon on its location (for our case, it is C:\Users\<username>\Desktop)
    


1. VBA_to_SAP.txt => WHY TO USE: manipulate SAP using Excel VBA codes
    
    HOW TO USE:
    * create a Macro-Enabled Excel file (.xlsm)
    * paste this code inside a module in that Excel file


