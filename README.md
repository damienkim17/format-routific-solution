# Step-by-Step Instructions

## 1. Open Excel and enable the Developer tab:
  - Go to File > Options > Customize Ribbon.
  - Check the box next to Developer in the right-hand panel and click OK.

## 2. Open the VBA Editor
  - Click on the Developer tab and then on Visual Basic.

## 3. Create a Personal Macro Workbook (if it doesn't exist):
  - In the VBA Editor, check if VBAProject (PERSONAL.XLSB) exists in the Project Explorer.
  - If it doesn't exist, create it by recording a dummy macro:
    - Close the VBA Editor.
    - Go to View > Macros > Record Macro.
    - Choose Store macro in: Personal Macro Workbook.
    - Click OK and then immediately stop the recording by clicking View > Macros > Stop Recording.
    - Reopen the VBA Editor.
   
## 4. Insert a New Module:
  - In the Project Explorer, locate VBAProject (PERSONAL.XLSB).
  - Right-click on VBAProject (PERSONAL.XLSB), select Insert > Module.

## 5. Copy and Paste the script.vba code

## 6. Save and Close the VBA Editor:
  - Save the macro by clicking File > Save.
  - Close the VBA Editor by clicking the X in the top-right corner or pressing Alt + Q.

## 7. Run the Macro:
  - Click on Macros in the Developer tab
  - Select Select FilterAndSplitData from the list of macros.
  - Click Run.
