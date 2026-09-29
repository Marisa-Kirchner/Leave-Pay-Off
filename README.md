# Annual Leave Script Usage

Generates personalized annual leave pay off letters for employees.

## Description

Generates personalized annual leave pay off letters from a given Excel sheet. Will choose the "CBA" or "Non-CBA" Word template based on whether the employee has a personal leave line on the Excel sheet. Fills out the Word doc with the employee name, leave amounts left per type, and annual leave hours that will be lost. Saves the filled out Word doc with a specified name and in a given output folder.

## Getting Started

### Dependencies

* Must have Microsoft Office downloaded

### Initial Setup

* Go to File Explorer, then go to your Documents folder
* Create a new folder named Leave
* Copy the Word doc templates for CBA & Non-CBA to the Leave folder
* Copy the Excel sheets you will be using to the Leave folder
* Copy the script Annual_Leave.ps1 to the Leave folder

### Getting Ready

* Go to Start & search "Make Me Admin"
* You will select make me admin & log in using your credentials
* Go to Start, search "Windows Powershell", & then click "Run as Admininstrator"

## Executing the Script

 Copy the command below & right click on the powershell window to paste there.

```
powershell.exe -ExecutionPolicy Bypass -File .\Annual_Leave.ps1 -Rename "file" -CBAPath "cba" -NonCBAPath "noncba" -ExcelPath "excel" -OutputFolder "ouput"
```

* The Word doc will be saved as "Employee_Name file_name (CBA/Non-CBA).docx". The file_name is currently "Annual Leave Pay Off 2026". If you want to rename it, replace "file" from the above box with a "new file name". If you do not want to rename it, remove -Rename "file" from your command. It would then look like the following: 

```
powershell.exe -ExecutionPolicy Bypass -File .\Annual_Leave.ps1 -CBAPath "cba" -NonCBAPath "noncba" -ExcelPath "excel" -OutputFolder "ouput"
```

* Go to your file explorer & right click on the CBA_Template doc. Then select copy as path. Now replace "cba" with the path you just copied.

* Go to your file explorer & right click on the Non_CBA_Template doc. Then select copy as path. Now replace "noncba" with the path you just copied.

* Go to your file explorer & right click on the Excel sheet you wish to use. Then select copy as path. Now replace "excel" with the path you just copied.

* Go to your file explorer & right click on your Leave folder. Then select copy as path. Now replace "output" with the path you copied. You will have to add on \FOLDER_NAME with a folder name of your choosing to the path. Make sure this is inside the end quotes.

#### Example Output Folder Path
```
"C:\Users\username\Documents\Leave\ADMIN"
```


## Help

To see the help message copy and paste into the Powershell window:
```
Get-Help .\Annual_Leave.ps1
```

### Windows Powershell Tips
* You cannot use your mouse to move your cursor in the Powershell window. You must use your left and right arrow keys to move the cursor in the window.
* To paste in the Powershell window, you only have to right click.
* To reuse the last command you entered, press the up arrow key. You can then edit it without retyping everything.

## Authors
Marisa Kirchner

[@Marisa-Kirchner](https://github.com/Marisa-Kirchner)

## Version History

* 0.1
    * Initial Release
