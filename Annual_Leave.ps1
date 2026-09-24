<#
.SYNOPSIS
    Generates personalized annual leave pay off letters from a given Excel sheet.
    Will choose the "CBA" or "Non-CBA" Word template based on whether the 
    employee has a personal leave line on the Excel sheet. Fills out the Word doc
    with the employee name, leave amounts left per type, and annual leave hours that 
    will be lost. Saves the filled out Word doc into a given output folder.

.DESCRIPTION
    Excel Sheet Layout:
        Row 1  : Title Row (ignored)
        Row 2  : Important Headers -> Worker, Time Off Plan, Total Balance as of Balance Period End Date, 
                 Current Balance at Risk of Forfeiture Excluding Pending Events as of Next Carryover Date
        Row 3+ : One row PER Leave Type PER Employee
                 (CLA Annual Time Off Plan / CLA Comp Time Time Off Plan /
                  CLA Sick Time Off Plan / Personal Leave Time Off Plan-if CBA)

    For each employee, this script:
        - Pulls the "Total Balance as of Balance Period End Date" value for each leave type
          into ANNUAL / COMP / SICK / PERSONAL
        - Pulls the "Current Balance at Risk of Forfeiture Excluding Pending Events as of Next Carryover Date" 
          value from the Annual Leave row into [WILL_LOSE]
        - Opens CBA_Doc_Name.doc  (CBA = "Yes", if employee has a Personal Leave line)
          or Non_CBA_Doc_Name.doc (CBA = "No", if there is no Personal Leave line)
        - Replaces [EMPLOYEE_NAME], [ANNUAL], [COMP], [SICK], [PERSONAL], [WILL_LOSE] in Word doc
        - Saves a copy named "<Employee Name> <OutputFileName> <CBA/Non-CBA>.doc" in the output folder

    Leave Types:
        ANNUAL
          - CLA Annual Time Off Plan on Excel sheet
          - Amount of leave that employee will lose if not used comes from annual leave
        COMP
          - CLA Comp Time Time Off Plan on Excel sheet
        SICK
          - CLA Sick Time Off Plan on Excel sheet
        PERSONAL (only if CBA)
          - Personal Leave Time Off Plan on Excel sheet

.EXAMPLE
    .\Annual_Leave.ps1 
        -ExcelPath "C:\Leave\Excel Sheet Name.xlsx"
        -Output "C:\Leave\TRADE FOLDER"
        -CBAPath "C:\Leave\CBA_Doc_Name.doc"
        -NonCBAPath "C:\Leave\Non_CBA_Doc_Name.doc"
#>

<#
The script's inputs:
    - Excel path is required.
    - The others all have a default value, so running the script with 
      no arguments for those, uses their hardcoded paths.
    - Can still override any of them
    - Edit the this block if the files ever move.
#>
# *** THESE ARE CURRENTLY MADE UP PATHS ***
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [string]$CBAPath        = "C:\HR\LEAVE\Annual_Leave_Pay_Off_(CBA).doc",
    [string]$NonCBAPath     = "C:\HR\LEAVE\Annual_Leave_Pay_Off_(Non_CBA).doc",
    [string]$OutputFolder   = "C:\HR\LEAVE\Pay_Off_Letters"
)

# Any error halts the script
$ErrorActionPreference = "Stop"

# Resolve to absolute paths & throw an error if the file doesn't exist
$ExcelPath  = (Resolve-Path $ExcelPath).Path
$CBAPath    = (Resolve-Path $CBAPath).Path
$NonCBAPath = (Resolve-Path $NonCBAPath).Path

<# 
Check if OutputFolder exists with Test-Path
If not, create it with New-Item, then resolve to absolute path
#>
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}
$OutputFolder = (Resolve-Path $OutputFolder).Path

# ---- Excel Columns ----
# Named constants for which excel column holds each field (1 = A, 2 = B, etc.)
# Edit here if the columns change on the excel sheet.
$COL_NAME = 3    # Employee Name
$COL_TYPE = 4    # Leave Type
$COL_LEFT = 10   # Leave hours left
$COL_LOSE = 14   # Leave hours that will be lost

# ---- Output File Naming ----
# File names come out as: "<FIRST LAST> $OutputFileName (CBA/Non-CBA).docx"
# Edit here to rename the output file.
$OutputFileName = "Annual Leave Pay Off 2026"

# Initialize $excel, $workbook, $word to $null before the try block,
# so the finally block at the bottom can safely check "did this get created"
$excel    = $null
$workbook = $null
$word     = $null

try {
    # ---------------- Read the Excel data ----------------
    Write-Host "Opening Excel workbook..."

    # Launch hidden instance of Excel & give PowerShell a handle to control it
    $excel = New-Object -ComObject Excel.Application
    # Keep it invisible
    $excel.Visible = $false
    # Suppress "do you want to save" popups that could freeze the script
    $excel.DisplayAlerts = $false

    # Open the file, get the first worksheet, and ask Excel how many rows actually contain data
    $workbook = $excel.Workbooks.Open($ExcelPath)
    $worksheets = $workbook.Worksheets
    $sheet = $worksheets.Item(1)
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $data = $usedRange.Value2   # Whole range pulled in one COM call

    # An ordered hashtable
    # Key = employee name, value = another hashtable holding their accumulated data
    $employees = [ordered]@{}

    # Data starts on row 3 (row 1 = title, row 2 = headers)
    for ($r = 3; $r -le $rowCount; $r++) {

        # Read the cell at (row, col)
        $name = $data[$r, $COL_NAME]            # Employee name
        
        # If the row is blank, skip it
        if ([string]::IsNullOrWhiteSpace([string]$name)) { continue }

        # Get all data & make them strings
        $name = [string]$name                   # Employee name
        $type = [string]$data[$r, $COL_TYPE]    # Leave type
        $left = [string]$data[$r, $COL_LEFT]    # Leave hours left
        $lose = [string]$data[$r, $COL_LOSE]    # Leave that will be lost

        # The first time seeing an employee's name, create their entry with blank defaults.
        if (-not $employees.Contains($name)) {
            $employees[$name] = @{
                HasPersonal = $false   # set to $true below if a "Personal" row is found
                Annual      = ""
                Comp        = ""
                Sick        = ""
                Personal    = ""
                Will_Lose   = ""
            }
        }

        # Edit here if the names of the leave types in the excel sheet change.
        switch ($type) {
            "CLA Annual Time Off Plan"      { $employees[$name].Annual   = $left; $employees[$name].Will_Lose = $lose }
            "CLA Comp Time Time Off Plan"   { $employees[$name].Comp     = $left }
            "CLA Sick Time Off Plan"        { $employees[$name].Sick     = $left }
            "Personal Leave Time Off Plan"  { $employees[$name].Personal = $left; $employees[$name].HasPersonal = $true }
        }
    }

    # Release the COM objects that were created while reading the Excel data
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheets) | Out-Null

    $workbook.Close($false)     # Close the workbook without saving changes
    $excel.Quit()               # Close the Excel application itself

    <#
    The ReleaseComObject calls are necessary because COM objects 
    aren't cleaned up by .NET's normal garbage collector
    #>
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Stop the finally block from trying to quit a second time
    $excel = $null
    $workbook = $null

    # Force .NET to drop the released COM references
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()

    # ------ Output Number of Employees in Hashtable ------
    Write-Host "Found $($employees.Count) employee(s) in the workbook."

    # ---------------- Generate the Letters ----------------
    Write-Host "Opening Word..."

    # Launch a hidden Word instance
    $word = New-Object -ComObject Word.Application
    # Keep it invisible
    $word.Visible = $false
    # wdAlertsNone - suppress prompts
    $word.DisplayAlerts = 0

    # Word's Find.Execute - replace all matches
    $wdReplaceAll = 2
    # Word's SaveAs - save as .docx format
    $wdFormatXMLDocument = 12

    foreach ($name in $employees.Keys) {
        $data = $employees[$name]
        $isCBA = $data.HasPersonal

        # Choose correct Word template for CBA/Non-CBA
        $templatePath = if ($isCBA) { $CBAPath } else { $NonCBAPath }

        Write-Host "Generating letter for '$name' (CBA = $isCBA) using template: $(Split-Path $templatePath -Leaf)"

        # If the name on the Excel sheet is "Last, First" - flip it to "First Last"
        $nameParts = $name -split ',\s*', 2
        if ($nameParts.Count -eq 2) {
            $empNameLetter = "$($nameParts[1].Trim()) $($nameParts[0].Trim())"
        } 
        # Otherwise use the name as is
        else {
            $empNameLetter = $name
        }


        # Open the template file as read-only (3rd arg) to guarantee the it remains unchanged
        $doc = $word.Documents.Open($templatePath, $false, $true)

        # Placeholder -> value map (set of replacements for the letter)
        $replacements = [ordered]@{
            "[EMPLOYEE_NAME]"   = $empNameLetter
            "[ANNUAL]"          = $data.Annual
            "[COMP]"            = $data.Comp
            "[SICK]"            = $data.Sick
            "[WILL_LOSE]"       = $data.Will_Lose
        }
        # Only has Personal Hours if CBA is Yes
        if ($isCBA) {
            $replacements["[PERSONAL]"] = $data.Personal
        }

        # Use Word's built-in Find object on the whole document body ($doc.Content)
        foreach ($key in $replacements.Keys) {
            $find = $doc.Content.Find
            $find.ClearFormatting()     # Make sure it's a plain text search/replace
            $find.Replacement.ClearFormatting()
            $find.Execute(
                $key,               # FindText
                $false,             # MatchCase
                $false,             # MatchWholeWord
                $false,             # MatchWildcards
                $false,             # MatchSoundsLike
                $false,             # MatchAllWordForms
                $true,              # Forward
                1,                  # Wrap (wdFindContinue)
                $false,             # Format
                $replacements[$key],# ReplaceWith
                $wdReplaceAll       # Replace
            ) | Out-Null            # Pipes Find.Execute()'s return value into nothing, discarding it silently
        }

        # ----- Build a safe output file name -----
        # Strip out characters that aren't legal in Windows filenames by replacing them with underscores
        $employeeFileName = ($empNameLetter -replace '[\\/:*?"<>|]', '_').ToUpper()
        $suffix = if ($isCBA) { "CBA" } else { "Non-CBA" }
        $outputPath = Join-Path $OutputFolder "$employeeFileName $OutputFileName ($suffix).docx"

        # Save the modified in-memory doc to a new file
        $doc.SaveAs([string]$outputPath, $wdFormatXMLDocument)
        # Close the in-memory copy without prompting to save again
        $doc.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    }

    Write-Host "Done. Letters saved to: $OutputFolder"
}

# Check each COM object variable; if it was created, quit the application and release it
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Call garbage collector to free the released COM references right away
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
