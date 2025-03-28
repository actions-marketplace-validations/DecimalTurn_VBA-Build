# Summary:
# This PowerShell script automates the process of importing VBA modules into an Excel workbook.
# It retrieves the current working directory, constructs the path to the Excel file,
# and imports all .bas files from a specified module folder into the workbook.
# It then saves and closes the workbook, and cleans up the COM objects.


# Get the current working directory
$currentDir = (Get-Location).Path
Write-Host "Current working directory: $currentDir"

# Define the Excel file path
$excelFile = Join-Path $currentDir "skeleton.xlsm"

# Create a new Excel application
$excelApp = New-Object -ComObject Excel.Application

# Open the workbook
$wb = $excelApp.Workbooks.Open($excelFile)

# Make Excel visible (uncomment if needed)
# $excelApp.Visible = $true

# Define the module folder path
$moduleFolder = Join-Path $currentDir "src\skeleton.xlsm\Modules"

# Check if the module folder exists
if (Test-Path $moduleFolder) {
    # Loop through each file in the module folder
    Get-ChildItem -Path $moduleFolder -Filter *.bas | ForEach-Object {
        # Import the module into the workbook
        $wb.VBProject.VBComponents.Import($_.FullName)
    }
} else {
    Write-Host "Module folder not found: $moduleFolder"
    exit 1
}

# Save the workbook
$wb.Save()

# Close the workbook
$wb.Close()

# Quit Excel
$excelApp.Quit()

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
Remove-Variable -Name wb, excelApp

Write-Host "Script completed successfully."