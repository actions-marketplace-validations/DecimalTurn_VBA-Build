# This script will list all the process that are part of the Office suite (and generic Office processes) and close them.

# It will also check if the processes are running and close them if they are.
# This is the code.

# Define an array of Office-related process names
$officeProcesses = @(
    "WINWORD",      # Microsoft Word
    "EXCEL",        # Microsoft Excel
    "POWERPNT",     # Microsoft PowerPoint
    "OUTLOOK",      # Microsoft Outlook
    "ONENOTE",      # Microsoft OneNote
    "MSACCESS",     # Microsoft Access
    "MSPUB",        # Microsoft Publisher
    "TEAMS",        # Microsoft Teams
    "lync",         # Skype for Business
    "ONENOTEM",     # OneNote quick notes
    "VISIO",        # Microsoft Visio
    "OfficeClickToRun", # Office Click-to-Run
    "OfficeC2RClient"   # Office Click-to-Run Client
)

# Function to check and close Office processes
function Close-OfficeProcesses {
    Write-Host "Checking for running Office processes..." -ForegroundColor Cyan
    
    $runningProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
    
    if ($runningProcesses.Count -eq 0) {
        Write-Host "No Office processes are currently running." -ForegroundColor Green
        return
    }
    
    Write-Host "Found $($runningProcesses.Count) running Office processes:" -ForegroundColor Yellow
    foreach ($process in $runningProcesses) {
        Write-Host "  - $($process.Name) (PID: $($process.Id))" -ForegroundColor Yellow
    }
    
    # Ask for confirmation
    $confirmation = Read-Host "`nDo you want to close these Office processes? (Y/N)"
    if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
        Write-Host "Operation cancelled." -ForegroundColor Cyan
        return
    }
    
    Write-Host "Closing Office processes..." -ForegroundColor Cyan
    foreach ($process in $runningProcesses) {
        try {
            $process | Stop-Process -Force -ErrorAction Stop
            Write-Host "  Successfully closed $($process.Name)." -ForegroundColor Green
        }
        catch {
            Write-Host "  Failed to close $($process.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Check if any processes remain running
    $remainingProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
    if ($remainingProcesses.Count -gt 0) {
        Write-Host "`nWARNING: $($remainingProcesses.Count) Office processes could not be closed:" -ForegroundColor Red
        foreach ($process in $remainingProcesses) {
            Write-Host "  - $($process.Name) (PID: $($process.Id))" -ForegroundColor Red
        }
    }
    else {
        Write-Host "`nAll Office processes have been successfully closed." -ForegroundColor Green
    }
}

# Execute the function
Close-OfficeProcesses