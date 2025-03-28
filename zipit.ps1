# This is a a powershell script that used 7zip to compress files and folders
# The files and folders to compress are located in the src/XMLSource folder
# It contains files such as src/skeleton.xlsm/XMLsource/[Content_Types].xml and folders such as src/skeleton.xlsm/XMLsource/_rels
# The script will compress the files and folders into a zip file located in the src/skeleton.xlsm/XMLsource folder
# The zip file will be named skeleton.zip
# The script will use 7zip to compress the files and folders
# The script will use the 7zip command line interface to compress the files and folders

# This script uses 7-Zip to compress files and folders in the src/XMLSource directory into a zip file named skeleton.zip.

Write-Host "Staring the compression process..."

# Define the source folder and the output zip file
$sourceFolder = "src/skeleton.xlsm/XMLSource"
$outputZipFile = "src/skeleton.xlsm/XMLOutput/skeleton.zip"

# Path to the 7-Zip executable
$sevenZipPath = "7z"  # Assumes 7-Zip is in the system PATH. Adjust if necessary.

# Check if the source folder exists
if (-not (Test-Path $sourceFolder)) {
    Write-Host "Source folder not found: $sourceFolder"
    exit 1
}

# Ensure the destination directory exists
$outputDir = Split-Path -Path $outputZipFile
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Compress the files and folders using 7-Zip
Write-Host "Compressing files in $sourceFolder to $outputZipFile..."
$command = "$sevenZipPath a -tzip `"$outputZipFile`" `"$sourceFolder\*`""
$exitCode = Invoke-Expression $command

# Check if the compression was successful
if ($exitCode -ne 0) {
    Write-Host "Error: Compression failed with exit code $exitCode"
    exit 1
}

Write-Host "Compression completed successfully. Zip file created at: $outputZipFile"