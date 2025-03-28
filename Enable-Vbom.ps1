function Enable-VBOM ($App) {
  Try {
    # Step 1: Check if the application registry key exists
    $AppKeyPath = "Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer"
    if (-not (Test-Path $AppKeyPath)) {
      Write-Output "Error: The registry path '$AppKeyPath' does not exist."
      return
    }

    # Step 2: Retrieve the current version
    $CurVer = Get-ItemProperty -Path $AppKeyPath -ErrorAction Stop
    $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    # Step 3: Check if the Office version registry key exists
    $OfficePath = "HKCU:\Software\Microsoft\Office"
    if (-not (Test-Path $OfficePath)) {
        Write-Output "Error: The registry path '$OfficePath' does not exist."
        return
      }

    $OfficeKeyPath = "HKCU:\Software\Microsoft\Office\$Version"
    if (-not (Test-Path $OfficeKeyPath)) {
      Write-Output "Error: The registry path '$OfficeKeyPath' does not exist."
      return
    }

    # Step 4: Check if the application-specific key exists
    $AppSecurityPath = "$OfficeKeyPath\$App\Security"
    if (-not (Test-Path $AppSecurityPath)) {
      Write-Output "Error: The registry path '$AppSecurityPath' does not exist."
      return
    }

    # Step 5: Set the AccessVBOM property
    Set-ItemProperty -Path $AppSecurityPath -Name AccessVBOM -Value 1 -ErrorAction Stop
    Write-Output "Successfully enabled access to VBA project object model for $App."
  } Catch {
    Write-Output "Failed to enable access to VBA project object model for $App."
    Write-Output "Error: $($_.Exception.Message)"
    Write-Output "StackTrace: $($_.Exception.StackTrace)"
  }
}

Enable-VBOM "Excel"