

function Enable-VBOM ($App) {
  Try {
    $CurVer = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer -ErrorAction Stop
    $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\$Version\$App\Security -Name AccessVBOM -Value 1 -ErrorAction Stop
  } Catch {
    Write-Output "Failed to enable access to VBA project object model for $App."
  }
}

Enable-VBOM "Excel"