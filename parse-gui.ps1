$src = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'receiver-gui.ps1'), [System.Text.Encoding]::UTF8)
$errs = $null
[System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$null, [ref]$errs) | Out-Null
if ($errs.Count -eq 0) {
  Write-Host "*** PARSE OK (no syntax errors, UTF-8) ***"
} else {
  Write-Host ("*** PARSE ERRORS: " + $errs.Count + " ***")
  $errs | Select-Object -First 25 | ForEach-Object { Write-Host ("line " + $_.Extent.StartLineNumber + ": " + $_.Message) }
  exit 1
}
