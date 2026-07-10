$ErrorActionPreference = 'Continue'
Write-Host ("PSVersion: " + $PSVersionTable.PSVersion)
$cs = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot 'enc.cs')
Write-Host ("C# length: " + $cs.Length)
try {
  Add-Type -TypeDefinition $cs -ReferencedAssemblies @('System.dll','System.Core.dll','System.Drawing.dll') -ErrorAction Stop
  Write-Host "*** Add-Type SUCCESS ***"
  $g = New-Object QRCoder.QRCodeGenerator
  $d = $g.CreateQrCode("Q1|deadbeef|0|1|SGVsbG8gV29ybGQ=", [QRCoder.QRCodeGenerator+ECCLevel]::M, $false, $false, [QRCoder.QRCodeGenerator+EciMode]::Default, -1)
  Write-Host ("QR modules: " + $d.ModuleMatrix.Count)
  Write-Host "*** ENCODER OK ***"
} catch {
  Write-Host "*** Add-Type FAILED ***"
  Write-Host ("Exception: " + $_.Exception.Message)
  Write-Host "--- errors ---"
  $Error | Select-Object -First 40 | ForEach-Object { Write-Host $_.ToString() }
  exit 1
}
