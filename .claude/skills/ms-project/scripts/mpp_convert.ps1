<#
Tukar fail MS Project melalui COM (perlukan MS Project dipasang).
  .mpp -> .xml (MSPDI)   : untuk dibaca oleh mspdi.py
  .xml -> .mpp           : serah fail yang dijana oleh build_mspdi.py
  .mpp -> .xlsx          : eksport jadual tugasan ke Excel
Format sasaran ikut sambungan fail -Out.

Contoh (dalam alat PowerShell):
  & ".claude\skills\ms-project\scripts\mpp_convert.ps1" -In "a.mpp" -Out "a.xml"
#>
param(
  [Parameter(Mandatory=$true)][string]$In,
  [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$In  = (Resolve-Path -LiteralPath $In).Path
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
if ($In -eq $Out) { throw "Fail -Out sama dengan -In. Guna nama lain." }

$fmt = switch ([System.IO.Path]::GetExtension($Out).ToLower()) {
  '.xml'  { 'MSProject.XML' }
  '.mpp'  { 'MSProject.MPP' }
  '.xlsx' { 'MSProject.XLSX' }
  default { throw "Sambungan -Out tidak disokong: $Out (guna .xml / .mpp / .xlsx)" }
}

$app = New-Object -ComObject MSProject.Application
try {
  $app.Visible = $false
  $app.DisplayAlerts = $false
  # ReadOnly supaya fail asal tidak dikunci/diubah
  [void]$app.FileOpenEx($In, $true)
  if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
  # FileSaveAs ada banyak parameter pilihan; guna argumen bernama melalui IDispatch
  [void][System.__ComObject].InvokeMember('FileSaveAs', [System.Reflection.BindingFlags]::InvokeMethod,
    $null, $app, [object[]]@($Out, $fmt), $null, $null, [string[]]@('Name', 'FormatID'))
  [void]$app.FileCloseEx(0)   # pjDoNotSave
  Write-Output "OK: $Out"
}
finally {
  try { [void]$app.Quit(0) } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
