<#
Kemas kini progres dalam fail .mpp melalui MS Project COM, simpan ke fail BAHARU.
Fail asal tidak diubah.

CSV (UTF-8, baris pertama tajuk). Lajur yang dikenali – isi yang perlu sahaja:
  UID            UniqueID tugasan (utamakan; atau guna ID)
  ID             ID baris tugasan
  ActualStart    YYYY-MM-DD
  ActualFinish   YYYY-MM-DD
  PercentComplete 0-100
  Notes          catatan (ditambah ke nota sedia ada)

Contoh:
  & ".claude\skills\ms-project\scripts\mpp_update.ps1" -In "a.mpp" -Csv "progres.csv" -StatusDate 2026-11-15 -Out "a_upd.mpp"
#>
param(
  [Parameter(Mandatory=$true)][string]$In,
  [Parameter(Mandatory=$true)][string]$Csv,
  [Parameter(Mandatory=$true)][string]$Out,
  [string]$StatusDate
)
$ErrorActionPreference = 'Stop'
$In  = (Resolve-Path -LiteralPath $In).Path
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
if ($In -eq $Out) { throw "Fail -Out mesti berbeza dari -In (fail asal tidak ditulis ganti)." }
$rows = Import-Csv -LiteralPath $Csv -Encoding UTF8
$fmt = if ([System.IO.Path]::GetExtension($Out).ToLower() -eq '.xml') { 'MSProject.XML' } else { 'MSProject.MPP' }

$app = New-Object -ComObject MSProject.Application
try {
  $app.Visible = $false
  $app.DisplayAlerts = $false
  [void]$app.FileOpenEx($In, $true)
  $proj = $app.ActiveProject
  if ($StatusDate) { $proj.StatusDate = [datetime]::Parse($StatusDate) }
  # Tasks.UniqueID(n) tidak boleh dipanggil dari PowerShell – bina peta sendiri
  $byUid = @{}; $byId = @{}
  foreach ($x in $proj.Tasks) { if ($x) { $byUid[[int]$x.UniqueID] = $x; $byId[[int]$x.ID] = $x } }
  $n = 0
  foreach ($r in $rows) {
    $t = $null
    if ($r.UID) { $t = $byUid[[int]$r.UID] }
    elseif ($r.ID) { $t = $byId[[int]$r.ID] }
    if (-not $t) { Write-Warning "Tugasan tidak dijumpai: UID=$($r.UID) ID=$($r.ID)"; continue }
    if ($r.ActualStart)     { $t.ActualStart = [datetime]::Parse($r.ActualStart).AddHours(8) }
    if ($r.PercentComplete) { $t.PercentComplete = [int]$r.PercentComplete }
    if ($r.ActualFinish)    { $t.ActualFinish = [datetime]::Parse($r.ActualFinish).AddHours(17) }
    if ($r.Notes)           { $t.Notes = ((@($t.Notes, $r.Notes) | Where-Object { $_ }) -join "`r`n") }
    $n++
    Write-Output ("{0,4} {1} -> {2}%" -f $t.ID, $t.Name, $t.PercentComplete)
  }
  if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
  [void][System.__ComObject].InvokeMember('FileSaveAs', [System.Reflection.BindingFlags]::InvokeMethod,
    $null, $app, [object[]]@($Out, $fmt), $null, $null, [string[]]@('Name', 'FormatID'))
  [void]$app.FileCloseEx(0)
  Write-Output "OK: $n tugasan dikemas kini -> $Out"
}
finally {
  try { [void]$app.Quit(0) } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
