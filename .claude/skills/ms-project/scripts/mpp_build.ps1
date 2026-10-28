<#
Bina fail .mpp baharu daripada JSON melalui MS Project COM (perlukan MS Project).
Import XML (MSPDI) tidak digunakan kerana MS Project membuang tempoh tugasan
yang tiada Start/Finish – COM menjadualkan semula dengan betul.

Contoh:
  & ".claude\skills\ms-project\scripts\mpp_build.ps1" -Json "program.json" -Out "program.mpp"

Format JSON:
{
  "title": "Recovery Programme – Fasa 1",
  "start": "2026-11-02",
  "working_days": [1,2,3,4,5,6],      # 1=Isnin ... 7=Ahad (lalai Isnin–Sabtu)
  "holidays": ["2026-12-25"],         # pilihan
  "status_date": "2026-11-15",        # pilihan
  "tasks": [
    {"id": 1, "name": "FASA 1", "level": 1},
    {"id": 2, "name": "Roboh siling", "level": 2, "duration": 10, "cost": 25000},
    {"id": 3, "name": "Pasang siling", "level": 2, "duration": 15,
     "predecessors": [{"id": 2, "type": "FS", "lag": 2}]},
    {"id": 4, "name": "Serah Fasa 1", "level": 2, "duration": 0, "predecessors": [3],
     "start_no_earlier": "2026-12-01", "notes": "Serahan kepada UTM"}
  ]
}
- id unik > 0, ikut susunan paparan. Tugasan ringkasan = tugasan yang diikuti aras lebih dalam.
- duration dalam hari bekerja; 0 = milestone. lag dalam hari (negatif = lead).
- Waktu kerja 08:00–12:00, 13:00–17:00 (8 jam sehari).
#>
param(
  [Parameter(Mandatory=$true)][string]$Json,
  [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$Json = (Resolve-Path -LiteralPath $Json).Path
$Out  = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
$spec = [System.IO.File]::ReadAllText($Json, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$types = @{ FF = 0; FS = 1; SF = 2; SS = 3 }   # PjTaskLinkType
$wd = if ($spec.working_days) { @($spec.working_days) } else { 1..6 }

# semak id & rujukan predecessor sebelum buka MS Project
$ids = @($spec.tasks | ForEach-Object { [int]$_.id })
if (($ids | Select-Object -Unique).Count -ne $ids.Count) { throw "id tugasan berulang" }
foreach ($t in $spec.tasks) { foreach ($p in @($t.predecessors)) {
  if ($null -eq $p) { continue }
  $pid_ = if ($p -is [int] -or $p -is [long]) { [int]$p } else { [int]$p.id }
  if ($ids -notcontains $pid_) { throw "Tugasan $($t.id) merujuk predecessor $pid_ yang tiada" } } }

$app = New-Object -ComObject MSProject.Application
try {
  $app.Visible = $false
  $app.DisplayAlerts = $false
  [void]$app.FileNew()
  $proj = $app.ActiveProject
  $proj.Title = [string]$spec.title
  $proj.ProjectStart = [datetime]::Parse($spec.start).AddHours(8)
  $proj.HoursPerWeek = 8 * $wd.Count
  $proj.DaysPerMonth = 4 * $wd.Count + [math]::Floor($wd.Count / 3)

  # kalendar projek: PjWeekday 1=Ahad ... 7=Sabtu; input 1=Isnin ... 7=Ahad
  $cal = $proj.Calendar
  foreach ($iso in 1..7) {
    $day = $cal.WeekDays.Item(($iso % 7) + 1)
    if ($wd -contains $iso) {
      $day.Working = $true
      $day.Shift1.Start = '08:00'; $day.Shift1.Finish = '12:00'
      $day.Shift2.Start = '13:00'; $day.Shift2.Finish = '17:00'
    } else { $day.Working = $false }
  }
  foreach ($h in @($spec.holidays)) {
    if (-not $h) { continue }
    $hd = [datetime]::Parse($h)
    [void]$cal.Exceptions.Add(1, $hd, $hd, 1, "Cuti $h")   # pjDaily, sekali
  }

  # tugasan
  $minPerDay = 480
  $byId = @{}
  $list = @($spec.tasks)
  for ($i = 0; $i -lt $list.Count; $i++) {
    $s = $list[$i]
    $t = $proj.Tasks.Add([string]$s.name)
    $t.Manual = $false
    $lvl = if ($s.level) { [int]$s.level } else { 1 }
    while ($t.OutlineLevel -lt $lvl) { $t.OutlineIndent() }
    $byId[[int]$s.id] = $t
  }
  for ($i = 0; $i -lt $list.Count; $i++) {
    $s = $list[$i]; $t = $byId[[int]$s.id]
    $lvl = if ($s.level) { [int]$s.level } else { 1 }
    $nextLvl = if ($i + 1 -lt $list.Count -and $list[$i + 1].level) { [int]$list[$i + 1].level } else { 0 }
    $summary = $nextLvl -gt $lvl
    if ($s.notes) { $t.Notes = [string]$s.notes }
    if ($summary) {
      if ($s.duration -or $s.predecessors) { Write-Warning "Ringkasan $($s.id) '$($s.name)': duration/predecessor diabaikan" }
      continue
    }
    $dur = if ($null -ne $s.duration) { [double]$s.duration } else { 1 }
    $t.Duration = [int]($dur * $minPerDay)
    if ($s.cost) { $t.FixedCost = [double]$s.cost }
    if ($s.start_no_earlier) {
      $t.ConstraintType = 4    # pjSNET
      $t.ConstraintDate = [datetime]::Parse($s.start_no_earlier).AddHours(8)
    }
    foreach ($p in @($s.predecessors)) {
      if ($null -eq $p) { continue }
      if ($p -is [int] -or $p -is [long]) { $p = [pscustomobject]@{ id = $p } }
      $type = if ($p.type) { $types[([string]$p.type).ToUpper()] } else { 1 }
      $lag  = if ($p.lag) { "{0}d" -f $p.lag } else { "0d" }
      $t.LinkPredecessors($byId[[int]$p.id], $type, $lag)
    }
    if ($s.percent_complete) { $t.PercentComplete = [int]$s.percent_complete }
  }
  if ($spec.status_date) { $proj.StatusDate = [datetime]::Parse($spec.status_date).AddHours(17) }

  if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
  $fmt = if ([System.IO.Path]::GetExtension($Out).ToLower() -eq '.xml') { 'MSProject.XML' } else { 'MSProject.MPP' }
  [void][System.__ComObject].InvokeMember('FileSaveAs', [System.Reflection.BindingFlags]::InvokeMethod,
    $null, $app, [object[]]@($Out, $fmt), $null, $null, [string[]]@('Name', 'FormatID'))
  Write-Output ("OK: {0} ({1} tugasan, siap {2:yyyy-MM-dd})" -f $Out, $list.Count, $proj.ProjectFinish)
  [void]$app.FileCloseEx(0)
}
finally {
  try { [void]$app.Quit(0) } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
