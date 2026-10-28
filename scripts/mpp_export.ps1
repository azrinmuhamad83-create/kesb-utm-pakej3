# Tukar fail MS Project (.mpp) ke XML (format MSPDI) guna MS Project yang dipasang.
# Guna:  powershell -ExecutionPolicy Bypass -File scripts/mpp_export.ps1 "docs/work program/fail.mpp"
# Hasil: output/programme/<nama>.xml
param([Parameter(Mandatory = $true)][string]$Mpp)

$src = (Resolve-Path $Mpp).Path
$outDir = Join-Path (Get-Location) "output\programme"
New-Item -ItemType Directory -Force $outDir | Out-Null
$xml = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($src) + ".xml")

$app = New-Object -ComObject MSProject.Application
try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    # ReadOnly = true supaya fail asal tidak diubah
    $app.FileOpenEx($src, $true) | Out-Null
    # FileSaveAs dengan parameter bernama Name + FormatID (PowerShell tak dapat pilih overload COM secara biasa)
    [System.__ComObject].InvokeMember("FileSaveAs", [Reflection.BindingFlags]::InvokeMethod, $null, $app,
        @($xml, "MSProject.XML"), $null, $null, @("Name", "FormatID")) | Out-Null
    $app.FileCloseEx(0) | Out-Null   # 0 = pjDoNotSave
    Write-Output "XML: $xml"
}
finally {
    $app.Quit(0)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
