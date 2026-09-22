# Exports named VBA components using the native VBIDE .Export() method -
# produces real, re-importable .bas/.cls/.frm(+.frx) files, unlike Export-VBA.ps1
# (which dumps CodeModule.Lines as plain text for diffing only, no .frx).
param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string[]]$Component,
    [Parameter(Mandatory)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$extByType = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false
    $wb = $excel.Workbooks.Open($Path, 0, $true)
    $proj = $wb.VBProject

    foreach ($name in $Component) {
        $c = $proj.VBComponents.Item($name)
        $ext = if ($extByType.ContainsKey([int]$c.Type)) { $extByType[[int]$c.Type] } else { '.bas' }
        $out = Join-Path $OutDir ($name + $ext)
        $c.Export($out)
        Write-Output "exported  $name -> $out"
    }
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
