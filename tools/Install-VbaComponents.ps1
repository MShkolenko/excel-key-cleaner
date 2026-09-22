# Replace VBA components in a workbook from native export files (.bas/.cls/.frm+.frx).
# Attaches to the running Excel when the workbook is already open there (PERSONAL.XLSB is loaded
# by every GUI instance, and a write from a second instance is lost when the GUI saves on exit);
# otherwise opens its own instance. Backs the workbook file up first.
param(
    [Parameter(Mandatory)][string]$Path,           # workbook, e.g. ...\XLSTART\PERSONAL.XLSB
    [Parameter(Mandatory)][string]$Files,          # component files to import, ';'-separated; a .frm needs its .frx beside it
    [Parameter(Mandatory)][string]$BackupDir
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$bak = Join-Path $BackupDir ((Split-Path $Path -Leaf) + ".$stamp.bak")
Copy-Item -LiteralPath $Path -Destination $bak
Write-Output "backup: $bak"

$excel = $null; $wb = $null; $attached = $false
try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($w in $excel.Workbooks) { if ($w.FullName -ieq $Path) { $wb = $w } }
        if ($wb) { $attached = $true; Write-Output "attached to running Excel (pid $((Get-Process EXCEL | Sort-Object StartTime | Select-Object -First 1).Id)), workbook already loaded" }
        else { $excel = $null }
    } catch { $excel = $null }
    if (-not $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Path)
        if ($wb.ReadOnly) { throw "'$Path' opened read-only - it is open elsewhere; close that Excel first" }
        Write-Output "opened own instance"
    }

    $proj = $wb.VBProject
    # one ';'-joined string: array parameters do not survive `powershell -File` invocation intact
    foreach ($f in ($Files -split ';')) {
        $name = [IO.Path]::GetFileNameWithoutExtension($f)
        foreach ($c in @($proj.VBComponents)) {
            if ($c.Name -eq $name) { $proj.VBComponents.Remove($c); Write-Output "removed  $name" }
        }
        $new = $proj.VBComponents.Import($f)
        Write-Output ("imported {0}  type={1} lines={2}" -f $new.Name, $new.Type, $new.CodeModule.CountOfLines)
    }
    $wb.Save()
    Write-Output "saved $($wb.FullName)"
}
finally {
    if (-not $attached) {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit() }
    }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
