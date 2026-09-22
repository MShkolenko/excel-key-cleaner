# Measures CleanKeys wall time vs. selection size, to check the O(n^2) claim from issue #1
# (Collection.Item(i) by positional index is a known VBA gotcha: it walks the list each call,
# so a loop indexing 1..N is O(n^2) if the Collection holds all N planned writes).
# Each size gets a FRESH sheet, ALL cells dirty, so plannedCells.Count == N every time -
# the worst case the issue describes, not a diluted incremental one.
# Imports the REAL dist/Cleaning.bas + CleaningForm unmodified except the .Show stub, same
# approach as tests/gen.py, so the timing reflects the actual shipped code.
param(
    [Parameter(Mandatory)][string]$DistDir,
    [int[]]$Sizes = @(2000, 8000, 20000, 40000)
)
$ErrorActionPreference = 'Stop'
$cyrR = [string][char]0x0420

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()

    $src = [IO.File]::ReadAllText((Join-Path $DistDir 'Cleaning.bas'), [Text.Encoding]::GetEncoding(1251))
    $src = $src.Replace('Attribute VB_Name = "Cleaning"', 'Attribute VB_Name = "Cleaning_Speed"')
    $src = $src.Replace('CleaningForm.Show', 'CleaningForm.Tag = "OK"')
    $tmpBas = Join-Path $env:TEMP 'Cleaning_Speed.bas'
    [IO.File]::WriteAllText($tmpBas, $src, [Text.Encoding]::GetEncoding(1251))
    [void]$wb.VBProject.VBComponents.Import($tmpBas)
    [void]$wb.VBProject.VBComponents.Import((Join-Path $DistDir 'CleaningForm.frm'))
    $mod = $wb.VBProject.VBComponents.Item('Cleaning_Speed').CodeModule
    # MsgBox is shadowed exactly as in gen.py: a headless run must never put a dialog on the
    # user's screen (Application.Visible = False does not hide a modal MsgBox).
    $mod.InsertLines($mod.CountOfLines + 1, @"
Private Function MsgBox(Prompt As Variant, Optional Buttons As Variant, Optional Title As Variant) As Long
End Function
Public Function RunNow(c1 As Boolean, c2 As Boolean, c3 As Boolean, c4 As Boolean) As String
    CleaningForm.CheckBox1.Value = c1
    CleaningForm.CheckBox2.Value = c2
    CleaningForm.CheckBox3.Value = c3
    CleaningForm.CheckBox4.Value = c4
    On Error Resume Next
    CleaningForm.Controls("CheckBox5").Value = False   ' absent before v3.0
    On Error GoTo 0
    CleanKeys
End Function
"@)

    $results = @()
    $prevN = 0; $prevT = 0.0
    foreach ($n in $Sizes) {
        $ws = $wb.Worksheets.Add()
        $ws.Name = "N$n"
        $ws.Activate()
        # write in chunks of 5000 (COM array write limit / comfort margin)
        $chunk = 5000
        for ($base = 0; $base -lt $n; $base += $chunk) {
            $len = [Math]::Min($chunk, $n - $base)
            $arr = New-Object 'object[,]' $len, 1
            for ($i = 0; $i -lt $len; $i++) { $arr[$i, 0] = "AB$($base + $i) $cyrR C  D" }
            $r = $ws.Range($ws.Cells.Item($base + 1, 2), $ws.Cells.Item($base + $len, 2))
            $r.NumberFormat = '@'
            $r.Value2 = $arr
        }
        $ws.Range($ws.Cells.Item(1, 2), $ws.Cells.Item($n, 2)).Select() | Out-Null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $excel.Run('Cleaning_Speed.RunNow', $true, $true, $false, $true) | Out-Null
        $sw.Stop()
        $t = $sw.Elapsed.TotalSeconds
        $ratioN = if ($prevN -gt 0) { [double]$n / $prevN } else { 0 }
        $ratioT = if ($prevT -gt 0) { $t / $prevT } else { 0 }
        $results += [pscustomobject]@{ N = $n; Seconds = [math]::Round($t, 3); RatioN = [math]::Round($ratioN, 2); RatioT = [math]::Round($ratioT, 2) }
        $prevN = $n; $prevT = $t
    }
    $results | Format-Table -AutoSize | Out-String | Write-Output
    Write-Output "If RatioT stays close to RatioN -> linear. If RatioT grows faster (e.g. RatioN=2 but RatioT=4) -> super-linear/quadratic."
}
finally {
    if ($wb)    { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
