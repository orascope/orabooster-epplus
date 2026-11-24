# Checks the filters of a workbook against Excel: for every sheet with an autofilter (of the sheet or of its first table)
# it lists the hidden rows as the file has them, lets Excel apply the filter again, and lists them again.
#   powershell -ExecutionPolicy Bypass -File filters-in-excel.ps1 -Path <issue17_autofilter.xlsx from "SaveBench issues">
# "same" means Excel hides the rows the file hides. Also printed: which columns Excel sees a filter on, and the sort.
# Excel runs hidden; only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path)

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$different = 0
function Get-HiddenRows($ws, $from, $to) {
    $h = @()
    for ($r = $from; $r -le $to; $r++) { if ($ws.Rows.Item($r).Hidden) { $h += $r } }
    return ($h -join ',')
}
try {
    $wb = $xl.Workbooks.Open($Path)
    "opened: $($wb.Worksheets.Count) sheets"
    foreach ($ws in $wb.Worksheets) {
        $af = $null; $kind = 'sheet'
        if ($ws.ListObjects.Count -gt 0) { $af = $ws.ListObjects.Item(1).AutoFilter; $kind = 'table' }
        elseif ($ws.AutoFilterMode) { $af = $ws.AutoFilter }
        if ($null -eq $af) { "{0,-24} no autofilter" -f $ws.Name; continue }
        $range = $af.Range
        $first = $range.Row + 1; $last = $range.Row + $range.Rows.Count - 1
        $on = @()
        for ($i = 1; $i -le $af.Filters.Count; $i++) { if ($af.Filters.Item($i).On) { $on += $i } }
        $sort = ''
        try { if ($af.Sort.SortFields.Count -gt 0) { $f = $af.Sort.SortFields.Item(1); $sort = " sort " + $f.Key.Address($false, $false) + $(if ($f.Order -eq 2) { ' desc' } else { ' asc' }) } } catch { }
        $col1 = @(); for ($r = $first; $r -le $last; $r++) { $col1 += [string]$ws.Cells.Item($r, $range.Column).Text }
        $fileHidden = Get-HiddenRows $ws $first $last
        $af.ApplyFilter()
        $excelHidden = Get-HiddenRows $ws $first $last
        $same = $fileHidden -eq $excelHidden
        if (-not $same) { $different++ }
        "{0,-24} {1,-5} filters on [{2}]{3}  hidden [{4}]  {5}" -f $ws.Name, $kind, ($on -join ','), $sort, $fileHidden, $(if ($same) { 'same' } else { "DIFFERENT, Excel hides [$excelHidden]" })
        if ($sort -ne '') { "{0,-24}       first column: {1}" -f '', ($col1 -join ' ') }
    }
    $wb.Close($false)
} catch {
    "ERROR $($_.Exception.Message)"
    $different = -1
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
"sheets where Excel hides other rows: $different"
exit $different
