# Opens a workbook in Excel (hidden, through COM) and prints every pivot table as Excel shows it after the refresh on load:
# its calculated fields, its data fields, its filters, and the cells of the table (labels and values, row by row).
#   powershell -ExecutionPolicy Bypass -File pivot-in-excel.ps1 -Path <xlsx>
# For the issue54_* (and later pivot) samples of "SaveBench issues". Only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path)
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    try { $wb = $xl.Workbooks.Open($Path) } catch { "{0}: DID NOT OPEN" -f (Split-Path $Path -Leaf); exit 1 }
    "{0}: opened" -f (Split-Path $Path -Leaf)
    foreach ($ws in $wb.Worksheets) {
        foreach ($pt in $ws.PivotTables()) {
            "  pivot table {0} on {1} at {2}" -f $pt.Name, $ws.Name, $pt.TableRange2.Address($false, $false)
            foreach ($f in $pt.CalculatedFields()) { "    calculated field {0} {1}" -f $f.Name, $f.Formula }
            $names = @(); foreach ($d in $pt.DataFields()) { $names += $d.Name }
            "    data fields: " + ($names -join ', ')
            foreach ($f in $pt.PivotFields()) {
                try { foreach ($flt in $f.PivotFilters) { "    filter on {0}: type {1} value [{2}]" -f $f.Name, $flt.FilterType, $flt.Value1 } } catch { }
            }
            $range = $pt.TableRange2
            for ($r = 1; $r -le $range.Rows.Count; $r++) {
                $cells = @()
                for ($c = 1; $c -le $range.Columns.Count; $c++) { $cells += [string]$range.Cells.Item($r, $c).Text }
                "    | " + ($cells -join ' | ')
            }
        }
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
