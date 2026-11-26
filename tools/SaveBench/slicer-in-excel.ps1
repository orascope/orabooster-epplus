# Opens a workbook in Excel (hidden, through COM) and prints its slicer caches and slicers as Excel reads them: the cache name,
# the source (table and column), and for each slicer its name, caption, style, columns, sheet and position, the items that are
# selected, and how many rows of the table are visible.
#   powershell -ExecutionPolicy Bypass -File slicer-in-excel.ps1 -Path <xlsx>
# For issue48_slicers.xlsx of "SaveBench issues". Only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path)
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    try { $wb = $xl.Workbooks.Open($Path) } catch { "{0}: DID NOT OPEN" -f (Split-Path $Path -Leaf); exit 1 }
    "{0}: opened, {1} slicer caches" -f (Split-Path $Path -Leaf), $wb.SlicerCaches.Count
    foreach ($cache in $wb.SlicerCaches) {
        $source = ''
        try { $source = $cache.ListObject.Name + '[' + $cache.SourceName + ']' } catch { $source = $cache.SourceName }
        $selected = @()
        foreach ($item in $cache.SlicerItems) { if ($item.Selected) { $selected += $item.Name } }
        "  cache {0}: {1}, {2} items, selected: {3}" -f $cache.Name, $source, $cache.SlicerItems.Count, ($selected -join ', ')
        foreach ($s in $cache.Slicers) {
            "    slicer {0}: caption [{1}] style {2} columns {3} on {4} at {5:0},{6:0} size {7:0}x{8:0}" -f $s.Name, $s.Caption, $s.Style.Name, $s.NumberOfColumns, $s.Parent.Name, $s.Left, $s.Top, $s.Width, $s.Height
        }
        try {
            $lo = $cache.ListObject
            $visible = 0
            foreach ($row in $lo.DataBodyRange.Rows) { if (-not $row.EntireRow.Hidden) { $visible++ } }
            "    table {0}: {1} of {2} rows visible" -f $lo.Name, $visible, $lo.DataBodyRange.Rows.Count
        } catch { }
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
