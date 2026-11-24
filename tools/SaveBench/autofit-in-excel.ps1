# Compares the column widths saved in a workbook with the widths Excel gives the same columns when it autofits them.
#   powershell -ExecutionPolicy Bypass -File autofit-in-excel.ps1 -Path <xlsx made by "SaveBench autofitsample">
# Per column: the header text, the saved width, Excel's autofit width, and saved / Excel.
# A ratio under 1 means the saved column is too narrow (text is cut off or a number shows as ####).
# Excel runs hidden; only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path)

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$narrow = 0
try {
    $wb = $xl.Workbooks.Open($Path)
    $ws = $wb.Worksheets.Item(1)
    $cols = $ws.UsedRange.Columns.Count
    $saved = @()
    for ($c = 1; $c -le $cols; $c++) { $saved += [double]$ws.Columns.Item($c).ColumnWidth }
    $ws.UsedRange.Columns.AutoFit() | Out-Null
    $sum = 0.0
    for ($c = 1; $c -le $cols; $c++) {
        $fit = [double]$ws.Columns.Item($c).ColumnWidth
        $ratio = $saved[$c - 1] / $fit
        $sum += $ratio
        if ($ratio -lt 0.999) { $narrow++ }
        "{0,-36} saved {1,7:0.00}  excel {2,7:0.00}  ratio {3:0.000}{4}" -f [string]$ws.Cells.Item(1, $c).Text, $saved[$c - 1], $fit, $ratio, $(if ($ratio -lt 0.999) { '  NARROW' } else { '' })
    }
    "columns {0}, narrower than Excel {1}, mean ratio {2:0.000}" -f $cols, $narrow, ($sum / $cols)
    $wb.Close($false)
} catch {
    "ERROR $($_.Exception.Message)"
    $narrow = -1
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
exit $narrow
