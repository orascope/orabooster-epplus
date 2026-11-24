# Opens every xlsx of a folder in Excel (hidden, through COM) and reports whether it opens without a repair.
#   powershell -ExecutionPolicy Bypass -File open-in-excel.ps1 -Dir <folder made by "SaveBench issues">
# With alerts off, a file that Excel would have to repair fails to open, so "opened" means no repair.
# For issue9_ignored_errors.xlsx it also prints, per cell, whether the error check fails and whether it is ignored.
# Only the Excel process started here is closed; an Excel the user has open is left alone.
param([Parameter(Mandatory = $true)][string]$Dir)

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$failed = 0
try {
    foreach ($f in Get-ChildItem $Dir -Filter *.xlsx | Sort-Object Name) {
        try {
            $wb = $xl.Workbooks.Open($f.FullName)
            $ws = $wb.Worksheets.Item(1)
            "{0}: opened (tables {1}, charts {2}, comments {3})" -f $f.Name, $ws.ListObjects.Count, $ws.ChartObjects().Count, $ws.Comments.Count
            if ($f.Name -eq 'issue9_ignored_errors.xlsx') {
                # XlErrorChecks: 1 = xlEvaluateToError, 3 = xlNumberAsText
                $cells = @(@('A2', 3), @('A21', 3), @('B2', 3), @('B21', 3), @('C2', 3), @('D2', 1), @('D21', 1), @('E2', 1), @('E21', 1), @('A2', 1), @('D2', 3))
                foreach ($c in $cells) {
                    $e = $ws.Range($c[0]).Errors.Item($c[1])
                    "  {0} check {1}: fails={2} ignored={3}" -f $c[0], $c[1], $e.Value, $e.Ignore
                }
            }
            $wb.Close($false)
        } catch {
            $failed++
            "{0}: DID NOT OPEN ({1})" -f $f.Name, $_.Exception.Message
        }
    }
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
"files that did not open: $failed"
exit $failed
