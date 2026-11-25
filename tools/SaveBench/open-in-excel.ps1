# Opens every xlsx of a folder in Excel (hidden, through COM) and reports whether it opens without a repair.
#   powershell -ExecutionPolicy Bypass -File open-in-excel.ps1 -Dir <folder made by "SaveBench issues">
# With alerts off, a file that Excel would have to repair fails to open, so "opened" means no repair.
# For every chart (on the first sheet and on chart sheets) it prints the chart type Excel gives it and its chart groups.
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
            # issue 35: the encrypted samples are opened with their password (UpdateLinks 0, ReadOnly, Format 5, Password)
            if ($f.Name -like 'issue35_encrypted*') { $wb = $xl.Workbooks.Open($f.FullName, 0, $true, 5, 'orabooster') }
            else { $wb = $xl.Workbooks.Open($f.FullName) }
            $ws = $wb.Worksheets.Item(1)
            if ($f.Name -like 'issue35_encrypted*') { "  rows Excel has in the encrypted workbook: {0}, A200 = {1}" -f $ws.UsedRange.Rows.Count, $ws.Range('A200').Text }
            "{0}: opened (tables {1}, charts {2}, comments {3})" -f $f.Name, $ws.ListObjects.Count, $ws.ChartObjects().Count, $ws.Comments.Count
            # Every chart: the type Excel gives it (XlChartType: 88-91 are the stock charts, 90 and 91 the ones with volume)
            # and, per chart group, the number of series and whether the high-low lines and up/down bars are on.
            $charts = @()
            for ($i = 1; $i -le $ws.ChartObjects().Count; $i++) { $charts += , @("chart $i", $ws.ChartObjects($i).Chart) }
            for ($i = 1; $i -le $wb.Charts.Count; $i++) { $charts += , @("chart sheet $($wb.Charts.Item($i).Name)", $wb.Charts.Item($i)) }
            foreach ($entry in $charts) {
                $ch = $entry[1]
                $groups = @()
                # the chart types of Excel 2016 (117 to 123) have no chart groups: the count of them is 0 or the call fails
                $groupCount = 0
                try { $groupCount = $ch.ChartGroups().Count } catch { }
                for ($g = 1; $g -le $groupCount; $g++) {
                    $cg = $ch.ChartGroups($g)
                    $hl = '-'; $ud = '-'
                    try { $hl = $cg.HasHiLoLines } catch { }
                    try { $ud = $cg.HasUpDownBars } catch { }
                    $groups += "group ${g}: axes $($cg.AxisGroup), series $($cg.SeriesCollection().Count), hiLoLines $hl, upDownBars $ud"
                }
                $seriesCount = '-'; $title = ''
                try { $seriesCount = $ch.SeriesCollection().Count } catch { }
                try { if ($ch.HasTitle) { $title = ", title [" + $ch.ChartTitle.Text + "]" } } catch { }
                "  {0}: type {1}, series {2}{3}; {4}" -f $entry[0], $ch.ChartType, $seriesCount, $title, ($groups -join '; ')
            }
            if ($f.Name -eq 'issue9_ignored_errors.xlsx') {
                # XlErrorChecks: 1 = xlEvaluateToError, 3 = xlNumberAsText
                $cells = @(@('A2', 3), @('A21', 3), @('B2', 3), @('B21', 3), @('C2', 3), @('D2', 1), @('D21', 1), @('E2', 1), @('E21', 1), @('A2', 1), @('D2', 3))
                foreach ($c in $cells) {
                    $e = $ws.Range($c[0]).Errors.Item($c[1])
                    "  {0} check {1}: fails={2} ignored={3}" -f $c[0], $c[1], $e.Value, $e.Ignore
                }
            }
            if ($f.Name -eq 'issue27_no_x_range.xlsx' -or $f.Name -eq 'issue29_scatter_loaded.xlsx') {
                # one sheet for each kind of chart: the type and the number of series of every chart on it
                # (issue 27: two charts, X range null and ""; issue 29: one loaded scatter chart with a series added)
                for ($s = 1; $s -le $wb.Worksheets.Count; $s++) {
                    $sheet = $wb.Worksheets.Item($s)
                    $parts = @()
                    for ($i = 1; $i -le $sheet.ChartObjects().Count; $i++) {
                        $ch = $sheet.ChartObjects($i).Chart
                        $part = "type $($ch.ChartType) series $($ch.SeriesCollection().Count)"
                        if ($f.Name -eq 'issue29_scatter_loaded.xlsx') {
                            # per series: the chart type Excel gives it (the added third series must have the type of the first two)
                            $each = @()
                            for ($k = 1; $k -le $ch.SeriesCollection().Count; $k++) { $each += $ch.SeriesCollection($k).ChartType }
                            $part += " [each: " + ($each -join ', ') + "]"
                        }
                        $parts += $part
                    }
                    "  sheet {0}: {1}" -f $sheet.Name, ($parts -join ' | ')
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
