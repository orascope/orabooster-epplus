# Reads the format of the charts of issue19_chart_format.xlsx back from Excel (one chart per sheet).
#   powershell -ExecutionPolicy Bypass -File chart-format-in-excel.ps1 -Path <issue19_chart_format.xlsx from "SaveBench issues">
# Per sheet: the fill and the line of series 1, the fill (or marker fill) of its third point, and whether series 1 and 2
# have error bars. The colours are printed as RRGGBB. What the sample sets: fill C00000, line 0070C0, point 00B050.
# Excel runs hidden; only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path)

function Format-Rgb($bgr) {
    # Excel gives a colour as blue * 65536 + green * 256 + red
    $v = [int]$bgr
    return '{0:X2}{1:X2}{2:X2}' -f ($v -band 0xFF), (($v -shr 8) -band 0xFF), (($v -shr 16) -band 0xFF)
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$failed = 0
try {
    $wb = $xl.Workbooks.Open($Path)
    "opened: $($wb.Worksheets.Count) sheets"
    foreach ($ws in $wb.Worksheets) {
        if ($ws.ChartObjects().Count -eq 0) { continue }
        $ch = $ws.ChartObjects(1).Chart
        $s = $ch.SeriesCollection(1)
        $fill = '-'; $line = '-'; $point = '-'; $marker = '-'; $bars1 = '-'; $bars2 = '-'
        try { if ($s.Format.Fill.Visible) { $fill = Format-Rgb $s.Format.Fill.ForeColor.RGB } } catch { }
        try { if ($s.Format.Line.Visible) { $line = Format-Rgb $s.Format.Line.ForeColor.RGB } } catch { }
        try { $pt = $s.Points(3); if ($pt.Format.Fill.Visible) { $point = Format-Rgb $pt.Format.Fill.ForeColor.RGB } } catch { }
        try { $marker = Format-Rgb $s.Points(3).MarkerBackgroundColor } catch { }
        try { $bars1 = $s.HasErrorBars } catch { }
        try { $bars2 = $ch.SeriesCollection(2).HasErrorBars } catch { }
        "{0,-20} type {1,6}  series {2}  fill {3,-6} line {4,-6}  point 3: fill {5,-6} marker {6,-6}  error bars {7}/{8}" -f $ws.Name, $ch.ChartType, $ch.SeriesCollection().Count, $fill, $line, $point, $marker, $bars1, $bars2
    }
    $wb.Close($false)
} catch {
    "ERROR $($_.Exception.Message)"
    $failed = 1
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
exit $failed
