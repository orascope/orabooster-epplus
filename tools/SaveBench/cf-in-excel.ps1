# Opens a workbook in Excel (hidden, COM) and prints, for every sheet, its conditional formats as Excel sees them (type,
# range, and for a data bar its options), and the DisplayFormat of some cells.
#   powershell -ExecutionPolicy Bypass -File cf-in-excel.ps1 -Path <xlsx> [-Cells "Refs!A8,Refs!A2"]
# For issue46_databars.xlsx of "SaveBench issues". A data bar is printed with: BarFillType (0 solid, 1 gradient), BarColor,
# BarBorder.Type (1 solid), the border colour, the negative fill and border colours, AxisPosition (0 automatic, 1 middle,
# 2 none), AxisColor, Direction (-5002 context, -5003 left to right, -5004 right to left), MinPoint and MaxPoint (type/value).
# Only the Excel process started here is closed.
param([Parameter(Mandatory = $true)][string]$Path, [string]$Cells = '')

function Get-Hex([double]$color) {
    $c = [int]$color
    "{0:X2}{1:X2}{2:X2}" -f ($c -band 0xFF), (($c -shr 8) -band 0xFF), (($c -shr 16) -band 0xFF)
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    try { $wb = $xl.Workbooks.Open($Path) } catch { "{0}: DID NOT OPEN" -f (Split-Path $Path -Leaf); return }
    "{0}: opened" -f (Split-Path $Path -Leaf)
    foreach ($ws in $wb.Worksheets) {
        $fcs = $ws.Cells.FormatConditions
        "  sheet {0}: {1} rules" -f $ws.Name, $fcs.Count
        for ($i = 1; $i -le $fcs.Count; $i++) {
            $fc = $fcs.Item($i)
            $line = "    {0}: type {1} on {2}" -f $i, $fc.Type, $fc.AppliesTo.Address($false, $false)
            if ($fc.Type -eq 4) {   # xlDatabar
                $line += " fill {0} color {1} border {2} borderColor {3} negColor {4} negBorder {5} axis {6} axisColor {7} direction {8} min {9}/{10} max {11}/{12}" -f `
                    $fc.BarFillType, (Get-Hex $fc.BarColor.Color), $fc.BarBorder.Type, $(if ($fc.BarBorder.Type -eq 1) { Get-Hex $fc.BarBorder.Color.Color } else { '-' }), `
                    (Get-Hex $fc.NegativeBarFormat.Color.Color), $(if ($fc.BarBorder.Type -eq 1) { Get-Hex $fc.NegativeBarFormat.BorderColor.Color } else { '-' }), `
                    $fc.AxisPosition, (Get-Hex $fc.AxisColor.Color), $fc.Direction, $fc.MinPoint.Type, $fc.MinPoint.Value, $fc.MaxPoint.Type, $fc.MaxPoint.Value
            } elseif ($fc.Type -eq 1 -or $fc.Type -eq 2) {
                $line += " formula [{0}]" -f $fc.Formula1
            }
            $line
        }
    }
    foreach ($c in @($Cells -split ',' | Where-Object { $_ })) {
        $parts = $c -split '!'
        $d = $wb.Worksheets.Item($parts[0]).Range($parts[1]).DisplayFormat
        $fill = if ($d.Interior.ColorIndex -eq -4142) { 'none' } else { Get-Hex $d.Interior.Color }
        "  {0}: fill {1} font {2} bold {3}" -f $c, $fill, (Get-Hex $d.Font.Color), $d.Font.Bold
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
