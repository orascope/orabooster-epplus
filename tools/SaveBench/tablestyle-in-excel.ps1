# Opens a workbook in Excel (hidden, through COM) and prints, for the first table of the first sheet, the name of its style
# and what Excel shows in some of its cells: the fill, the font colour, bold and italic (Range.DisplayFormat, which has the
# format a table style gives).
#   powershell -ExecutionPolicy Bypass -File tablestyle-in-excel.ps1 -Path <xlsx made by "SaveBench issues", issue38_*>
# Only the Excel process started here is closed; an Excel the user has open is left alone.
param([Parameter(Mandatory = $true)][string]$Path)

function Get-Hex([double]$color) {
    $c = [int]$color   # Excel gives blue*65536 + green*256 + red
    "{0:X2}{1:X2}{2:X2}" -f ($c -band 0xFF), (($c -shr 8) -band 0xFF), (($c -shr 16) -band 0xFF)
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    $ws = $wb.Worksheets.Item(1)
    $table = $ws.ListObjects.Item(1)
    $range = $table.Range
    "{0}: opened; table {1}, style [{2}], rows {3}, columns {4}" -f (Split-Path $Path -Leaf), $table.Name, $table.TableStyle.Name, $range.Rows.Count, $range.Columns.Count
    $custom = @()
    foreach ($s in $wb.TableStyles) { if (-not $s.BuiltIn) { $custom += $s.Name } }
    "  table styles defined in the workbook: " + ($custom -join ', ')
    $lastRow = $range.Rows.Count
    $lastCol = $range.Columns.Count
    # the first and the last column of: the header row, the first six data rows, the total row
    foreach ($r in @(1, 2, 3, 4, 5, 6, 7, $lastRow)) {
        $cells = @()
        foreach ($c in @(1, 2, $lastCol)) {
            $d = $range.Cells.Item($r, $c).DisplayFormat
            $fill = if ($d.Interior.ColorIndex -eq -4142) { 'none' } else { Get-Hex $d.Interior.Color }
            $cells += ("col {0}: fill {1} font {2}{3}{4}" -f $c, $fill, (Get-Hex $d.Font.Color), $(if ($d.Font.Bold) { ' bold' } else { '' }), $(if ($d.Font.Italic) { ' italic' } else { '' }))
        }
        "  row {0,2}: {1}" -f $r, ($cells -join ' | ')
    }
    $wb.Close($false)
} catch {
    "{0}: DID NOT OPEN ({1})" -f (Split-Path $Path -Leaf), $_.Exception.Message
    exit 1
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Start-Sleep -Seconds 2
    Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { Stop-Process -Id $_.Id -Force }
}
