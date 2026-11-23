# Compares two xlsx files part by part (entry names and uncompressed bytes).
# The compressed bytes may differ when the deflater changes; the parts must not.
param([Parameter(Mandatory = $true)][string]$A, [Parameter(Mandatory = $true)][string]$B)
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
function Read-Parts([string]$p) {
    $z = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $p).Path)
    try {
        $d = [ordered]@{}
        foreach ($e in $z.Entries) {
            $s = $e.Open(); $m = New-Object IO.MemoryStream; $s.CopyTo($m); $s.Close()
            $d[$e.FullName.TrimStart('/')] = $m.ToArray()
        }
        return $d
    } finally { $z.Dispose() }
}
$pa = Read-Parts $A; $pb = Read-Parts $B
$diff = 0
$names = @($pa.Keys) + @($pb.Keys) | Select-Object -Unique
foreach ($n in $names) {
    if (-not $pa.Contains($n)) { "only in B: $n"; $diff++; continue }
    if (-not $pb.Contains($n)) { "only in A: $n"; $diff++; continue }
    $x = $pa[$n]; $y = $pb[$n]
    if ($x.Length -ne $y.Length -or -not [Linq.Enumerable]::SequenceEqual([byte[]]$x, [byte[]]$y)) {
        $i = 0; $max = [Math]::Min($x.Length, $y.Length); while ($i -lt $max -and $x[$i] -eq $y[$i]) { $i++ }
        $ta = [Text.Encoding]::UTF8.GetString($x, [Math]::Max(0, $i - 40), [Math]::Min(120, $x.Length - [Math]::Max(0, $i - 40)))
        $tb = [Text.Encoding]::UTF8.GetString($y, [Math]::Max(0, $i - 40), [Math]::Min(120, $y.Length - [Math]::Max(0, $i - 40)))
        "DIFF $n (len $($x.Length) vs $($y.Length), first at $i)`n  A: $ta`n  B: $tb"; $diff++
    }
}
$orderA = (@($pa.Keys) -join ','); $orderB = (@($pb.Keys) -join ',')
"parts A=$($pa.Count) B=$($pb.Count) differing=$diff entryOrderSame=$($orderA -eq $orderB)"
