Param (
   [int] $count
)

$i = 3
if($count) { $i = $count }

while ($i) {
    Write-Host "Line counter [$i]`n`r"
    Start-Sleep -Seconds 1
    $i --
}