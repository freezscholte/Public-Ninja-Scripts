#Modification of Chris White his script posted in the Dojo

$HBTempFolderName = "$env:TEMP\hbtemp\"
$URI = "https://download.sysinternals.com/files/DU.zip"
$OUTFILE = "$env:TEMP\hbtemp\DU.zip"
$DestinationPath = "$env:TEMP\hbtemp\DU"
$scan = "C:\"

if (Test-Path "$env:TEMP\hbtemp\DU\du.exe"){
    Write-Host "DU.exe already there"
}else{
    New-Item $HBTempFolderName -ItemType Directory -Force
    Write-Host "hbtemp created succesfully"
    Invoke-WebRequest -URI $URI -OUTFILE $OUTFILE
    Expand-Archive -LiteralPath $OUTFILE -DestinationPath $DestinationPath
}

$StartTime = Get-Date
cd "$env:TEMP\hbtemp\DU"

$forest = .\du.exe -accepteula -nobanner -c -l 5 $scan | ConvertFrom-Csv `
| Select-Object Path,@{Name="DirectorySizeOnDisk";expression={[Math]::Round($_.DirectorySizeOnDisk / 1GB) }} `
| Where-Object { $_.DirectorySizeOnDisk -gt 1 } `
| Sort-Object { $_.DirectorySizeOnDisk } -descending | Select-Object -First 20

$Output = foreach ($tree in $forest){
    [PSCustomObject]@{
        TreeSize = "$($Tree.Path) - $($Tree.DirectorySizeOnDisk) GB"
    }
}

$stopwatch = "Total scan time in $((New-Timespan -Start $StartTime -End $(Get-Date)).TotalSeconds) seconds"

$CustomField = $Output + $stopwatch | format-table | Out-String

$CustomField

Ninja-Property-Set Treesize $CustomField
