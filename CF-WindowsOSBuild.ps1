#Get OS Version Info
#V2.0
#JS
$minimumBuild = "19042"
$actualbuild = [System.Environment]::OSVersion.Version

if ($actualbuild.Build -lt $minimumBuild){
    Write-Host "Build To Old Not Supported"
    $buidinfo = "Build to old, update!"
}
else {
    $buidinfo = "Build Up To Date and Supported"
}

$Osversion = Get-ComputerInfo | select WindowsProductName, WindowsVersion, OsHardwareAbstractionLayer

$Output = [pscustomobject][ordered]@{
    Product = $Osversion.WindowsProductName
    Version = $Osversion.WindowsVersion
    Build = $actualbuild
    BuildStatus = $buidinfo 
}

$Output = $Output | Format-List | Out-String

Ninja-Property-Set windowsOsBuild $Output

exit
