$SMB = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol
$WinEvent = Get-WinEvent -LogName Microsoft-Windows-SMBServer/Audit -MaxEvents 10

If ($SMB.State -eq "Enabled"){
    $Status = "Warning-1001: SMB is $($SMB.State)"
}else{
    $Status = "Info-1000: SMB is $($SMB.State)"
}

$message = foreach ($Event in $WinEvent.Message){
  $output = $Event -split "`n" | Select-String -pattern "^Client\sAddress:\s(.*)"
  $output
}

$Customfield = [PSCustomObject]@{
    "Status" = $Status
    "SMBv1 Clients" = $message
    "Remediation" = "LinktoWiki"
}

$Customfield = $Customfield | Format-List | Out-String

$Customfield


Ninja-Property-Set smbv1Status $Customfield
