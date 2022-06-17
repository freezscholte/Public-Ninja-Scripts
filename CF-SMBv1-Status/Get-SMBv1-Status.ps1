$SMB = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol

If ($SMB.State -eq "Enabled"){
    $Status = "Warning-1001: SMB is $($SMB.State)"
}else{
    $Status = "Info-1000: SMB is $($SMB.State)"
}

$Customfield = [PSCustomObject]@{
    "Status" = $Status
    "Remediation" = "We have a link to our wiki with more information"
}

$Customfield = $Customfield | Format-List | Out-String

$Customfield


Ninja-Property-Set smbv1Status $Customfield
