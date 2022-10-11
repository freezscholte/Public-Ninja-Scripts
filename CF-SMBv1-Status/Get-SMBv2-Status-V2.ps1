$SMB = Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol
$WinEvent = Get-WinEvent -LogName Microsoft-Windows-SMBServer/Audit -MaxEvents 10

$i = (Get-SmbServerConfiguration).AuditSMB1Access

#Checks if SMBv1 Audit is Enabled
if ($i -eq $True){
  Write-Host "SMB Audit Enabled is $($i)"
}else{
  Set-SmbServerConfiguration –AuditSmb1Access $true -Force
}

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
    "Remediation" = "We have a link to our wiki with more information"
}

$Customfield = $Customfield | Format-List | Out-String

$Customfield


Ninja-Property-Set smbv1Status $Customfield
