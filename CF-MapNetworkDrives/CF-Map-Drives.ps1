
$Drives = Get-ItemProperty "Registry::HKEY_USERS\*\Network\*"

# See if any drives were found
if ( $Drives ) {

   $OutputArray = ForEach ( $Drive in $Drives ) {

        # PSParentPath looks like this: Microsoft.PowerShell.Core\Registry::HKEY_USERS\S-1-5-21-##########-##########-##########-####\Network
        $SID = ($Drive.PSParentPath -split '\\')[2]

           $data = [PSCustomObject]@{
            # Use .NET to look up the username from the SID
            Username            = ([System.Security.Principal.SecurityIdentifier]"$SID").Translate([System.Security.Principal.NTAccount])
            DriveLetter         = $Drive.PSChildName
            RemotePath          = $Drive.RemotePath
            SID                 = $SID
        }
        
        $Data

    }

} else {

    $Customfield = "No mapped drives were found"
    $Customfield

}

<#
(Get-CimInstance Win32_OperatingSystem).ProductType
1 - Work Station
2 - Domain Controller
3 - Server
#>

$MyOS = (Get-CimInstance Win32_OperatingSystem).ProductType
#If server only list mapped drives not users included
if ($MyOS -eq 3 ){
  
  $Customfield = $OutputArray.RemotePath | Sort-Object -Unique | Format-List | Out-String
  
} else {
  
  $Customfield = "User: $($OutputArray.Username) | Driveletter: $($OutputArray.DriveLetter) | Path: $($OutputArray.Remotepath)" | Format-List | Out-String
  
  
}

#Write Customfield for Console Ouput
$Customfield

Ninja-Property-Set networkDrives $Customfield



