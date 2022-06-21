$Drives = Get-ItemProperty "Registry::HKEY_USERS\*\Network\*"

# See if any drives were found
if ( $Drives ) {

   $OutputArray = ForEach ( $Drive in $Drives ) {

        $SID = ($Drive.PSParentPath -split '\\')[2]
        $Username = ([System.Security.Principal.SecurityIdentifier]"$SID").Translate([System.Security.Principal.NTAccount])
           $Data = [PSCustomObject]@{
            Output  = "User: $($Username) - Driveletter: $($Drive.PSChildName) - Path: $($Drive.RemotePath)"
        }
        
        $Data

    }

} else {

    $Customfield = "No mapped drives were found"
    $Customfield

}


$Customfield = $OutputArray.Output | Format-List | Out-String

$Customfield

Ninja-Property-Set networkDrives $Customfield
