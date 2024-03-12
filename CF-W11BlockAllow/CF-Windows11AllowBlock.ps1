#Get Ninja current documentation field value
$ErrorActionPreference = 'silentlycontinue'

try {
    $W10DeviceOnly = (Ninja-Property-Docs-Get-Single "Configuraties" "windows10OnlyDevices").split(",").ToUpper()
    $W11ApproveDeny = Ninja-Property-Docs-Get-Single "Configuraties" "windows11ApproveDeny"
    $W11ReadyStatus = (Ninja-Property-Get windows11Ready) -eq "Status : Warning-1001: Not Compatible with Windows 11"
}
catch {
    Write-Output 'An error occurred while getting the data from the Documentation fields.'
    $_.Exception.Message

}


#Functions
function BlockWindows11Regkeys {
    # Set the necessary registry keys
    $windowsUpdateRegPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $windowsUpdateRegKeys = @{
        'TargetReleaseVersion'     = 1
        'TargetReleaseVersionInfo' = '22H2'
        'ProductVersion'           = 'Windows 10'
    }

    # Create the registry path if it does not exist
    if (!(Test-Path $windowsUpdateRegPath)) {
        New-Item -Path $windowsUpdateRegPath -Force | Out-Null
    }

    # Set the registry values
    foreach ($key in $windowsUpdateRegKeys.Keys) {
        Set-ItemProperty -Path $windowsUpdateRegPath -Name $key -Value $windowsUpdateRegKeys[$key]
    }

    # Information message
    Write-Output 'Registry keys have been set to block Windows 11 and allow Windows 10 21H2.'
}

function Test-WindowsUpdateRegKeys {
    $windowsUpdateRegPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $windowsUpdateRegKeys = @{
        'TargetReleaseVersion'     = 1
        'TargetReleaseVersionInfo' = '22H2'
        'ProductVersion'           = 'Windows 10'
    }

    # Check if the registry path exists
    if (!(Test-Path $windowsUpdateRegPath)) {
        return $false
    }

    # Check if the registry values are set as expected
    foreach ($key in $windowsUpdateRegKeys.Keys) {
        $currentValue = (Get-ItemProperty -Path $windowsUpdateRegPath -Name $key -ErrorAction SilentlyContinue).$key
        if ($currentValue -ne $windowsUpdateRegKeys[$key]) {
            return $false
        }
    }

    return $true
}

function RemoveWindowsUpdateRegKeys {
    $windowsUpdateRegPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $windowsUpdateRegKeys = @(
        'TargetReleaseVersion',
        'TargetReleaseVersionInfo',
        'ProductVersion'
    )

    # Check if the registry path exists
    if (!(Test-Path $windowsUpdateRegPath)) {
        Write-Host "Registry path not found. No keys to remove." -ForegroundColor Yellow
        return
    }

    # Remove the registry values if they exist
    $keysRemoved = $false
    foreach ($key in $windowsUpdateRegKeys) {
        if ((Get-ItemProperty -Path $windowsUpdateRegPath -Name $key -ErrorAction SilentlyContinue).$key) {
            Remove-ItemProperty -Path $windowsUpdateRegPath -Name $key
            $keysRemoved = $true
        }
    }

    if ($keysRemoved) {
        Write-Host "Registry keys have been removed." -ForegroundColor Green
    }
    else {
        Write-Host "No matching registry keys found to remove." -ForegroundColor Yellow
    }
}

#Logic

try {
    $osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    $regKeysSet = (Test-WindowsUpdateRegKeys)
    $blockWindows11 = ($W11ApproveDeny -eq 'Deny') -or ($W10DeviceOnly -contains $env:computername)

    if ($osVersion -ge '10.0.22000' -or $w11ReadyStatus -eq $True) {
        Write-Output 'Windows 11 is already installed, or device is not compatible exit script'
        exit
    }

    if ($regKeysSet -eq $True -and $blockWindows11 -eq $True) {
        Write-Output 'Registry keys are already set correctly and Windows 11 is blocked, exit script'
        exit
    }
    else {
        Write-Output 'Registry keys are not set correctly, continue with script to check if Windows 11 is allowed on this device.'
    }

    if ($W11ApproveDeny -eq 'Deny' -or $W10DeviceOnly -contains $env:computername) {
        BlockWindows11Regkeys
        Write-Output 'Windows 11 is not allowed on this device, registry keys have been set to block Windows 11 and allow Windows 10 22H2.'
        exit
    }
    else {
        Write-Output 'Windows 11 is allowed on this device, continue with script to check if registry keys are set to block Windows 11 and only allow Windows 10 22H2.'
    }

    if ($W11ApproveDeny -ne 'Deny' -and $W10DeviceOnly -notcontains $env:computername -and ((Test-WindowsUpdateRegKeys) -eq $True)) {
        Write-Output 'Windows 11 is allowed on this device, but the registry keys are set to block Windows 11 and allow Windows 10 22H2.'
        RemoveWindowsUpdateRegKeys
        exit
    }


}
catch {
    Write-Output 'An error occurred while setting the registry keys.'
    $_.Exception.Message
    exit 1002
}
