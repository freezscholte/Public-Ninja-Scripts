##This script will download and update a machine to the latest version of windows 10
##
## -whereToDownload "http://www.myurl.com/windows10.iso"
## -fileLocation "c:\mypath\windows10.iso" or "\\networkpath\windows10.iso"

param([String]$whereToDownload='',[String]$fileLocation='')

##If the filepath is a network drive, copy locally first
##
if($fileLocation -ne '' -And $fileLocation.StartsWith("\\"))
{
	New-Item -ItemType Directory -Force -Path C:\Temp\Win10
	$fileDest = 'C:\Temp\Win10\Windows10.iso';
	Copy-Item -Path $fileLocation -Destination $fileDest -Force
	$fileLocation = $fileDest
}

# Enable TLS 1.2 as Security Protocol
[Net.ServicePointManager]::SecurityProtocol = `
    [Net.SecurityProtocolType]::Tls12 ;

##Code to download iso here
##
if($whereToDownload -ne '')
{
	$Username = "public" 
	$Password = "public" 
	New-Item -ItemType Directory -Force -Path C:\Temp\Win10
	$fileLocation = 'C:\Temp\Win10\Windows10.iso';
	$down = New-Object System.Net.WebClient
	$down.Credentials = New-Object System.Net.Networkcredential($Username, $Password) 
	$down.DownloadFile($whereToDownload,$fileLocation);
}

##Get path to downloaded iso
##
$ImagePath= $fileLocation ## Path of ISO image to be mounted 

##Get the drive letter if not already mounted
##
$ISODrive = (Get-DiskImage -ImagePath $ImagePath | Get-Volume).DriveLetter

if(!$ISODrive) 
{
	Mount-DiskImage -ImagePath $ImagePath -StorageType ISO
}

#Get the drive letter of the mounted iso
#
$ISODrive = (Get-DiskImage -ImagePath $ImagePath | Get-Volume).DriveLetter

#Build the final command to perform the upgrade
#
$command = "$ISODrive" + ":" + "\setup.exe"

#Run the update command
#
Write-host "Start Installation"
& $command /auto upgrade /quiet /noreboot /showoobe None
