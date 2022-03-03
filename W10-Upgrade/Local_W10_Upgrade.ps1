##Get path to downloaded iso
##
$fileLocation = "\\fileshare\Window10.iso"

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
& $command /auto upgrade /quiet /noreboot /showoobe None
