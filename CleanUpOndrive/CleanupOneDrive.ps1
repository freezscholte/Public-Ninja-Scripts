# Import the GetCompressedFileSize function from kernel32.dll
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Kernel32 {
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern uint GetCompressedFileSize(string lpFileName, out uint lpFileSizeHigh);
}
"@

#Check if a file is pinned or unpinned
function Get-FilePinnedStatus {
    param (
        [string]$filePath
    )

    # Correct usage of 'attrib' command
    $attribOutput = & cmd /c "attrib" "$filePath"
    
    # Check for 'P' (Pinned) or 'U' (Unpinned) attributes
    if ($attribOutput -match "\sP\s") {
        return $true  # File is pinned
    }
    elseif ($attribOutput -match "\sU\s") {
        return $false # File is unpinned
    }
    else {
        Write-Warning "Unable to determine the status of $filePath"
        return $null
    }
}
# Function to get the actual size on disk
function Get-ActualSizeOnDisk {
    param (
        [string]$filePath
    )

    $highSize = 0
    $lowSize = [Kernel32]::GetCompressedFileSize($filePath, [ref]$highSize)

    if ($lowSize -eq 0xFFFFFFFF) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if ($errorCode -ne 0) {
            Write-Output "Error getting size for $filePath : $errorCode"
            return 0
        }
    }

    return ($highSize -shl 32) -bor $lowSize
}

# Function to log unpinned files
function Log-UnpinnedFile {
    param (
        [string]$FilePath
    )
    $logPath = "$env:TEMP\OneDriveCleanup.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "$timestamp : Unpinned $FilePath"
}

# Function to cleanup old logs
function Cleanup-OldLogs {
    param (
        [string]$LogPath,
        [int]$Days = 30
    )
    if (Test-Path $LogPath) {
        $cutoffDate = (Get-Date).AddDays(-$Days)
        $logEntries = Get-Content $LogPath
        $newLogEntries = foreach ($entry in $logEntries) {
            if ($entry -match "^\d{4}-\d{2}-\d{2}") {
                $logDate = $entry.Substring(0, 19) -as [datetime]
                if ($logDate -ge $cutoffDate) {
                    $entry
                }
            }
        }
        $newLogEntries | Set-Content $LogPath
    }
}

# Function to unpin a file (set to cloud)
function UnpinOneDriveFile {
    param (
        [string]$FilePath
    )
    # Using attrib command to set file to cloud
    attrib +U -P $FilePath
    if ($EnableLogging) {
        Log-UnpinnedFile -FilePath $FilePath
    }
}

function SetOneDriveFilesToCloud {
    param (
        [int]$MaxLocalSizeMB = 5000, # Maximum allowed size on local disk in MB
        [int]$MaxFileAgeDays = [int]::MaxValue, # Maximum age of files to be unpinned
        [bool]$EnableLogging = $false, # Enable logging of unpinned files
        [string[]]$ExclusionListExtensions = @(), # List of file extensions to exclude
        [string[]]$ExclusionListFilenames = @()   # List of filenames to exclude
    )

    # Convert MB to Bytes
    $MaxLocalSizeBytes = $MaxLocalSizeMB * 1MB

    # Cleanup old logs
    if ($EnableLogging) {
        Cleanup-OldLogs -LogPath "$env:TEMP\OneDriveCleanup.log" -Days 30
    }

    # Retrieve and filter the file list once
    $oneDrivePath = $env:OneDrive  # Adjust the OneDrive path if necessary

    $allFiles = [System.IO.Directory]::EnumerateFiles($oneDrivePath, "*", [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
        try {
            $fileInfo = Get-Item -Path $_ -ErrorAction Stop
            if (-not $fileInfo.Attributes.ToString().Contains("Hidden,System") -and $fileInfo.Length -gt 0) {
                $actualSize = Get-ActualSizeOnDisk -filePath $_
                if ($actualSize -gt 0) {
                    $fileInfo | Add-Member -MemberType NoteProperty -Name ActualSize -Value $actualSize -PassThru
                }
            }
        }
        catch {
            Write-Output "Skipping file due to error: $_"
        }
    }
    
    try {
        # Initialize a list to store file sizes and paths
        $fileDetails = [System.Collections.ArrayList]::new()
    
        # Populate the file details list
        foreach ($file in $allFiles) {
            if ($null -ne $file -and $file.PSObject.Properties['ActualSize']) {
                [void]$fileDetails.Add([PSCustomObject]@{
                        Path           = $file.FullName
                        Size           = $file.ActualSize
                        LastAccessTime = $file.LastAccessTime
                    })
            }
        }

        # Calculate the total size by summing the elements of the list
        $totalSize = ($fileDetails | Measure-Object -Property Size -Sum).Sum
        Write-Output ("Initial local size: {0} GB" -f ([math]::Round($totalSize / 1GB, 2)))
    }
    catch {
        $_.Exception.Message
    }

    try {
        # Main processing loop to unpin files until the size limit is met
        if ($totalSize -gt $MaxLocalSizeBytes) {
            $files = $fileDetails | Sort-Object LastAccessTime
            foreach ($fileDetail in $files) {
                if ($totalSize -le $MaxLocalSizeBytes) {
                    Write-Output "OneDrive Cleaned Up"
                    break
                }

                $filePath = $fileDetail.Path
                $fileName = [System.IO.Path]::GetFileName($filePath)
                $extension = [System.IO.Path]::GetExtension($filePath)

                # Check if file is excluded
                $excludeFile = $ExclusionListExtensions -contains $extension -or $ExclusionListFilenames -contains $fileName

                # Check file age
                $fileAge = (Get-Date) - $fileDetail.LastAccessTime
                if ($fileAge.TotalDays -ge $MaxFileAgeDays -and -not $excludeFile) {
                    Write-Output "Unpinning file: $filePath"
                    UnpinOneDriveFile -FilePath $filePath
                
                    # Recheck the actual status of the file pinned or unpinned
                    $actualSize = Get-ActualSizeOnDisk -filePath $filePath
                    if ((Get-FilePinnedStatus -filePath $filePath) -eq $false) {
                        [void]$fileDetails.Remove($fileDetail)
                        $totalSize = ($fileDetails | Measure-Object -Property Size -Sum).Sum
                        Write-Output "Updated local size: $totalSize bytes"
                    }
                }
            }
        }

        if ($totalSize -le $MaxLocalSizeBytes) {
            Write-Output "OneDrive Cleaned Up"
        }
    }
    catch {
        $_.Exception.Message
    }
}

#Example of how to run the script or just set the params statically
$parameters = @{
    MaxLocalSizeMB          = 3000
    MaxFileAgeDays          = 0
    EnableLogging           = $true
    ExclusionListExtensions = @(".txt", ".docx")
    ExclusionListFilenames  = @("DummyFile_1GB.dat")
}

SetOneDriveFilesToCloud @parameters


