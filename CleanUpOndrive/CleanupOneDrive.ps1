#Check if a file is pinned or unpinned
function Get-FilePinnedStatus {
    param (
        [string]$filePath
    )

    $fileInfo = [System.IO.FileInfo]::new($filePath)
    $attributes = $fileInfo.Attributes

    if ($attributes.HasFlag([System.IO.FileAttributes]::SparseFile)) {
        return $false # File is unpinned
    }
    else {
        return $true  # File is pinned
    }
}


# Function to get the actual size on disk
function Get-ActualSizeOnDisk {
    param (
        [string]$filePath
    )

    # Define the PInvoke signature for GetCompressedFileSize if it doesn't already exist
    if (-not ([System.Management.Automation.PSTypeName]'Kernel32').Type) {
        $signature = @"
        using System;
        using System.Runtime.InteropServices;

        public class Kernel32 {
            [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern uint GetCompressedFileSize(string lpFileName, out uint lpFileSizeHigh);
        }
"@

        Add-Type -TypeDefinition $signature -Language CSharp -PassThru
    }

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
# Define the function to set file attributes using PInvoke
function UnpinOneDriveFile {
    param (
        [string]$FilePath
    )

    # Resolve the full path of the file
    $resolvedPath = (Resolve-Path -Path $FilePath).ProviderPath

    try {
        # Define the PInvoke signature for SetFileAttributes if it doesn't already exist
        if (-not ([System.Management.Automation.PSTypeName]'WinAPI.Kernel32').Type) {
            $signature = @"
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            public static extern bool SetFileAttributes(string lpFileName, int dwFileAttributes);
"@

            Add-Type -MemberDefinition $signature -Name "Kernel32" -Namespace "WinAPI"
        }

        # Ensure the file exists
        if (-not (Test-Path -Path $resolvedPath)) {
            Write-Error "File does not exist: $resolvedPath"
            return
        }

        # Set the file attribute to 5248544 (combination of flags)
        $result = [WinAPI.Kernel32]::SetFileAttributes($resolvedPath, 5248544)

        if (-not $result) {
            $errorId = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            throw "Failed to set file attributes. Win32 Error Code: $errorId"
        }

        if ($EnableLogging) {
            Log-UnpinnedFile -FilePath $resolvedPath
        }
    }
    catch {
        Write-Error "Failed to unpin file: $_"
    }
}

# Function to calculate total size of OneDrive files on local disk

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

                    [void]$fileDetails.Remove($fileDetail)
                    $totalSize = ($fileDetails | Measure-Object -Property Size -Sum).Sum
                    Write-Output ("Initial local size: {0} GB" -f ([math]::Round($totalSize / 1GB, 2)))
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

# Example of how to run the script or just set the params statically

$parameters = @{
    MaxLocalSizeMB          = 4500
    MaxFileAgeDays          = 0
    EnableLogging           = $true
    ExclusionListExtensions = @(".txt", ".docx")
    ExclusionListFilenames  = @("DummyFile_1GB.dat")
}

SetOneDriveFilesToCloud @parameters