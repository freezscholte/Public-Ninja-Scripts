<#
.SYNOPSIS
    Calculate the sizes of folders from the Master File Table using C# on a drive and output the results in an HTML table.

.DESCRIPTION
    This script calculates the sizes of folders from the Master File Table (MFT) using C# on a specified drive and outputs the results in an HTML table.
    The script uses a C# Library to read the MFT records and calculate the sizes of folders recursively.
    Duplicate files are then found based on the file size and MD5 hash, and files bigger then 25MB are hashed using a preliminary hash.

.NOTES
    File Name      : CF-FindDuplicateFiles.ps1
    Author         : Jan Scholte
    Version        : 0.2 Beta
#>

Add-Type -TypeDefinition @"
using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace FolderSizeCalculatorNamespace
{
    public class FileSystemItem
    {
        public string Path { get; set; }
        public long SizeOnDisk { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime LastWriteTime { get; set; }
        public bool IsDirectory { get; set; }
    }

    public class FolderSizeCalculator
    {
        private string driveLetter;
        private int maxDepth;
        private bool verboseOutput;
        public ConcurrentBag<FileSystemItem> Items { get; private set; }

        public FolderSizeCalculator(string driveLetter, int maxDepth, bool verboseOutput)
        {
            this.driveLetter = driveLetter;
            this.maxDepth = maxDepth;
            this.verboseOutput = verboseOutput;
            this.Items = new ConcurrentBag<FileSystemItem>();
        }

        public void CalculateFolderSizes()
        {
            DirectoryInfo rootDir = new DirectoryInfo(this.driveLetter + ":\\");
            CalculateFolderSize(rootDir, 0);
        }

        private long CalculateFolderSize(DirectoryInfo dirInfo, int currentDepth)
        {
            if (currentDepth > this.maxDepth)
            {
                return 0;
            }

            long folderSizeOnDisk = 0;

            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

            try
            {
                // Process files in the current directory
                var files = Enumerable.Empty<FileInfo>();
                try
                {
                    files = dirInfo.EnumerateFiles();
                }
                catch (Exception ex)
                {
                    if (this.verboseOutput)
                    {
                        Console.WriteLine("Error accessing files in directory {0}: {1}", dirInfo.FullName, ex.Message);
                    }
                }

                Parallel.ForEach(files, parallelOptions, file =>
                {
                    try
                    {
                        long fileSizeOnDisk = GetSizeOnDisk(file.FullName);
                        var item = new FileSystemItem
                        {
                            Path = file.FullName,
                            SizeOnDisk = fileSizeOnDisk,
                            CreationTime = file.CreationTime,
                            LastWriteTime = file.LastWriteTime,
                            IsDirectory = false
                        };
                        Items.Add(item);

                        Interlocked.Add(ref folderSizeOnDisk, fileSizeOnDisk);
                    }
                    catch (Exception ex)
                    {
                        if (this.verboseOutput)
                        {
                            Console.WriteLine("Error processing file {0}: {1}", file.FullName, ex.Message);
                        }
                    }
                });

                // Process subdirectories
                var subDirs = Enumerable.Empty<DirectoryInfo>();
                try
                {
                    subDirs = dirInfo.EnumerateDirectories();
                }
                catch (Exception ex)
                {
                    if (this.verboseOutput)
                    {
                        Console.WriteLine("Error accessing subdirectories in directory {0}: {1}", dirInfo.FullName, ex.Message);
                    }
                }

                Parallel.ForEach(subDirs, parallelOptions, subDir =>
                {
                    try
                    {
                        long subDirSizeOnDisk = CalculateFolderSize(subDir, currentDepth + 1);
                        Interlocked.Add(ref folderSizeOnDisk, subDirSizeOnDisk);
                    }
                    catch (Exception ex)
                    {
                        if (this.verboseOutput)
                        {
                            Console.WriteLine("Error processing directory {0}: {1}", subDir.FullName, ex.Message);
                        }
                    }
                });

                // Create FileSystemItem for the current directory
                var dirItem = new FileSystemItem
                {
                    Path = dirInfo.FullName,
                    SizeOnDisk = folderSizeOnDisk,
                    CreationTime = dirInfo.CreationTime,
                    LastWriteTime = dirInfo.LastWriteTime,
                    IsDirectory = true
                };
                Items.Add(dirItem);
            }
            catch (Exception ex)
            {
                if (this.verboseOutput)
                {
                    Console.WriteLine("Error accessing directory {0}: {1}", dirInfo.FullName, ex.Message);
                }
            }

            return folderSizeOnDisk;
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern uint GetCompressedFileSize(string lpFileName, out uint lpFileSizeHigh);

        public static long GetSizeOnDisk(string filename)
        {
            uint highOrder;
            uint lowOrder = GetCompressedFileSize(filename, out highOrder);
            if (lowOrder == 0xFFFFFFFF)
            {
                int error = Marshal.GetLastWin32Error();
                if (error != 0)
                {
                    throw new Win32Exception(error);
                }
            }
            return ((long)highOrder << 32) + lowOrder;
        }
    }
}
"@
function GetFilesMFTTable {
    param (
        [Parameter(Mandatory = $false)]
        [string]$DriveLetter,
        [int]$MaxDepth = 5,
        [Switch]$FolderSize,
        [Switch]$FileSize,
        [Switch]$VerboseOutput,
        [Switch]$AllDrives
    )

    # Validate parameters
    if (-not $AllDrives -and -not $DriveLetter) {
        throw "You must specify either -DriveLetter or -AllDrives."
    }

    # Get list of drives to process
    $drivesToProcess = @()
    if ($AllDrives) {
        # Get all local drives (excluding removable and network drives)
        $drives = Get-CimInstance Win32_LogicalDisk | Where-Object {
            $_.DriveType -eq 3 # DriveType 3 = Local Disk
        }
        $drivesToProcess = $drives.DeviceID
    }
    else {
        $drivesToProcess = @("$DriveLetter`:") # Add colon to match the format (e.g., "C:")
    }

    $allSortedItems = [System.Collections.Generic.List[object]]::new()

    foreach ($drive in $drivesToProcess) {
        if ($VerboseOutput) {
            Write-Output "Processing drive $drive"
        }

        # Extract the drive letter without colon
        $driveLetterOnly = $drive.TrimEnd(':')

        try {
            $folderSizeCalculator = New-Object FolderSizeCalculatorNamespace.FolderSizeCalculator($driveLetterOnly, $MaxDepth, [bool]$VerboseOutput)
            $folderSizeCalculator.CalculateFolderSizes()
            $items = $folderSizeCalculator.Items
        }
        catch {
            Write-Warning "Failed to calculate folder sizes for drive $drive : $_"
            continue
        }

        if ($items.Count -eq 0) {
            Write-Warning "No items were found on drive $drive. Ensure that the drive is accessible."
            continue
        }

        # Filter items based on parameters
        $selectedItems = $items

        if ($FolderSize -and -not $FileSize) {
            $selectedItems = $items | Where-Object { $_.IsDirectory }
        }
        elseif ($FileSize -and -not $FolderSize) {
            $selectedItems = $items | Where-Object { -not $_.IsDirectory }
        }
        elseif (-not $FolderSize -and -not $FileSize) {
            # If neither is specified, default to folders only
            $selectedItems = $items | Where-Object { $_.IsDirectory }
        }
        else {
            # Both FolderSize and FileSize are specified; include all items
            $selectedItems = $items
        }

        # Add the sorted items for this drive to the list of all items
        $allSortedItems.AddRange($selectedItems )
    }

    # Return all sorted items
    return $allSortedItems
}
function Get-DuplicateFilesBySizeAndHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Object[]]$Items,

        [Parameter()]
        [switch]$ExcludeWindowsOS,

        [Parameter()]
        [string[]]$ExcludePaths,

        [Parameter()]
        [double]$MinimumFileSizeMB = 0,

        [Parameter()]
        [string[]]$ExcludeExtensions
    )

    # Initialize an ArrayList to hold results
    $results = New-Object System.Collections.ArrayList

    # Step 1: Build the list of paths to exclude
    $excludedPaths = New-Object System.Collections.ArrayList

    if ($ExcludeWindowsOS) {
        # Get list of fixed drives
        $fixedDrives = (Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3").DeviceID

        # Build system directories
        $systemDirs = @(
            '\Windows',
            '\ProgramData',
            '\Users\All Users',
            '\Program Files (x86)',
            '\Program Files',
            '\Documents and Settings'
        )

        # Build paths for system directories on all fixed drives
        foreach ($drive in $fixedDrives) {
            foreach ($dir in $systemDirs) {
                [void]$excludedPaths.Add("$drive$dir")
            }
        }
    }

    if ($ExcludePaths) {
        $excludedPaths.AddRange($ExcludePaths)
    }

    # Convert MinimumFileSizeMB to bytes
    $minimumFileSizeBytes = $MinimumFileSizeMB * 1MB

    # Step 2: Filter items based on excluded paths
    if ($excludedPaths.Count -gt 0) {
        $filteredItems = $Items | Where-Object {
            $itemPath = $_.Path

            # Normalize the item path
            $normalizedItemPath = [System.IO.Path]::GetFullPath($itemPath)

            # Check if the item's path does not start with any of the excluded paths
            $exclude = $false
            foreach ($path in $excludedPaths) {
                # Normalize the excluded path
                $normalizedExcludedPath = [System.IO.Path]::GetFullPath($path)

                if ($normalizedItemPath -like "$normalizedExcludedPath*") {
                    $exclude = $true
                    break
                }
            }
            return -not $exclude
        }
    }
    else {
        $filteredItems = $Items
    }

    # Step 3: Further filter items based on file size and extension
    $filteredItems = $filteredItems | Where-Object {
        $_.SizeOnDisk -gt 0 -and
        $_.SizeOnDisk -ge $minimumFileSizeBytes -and
        $_.IsDirectory -eq $false -and
        (
            -not ($ExcludeExtensions -and ($ExcludeExtensions -contains [System.IO.Path]::GetExtension($_.Path)))
        )
    }

    # Step 4: Group files by SizeOnDisk where more than one file shares the same size
    $duplicateSizeGroups = $filteredItems | Group-Object -Property SizeOnDisk | Where-Object { $_.Count -gt 1 }

    # Step 5: For each group of files with the same size
    foreach ($group in $duplicateSizeGroups) {
        #$size = $group.Name
        $files = $group.Group

        # Step 6: Calculate MD5 hash for each file in the group
        foreach ($item in $files) {
            try {
                
                if ($item.SizeOnDisk -lt 25000000) {
                    # Calculate the MD5 hash of the file
                    $hash = Get-FileHash -Algorithm MD5 -Path $item.Path -ErrorAction Stop

                    # Create a new object with the original properties plus the MD5 hash
                    $itemWithHash = $item | Select-Object *, @{Name = 'MD5Hash'; Expression = { $hash.Hash } }

                    # Add the new object to the results array
                    [void]$results.Add($itemWithHash)
                } else {
                    # Generate preliminary hash
                    $prelimHash = Get-FilePreliminaryHash -Path $item.Path -ErrorAction Stop

                    # Create a new object with the original properties plus the preliminary hash
                    $itemWithHash = $item | Select-Object *, @{Name = 'MD5Hash'; Expression = { $prelimHash } }

                    # Add the new object to the results array
                    [void]$results.Add($itemWithHash)
                }
                    
                # Calculate the MD5 hash of the file
                #$hash = Get-FileHash -Algorithm MD5 -Path $item.Path -ErrorAction Stop

                # Create a new object with the original properties plus the MD5 hash
                #$itemWithHash = $item | Select-Object *, @{Name = 'MD5Hash'; Expression = { $hash.Hash } }

                # Add the new object to the results array
                #[void]$results.Add($itemWithHash)
            }
            catch {
                Write-Warning "Could not compute hash for $($item.Path): $_"
            }
        }
    }

    # Return the list of items with MD5 hashes

    $duplicates = $results `
    | Group-Object -Property MD5Hash `
    | Where-Object { $_.Count -gt 1 } `
    | ForEach-Object { $_.Group }

    # Output the duplicates
    #$duplicates

    return $duplicates
}

function Get-FilePreliminaryHash {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$ChunkSize = 4096 # 4KB
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "File does not exist: $Path"
    }

    $buffer = New-Object byte[] ($ChunkSize * 2)

    $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
    try {
        # Read first chunk
        $bytesRead = $stream.Read($buffer, 0, $ChunkSize)

        # Read last chunk if file is larger than ChunkSize
        if ($stream.Length -gt $ChunkSize) {
            $stream.Seek(-$ChunkSize, [System.IO.SeekOrigin]::End) | Out-Null
            $stream.Read($buffer, $ChunkSize, $ChunkSize) | Out-Null
        }

        $hashAlgorithm = [System.Security.Cryptography.MD5]::Create()
        $prelimHash = [BitConverter]::ToString($hashAlgorithm.ComputeHash($buffer, 0, $bytesRead + ($stream.Length -gt $ChunkSize ? $ChunkSize : 0))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $stream.Close()
    }
    
    return $prelimHash
}

$allFiles = GetFilesMFTTable -AllDrives -FileSize -MaxDepth 5

# Measure the time it takes to find duplicate files
Measure-Command { $duplicates = Get-DuplicateFilesBySizeAndHash -Items $allFiles -ExcludeExtensions '.vmgs', '.vhdx', '.vmdk', '.dat', '.tmp', '.log', '.dll', '.evtx' -ExcludeWindowsOS -MinimumFileSizeMB 10 }

Write-Output "Found $($duplicates.Count) duplicate files."



