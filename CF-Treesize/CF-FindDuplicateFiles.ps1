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
    Version        : 0.9 RC
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$CustomFieldName = "duplicateFiles",

    [Parameter()]
    [string[]]$ExcludeExtensions = @('.vmgs', '.vhdx', '.vhd', '.vmrs', '.vmdk', '.dat', '.tmp', '.log', '.dll', '.evtx'),

    [Parameter()]
    [string[]]$ExcludePaths = @(),
    [Parameter()]
    [int]$MinimumFileSizeMB = 10,

    [Parameter()]
    [int]$MaxDepth = 5
)

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
        private int maxDegreeOfParallelism;
        public ConcurrentBag<FileSystemItem> Items { get; private set; }

        public FolderSizeCalculator(string driveLetter, int maxDepth, bool verboseOutput, int? maxParallelism = null)
        {
            this.driveLetter = driveLetter;
            this.maxDepth = maxDepth;
            this.verboseOutput = verboseOutput;
            this.maxDegreeOfParallelism = maxParallelism ?? (Environment.ProcessorCount * 4);
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

            var parallelOptions = new ParallelOptions { MaxDegreeOfParallelism = this.maxDegreeOfParallelism };

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

    # Use generic List instead of ArrayList for better performance
    $results = [System.Collections.Generic.List[object]]::new()
    $excludedPaths = [System.Collections.Generic.List[string]]::new()

    if ($ExcludeWindowsOS) {
        $fixedDrives = (Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3").DeviceID
        $systemDirs = @(
            '\Windows',
            '\ProgramData',
            '\Users\All Users',
            '\Program Files (x86)',
            '\Program Files',
            '\Documents and Settings'
        )

        foreach ($drive in $fixedDrives) {
            foreach ($dir in $systemDirs) {
                $excludedPaths.Add("$drive$dir")
            }
        }
    }

    if ($ExcludePaths) {
        $excludedPaths.AddRange($ExcludePaths)
    }

    $minimumFileSizeBytes = $MinimumFileSizeMB * 1MB

    # Optimize filtering by combining conditions
    $filteredItems = $Items.Where({
            $item = $_
            $include = $_.SizeOnDisk -gt 0 -and
            $_.SizeOnDisk -ge $minimumFileSizeBytes -and
            -not $_.IsDirectory -and
                  (-not ($ExcludeExtensions -and ($ExcludeExtensions -contains [System.IO.Path]::GetExtension($_.Path))))

            if (-not $include) { return $false }

            if ($excludedPaths.Count -gt 0) {
                $normalizedItemPath = [System.IO.Path]::GetFullPath($item.Path)
                foreach ($path in $excludedPaths) {
                    if ($normalizedItemPath.StartsWith([System.IO.Path]::GetFullPath($path), [StringComparison]::OrdinalIgnoreCase)) {
                        return $false
                    }
                }
            }
            return $true
        })

    # Group files by size
    $duplicateSizeGroups = $filteredItems | Group-Object -Property SizeOnDisk | Where-Object { $_.Count -gt 1 }

    # Process files in batches for better memory management
    foreach ($group in $duplicateSizeGroups) {
        $files = $group.Group
        
        foreach ($item in $files) {
            try {
                if ($item.SizeOnDisk -lt 25MB) {
                    $hash = Get-FileHash -Algorithm MD5 -Path $item.Path -ErrorAction Stop
                    $itemWithHash = $item | Select-Object *, 
                    @{Name = 'MD5Hash'; Expression = { $hash.Hash } },
                    @{Name = 'Size'; Expression = { Convert-BytesToSize -Bytes $_.SizeOnDisk } },
                    @{Name = 'RowColour'; Expression = {
                            switch ($_.SizeOnDisk) {
                                { $_ -gt 1073741824 } { "danger"; break }  # 1GB in bytes
                                { $_ -gt 524288000 } { "warning"; break } # 500MB in bytes
                                { $_ -gt 10485760 } { "info"; break }    # 10MB in bytes
                                default { "default" }
                            }
                        }
                    }
                }
                else {
                    $prelimHash = Get-FilePreliminaryHash -Path $item.Path -ChunkSize 1MB -ErrorAction Stop
                    $itemWithHash = $item | Select-Object *, 
                    @{Name = 'MD5Hash'; Expression = { $prelimHash } },
                    @{Name = 'Size'; Expression = { Convert-BytesToSize -Bytes $_.SizeOnDisk } },
                    @{Name = 'RowColour'; Expression = {
                            switch ($_.SizeOnDisk) {
                                { $_ -gt 1073741824 } { "danger"; break }  # 1GB in bytes
                                { $_ -gt 524288000 } { "warning"; break } # 500MB in bytes
                                { $_ -gt 10485760 } { "info"; break }    # 10MB in bytes
                                default { "default" }
                            }
                        }
                    }
                }
                $results.Add($itemWithHash)
            }
            catch {
                Write-Warning "Could not compute hash for $($item.Path): $_"
            }
        }
    }

    # Return duplicates with formatted size
    return $results | 
    Group-Object -Property MD5Hash | 
    Where-Object { $_.Count -gt 1 } |
    Sort-Object { ($_.Group | Measure-Object -Property SizeOnDisk -Sum).Sum } -Descending |
    ForEach-Object { $_.Group }
}
function Get-FilePreliminaryHash {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$ChunkSize = 1MB
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
        
        # Replace ternary operator with if-else
        $additionalBytes = if ($stream.Length -gt $ChunkSize) { $ChunkSize } else { 0 }
        $prelimHash = [BitConverter]::ToString(
            $hashAlgorithm.ComputeHash($buffer, 0, $bytesRead + $additionalBytes)
        ).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $stream.Close()
    }
    
    return $prelimHash
}
function Convert-BytesToSize {
    param (
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    $sizes = "bytes", "KB", "MB", "GB", "TB", "PB", "EB"
    $factor = 0

    while ($Bytes -ge 1KB -and $factor -lt $sizes.Length - 1) {
        $Bytes /= 1KB
        $factor++
    }

    return "{0:N2} {1}" -f $Bytes, $sizes[$factor]
}
function ConvertTo-ObjectToHtmlTable {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[Object]]$Objects
    )
    $sb = New-Object System.Text.StringBuilder
    # Start the HTML table
    [void]$sb.Append('<table><thead><tr>')
    # Add column headers based on the properties of the first object, excluding "RowColour"
    $Objects[0].PSObject.Properties.Name |
    Where-Object { $_ -ne 'RowColour' } |
    ForEach-Object { [void]$sb.Append("<th>$_</th>") }

    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($obj in $Objects) {
        # Use the RowColour property from the object to set the class for the row
        $rowClass = if ($obj.RowColour) { $obj.RowColour } else { "" }

        [void]$sb.Append("<tr class=`"$rowClass`">")
        # Generate table cells, excluding "RowColour"
        foreach ($propName in $obj.PSObject.Properties.Name | Where-Object { $_ -ne 'RowColour' }) {
            [void]$sb.Append("<td>$($obj.$propName)</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    $OutputLength = $sb.ToString() | Measure-Object -Character -IgnoreWhiteSpace | Select-Object -ExpandProperty Characters
    if ($OutputLength -gt 200000) {
        Write-Warning ('Output appears to be over the NinjaOne WYSIWYG field limit of 200,000 characters. Actual length was: {0}' -f $OutputLength)
    }
    return $sb.ToString()
}

#Grab all files from the MFT table
$allFiles = GetFilesMFTTable -AllDrives -FileSize -MaxDepth $MaxDepth

#Set the parameters for the Get-DuplicateFilesBySizeAndHash function
$params = @{
    Items             = $allFiles
    ExcludeExtensions = $ExcludeExtensions 
    MinimumFileSizeMB = $MinimumFileSizeMB
    ExcludeWindowsOS  = [switch]::Present
}

# Get the duplicate files
$duplicates = Get-DuplicateFilesBySizeAndHash @params

Write-Output "Found $($duplicates.Count) duplicate files."

#Convert the duplicate files to an HTML table and set the property in NinjaOne
ConvertTo-ObjectToHtmlTable -Objects $duplicates | Ninja-Property-Set-Piped $CustomFieldName