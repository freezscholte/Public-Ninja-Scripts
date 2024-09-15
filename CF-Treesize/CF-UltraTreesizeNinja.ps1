<#
.SYNOPSIS
    Calculate the sizes of folders from the Master File Table using C# on a drive and output the results in an HTML table.

.DESCRIPTION
    This script calculates the sizes of folders from the Master File Table (MFT) using C# on a specified drive and outputs the results in an HTML table.
    The script uses a compiled C# class to read the MFT records and calculate the sizes of folders recursively.
    The results are displayed in an HTML table with the top folders by size, and the row color is based on the size of the folder.

.NOTES
    File Name      : CF-UltraTreesizeNinja.ps1
    Author         : Jan Scholte
    Version        : 0.9.1 RC
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
function Get-FolderSizes {
    param (
        [Parameter(Mandatory = $false)]
        [string]$DriveLetter,
        [int]$MaxDepth = 5,
        [int]$Top = 20,
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
    } else {
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
        } elseif ($FileSize -and -not $FolderSize) {
            $selectedItems = $items | Where-Object { -not $_.IsDirectory }
        } elseif (-not $FolderSize -and -not $FileSize) {
            # If neither is specified, default to folders only
            $selectedItems = $items | Where-Object { $_.IsDirectory }
        } else {
            # Both FolderSize and FileSize are specified; include all items
            $selectedItems = $items
        }

        # Process and sort the selected items
        $sortedItems = $selectedItems | Sort-Object -Property SizeOnDisk -Descending | Select-Object -First $Top | ForEach-Object {
            [PSCustomObject]@{
                Drive         = $drive
                Path          = $_.Path
                Size          = Convert-BytesToSize -Bytes $_.SizeOnDisk
                CreationTime  = $_.CreationTime
                LastWriteTime = $_.LastWriteTime
                IsDirectory   = $_.IsDirectory
                RowColour     = switch ($_.SizeOnDisk) {
                    { $_ -gt 30GB } { "danger"; break }
                    { $_ -gt 5GB }  { "warning"; break }
                    { $_ -gt 1GB }  { "info"; break }
                    default         { "default" }
                }
            }
        }

        # Add the sorted items for this drive to the list of all items
        $allSortedItems.AddRange($sortedItems)
    }

    # Return all sorted items
    return $allSortedItems
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

# Generate the Treesize report
$results = Get-FolderSizes -AllDrives -MaxDepth 5 -Top 40 -FolderSize -FileSize

# Convert the results to an HTML table
ConvertTo-ObjectToHtmlTable -Objects $results | Ninja-Property-Set-Piped devhtml