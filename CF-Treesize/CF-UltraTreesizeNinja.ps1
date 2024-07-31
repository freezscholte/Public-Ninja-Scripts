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
    Version        : 0.5 Beta
#>

Add-Type -TypeDefinition @"
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace MftReader
{
    public class FolderSizeCalculator
    {
        private string driveLetter;
        private int maxDepth;

        public FolderSizeCalculator(string driveLetter, int maxDepth)
        {
            this.driveLetter = driveLetter;
            this.maxDepth = maxDepth;
        }

        public ConcurrentDictionary<string, long> CalculateFolderSizes()
        {
            ConcurrentDictionary<string, long> folderSizes = new ConcurrentDictionary<string, long>();
            DirectoryInfo rootDir = new DirectoryInfo(this.driveLetter + ":\\");
            CalculateFolderSize(rootDir, 0, folderSizes);
            return folderSizes;
        }

        private void CalculateFolderSize(DirectoryInfo dirInfo, int currentDepth, ConcurrentDictionary<string, long> folderSizes)
        {
            if (currentDepth > this.maxDepth)
            {
                return;
            }

            try
            {
                long folderSize = 0;

                Parallel.ForEach(dirInfo.GetFiles(), file =>
                {
                    try
                    {
                        folderSize += file.Length;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error accessing file {0}: {1}", file.FullName, ex.Message);
                    }
                });

                Parallel.ForEach(dirInfo.GetDirectories(), subDir =>
                {
                    try
                    {
                        CalculateFolderSize(subDir, currentDepth + 1, folderSizes);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Console.WriteLine("Access denied to directory {0}: {1}", subDir.FullName, ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unexpected error with directory {0}: {1}", subDir.FullName, ex.Message);
                    }
                });

                folderSizes[dirInfo.FullName] = folderSize;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Access denied to directory {0}: {1}", dirInfo.FullName, ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error with directory {0}: {1}", dirInfo.FullName, ex.Message);
            }
        }
    }

    public class Utils
    {
        public static void WriteToFile(string content, string path)
        {
            File.WriteAllText(path, content);
        }
    }
}
"@

function Convert-BytesToSize {
    param (
        [Parameter(Mandatory = $true)]
        [long]$Bytes
    )

    $Kilobytes = $Bytes / 1KB
    $Megabytes = $Bytes / 1MB
    $Gigabytes = $Bytes / 1GB

    if ($Gigabytes -ge 1) {
        return "{0:N2} GB" -f $Gigabytes
    }
    elseif ($Megabytes -ge 1) {
        return "{0:N2} MB" -f $Megabytes
    }
    elseif ($Kilobytes -ge 1) {
        return "{0:N2} KB" -f $Kilobytes
    }
    else {
        return "{0:N2} bytes" -f $Bytes
    }
}

function Get-FolderSizes {
    param (
        [string]$driveLetter,
        [int]$maxDepth = 5,
        [int]$Top = 20
    )

    $folderSizeCalculator = New-Object MftReader.FolderSizeCalculator($driveLetter, $maxDepth)
    $folderSizes = $folderSizeCalculator.CalculateFolderSizes()

    if ($folderSizes.Count -eq 0) {
        throw "No folder sizes were calculated. Ensure that the drive letter is correct and accessible."
    }

    $sortedFolderSizes = $folderSizes.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $Top | ForEach-Object {
        [PSCustomObject]@{
            Folder = $_.Key
            Size     = Convert-BytesToSize -Bytes $_.Value
            RowColour = switch ($_.Value) {
                { $_ -gt 1GB } { "danger"; break }
                { $_ -gt 500MB } { "warning"; break }
                { $_ -gt 100MB } { "other"; break }
                default { "unknown" }
            }
        }
    }

    return $sortedFolderSizes
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

# Example usage
$results = Get-FolderSizes -driveLetter "C"


# Convert the results to an HTML table
ConvertTo-ObjectToHtmlTable -Objects $results | Ninja-Property-Set-Piped devhtml



