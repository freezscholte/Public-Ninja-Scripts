<#
.SYNOPSIS
This function converts an HTML table to an array of objects, that is compatible with the WYSIWYG Styling in Ninja Customfields.

.PARAMETER Html
Convert an HTML table as string to an array of objects.

.EXAMPLE
Example usage of the code:

    $i = (Ninja-Property-Get devhtml | ConvertFrom-Json).html

    $Object = ConvertTo-HtmlTableToObject -Html $i

.NOTES
Feel free to use and modify this function as you see fit. If you have any questions or suggestions, please feel free to reach out to me.
#>

function ConvertTo-HtmlTableToObject {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Html
    )

    $objects = [System.Collections.Generic.List[Object]]::new()

    # Extract headers without HTML tags
    $headers = ([regex]::Matches($Html, '<th>\s*(.*?)\s*</th>')).Captures | ForEach-Object { $_.Value -replace '<[^>]+>', '' }

    # Extract rows
    $rows = ([regex]::Matches($Html, '<tr.*?>\s*(.*?)\s*</tr>', 'Singleline')).Captures | ForEach-Object { $_.Value }

    # Skip the first row (header row) by starting iteration from the second row
    for ($j = 1; $j -lt $rows.Count; $j++) {
        $row = $rows[$j]

        # Extract cell values without HTML tags
        $cellValues = ([regex]::Matches($row, '<td>\s*(.*?)\s*</td>')).Captures | ForEach-Object { $_.Value -replace '<[^>]+>', '' }

        # Use the previously determined regex to capture the class attribute from <tr>
        $rowClassMatch = [regex]::Match($row, '<tr class="?([^"\s]+)"?>')
        $rowClass = if ($rowClassMatch.Success) { $rowClassMatch.Groups[1].Value } else { "" }

        # Cleanup $RowClass of all HTML leftovers using a regex replacement
        #$rowClass = $rowClass -replace '<[^>]+>', '' # This removes any HTML tags
        $rowClass = $rowClass -replace '>.+', ''

        $obj = New-Object PSObject

        for ($i = 0; $i -lt $headers.Count; $i++) {
            # Assign clean header and cell values to the object
            $obj | Add-Member -Type NoteProperty -Name $headers[$i] -Value $cellValues[$i]
        }

        # Add RowColour property to the object after cleanup
        $obj | Add-Member -Type NoteProperty -Name "RowColour" -Value $rowClass

        [void]$objects.Add($obj)
    }

    return $objects
}
