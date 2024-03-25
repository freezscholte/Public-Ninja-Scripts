function ConvertFrom-HtmlTable {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Html
    )

    $objects = New-Object System.Collections.ArrayList

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

#Example code of how to use the function

$i = (Ninja-Property-Get devhtml | ConvertFrom-Json).html

$Object = ConvertFrom-HtmlTable -Html $i