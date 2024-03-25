function ConvertFrom-HtmlTable {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Html
    )

    $objects = New-Object System.Collections.ArrayList

    # Extract headers without HTML tags
    $headers = ([regex]::Matches($Html, '<th>\s*(.*?)\s*</th>')).Captures | ForEach-Object { $_.Value -replace '<[^>]+>', '' }

    # Ensure headers were found
    if ($headers.Count -eq 0) {
        Write-Error "No headers found in the HTML table."
        return $null
    }

    # Extract rows, excluding the header row
    $rows = ([regex]::Matches($Html, '<tr[^>]*>\s*(.*?)\s*</tr>', 'Singleline')).Captures | ForEach-Object { $_.Value }

    # Skip the first row (header row) by starting iteration from the second row
    for ($j = 1; $j -lt $rows.Count; $j++) {
        $row = $rows[$j]

        # Extract cell values without HTML tags
        $cellValues = ([regex]::Matches($row, '<td>\s*(.*?)\s*</td>')).Captures | ForEach-Object { $_.Value -replace '<[^>]+>', '' }

        # Extract the class attribute for RowColour
        $rowClassMatch = [regex]::Match($row, 'class="([^"]*)"')
        $rowClass = if ($rowClassMatch.Success) { $rowClassMatch.Groups[1].Value } else { $null }

        $obj = New-Object PSObject -Property @{ RowColour = $rowClass }

        for ($i = 0; $i -lt $headers.Count; $i++) {
            # Assign clean header and cell values to the object
            $obj | Add-Member -Type NoteProperty -Name $headers[$i] -Value $cellValues[$i]
        }

        [void]$objects.Add($obj)
    }

    return $objects
}