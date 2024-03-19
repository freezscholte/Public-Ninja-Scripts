param(
    [string[]]$FilterExclude = @("C:\Windows"), # Default exclude path
    [string]$FilterSize = ">100mb", # Default size threshold is 100MB and greater, supported operators are >, <, >=, <=, =
    [int]$listResults = 40, # Default number of results to display
    [ValidateSet("FullName", "CreationTime", "LastAccessTime", "LastWriteTime", "SizeInGB")]
    [string]$PropertyToSort = "SizeInGB", # Default properties to display
    [string]$PortableEverythingURL = "https://www.voidtools.com/Everything-1.4.1.1024.x64.zip" #Portable Everything Download URL
)

# Function to convert the results to HTML table
function ConvertTo-HtmlTable {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.ArrayList]$Objects
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

    return $sb.ToString()
}

# Install Everything if not installed

try {
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Invoke-WebRequest -UseBasicParsing -Uri $PortableEverythingURL -OutFile "$($ENV:TEMP)\Everything.zip"
    Expand-Archive "$($ENV:TEMP)\Everything.zip" -DestinationPath $($ENV:Temp) -Force
    if (!(Get-Service "Everything Client" -ErrorAction SilentlyContinue)) {
        & "$($ENV:TEMP)\everything.exe" -install-client-service
        & "$($ENV:TEMP)\everything.exe" -reindex
        start-sleep 3
        Install-Module PSEverything
    }
    else {
        & "$($ENV:TEMP)\everything.exe" -reindex
        Install-Module PSEverything

    }
}
catch {
    $_.Exception.Message
}

try {
    #Scan the system and get the results

    $ScanResults = Search-Everything -Global -PathExclude $FilterExclude -AsArray -Filter "size:$($FilterSize)" | 
    ForEach-Object { Get-Item $_ -ErrorAction SilentlyContinue } | 
    Select-Object FullName, CreationTime, LastAccessTime, LastWriteTime, @{Name = "SizeInGB"; Expression = { [math]::Round($_.Length / 1GB, 1) } } |
    Sort-Object $PropertyToSort -Descending | Select-Object -First $listResults

    $CustomField = [System.Collections.Generic.List[object]]::new()

    foreach ($Object in $ScanResults) {
    
        $RowColour = switch ($Object.SizeInGB) {
            { $_ -gt 4 } { "danger"; break }
            { $_ -gt 1 } { "warning"; break }
            { $_ -gt 0.5 } { "other"; break }
            default { "unknown" } 
        }

        [void]$Customfield.Add([PSCustomObject]@{
                Path         = $Object.FullName
                Size         = "$($Object.SizeInGB) GB"
                Created      = $Object.CreationTime.Date.ToString("dd/MM/yyyy")
                LastModified = $Object.LastWriteTime.Date.ToString("dd/MM/yyyy")
                LastAccessed = $Object.LastAccessTime.Date.ToString("dd/MM/yyyy")
                RowColour    = $RowColour
            })
    }

    if (-not $CustomField.Count) {
        Write-Host "Did not find files bigger then 100MB"
    }


    $htmlTable = ConvertTo-HtmlTable -Objects $CustomField

    Ninja-Property-Set devhtml $htmlTable
}
catch {
    $_.Exception.Message
}

