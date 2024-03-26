<#
.SYNOPSIS
This function converts an array of objects to an HTML table, that is compatible with the WYSIWYG Styling in Ninja Customfields.

.PARAMETER Objects
An array of objects to convert to an HTML table. You can set the "RowColour" property on each object to set the <tr class=`"$rowClass`">

.EXAMPLE
Example usage of the code:

    $htmlTable = ConvertTo-ObjectToHtmlTable -Objects $ArrayObject

    Ninja-Property-Set cveTable $output

    I have seen some rare cases to wrap the $output like this:

      $output = "@'
    $($htmlTable)
    '@"

.NOTES
Feel free to use and modify this function as you see fit. If you have any questions or suggestions, please feel free to reach out to me.
#>


function ConvertTo-ObjectToHtmlTable {
    param (
        [Parameter(Mandatory=$true)]
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

