<#
.SYNOPSIS
This function converts an array of objects to an HTML table, that is compatible with the WYSIWYG Styling in Ninja Customfields.

.PARAMETER Objects
An array of objects to convert to an HTML table. You can set the "RowColour" property on each object to set the <tr class=`"$rowClass`">

.EXAMPLE
Example usage of the code:

    $htmlTable = ConvertTo-HtmlTable -Objects $ArrayObject

    $output = "@'
    $($htmlTable)
    '@"

    Ninja-Property-Set cveTable $output

.NOTES
Feel free to use and modify this function as you see fit. If you have any questions or suggestions, please feel free to reach out to me.
#>


function ConvertTo-HtmlTable {
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

#Example code of how to use the function

$RowColour = switch ($Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.baseSeverity) {
    "HIGH" { "danger" } # Assuming you meant to map "HIGH" to "danger" as per your initial description
    "MEDIUM" { "warning" }
    "LOW" { "other" }
    default { $null } # Fallback color in case it doesn't match any case
}

$CVEResult = [PSCustomObject]@{
    'CVE'                 = $cveId
    'ExploitabilityScore' = $Object.vulnerabilities.cve.metrics.cvssMetricV31.exploitabilityScore
    'ImpactScore'         = $Object.vulnerabilities.cve.metrics.cvssMetricV31.impactScore
    'AttackVector'        = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.attackVector
    'AttackComplexity'    = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.attackComplexity
    'PrivilegesRequired'  = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.privilegesRequired
    'IntegrityImpact'     = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.integrityImpact
    'UserInteraction'     = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.userInteraction
    'BaseScore'           = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.baseScore
    'BaseSeverity'        = $Object.vulnerabilities.cve.metrics.cvssMetricV31.cvssdata.baseSeverity
    'RowColour'           = $RowColour
}

$htmlTable = ConvertTo-HtmlTable -Objects $ArrayObject

$output = "@'
$($htmlTable)
'@"

Ninja-Property-Set cveTable $output
