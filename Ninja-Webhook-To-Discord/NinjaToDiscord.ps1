#Place this PS Function in a HTTP Trigger Azure Function App and add the Function URL to the NinjaOne Webhook Notification Channels

using namespace System.Net

param($Request, $TriggerMetadata)

Function Process-Request {
    Param([object]$Request)

    Try {
        # Check if the body is a hashtable and use it directly
        if ($Request.body -is [System.Collections.Hashtable]) {
            $IncomingWebhookData = $Request.body
        }
        else {
            $IncomingWebhookData = $Request.body | ConvertFrom-Json
        }

        if ($null -eq $IncomingWebhookData) {
            Throw "Incoming webhook data is null."
        }

        # Initialize variables and setup
        $DiscordWebhookUrl = "DISCORD_WEBHOOK_URL_HERE"
       
        # Severity to color mapping
        $SeverityColorMap = @{
            "CRITICAL" = 0xFF0000
            "MAJOR"    = 0xFFA500
            "MODERATE" = 0xFFFF00
            "MINOR"    = 0x00FF00
            "NONE"     = 0x808080  # Default color if severity is unrecognized
        }

        # Validate severity and determine color
        $Severity = $IncomingWebhookData.severity.ToUpper()
        if (-not $SeverityColorMap.ContainsKey($Severity)) {
            $Severity = "NONE"
        }
        $EmbedColor = $SeverityColorMap[$Severity]

        # Construct Discord message
        $Fields = @(
            @{ name = "Type"; value = $IncomingWebhookData.type; inline = $true }
            @{ name = "Status"; value = $IncomingWebhookData.status; inline = $true }
            @{ name = "Severity"; value = $IncomingWebhookData.severity; inline = $true }
            @{ name = "Priority"; value = $IncomingWebhookData.priority; inline = $true }
            @{ name = "Result"; value = $IncomingWebhookData.activityResult; inline = $true }
            @{ name = "User"; value = $IncomingWebhookData.data.message.params.appUserName; inline = $true }
            @{ name = "Message"; value = $IncomingWebhookData.message; inline = $false }
        )

        # Create the embed object for Discord
        $DiscordEmbed = [PSCustomObject]@{
            title       = "Alert"
            description = "A new event has occurred."
            color       = $EmbedColor
            fields      = $Fields
            timestamp   = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"
        }

        # Wrap the embed in an array
        [System.Collections.ArrayList]$EmbedArray = @()
        $EmbedArray.Add($DiscordEmbed)

        # Create the payload
        $Payload = [PSCustomObject]@{
            content  = if ($Severity -eq "CRITICAL") { "@everyone - A critical event has occurred" } else { "" }
            username = "NinjaOne"
            embeds   = $EmbedArray

        }

        # Convert the object to JSON with increased depth
        $DiscordWebhookData = $Payload | ConvertTo-Json -Depth 4

        # Send data to Discord webhook
        $Response = Invoke-RestMethod -Uri $DiscordWebhookUrl -Method Post -Body $DiscordWebhookData -ContentType "application/json"

        return "Webhook sent successfully"
    }
    Catch {
        Write-Error "Failed to process request: $_"
        return $_.Exception.Message
    }
}


$body = Process-Request -Request $Request

#Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $body
    })

