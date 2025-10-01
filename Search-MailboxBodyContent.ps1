[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]] $UserId,

    [Parameter(Mandatory)]
    [string] $SearchString,

    [int] $Top
)

foreach ($id in $UserId) {
    $searchParam = @{
        Search   = "body:$SearchString"
        UserId   = $id
        Property = "subject,from,receivedDateTime,body,HasAttachments"
    }

    if ($Top) {
        $searchParam.Top = $Top
    }
    else {
        $searchParam.All = $true
        $searchParam.PageSize = 100
    }

    try {
        $messages = Get-MgUserMessage @searchParam -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to retrieve messages for $id : $_"
        continue
    }

    foreach ($msg in $messages) {
        $bodyText = [System.Net.WebUtility]::HtmlDecode(
            ($msg.Body.Content -replace '<[^>]+>', ' ')
        )

        if (-not $msg.HasAttachments -or $bodyText -match [regex]::Escape($SearchString)) {
            $matchLocation = "Body"
            $attachments = @()
        }
        else {
            $matchLocation = "Attachment"
            $attachments = Get-MgUserMessageAttachment -UserId $id -MessageId $msg.Id
        }

        [PSCustomObject]@{
            Mailbox            = $id
            Subject            = $msg.Subject
            From               = $msg.From.EmailAddress.Address
            ReceivedDate       = $msg.ReceivedDateTime
            FirstMatchLocation = $matchLocation
            MessageId          = $msg.Id
            AttachmentNames    = ($attachments).Name -join ","
            AttachmentIds      = ($attachments).Id -join ","
        }
    }
}
