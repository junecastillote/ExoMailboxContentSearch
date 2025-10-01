[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]] $UserId,

    [Parameter(Mandatory)]
    [string] $SearchString,

    [Parameter()]
    [int] $Top,

    [Parameter()]
    [switch]
    $ShowAttachment
)

$startTime = (Get-Date)

$UserId = $UserId | Select-Object -Unique

if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

Write-Verbose "Search string = '$($SearchString)'"
if ($Top) {
    Write-Verbose "Only the first [$($Top)] matching messages per mailbox will be returned."
}

$counter = 0
$total = $UserId.Count
$totalMessageCount = 0
$totalMailboxMatchCount = 0

Write-Verbose "Start searching mailboxes..."
foreach ($id in $UserId) {
    $preCounter = $counter
    $prePercentComplete = ($preCounter / $total) * 100

    $counter++
    $percentComplete = ($counter / $total) * 100

    Write-Verbose "Searching mailbox [$($id)]..."
    Write-Progress -Activity "Searching mailbox body for [$($SearchString)]... $([math]::Round($prePercentComplete,2))%" -Status "[$($counter)/$($total))] $($id)" -PercentComplete $prePercentComplete


    try {
        $null = Get-MgUser -UserId $id -ErrorAction Stop -Property UserPrincipalName
    }
    catch {
        # Write-Verbose "  -> [$([math]::Round($percentComplete,2))%] User [$id] not found."
        Write-Verbose "  -> [$([math]::Round($percentComplete,2))%] $($_.Exception.Message)."
        continue
        # Write-Verbose "NOT SHOWN"
    }

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

    $messages = @()
    try {
        $messages = @(Get-MgUserMessage @searchParam -ErrorAction Stop)
    }
    catch {
        Write-Warning "Failed to retrieve messages for $id"
        Write-Warning $_.Exception.Message
        # break
    }

    if (-not $messages) {
        Write-Verbose "  -> [$([math]::Round($percentComplete,2))%] No match in [$($id)]"
        # continue
    }
    else {
        $totalMessageCount += $messages.Count
        $totalMailboxMatchCount += 1
        Write-Verbose "  -> [$([math]::Round($percentComplete,2))%] Found $($messages.Count) matches in [$($id)]"
    }

    Write-Progress -Activity "Searching mailbox body for [$($SearchString)]... $([math]::Round($percentComplete,2))%" -Status "[$($counter)/$($total))] $($id)" -PercentComplete $prePercentComplete

    foreach ($msg in $messages) {
        $result = @()
        $attachments = @()
        $bodyText = [System.Net.WebUtility]::HtmlDecode(
            ($msg.Body.Content -replace '<[^>]+>', ' ')
        )

        if (-not $msg.HasAttachments -or $bodyText -match [regex]::Escape($SearchString)) {
            $matchLocation = "Body"
        }
        else {
            $matchLocation = "Attachment"
            if ($ShowAttachment) {
                Write-Debug "Getting attachments..."
                $attachments = Get-MgUserMessageAttachment -UserId $id -MessageId $msg.Id
            }
        }

        $result = [ordered]@{
            Mailbox            = $id
            Subject            = $msg.Subject
            From               = $msg.From.EmailAddress.Address
            ReceivedDate       = $msg.ReceivedDateTime
            FirstMatchLocation = $matchLocation
            MessageId          = $msg.Id
        }

        if ($ShowAttachment) {
            $result += @{
                AttachmentNames = ($attachments).Name -join ","
                AttachmentIds   = ($attachments).Id -join ","
            }
        }
        New-Object psobject -Property $result
    }
}
Write-Progress -Activity '' -Completed -PercentComplete 100
$endTime = Get-Date
$totalRunTime = New-TimeSpan -Start $startTime -End $endTime

"" | Out-Default
"===============================================" | Out-Default
"SUMMARY" | Out-Default
"===============================================" | Out-Default
"Total mailbox searched : $($UserId.Count)" | Out-Default
"Total mailbox matched  : $($totalMailboxMatchCount)" | Out-Default
"Total message matched  : $($totalMessageCount)" | Out-Default
"Total run time         : $($totalRunTime.ToString())" | Out-Default