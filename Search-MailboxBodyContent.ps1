[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]
    $UserId,

    # Parameter help description
    [Parameter(Mandatory)]
    [string]
    $SearchString
)

Get-MgUserMessage -Search "body:$($searchString)" -UserId $UserId -All | Select-Object subject,
@{n = "From"; e = { $_.from.emailAddress.address } },
receivedDateTime,
bodyPreview