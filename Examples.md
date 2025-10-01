# Examples

## Search within single mailbox

```PowerShell
$searchString = "confidential"
$mailboxToSearch = "dummy@poshlab.xyz"
.\Search-MailboxBodyContent.ps1 -UserId $mailboxToSearch -SearchString $searchString
```

## Search within a single mailbox and display attachment details when a match is found inside the attachment

```PowerShell
$searchString = "confidential"
$mailboxToSearch = "dummy@poshlab.xyz"
.\Search-MailboxBodyContent.ps1 -UserId $mailboxToSearch -SearchString $searchString -ShowAttachment
```

## Search within multiple mailbox

```PowerShell
$searchString = "confidential"
$mailboxToSearch = @("dummy@poshlab.xyz", "patrick@poshlab.xyz")
.\Search-MailboxBodyContent.ps1 -UserId $mailboxToSearch -SearchString $searchString
```

## Search within all user mailbox

```PowerShell
$searchString = "confidential"
$mailboxToSearch = (Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited).UserPrincipalName
.\Search-MailboxBodyContent.ps1 -UserId $mailboxToSearch -SearchString $searchString
```
