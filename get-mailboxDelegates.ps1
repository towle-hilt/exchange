# get-mailboxDelegates.ps1
Param(
    [Parameter(Position=0, Mandatory=$true)]
    [string] $ConnectionUri
)

$exchsession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri
Import-PSSession $exchsession
Import-Module ActiveDirectory

Write-Progress -Activity "Delegate mailbox search initiated" -Status "Finding mailboxes"

$mailboxes = Get-Mailbox -ResultSize Unlimited | Sort-Object displayName

$i = 1
$delegates = @()

foreach ($mailbox in $mailboxes) {
    Write-Progress -Activity "Searching each mailbox for delegates" -PercentComplete ($i / $mailboxes.count * 100) -CurrentOperation $mailbox.displayName

    $delegates += Get-MailboxPermission $mailbox.SamAccountName | ?{ `
        $_.IsInherited -eq $False -and `
        $_.User -notlike "NT AUTHORITY\SELF" -and `
        $_.User -notlike "*\Exchange*" -and `
        $_.User -notlike "S-1-5*" } `
        | Select-Object `
            @{E={$mailbox.displayName};L="Name"}, `
            @{E={$mailbox.SamAccountName.ToLower()};L="Account"}, `
            @{E={$mailbox.PrimarySmtpAddress.ToLower()};L="Email"}, `
            @{E={$mailbox.RecipientTypeDetails};L="MailboxType"}, `
            @{E={(Get-ADUser ($_.User.Split("\")[-1])).Name};L="Delegate"}, `
            @{E={$_.User.Split("\")[-1].toLower()};L="DelegateAccount"}, `
            @{E={$_.AccessRights -join ","};L="Permissions"}

    $i++
}

$delegates | Sort-Object DisplayName,Delegate
