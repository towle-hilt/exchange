<#
.PARAMETER Group
Name of group to recreate.
.PARAMETER ExchangeServer
Name of on-prem Microsoft Exchange server.
.PARAMETER ContactOU
DistinguishedName for an OU to place contact records. This OU should be excluded from Office 365 syncronizaion.
.PARAMETER AzureADSyncServer
Name of on-prem Azure AD Synchronization server.
.PARAMETER CloudCredential
.PARAMETER CloundDomain
Office 365 tenant domain tenant.mail.onmicrosoft.com
.PARAMETER OnPremCredential
.PARAMETER emailDomain
Email domain domain.com
.PARAMETER Stage
Create placeholder staging group.
.PARAMETER Sync
Sync clound group settings. Use if onprem group has changed since initial staging.
.PARAMETER Finalize
Convert staging group to final group.
.PARAMETER UpdateOffice365Group
Update Office 365 group naming scheme
.EXAMPLE #1
.\migrate-DistributionGroup.ps1 -Group "Marketing" -Verbose -cloudCredential $cloud -onPremCredential $onprem -Stage -Finalize
.EXAMPLE #2
.\migrate-DistributionGroup.ps1 -Group "Marketing" -Finalize
#>

Param(
    [Parameter(Mandatory=$True)]
    [string]$Group,
    [Parameter(Mandatory=$False)]
    [string]$ExchangeServer,
    [Parameter(Mandatory=$False)]
    [string]$ContactOU,
    [Parameter(Mandatory=$False)]
    [string]$AzureADSyncServer,
    [Parameter(Mandatory=$false)]
    [PSObject]$cloudCredential,
    [Parameter(Mandatory=$true)]
    [string]$cloudDomain,    
    [Parameter(Mandatory=$false)]
    [PSObject]$onPremCredential,
    [Parameter(Mandatory=$true)]
    [string]$onPremDomain,    
    [Parameter(Mandatory=$False)]
    [switch]$Stage,
    [Parameter(Mandatory=$False)]
    [switch]$Sync,
    [Parameter(Mandatory=$False)]
    [switch]$Finalize
)

$ExportDirectory = ".\ExportedAddresses\"

Start-Transcript -Path ("$ExportDirectory\$Group" + "_transcript.txt")

function get-PrimarySmtpAddress {
    param( [string]$Name )
    
    $mb = get-Mailbox -Identity $name -ErrorAction SilentlyContinue
    if ($mb) {
        $mb.PrimarySmtpAddress
    } else {
        $dl = Get-DistributionGroup -Identity $name -ErrorAction SilentlyContinue
        if ($dl) {
            $dl.PrimarySmtpAddress
        } else {
            $c = Get-MailContact -Identity $Name -ErrorAction SilentlyContinue
            if ($c) {
                $c.PrimarySmtpAddress
            } else {
                Write-Verbose "$Name does not have a valid mail object, check license"
            }
        }
    }
}

if (!($cloudCredential)) {
    $cloudCredential = Get-Credential -Message "Office 365 Credentials (outlook.office365.com)."
}

if ($Stage.IsPresent -or $Sync.IsPresent) {
    
    if ($Stage.IsPresent) { $Option = "Stage" } else { $Option = "Sync" }
    
    $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cloudCredential -Authentication Basic -AllowRedirection
    Import-PSSession $cloudSession -AllowClobber

    if (((Get-DistributionGroup $Group -ErrorAction 'SilentlyContinue').IsValid) -eq $true) {

        Write-Verbose "$Option process initiated"

        $OldDG = Get-DistributionGroup -Identity $Group

        $OldDG | Export-Csv ("$ExportDirectory\" + $OldDG.Name + "_backup_$Option.csv") -NoClobber -NoTypeInformation -Append
        Get-DistributionGroupMember -Identity $OldDG.Name | Export-Csv ("$ExportDirectory\" + $OldDG.Name + "_members_backup_$Option.csv") -NoClobber -NoTypeInformation -Append

        [System.IO.Path]::GetInvalidFileNameChars() | % { $Group = $Group.Replace($_,'_') }
        
        $OldName = [string]$OldDG.Name
        $OldDisplayName = [string]$OldDG.DisplayName
        $OldPrimarySmtpAddress = [string]$OldDG.PrimarySmtpAddress
        $OldAlias = [string]$OldDG.Alias
        $OldMembers = (Get-DistributionGroupMember -Identity $OldDG.Name | ?{$_.PrimarySmtpAddress -ne ""}).PrimarySmtpAddress
        
        if ($OldDG.ManagedBy) {
            $ManagedBy = $OldDG.ManagedBy | %{ Get-PrimarySmtpAddress $_ }
            if (!($ManagedBy)) {
                $ManagedBy = Get-PrimarySmtpAddress -Name administrator
            }
        } else {
            $ManagedBy = Get-PrimarySmtpAddress -Name administrator
        }

        if(!(Test-Path -Path $ExportDirectory )){
            Write-Verbose "... Creating directory $ExportDirectory"
            New-Item -ItemType directory -Path $ExportDirectory | Out-Null
        }

        "EmailAddress" > "$ExportDirectory\$Group.csv"
        $OldDG.EmailAddresses >> "$ExportDirectory\$Group.csv"
        "x500:"+$OldDG.LegacyExchangeDN >> "$ExportDirectory\$Group.csv"

        if ($Sync.IsPresent) {
            Write-Verbose "... Syncronizing group membership"
            $LocalMembers = Get-DistributionGroupMember -Identity $OldDG.Name
            $CloudMembers = Get-DistributionGroupMember -Identity "cloud-$OldName"

            $AddMembers = (Compare-Object -ReferenceObject $CloudMembers -DifferenceObject $LocalMembers | ?{$_.SideIndicator -eq "=>"}).InputObject
            $RemMembers = (Compare-Object -ReferenceObject $CloudMembers -DifferenceObject $LocalMembers | ?{$_.SideIndicator -eq "<="}).InputObject

            if ($AddMembers.count -gt 0) {
                $AddMembers | % { Add-DistributionGroupMember `
                    -Identity "cloud-$OldName" `
                    -Member $_.Name
                }
            } else {
                Write-Verbose "... no new members"
            }

            if ($RemMembers.count -gt 0) {
                $RemMembers | % { Remove-DistributionGroupMember `
                    -Identity "cloud-$OldName" `
                    -Member $_.Name
                }
            } else {
                Write-Verbose "... no members to remove"
            }

        } else {
            Write-Verbose "... Creating group cloud-$OldDisplayName"
    
            New-DistributionGroup `
                -Name "cloud-$OldName" `
                -Alias "cloud-$OldAlias" `
                -DisplayName "cloud-$OldDisplayName" `
                -ManagedBy $ManagedBy `
                -Members $OldMembers | Out-Null
        }

        while ($null -eq (Get-DistributionGroup -Identity "cloud-$OldName" -ErrorAction SilentlyContinue)) {
            Write-Verbose "... Waiting to connect to distribution group"
            Start-Sleep -Seconds 3
        }

        Write-Verbose "... Setting values for cloud-$OldDisplayName"

        Set-DistributionGroup `
            -Identity "cloud-$OldName" `
            -AcceptMessagesOnlyFromSendersOrMembers ($OldDG.AcceptMessagesOnlyFromSendersOrMembers | %{ get-PrimarySmtpAddress -Name $_ -Verbose }) `
            -RejectMessagesFromSendersOrMembers ($OldDG.RejectMessagesFromSendersOrMembers | %{get-PrimarySmtpAddress -Name $_ -Verbose}) `

        Set-DistributionGroup `
            -Identity "cloud-$OldName" `
            -AcceptMessagesOnlyFrom ($OldDG.AcceptMessagesOnlyFrom | %{get-PrimarySmtpAddress -Name $_}) `
            -AcceptMessagesOnlyFromDLMembers ($OldDG.AcceptMessagesOnlyFromDLMembers | %{get-PrimarySmtpAddress -Name $_}) `
            -BypassModerationFromSendersOrMembers ($OldDG.BypassModerationFromSendersOrMembers | %{get-PrimarySmtpAddress -Name $_}) `
            -BypassNestedModerationEnabled $OldDG.BypassNestedModerationEnabled `
            -CustomAttribute1 $OldDG.CustomAttribute1 `
            -CustomAttribute2 $OldDG.CustomAttribute2 `
            -CustomAttribute3 $OldDG.CustomAttribute3 `
            -CustomAttribute4 $OldDG.CustomAttribute4 `
            -CustomAttribute5 $OldDG.CustomAttribute5 `
            -CustomAttribute6 $OldDG.CustomAttribute6 `
            -CustomAttribute7 $OldDG.CustomAttribute7 `
            -CustomAttribute8 $OldDG.CustomAttribute8 `
            -CustomAttribute9 $OldDG.CustomAttribute9 `
            -CustomAttribute10 $OldDG.CustomAttribute10 `
            -CustomAttribute11 $OldDG.CustomAttribute11 `
            -CustomAttribute12 $OldDG.CustomAttribute12 `
            -CustomAttribute13 $OldDG.CustomAttribute13 `
            -CustomAttribute14 $OldDG.CustomAttribute14 `
            -CustomAttribute15 $OldDG.CustomAttribute15 `
            -ExtensionCustomAttribute1 $OldDG.ExtensionCustomAttribute1 `
            -ExtensionCustomAttribute2 $OldDG.ExtensionCustomAttribute2 `
            -ExtensionCustomAttribute3 $OldDG.ExtensionCustomAttribute3 `
            -ExtensionCustomAttribute4 $OldDG.ExtensionCustomAttribute4 `
            -ExtensionCustomAttribute5 $OldDG.ExtensionCustomAttribute5 `
            -GrantSendOnBehalfTo $OldDG.GrantSendOnBehalfTo `
            -HiddenFromAddressListsEnabled $True `
            -MailTip $OldDG.MailTip `
            -MailTipTranslations $OldDG.MailTipTranslations `
            -MemberDepartRestriction $OldDG.MemberDepartRestriction `
            -MemberJoinRestriction $OldDG.MemberJoinRestriction `
            -ModeratedBy $OldDG.ModeratedBy `
            -ModerationEnabled $OldDG.ModerationEnabled `
            -RejectMessagesFrom $OldDG.RejectMessagesFrom `
            -RejectMessagesFromDLMembers $OldDG.RejectMessagesFromDLMembers `
            -ReportToManagerEnabled $OldDG.ReportToManagerEnabled `
            -ReportToOriginatorEnabled $OldDG.ReportToOriginatorEnabled `
            -RequireSenderAuthenticationEnabled $OldDG.RequireSenderAuthenticationEnabled `
            -SendModerationNotifications $OldDG.SendModerationNotifications `
            -SendOofMessageToOriginatorEnabled $OldDG.SendOofMessageToOriginatorEnabled `
            -BypassSecurityGroupManagerCheck

        Write-Verbose "... $Option complete"

    } else {
        Throw "ERROR: The distribution group '$Group' was not found"
    }
} 

if ($Finalize.IsPresent) {
    if ($Stage.IsPresent -or $Sync.IsPresent) { Sleep -Seconds 60 }

    Write-Verbose "Finalize process initiated"
    
    if (!($onPremCredential)) {
        $onPremCredential = Get-Credential -Message "OnPrem Exchange Credentials ($ExchangeServer)."
    }

    $onpremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $ExchangeServer + "/PowerShell/") -Credential $onpremCredential
    Import-PSSession $onpremSession -AllowClobber | Out-Null

        $OldDG = Get-DistributionGroup -Identity $Group -ErrorAction SilentlyContinue

        if (!($OldDG)) { Throw "ERROR: No group detected" }

        $OldDG | Export-Csv ("$ExportDirectory\" + $OldDG.Name + "_backup.csv") -NoClobber -NoTypeInformation -Append
        Get-DistributionGroupMember -Identity $OldDG.Name | Export-Csv ("$ExportDirectory\" + $OldDG.Name + "_members_backup.csv") -NoClobber -NoTypeInformation -Append

        if ($OldDG.count -gt 1) { Throw "ERROR: Multiple group items detected" }

        if ($OldDG) {
            Write-Verbose "... Cleanup onprem active directory."    
            if ($OldDG.GroupType -match "SecurityEnabled") {
                Write-Verbose "... Disable distribution group"
                $OldDG | Disable-DistributionGroup -Confirm:$False
                # might need to also rename and move the old group.
            } else {
                Write-Verbose "... Remove distribution group"
                $OldDG | Remove-DistributionGroup -Confirm:$False
            }
            Start-Sleep -Seconds 15
            Do {
                Write-Verbose "... Waiting for group removal to complete"
                Start-Sleep -Seconds 3
            } While (Get-DistributionGroup -Identity $OldDG.Name -ErrorAction SilentlyContinue)
        } else {
            Write-Verbose "... Unable to locate onprem group"
        }

        Write-Verbose "... Creating local contact record"
        $ExternalSMTPAddress = $OldDG.PrimarySmtpAddress.Replace("@$emailDomain","@$cloudDomain")
        Do {
            New-MailContact `
                -Name $OldDG.DisplayName `
                -DisplayName $OldDG.DisplayName `
                -ExternalEmailAddress $ExternalSMTPAddress `
                -Alias $OldDG.PrimarySmtpAddress.Replace("@emailDomain","") `
                -OrganizationalUnit $ContactOU | Out-Null

            Start-Sleep -Seconds 15
        }
        While (!(Get-MailContact -Identity $OldDG.DisplayName -ErrorAction SilentlyContinue))

        Set-MailContact `
            -Identity $OldDG.DisplayName `
            -EmailAddressPolicyEnabled $true

    Write-Verbose "... Initiating Azure AD sync cycle (this process will take a few moments)"
    Invoke-Command -ComputerName $AzureADSyncServer -Credential $onpremCredential -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }

    Start-Sleep -Seconds 180

    Write-Verbose "... Connecting to Office 365"
    Remove-PSSession $onpremSession
    $cloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cloudCredential -Authentication Basic -AllowRedirection
    Import-PSSession $cloudSession -AllowClobber | Out-Null

    Do {
        # pause until the old distribution group is no longer in Office 365
        Write-Verbose "... Waiting for old group to be removed from Office 365"
        Start-Sleep -Seconds 30
    } 
    While (Get-DistributionGroup -Identity $OldDG.Name -ErrorAction SilentlyContinue)

    Write-Verbose "... Updating temporary cloud group"
    $TempDG = Get-DistributionGroup -Identity "cloud-$Group"
    $TempPrimarySmtpAddress = $TempDG.PrimarySmtpAddress

    $OldAddresses = $OldDG.EmailAddresses
    $NewAddresses = $OldAddresses | ForEach {$_.Replace("X500","x500")}
    $NewDisplayName = $TempDG.displayName
    $NewDisplayName = ($NewDisplayName -creplace "[A-Z][a-z]|(?<!\s)\d{2,}"," $&").Trim()
    $NewDisplayName = $NewDisplayName -replace "^cloud-",""
    $NewDisplayName = $NewDisplayName -replace " (Department|List)$",""
    $NewDisplayName = $NewDisplayName -replace "[\'\.]",""
    $NewDisplayName = $NewDisplayName -replace "[\-\,\&_]"," "
    $NewDisplayName = $NewDisplayName -replace "\s{1,6}"," "
    $NewDisplayName = $NewDisplayName.Trim()
    $NewDisplayName = "DL " + $NewDisplayName + " List"
    $NewAlias = $NewDisplayName.ToLower().Replace(" ",".") -replace "^dl.",""
    $NewPrimarySmtpAddress = ($NewAddresses | Where-Object {$_ -clike "SMTP:*"}).Replace("SMTP:","")

    Set-DistributionGroup `
        -Identity $TempDG.Name `
        -Name $NewDisplayName `
        -Alias $NewAlias `
        -DisplayName $NewDisplayName `
        -PrimarySmtpAddress $NewPrimarySmtpAddress `
        -HiddenFromAddressListsEnabled $False `
        -BypassSecurityGroupManagerCheck
    Set-DistributionGroup `
        -Identity $NewDisplayName `
        -EmailAddresses @{Add=$NewAddresses} `
        -BypassSecurityGroupManagerCheck
    Set-DistributionGroup `
        -Identity $NewDisplayName `
        -EmailAddresses @{Remove=$TempPrimarySmtpAddress} `
        -BypassSecurityGroupManagerCheck

    Write-Verbose "... Group migration completed!"

}

Remove-PSSession $cloudSession

Stop-Transcript
