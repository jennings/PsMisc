# .SYNOPSIS
#
# Provides helper functions to manage Exchange Online
#
# .NOTES
#
# This module is licensed under the Apache License, Version 2.0.
# http://www.apache.org/licenses/LICENSE-2.0


function New-SharedMailbox($Alias = (Read-Host -Prompt "Shared Mailbox Alias"),
                           $DisplayName = (Read-Host -Prompt "Shared Mailbox Display Name"))
{
    # .SYNOPSIS
    #
    # Creates a new shared mailbox

    if (!$Alias)
    {
        Write-Error "No alias specified."
        return
    }

    if (!$DisplayName)
    {
        Write-Error "No display name specified."
        return
    }

    # Create the mailbox

    Write-Host "Creating mailbox..."
    New-Mailbox -Name $DisplayName -Alias $Alias -Shared

    Write-Host "Setting quotas on mailbox..."
    Set-Mailbox $Alias -ProhibitSendReceiveQuota 5GB `
                       -ProhibitSendQuota 4.75GB `
                       -IssueWarningQuota 4.5GB `
                       | Out-Null

    # Set up a security group for granting access to this mailbox

    $GroupDisplayName = "Security Group - $DisplayName"
    $GroupAlias = "SecurityGroup-$($Alias)"

    Write-Host "Creating security group..."
    New-DistributionGroup -Name $GroupDisplayName `
                          -Alias $GroupAlias `
                          -Type Security `
                          | Out-Null

    Write-Host "Setting options on the security group..."
    Set-DistributionGroup $GroupAlias -RequireSenderAuthenticationEnabled $true `
                                      -AcceptMessagesOnlyFrom $Alias `
                                      -HiddenFromAddressListsEnabled $true

    Write-Host "Setting permissions on mailbox..."
    Add-MailboxPermission $DisplayName -User $GroupAlias -AccessRights FullAccess
    Add-RecipientPermission $DisplayName -Trustee $GroupAlias -AccessRights SendAs

    # Tell the user what we did

    Write-Host "Shared mailbox '$DisplayName' created."
    Write-Host "Add users as members to the '$GroupDisplayName' distribution group to grant access."
}
