# .SYNOPSIS
#
# Provides helper functions to manage Exchange Online
#
# .NOTES
#
# This module is licensed under the Apache License, Version 2.0.
# http://www.apache.org/licenses/LICENSE-2.0


function New-SharedMailbox
{
    # .SYNOPSIS
    #
    # Creates a new shared mailbox

    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Shared Mailbox Alias")]
        [Alias("Name")]
        [string]
        $Alias,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Shared Mailbox Display Name")]
        [string]
        $DisplayName
    )

    PROCESS {

        if ($pscmdlet.ShouldProcess())
        {
            # Create the mailbox
            Write-Verbose "Creating mailbox..."
            New-Mailbox -Name $DisplayName -Alias $Alias -Shared
    
            Write-Verbose "Setting quotas on mailbox..."
            Set-Mailbox $Alias -ProhibitSendReceiveQuota 5GB `
                               -ProhibitSendQuota 4.75GB `
                               -IssueWarningQuota 4.5GB `
                               | Out-Null
    
            # Set up a security group for granting access to this mailbox
    
            $GroupDisplayName = "Security Group - $DisplayName"
            $GroupAlias = "SecurityGroup-$($Alias)"
    
            Write-Verbose "Creating security group..."
            New-DistributionGroup -Name $GroupDisplayName `
                                  -Alias $GroupAlias `
                                  -Type Security `
                                  | Out-Null
    
            Write-Verbose "Setting options on the security group..."
            Set-DistributionGroup $GroupAlias -RequireSenderAuthenticationEnabled $true `
                                              -AcceptMessagesOnlyFrom $Alias `
                                              -HiddenFromAddressListsEnabled $true
    
            Write-Verbose "Setting permissions on mailbox..."
            Add-MailboxPermission $DisplayName -User $GroupAlias -AccessRights FullAccess
            Add-RecipientPermission $DisplayName -Trustee $GroupAlias -AccessRights SendAs
    
            # Tell the user what we did
    
            Write-Verbose "Shared mailbox '$DisplayName' created."
            Write-Verbose "Add users as members to the '$GroupDisplayName' distribution group to grant access."
        }

    }
}
