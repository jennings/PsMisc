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
    # Creates a new shared mailbox. Optionally sets the domain names
    # that are valid for the mailbox.

    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Alias for the shared mailbox (alias@example.com).")]
        [Alias("Name")]
        [string]$Alias,

        [Parameter(
            Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Display name for the shared mailbox.")]
        [string]$DisplayName,

        [Parameter(
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="Primary domain name for this mailbox.")]
        [string]$PrimaryDomainName,

        [Parameter(
            ValueFromPipelineByPropertyName=$true,
            HelpMessage="List of additional domain names for this mailbox.")]
        [string[]]$AdditionalDomainNames
    )

    PROCESS
    {
        $myPrimaryDomainName = $PrimaryDomainName
        $myAdditionalDomainNames = $AdditionalDomainNames
        $myAlias = $Alias
        $myDisplayName = $DisplayName
        $myGroupDisplayName = "Access Group for $myDisplayName@$myPrimaryDomainName"
        $myGroupAlias = "$myAlias-accessgroup"

        if (!$myAlias)
        {
            Write-Error "No alias specified."
            return
        }

        if (!$myDisplayName)
        {
            Write-Error "No display name specified."
            return
        }


        # Create the mailbox and set domain names

        Write-Verbose "Creating shared mailbox '$myAlias'."

        New-Mailbox -Name $myDisplayName -Alias $myAlias -Shared

        if ($myPrimaryDomainName)
        {
            $addresslist = ,"SMTP:$myAlias@$myPrimaryDomainName"
            $myAdditionalDomainNames | ForEach-Object { $addresslist += "$myAlias@$_" }
            Set-Mailbox $myAlias -EmailAddresses $addresslist
        }



        # Set quotas (Office 365 limits shared mailboxes to 5 GB)

        Write-Verbose "Setting quotas on mailbox '$myAlias'."

        Set-Mailbox $myAlias -ProhibitSendReceiveQuota 5GB `
                             -ProhibitSendQuota 4.75GB `
                             -IssueWarningQuota 4.5GB `
                             | Out-Null



        # Set up a security group for granting access to this mailbox

        Write-Verbose "Creating security group '$myGroupAlias'."

        New-DistributionGroup -Name $myGroupDisplayName `
                              -Alias $myGroupAlias `
                              -Type Security `
                              | Out-Null

        Write-Verbose "Setting options on security group '$myGroupAlias'."

        Set-DistributionGroup $myGroupAlias -RequireSenderAuthenticationEnabled $true `
                                            -AcceptMessagesOnlyFrom $myAlias `
                                            -HiddenFromAddressListsEnabled $true



        # Finally, grant the security group access to the mailbox

        Write-Verbose "Setting permissions on mailbox '$myAlias'."

        Add-MailboxPermission $myDisplayName -User $myGroupAlias -AccessRights FullAccess
        Add-RecipientPermission $myDisplayName -Trustee $myGroupAlias -AccessRights SendAs -Confirm:$false
    }
}
