# .SYNOPSIS
#
# Provides the Connect-Msol function to connect to Microsoft Online
#
# .NOTES
#
# This module is licensed under the Apache License, Version 2.0.
# http://www.apache.org/licenses/LICENSE-2.0


function Connect-Msol($Credential = (Get-Credential -Message "Enter your Office 365 email address and password."))
{
    # .SYNOPSIS
    #
    # Connects a PowerShell session to Office 365 and Exchange Online.
    # 
    # .DESCRIPTION
    #
    # Connecting a PowerShell session to both Microsoft Online and
    # Exchange Online requires a few magic incantations. This function
    # simply does them for you.
    #
    # .EXAMPLE
    #
    # Connect-Msol
    #
    # Prompts for credentials then connects to Office 365.
    #
    # .EXAMPLE
    #
    # Connect-Msol -Credential $credential
    #
    # Connects using a previously-created credential.

    if (!$Credential)
    {
        Write-Error "No credential specified."
        return
    }

    Copy-Item Function:prompt Function:prompt-beforemsol
    (Get-Host).UI.RawUI.WindowTitle = "Office 365: " + $Credential.UserName
    
    Import-Module MSOnline -Scope Global
    Connect-MsolService -Credential $Credential
    
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri https://ps.outlook.com/powershell/ `
                             -Credential $Credential `
                             -Authentication Basic `
                             -AllowRedirection
    Import-Module (Import-PSSession $session) -Scope Global
}

function Disconnect-Msol
{
    # .SYNOPSIS
    #
    # Disconnects a PowerShell session from Office 365.
    # 
    # .DESCRIPTION
    #
    # Disconnects a PowerShell session started with Connect-Msol. Removes the
    # cmdlet modules and restores the prompt.
    #

    if (Function:prompt-beforemsol)
    {
        Copy-Item Function:prompt-beforemsol Function:prompt
        Remove-Item Function:prompt-beforemsol
    }

    Remove-Module MSOnline
    Remove-Module "tmp_*"
    Remove-PSSession
}
