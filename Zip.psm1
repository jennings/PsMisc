
# Thanks to Dave Aiken
# http://blogs.msdn.com/b/daiken/archive/2007/02/12/compress-files-with-windows-powershell-then-package-a-windows-vista-sidebar-gadget.aspx

function New-Zip
{
    param([string]$zipfilename)

    Set-Content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
    (Get-ChildItem $zipfilename).IsReadOnly = $false
}

function Add-Zip
{
    param([string]$zipfilename)

    if(-not (Test-Path($zipfilename)))
    {
        Set-Content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (Get-ChildItem $zipfilename).IsReadOnly = $false
    }

    $shellApplication = New-Object -ComObject Shell.Application
    $zipPackage = $shellApplication.NameSpace("$zipfilename")

    foreach($file in $input) 
    {
        $zipPackage.CopyHere($file.FullName)
        $size = $zipPackage.Items().Item($file.FullName).Size
        while ($zipPackage.Items().Item($file.Name) -Eq $null)
        {
            Start-Sleep -Seconds 1
            Write-Host "." -NoNewline
        }
    }
}

function Get-Zip
{
    param([string]$zipfilename)

    if(Test-Path($zipfilename))
    {
        $shellApplication = New-Object -ComObject Shell.Application
        $zipPackage = $shellApplication.NameSpace($zipfilename)
        $zipPackage.Items() | Select-Object Path
    }
}

function Extract-Zip
{
    param([string]$zipfilename, [string] $destination)

    if(Test-Path($zipfilename))
    {
        $shellApplication = New-Object -ComObject Shell.Application
        $zipPackage = $shellApplication.NameSpace($zipfilename)
        $destinationFolder = $shellApplication.NameSpace($destination)
        $destinationFolder.CopyHere($zipPackage.Items())
    }
}

Export-ModuleMember Get-Zip, New-Zip, Add-Zip, Extract-Zip
