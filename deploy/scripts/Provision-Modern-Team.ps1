<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/New-PnPSite.md

.SYNOPSIS
Provision a new Modern Team Site - does not yet set permission

Based on PnP: New-PnPSite commandlet
	New-PnPSite -Type <SiteType>
            -Title <String>
            -Alias <String>
            [-Description <String>]
            [-Classification <String>]
            [-IsPublic <String>]
            [-Connection <SPOnlineConnection>]

.EXAMPLE
PS C:\> .\Provision-Modern-Team.ps1 -Tenant "pixelmilldev1" -Title "Modern Team Site 1" -Alias "ModernTeam1"

.EXAMPLE
PS C:\> .\Provision-Modern-Team.ps1 -Tenant "pixelmilldev1" -Title "Modern Team Site 2" -Alias "ModernTeam2" -Description "A modern team site" -Classification "HB1" -IsPublic -Credentials $creds
#>


[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

    [Parameter(Mandatory = $true, HelpMessage="Team site Title")]
    [String]
    $Title,

	[Parameter(Mandatory = $true, HelpMessage="Team site Alias")]
    [String]
    $Alias,

	[Parameter(Mandatory = $false, HelpMessage="Team site Description")]
    [String]
    $Description,

	[Parameter(Mandatory = $false, HelpMessage="Team site Classification")]
    [String]
    $Classification,

	[Parameter(Mandatory = $false, HelpMessage="Public Site?")]
    [Switch]
    $IsPublic,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url "https://$($Tenant).sharepoint.com"

Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host -ForegroundColor White "|              Provision Modern Team Site              |"
Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Provisioning a Modern Team Site: $($Title)"
Write-Host ""

try
{
	$TenantUrl = "https://" + $Tenant + ".sharepoint.com"

	#connent to site
	Write-Host -ForegroundColor White "Connecting to the tenant: $($TenantUrl)"
	Connect-PnPOnline -Url $TenantUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to the tenant"

	$exp = '$web = New-PnPSite -Type TeamSite -Title $($Title) -Alias $($Alias)'

	#Description
	if (![string]::IsNullOrEmpty($Description))
	{
		$exp += ' -Description "' + $Description + '"'
	}

	#Classification
	if (![string]::IsNullOrEmpty($Classification))
	{
		$exp += ' -Classification "' + $Classification + '"'
	}

	#IsPublic
	if ($IsPublic)
	{
		$exp += ' -IsPublic'
	}

	Write-Host ""
	Write-Host -ForegroundColor White "Going to create a modern team site using the following command:"
	Write-Host $exp

	Invoke-Expression $exp

	#display the newly created Communcation Site
	$web

	Write-Host -ForegroundColor White "Site Created"

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}