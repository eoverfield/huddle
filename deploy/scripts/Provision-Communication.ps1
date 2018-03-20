<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/New-PnPSite.md

.SYNOPSIS
Provision a new Communication Site

Based on PnP: New-PnPSite commandlet
	New-PnPSite -Type <SiteType>
            -Title <String>
            -Url <String>
            [-Description <String>]
            [-Classification <String>]
            [-AllowFileSharingForGuestUsers [<SwitchParameter>]]
            [-SiteDesign <CommunicationSiteDesign>] //Topic Showcase Blank
			-SiteDesignId <GuidPipeBind>
            [-Lcid <UInt32>]
            [-Connection <SPOnlineConnection>]

.EXAMPLE
PS C:\> .\Provision-Communication.ps1 -Tenant "pixelmilldev1" -Title "Comm Site 1" -Alias "CommSite1"

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Provision-Communication.ps1 -Tenant "pixelmilldev1" -Title "Comm Site 20" -Alias "CommSite20" -Description "A comm site" -Credentials $creds
PS C:\> .\Provision-Communication.ps1 -Tenant "pixelmilldev1" -Title "Comm Site 21" -Alias "CommSite21" -Description "A comm site" -SiteDesign Topic -Classification "HB1" -AllowFileSharingForGuestUsers -Credentials $creds
#>


[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

    [Parameter(Mandatory = $true, HelpMessage="Communication site Title")]
    [String]
    $Title,

	[Parameter(Mandatory = $true, HelpMessage="Communication site Alias")]
    [String]
    $Alias,

	[Parameter(Mandatory = $false, HelpMessage="Communication site Description")]
    [String]
    $Description,

	[Parameter(Mandatory = $false, HelpMessage="Communication site Classification")]
    [String]
    $Classification,

	[Parameter(Mandatory = $false, HelpMessage="Allow File Sharing For Guest Users?")]
    [Switch]
    $AllowFileSharingForGuestUsers,

	[Parameter(Mandatory = $false, HelpMessage="Communication site design")]
    [String]
    $SiteDesign,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url "https://$($Tenant).sharepoint.com"

if([string]::IsNullOrEmpty($SiteDesign))
{
	$SiteDesign = "Blank"
}

Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host -ForegroundColor White "|             Provision Communication Site             |"
Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Provisioning a Communication Site: $($Title)"
Write-Host ""

try
{
	$TenantUrl = "https://" + $Tenant + ".sharepoint.com"
	$Url = "https://" + $Tenant + ".sharepoint.com/sites/" + $Alias

	#connent to site
	Write-Host -ForegroundColor White "Connecting to the tenant: $($TenantUrl)"
	Connect-PnPOnline -Url $TenantUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to the tenant"

	$exp = '$web = New-PnPSite -Type CommunicationSite -Title $($Title) -Url $($Url)'

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

	#SiteDesign
	if (![string]::IsNullOrEmpty($SiteDesign))
	{
		$exp += ' -SiteDesign "' + $SiteDesign + '"'
	}

	#AllowFileSharingForGuestUsers
	if ($AllowFileSharingForGuestUsers)
	{
		$exp += ' -AllowFileSharingForGuestUsers'
	}

	Write-Host ""
	Write-Host -ForegroundColor White "Going to create a communcation site using the following command:"
	Write-Host $exp

	Invoke-Expression $exp

	#display the newly created Communcation Site
	$web

	Write-Host -ForegroundColor Green "Site Created"

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}