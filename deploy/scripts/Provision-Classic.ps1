<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/New-PnPTenantSite.md

.SYNOPSIS
Provision a new classic Team Site - team or publishing

Based on PnP: New-PnPTenantSite commandlet
	New-PnPTenantSite -Title <String>
			-Url <String>
			-Owner <String>
			-TimeZone <Int>
			[-Description <String>]
			[-Lcid <UInt32>]
			[-Template <String>]
			[-ResourceQuota <Double>]
			[-ResourceQuotaWarningLevel <Double>]
			[-StorageQuota <Int>]
			[-StorageQuotaWarningLevel <Int>]
			[-RemoveDeletedSite [<SwitchParameter>]]
			[-Wait [<SwitchParameter>]]
			[-Force [<SwitchParameter>]]
			[-Connection <SPOnlineConnection>]

.EXAMPLE
PS C:\> .\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Modern Team Site 1" -Alias "ClassicTeam1" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0"

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Team Site 1" -Alias "ClassicTeam3" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0" -StorageQuota 100 -ResourceQuota 30 -Force -Credentials $creds
PS C:\> .\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Team Site 1" -Alias "ClassicTeam2" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0" -Force -Wait -Credentials $creds
PS C:\> .\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Publishing Site 1" -Alias "ClassicPub1" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "BLANKINTERNETCONTAINER#0" -Force -Wait -Credentials $creds
#>


[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Tenant Name")]
    [String]
    $Tenant,

	[Parameter(Mandatory = $true, HelpMessage="Template")]
    [String]
    $Template,

    [Parameter(Mandatory = $true, HelpMessage="Site Title")]
    [String]
    $Title,

	[Parameter(Mandatory = $true, HelpMessage="Site Alias")]
    [String]
    $Alias,

	[Parameter(Mandatory = $true, HelpMessage="Site Owner")]
    [String]
    $Owner,

	[Parameter(Mandatory = $true, HelpMessage="Site Time Zone")]
    [Int]
    $TimeZone,

	[Parameter(Mandatory = $false, HelpMessage="Team site Description")]
    [String]
    $Description,

	[Parameter(Mandatory = $false, HelpMessage="Resource Quota")]
    [Double]
    $ResourceQuota,

	[Parameter(Mandatory = $false, HelpMessage="Storage Quota")]
    [Double]
    $StorageQuota,

	[Parameter(Mandatory = $false, HelpMessage="Wait?")]
    [Switch]
    $Wait,

	[Parameter(Mandatory = $false, HelpMessage="Force?")]
    [Switch]
    $Force,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url "https://$($Tenant).sharepoint.com"

if([string]::IsNullOrEmpty($Description))
{
	$Description = ""
}
if([string]::IsNullOrEmpty($Template))
{
	#BLANKINTERNETCONTAINER#0
	$Template = "STS#0"
}
if([string]::IsNullOrEmpty($TimeZone))
{
	$TimeZone = 13
}
if([string]::IsNullOrEmpty($StorageQuota))
{
	$StorageQuota = 0
}
if([string]::IsNullOrEmpty($ResourceQuota))
{
	$ResourceQuota = 0
}

Write-Host -ForegroundColor White "--------------------------------------------------"
Write-Host -ForegroundColor White "|             Provision Classic Site             |"
Write-Host -ForegroundColor White "--------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target tenant: $($Tenant)"
Write-Host ""

try
{
	$TenantUrl = "https://" + $Tenant + ".sharepoint.com"
	$Url = "https://" + $Tenant + ".sharepoint.com/sites/" + $Alias

	#connent to site
	Write-Host -ForegroundColor White "Connecting to the tenant: $($TenantUrl)"
	Connect-PnPOnline -Url $TenantUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to the tenant"

	$exp = '$web = New-PnPTenantSite -Title $($Title) -Url $($Url) -Owner $($Owner) -TimeZone $($TimeZone)'

	#Description
	if(![string]::IsNullOrEmpty($Description))
	{
		$exp += ' -Description "' + $Description + '"'
	}

	#Template
	if(![string]::IsNullOrEmpty($Template))
	{
		$exp += ' -Template "' + $Template + '"'
	}

	#StorageQuota
	if ($StorageQuota -gt 0)
	{
		$exp += ' -StorageQuota ' + $StorageQuota
	}

	#ResourceQuota
	if ($ResourceQuota -gt 0)
	{
		$exp += ' -ResourceQuota ' + $ResourceQuota
	}

	#Wait
	if ($Wait)
	{
		$exp += ' -Wait'
	}

	#Force
	if ($Force)
	{
		$exp += ' -Force'
	}

	Write-Host ""
	Write-Host -ForegroundColor White "Going to create site using the following command:"
	Write-Host $exp

	Invoke-Expression $exp

	Write-Host -ForegroundColor Green "Site Created"
	$web

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}