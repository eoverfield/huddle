<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://github.com/SharePoint/PnP-PowerShell/blob/master/Documentation/Add-PnPOffice365GroupToSite.md

as of 11/29/2017 does not appear to be available just yet

.SYNOPSIS
Groupify a Classic site

Based on PnP: Add-PnPOffice365GroupToSite commandlet
	Add-PnPOffice365GroupToSite -Alias <String>
			-DisplayName <String>
			[-Description <String>]
			[-Classification <String>]
			[-IsPublic [<SwitchParameter>]]
			[-Connection <SPOnlineConnection>]

.EXAMPLE
PS C:\> .\Groupify-Classic.ps1 -TargetWebUrl "" -Alias "" -DisplayName "" -Description "" -Classification "" -IsPublic 

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Groupify-Classic.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam2" -Alias "ClassicTeam2" -DisplayName "Classic Team 2" -Description "ClassicTeam2 desc" -IsPublic -Credentials $creds
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $true, HelpMessage="Url")]
    [String]
    $TargetWebUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $true, HelpMessage="Alias")]
    [String]
    $Alias,

	[Parameter(Mandatory = $true, HelpMessage="Display Name")]
    [String]
    $DisplayName,

    [Parameter(Mandatory = $false, HelpMessage="Description")]
    [String]
    $Description,

	[Parameter(Mandatory = $false, HelpMessage="Classification")]
    [String]
    $Classification,

	[Parameter(Mandatory = $false, HelpMessage="Is Public?")]
    [Switch]
    $IsPublic
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetWebUrl

Write-Host -ForegroundColor White "-------------------------------------------------------"
Write-Host -ForegroundColor White "|              Groupify an existing site              |"
Write-Host -ForegroundColor White "-------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target site: $($TargetWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to the tenant: $($TargetWebUrl)"
	Connect-PnPOnline -Url $TargetWebUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to the tenant"

	$exp = '$web = Add-PnPOffice365GroupToSite -Alias $($Alias) -DisplayName $($DisplayName)'

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
	Write-Host -ForegroundColor White "Going to groupify site using the following command:"
	Write-Host $exp

	Invoke-Expression $exp

	Write-Host -ForegroundColor Green "Group created"
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