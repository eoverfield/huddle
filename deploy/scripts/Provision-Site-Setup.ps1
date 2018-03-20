<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Configure an existing site with solutions

.EXAMPLE
PS C:\> .\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -ProvisionIA -ProvisionAssets -ProvisionPages -ActivateExtensions -ActivateSiteSettings -ProvisionTheme -ProvisionLogo

.EXAMPLE
PS C:\> $creds = Get-Credential
PS C:\> $creds = Get-PnPStoredCredential -Name "yourStoredCredentialName" -Type PSCredential 
PS C:\> .\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials $creds
PS C:\> .\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials "yourStoredCredentialName"
PS C:\> .\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/collabcomm194" -Credentials $creds -ProvisionIA -ProvisionAssets -ProvisionPages -ActivateExtensions -ActivateSiteSettings -ProvisionTheme -ProvisionLogo
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the URL of the target site to set up, e.g. 'https://intranet.mydomain.com/sites/apps'")]
    [String]
    $TargetWebUrl,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Provision Information Architecture")]
    [Switch]
    $ProvisionIA,

	[Parameter(Mandatory = $false, HelpMessage="Provision site and branding assets")]
    [Switch]
    $ProvisionAssets,

	[Parameter(Mandatory = $false, HelpMessage="Provision content pages")]
    [Switch]
    $ProvisionPages,

	[Parameter(Mandatory = $false, HelpMessage="Activate SPFx Extensions")]
    [Switch]
    $ActivateExtensions,

	[Parameter(Mandatory = $false, HelpMessage="Activate Site Setttings")]
    [Switch]
    $ActivateSiteSettings,

	[Parameter(Mandatory = $false, HelpMessage="Provision themeing")]
    [Switch]
    $ProvisionTheme,

	[Parameter(Mandatory = $false, HelpMessage="Provision logo")]
    [Switch]
    $ProvisionLogo
)

#includes
. "./utilities.ps1"

$Credentials = Verify-Credentials -Credentials $Credentials -Url $TargetWebUrl

Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host -ForegroundColor White "|             Site Setup and Configuration             |"
Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target site: $($TargetWebUrl)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to web"
	Connect-PnPOnline $TargetWebUrl -Credentials $Credentials
	Write-Host -ForegroundColor Green "Connected to $($TargetWebUrl)"

	#Information Architecture
	if ($ProvisionIA -eq $true)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Provisioning Information Architecture"
		Apply-PnPProvisioningTemplate .\templates\Site.Setup.IA.xml
		Write-Host -ForegroundColor Green "Provisioning Information Architecture: complete"
	}

	#Assets
	if ($ProvisionAssets -eq $true)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Provisioning Assets"
		Apply-PnPProvisioningTemplate .\templates\Site.Setup.Assets.xml
		Write-Host -ForegroundColor Green "Provisioning Assets: complete"
	}

	#Pages
	if ($ProvisionPages -eq $true)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Add pages"
		Apply-PnPProvisioningTemplate .\templates\Site.Setup.Pages.xml
		Write-Host -ForegroundColor Green "Add pages: complete"
	}

	#Extensions
	if ($ActivateExtensions -eq $true)
	{	
		Write-Host ""
		Write-Host -ForegroundColor White "Activating SPFx Extensions"
		Apply-PnPProvisioningTemplate .\templates\Site.Setup.Extensions.xml
		Write-Host -ForegroundColor Green "Activating SPFx Extensions: complete"
	}

	#Site Settings
	if ($ActivateSiteSettings -eq $true)
	{	
		Write-Host ""
		Write-Host -ForegroundColor White "Update site settings"
		Apply-PnPProvisioningTemplate .\templates\Site.Setup.Settings.xml
		Write-Host -ForegroundColor Green "Update site settings: complete"
	}

	#Set up theme
	if ($ProvisionTheme -eq $true)
	{	
		Write-Host ""
		Write-Host -ForegroundColor White "Applying classic custom theme"
		$web = Get-PnPWeb
		$palette = $web.ServerRelativeUrl + "/SiteAssets/PaletteCustom001.spcolor"
		#$background = $web.ServerRelativeUrl + "/SiteAssets/sppnp-bg.png"
		$logo = $web.ServerRelativeUrl + "/SiteAssets/PnP.png"

		#classic logo
		Set-PnPWeb -SiteLogoUrl $logo

		# We use OOTB CSOM operation for this
		#$web.ApplyTheme($palette, [NullString]::Value, $background, $true)
		$web.ApplyTheme($palette, [NullString]::Value, [NullString]::Value, $true)
		$web.Update()
		# Set timeout as high as possible and execute
		$web.Context.RequestTimeout = [System.Threading.Timeout]::Infinite
		$web.Context.ExecuteQuery()
		Write-Host -ForegroundColor Green "Applying classic custom theme: complete"
	}

	#Set up Group based Logo
	if ($ProvisionLogo -eq $true)
	{	
		Write-Host -ForegroundColor White "Set the group based logo"

		#Connect-PnPMicrosoftGraph -Scopes "Group.ReadWrite.All","User.Read.All"
		Connect-PnPOnline -Scopes "Group.ReadWrite.All","User.Read.All"
		#Connect-PnPOnline -AppId '<id>' -AppSecret '<secrect>' -AADDomain 'contoso.onmicrosoft.com'
		#Connect-PnPOnline -AppId 368e2766-c85b-434e-a222-69f238b6a7d1 -AppSecret f+P5zmTvBalckI5iKq3/hazN3vfbJCR+dEJUAbDDyA8= -AADDomain pixelmilldev1.onmicrosoft.com
		#Connect-PnPOnline -Url $TargetWebUrl -AppId 368e2766-c85b-434e-a222-69f238b6a7d1 -AppSecret f+P5zmTvBalckI5iKq3/hazN3vfbJCR+dEJUAbDDyA8=

		#https://login.microsoftonline.com/pixelmilldev1.onmicrosoft.com/oauth2/authorize?client_id=368e2766-c85b-434e-a222-69f238b6a7d1&resource=https://graph.microsoft.com&redirect_uri=https://pixelmilldev1.sharepoint.com/CollabCommPlaygroup&response_type=code&prompt=admin_consent

		Write-Host -ForegroundColor White "1"
		$groups = Get-PnPUnifiedGroup 
		Write-Host -ForegroundColor White "2"
		$group = $groups | Where {$_.SiteURL -eq $web.Url}
		Write-Host -ForegroundColor White "3"
		Set-PnPUnifiedGroup -Identity $group -GroupLogoPath ".\templates\SiteAssets\PnP.png"

		Write-Host -ForegroundColor White "Set the group based logo: complete"
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Provisioning Completed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}