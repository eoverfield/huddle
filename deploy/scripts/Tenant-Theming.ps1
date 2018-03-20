<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Based on documentation:
https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-powershell

.SYNOPSIS
Add, remove, and work with modern themes at tenant level
Requires tenant admin level access

.EXAMPLE
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" 

.EXAMPLE
PS C:\> $credsSPO = Get-Credential
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -HideDefaultThemes
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -ShowDefaultThemes
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddTheme
PS C:\> .\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RemoveTheme
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Enter the organization name, e.g. 'pixelmilldev1'")]
    [String]
    $OrgName,

	[Parameter(Mandatory = $false, HelpMessage="Optional administration credentials")]
    [PSCredential]
    $Credentials,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Hide default themes?")]
    [Switch]
    $HideDefaultThemes,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Show default themes?")]
    [Switch]
    $ShowDefaultThemes,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Add a custom theme?")]
    [Switch]
    $AddTheme,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Remove custom theme?")]
    [Switch]
    $RemoveTheme
)

#includes
. "./utilities.ps1"

$Credentials = Verify-SPOCredentials -Credentials $Credentials -Url "https://$($orgName)-admin.sharepoint.com"

#custom theme palette - get from: aka.ms/spthemebuilder
$themepalette = HashToDictionary(
	@{
	"themePrimary" = "#00a6cb";
	"themeLighterAlt" = "#f0fcff";
	"themeLighter" = "#e0f9ff";
	"themeLight" = "#c2f4ff";
	"themeTertiary" = "#7ee7ff";
	"themeSecondary" = "#f4971f";
	"themeDarkAlt" = "#0096b8";
	"themeDark" = "#00758f";
	"themeDarker" = "#005c70";
	"neutralLighterAlt" = "#f8f8f8";
	"neutralLighter" = "#f4f4f4";
	"neutralLight" = "#eaeaea";
	"neutralQuaternaryAlt" = "#dadada";
	"neutralQuaternary" = "#d0d0d0";
	"neutralTertiaryAlt" = "#c8c8c8";
	"neutralTertiary" = "#d6d6d6";
	"neutralSecondary" = "#474747";
	"neutralPrimaryAlt" = "#2e2e2e";
	"neutralPrimary" = "#333333";
	"neutralDark" = "#242424";
	"black" = "#1c1c1c";
	"white" = "#ffffff";
	"primaryBackground" = "#ffffff";
	"primaryText" = "#333333";
	"bodyBackground" = "#ffffff";
	"bodyText" = "#333333";
	"disabledBackground" = "#f4f4f4";
	"disabledText" = "#c8c8c8";
	}
)

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|          Tenant Level Theming Administration          |"
Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Target tenant: $($orgName)"
Write-Host ""

try
{
	#connent to site
	Write-Host -ForegroundColor White "Connecting to tenant"
	Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $Credentials
	Write-Host -ForegroundColor Green "Connected to $($orgName) tenant"
	
	#hide default themes if requested
	if($HideDefaultThemes)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Hiding default themes"
		Set-HideDefaultThemes $true
		Write-Host -ForegroundColor Green "Hiding default themes: complete"
	}

	#show default themes if requested
	if($ShowDefaultThemes)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Enabling default themes"
		Set-HideDefaultThemes $false
		Write-Host -ForegroundColor Green "Enabling default themes: complete"
	}

	#add custom theme based on palette
	if($AddTheme)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Adding new theme"
		Add-SPOTheme -Name "Custom 1" -Palette $themepalette -IsInverted $false
		Write-Host -ForegroundColor Green "Adding new theme: complete"
	}

	#remove custom theme
	if($RemoveTheme)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Removing custom theme"
		Remove-SPOTheme -Name "Custom 1"
		Write-Host -ForegroundColor Green "Removing custom theme: complete"
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Tenant theme administration complete"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}