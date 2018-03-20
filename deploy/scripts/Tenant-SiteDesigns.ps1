<#
.REQUIREMENTS
Requires SharePoint Online Management Shell
https://www.microsoft.com/en-us/download/details.aspx?id=35588

Based on documentation:
https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview

.SYNOPSIS
Add, remove, and work with site designs at tenant level
Requires tenant admin level access

.EXAMPLE
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" 

.EXAMPLE
PS C:\> $credsSPO = Get-Credential
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GetSiteScript asd
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteScript -Title "site script 1" -Description "a desc" -Script ".\templates\sitescript1.json"
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RemoveSiteScript -Id asd

PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -getSiteDesign asd
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteDesign -Title "Site Design 1" -WebTemplate "64" -SiteScripts ""
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteDesign -Title "Site Design 1" -WebTemplate "64" -SiteScripts "" -PreviewImageUrl "" -PreviewImageAltText "" -IsDefault

PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GetSiteDesignRights asd
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GrantSiteDesignRights asd -Priciples "eoverfield@pixelmilldev1.onmicrosoft.com" -Rights "view"
PS C:\> .\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RevokeSiteDesignRights asd -Priciples "eoverfield@pixelmilldev1.onmicrosoft.com"
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

	[Parameter(Mandatory = $false, HelpMessage="Enter a title")]
    [String]
    $Title,

	[Parameter(Mandatory = $false, HelpMessage="Enter a description")]
    [String]
    $Description,

	[Parameter(Mandatory = $false, HelpMessage="Enter a Identifier / GUID")]
    [String]
    $Id,

	[Parameter(Mandatory = $false, HelpMessage="Enter location (relative or absolute) to a json script")]
    [String]
    $Script,

	[Parameter(Mandatory = $false, HelpMessage="Web Template Id - 64: Modern Team 68: Communication Site")]
    [String]
    $WebTemplate,

	[Parameter(Mandatory = $false, HelpMessage="Array of Site Scripts")]
    [String[]]
    $SiteScripts,

	[Parameter(Mandatory = $false, HelpMessage="Preview Image Url")]
    [String]
    $PreviewImageUrl,

	[Parameter(Mandatory = $false, HelpMessage="Preview Image Alt Text")]
    [String]
    $PreviewImageAltText,

	[Parameter(Mandatory = $false, HelpMessage="Set as a Default Option?")]
    [Switch]
    $IsDefault,

	[Parameter(Mandatory = $false, HelpMessage="Array of Principals")]
    [String[]]
    $Principals,
	
	[Parameter(Mandatory = $false, HelpMessage="Rights")]
    [String]
    $Rights,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Show Site Scripts?")]
    [Switch]
    $GetSiteScript,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Delete Site Script?")]
    [Switch]
    $RemoveSiteScript,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Add Site Design?")]
    [Switch]
    $AddSiteScript,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Show Site Designs?")]
    [Switch]
    $GetSiteDesign,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Remove Site Design?")]
    [Switch]
    $RemoveSiteDesign,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Add Site Design?")]
    [Switch]
    $AddSiteDesign,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Get Site Design Rights?")]
    [Switch]
    $GetSiteDesignRights,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Grant Site Design Rights?")]
    [Switch]
    $GrantSiteDesignRights,

	[Parameter(Mandatory = $false, HelpMessage="Optional - Revoke Site Design Rights?")]
    [Switch]
    $RevokeSiteDesignRights
)

#includes
. "./utilities.ps1"

$Credentials = Verify-SPOCredentials -Credentials $Credentials -Url "https://$($orgName)-admin.sharepoint.com"

Write-Host -ForegroundColor White "---------------------------------------------------------"
Write-Host -ForegroundColor White "|        Tenant Level Site Design Administration        |"
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
	
	#Show available site scripts if requested
	if($GetSiteScript)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Show Site Scripts"
		
		if([string]::IsNullOrEmpty($Id)) {
			Get-SPOSiteScript
		}
		else {
			Get-SPOSiteScript $Id
		}
		Write-Host -ForegroundColor Green "Show Site Scripts: complete"
	}

	#Remove a site scripts if requested
	if($RemoveSiteScript)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Remove a Site Script"
		
		if([string]::IsNullOrEmpty($Id)) {
			Write-Host -ForegroundColor Yellow "Identity required to delete a site script"
		}
		else {
			$script = Remove-SPOSiteScript $Id
			Write-Host -ForegroundColor Green "Remove Site Script: complete"
		}
		
	}

	#Add a site script if requested
	if($AddSiteScript)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Add Site Script"

		#open content file
		if(![System.IO.File]::Exists($Script)) {
			$fullPath = Resolve-Path $Script 

			Write-Host -ForegroundColor White "Site Script found"
			$content = [System.IO.File]::ReadAllText($fullPath)
		}
		else {
			Write-Host -ForegroundColor White "Site Script unavailable"
			Break
		}
		
		Add-SPOSiteScript -Title $Title -Content $Content -Description $Description
		
		Write-Host -ForegroundColor Green "Add Site Script: complete"
	}

	#Show available site designs if requested
	if($GetSiteDesign)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Show Site Designs"
		Get-SPOSiteDesign
		Write-Host -ForegroundColor Green "Show Site Designs: complete"
	}

	#Remove a site design if requested
	if($RemoveSiteDesign)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Remove a Site Design"
		
		if([string]::IsNullOrEmpty($Id)) {
			Write-Host -ForegroundColor Yellow "Identity required to delete a site design"
		}
		else {
			$script = Remove-SPOSiteDesign $Id
			Write-Host -ForegroundColor Green "Remove Site Design: complete"
		}
	}

	#Add a site design if requested
	if($AddSiteDesign)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Add Site Design"

		if([string]::IsNullOrEmpty($Title)) {
			Write-Host -ForegroundColor Red "Title required"
			Break
		}
		if([string]::IsNullOrEmpty($WebTemplate)) {
			Write-Host -ForegroundColor Red "Web Template required. 64 = Modern Team Site - 68 = Communication Site"
			Break
		}
		if([string]::IsNullOrEmpty($SiteScripts)) {
			Write-Host -ForegroundColor Red "Array of one or more site script ids required"
			Break
		}

		#set up the command we will use to create a group
		$exp = '$siteDesign = Add-SPOSiteDesign -Title $($Title) -WebTemplate $($WebTemplate)'

		Write-Host "Setting Owners" $Owners
		$exp += " -SiteScripts "
		for ($i=0; $i -lt $SiteScripts.length; $i++) {
			if ($i -gt 0) {
				$exp += ', '
			}
			$exp += '"' + $SiteScripts[$i] + '"'
		}

		if(![string]::IsNullOrEmpty($Description)) {
			$exp += ' -Description "' + $Description + '"'
		}

		if(![string]::IsNullOrEmpty($PreviewImageUrl)) {
			$exp += ' -PreviewImageUrl "' + $PreviewImageUrl + '"'
		}

		if(![string]::IsNullOrEmpty($PreviewImageAltText)) {
			$exp += ' -PreviewImageAltText "' + $PreviewImageAltText + '"'
		}

		if($IsDefault)
		{
			$exp += " -IsDefault"
		}

		Write-Host ""
		Write-Host -ForegroundColor White "Going to create site design using the following command:"
		Write-Host $exp

		Invoke-Expression $exp

		if ($siteDesign -ne $null)
		{
			Write-Host "Site Design created " $siteDesign.Id
			$siteDesign
		}

		Write-Host -ForegroundColor Green "Add Site Script: complete"
	}

	#Get Site Design Rights
	if($GetSiteDesignRights)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Get Site Design Rights"
		
		if([string]::IsNullOrEmpty($Id)) {
			Write-Host -ForegroundColor Yellow "Site Design Id required to get Rights"
			break
		}
		else {
			Get-SPOSiteDesignRights $Id
		}
		Write-Host -ForegroundColor Green "Get Site Design Rights: complete"
	}

	#Grant Site Design Rights
	if($GrantSiteDesignRights)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Grant Site Design Rights"
		
		if([string]::IsNullOrEmpty($Id)) {
			Write-Host -ForegroundColor Red "Id required of site design to apply rights"
			Break
		}

		if([string]::IsNullOrEmpty($Principals)) {
			Write-Host -ForegroundColor Red "Array of one or more Principals required"
			Break
		}
		if([string]::IsNullOrEmpty($Rights)) {
			$Rights = 'view'
		}

		#set up the command we will use to create a group
		$exp = '$siteDesignRights = Grant-SPOSiteDesignRights $($Id) -Rights $($Rights)'

		Write-Host "Setting Priciples" $Principals
		$exp += " -Principals "
		for ($i=0; $i -lt $Principals.length; $i++) {
			if ($i -gt 0) {
				$exp += ', '
			}
			$exp += '"' + $Principals[$i] + '"'
		}

		Write-Host ""
		Write-Host -ForegroundColor White "Going to grant site design rights using the following command:"
		Write-Host $exp

		Invoke-Expression $exp

		if ($siteDesignRights -ne $null)
		{
			Write-Host "Site Design Rights applied "
			$siteDesignRights
		}
		
		Write-Host -ForegroundColor Green "Grant Site Design Rights: complete"
	}

	#Revoke Site Design Rights
	if($RevokeSiteDesignRights)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Revoke Site Design Rights"
		
		if([string]::IsNullOrEmpty($Id)) {
			Write-Host -ForegroundColor Red "Id required of site design to revoke rights"
			Break
		}

		if([string]::IsNullOrEmpty($Principals)) {
			Write-Host -ForegroundColor Red "Array of one or more Principals required"
			Break
		}

		#set up the command we will use to create a group
		$exp = '$siteDesignRights = Revoke-SPOSiteDesignRights $($Id)'

		Write-Host "Setting Priciples" $Principals
		$exp += " -Principals "
		for ($i=0; $i -lt $Principals.length; $i++) {
			if ($i -gt 0) {
				$exp += ', '
			}
			$exp += '"' + $Principals[$i] + '"'
		}

		Write-Host ""
		Write-Host -ForegroundColor White "Going to revoke site design rights using the following command:"
		Write-Host $exp

		Invoke-Expression $exp

		if ($siteDesignRights -ne $null)
		{
			Write-Host "Site Design Rights revoked "
			$siteDesignRights
		}
		
		Write-Host -ForegroundColor Green "Revoke Site Design Rights: complete"
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Tenant site designs administration complete"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}