<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Creation documenation:
https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/modern-experience-customizations-provisioning-sites

Must create an AAD App per:
https://github.com/SharePoint/PnP-PowerShell/tree/master/Samples/Graph.ConnectUsingAppPermissions

Once an app is created, it must be consented:
https://login.microsoftonline.com/<tenant>/adminconsent?client_id=<clientid>&state=<something>

i.e.
https://login.microsoftonline.com/pixelmilldev1.onmicrosoft.com/adminconsent?client_id=8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e&state=CA

Known issue, when setting the logo, IsPrivate is ignored: 11/28/17
https://github.com/SharePoint/PnP-PowerShell/issues/1206

Unable to provision a group logo via an app id / secret because of issue with Graph, required a user context

Based on PnP: New-PnPUnifiedGroup commandlet
	New-PnPUnifiedGroup -DisplayName <String>
				-Description <String>
				-MailNickname <String>
				[-Owners <String[]>]
				[-Members <String[]>]
				[-IsPrivate [<SwitchParameter>]]
				[-GroupLogoPath <String>]
				[-Force [<SwitchParameter>]]


.SYNOPSIS
Provision a new group

.EXAMPLE
PS C:\> .\Provision-Group.ps1 -DisplayName "CollabComm1" -Description "Collab Comm 1 Group" -MailNickname "collabcomm1" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com

.EXAMPLE
PS C:\> .\Provision-Group.ps1 -DisplayName "CollabComm20" -Description "Collab Comm 20 Group" -MailNickname "collabcomm20" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com","admin@pixelmilldev1.pixelmill.com" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Provision-Group.ps1 -DisplayName "CollabComm1" -Description "Collab Comm 1 Group" -MailNickname "collabcomm1" -GroupLogoPath ".\templates\SiteAssets\PnP.png" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Provision-Group.ps1 -DisplayName "CollabComm1" -Description "Collab Comm 1 Group" -MailNickname "collabcomm1" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com","admin@pixelmilldev1.pixelmill.com" -IsPrivate -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Provision-Group.ps1 -DisplayName "CollabComm1" -Description "Collab Comm 1 Group" -MailNickname "collabcomm1" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com" -Members "admin@pixelmilldev1.pixelmill.com" -GroupLogoPath "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts\templates\SiteAssets\PnP.png" -IsPrivate -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Group Display Name")]
    [String]
    $DisplayName,

	[Parameter(Mandatory = $true, HelpMessage="Group Description")]
    [String]
    $Description,

	[Parameter(Mandatory = $true, HelpMessage="Group Mail Nickname")]
    [String]
    $MailNickname,

	[Parameter(Mandatory = $false, HelpMessage="Group Owners")]
    [String[]]
    $Owners,

	[Parameter(Mandatory = $false, HelpMessage="Group Members")]
    [String[]]
    $Members,

	[Parameter(Mandatory = $false, HelpMessage="Make this a private group?")]
    [Switch]
    $IsPrivate,

	[Parameter(Mandatory = $false, HelpMessage="Group Logo Path")]
    [String]
    $GroupLogoPath,

	[Parameter(Mandatory = $false, HelpMessage="Force create the group?")]
    [Switch]
    $Force,

	#AAD app information
	[Parameter(Mandatory = $false, HelpMessage="AAD App ID")]
    [String]
    $AppId,
	[Parameter(Mandatory = $false, HelpMessage="AAD App Secret")]
    [String]
    $AppSecret,
	[Parameter(Mandatory = $false, HelpMessage="AAD App Domain")]
    [String]
    $AADDomain
)

if($GroupLogoPath -eq $null)
{
	$GroupLogoPath = ""
}

Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host -ForegroundColor White "|                 Provisioin New Group                 |"
Write-Host -ForegroundColor White "--------------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Provisioning a Group: $($DisplayName)"
Write-Host ""

try
{
	Write-Host -ForegroundColor White "Connecting to graph"

	#connent to graph to add group
	#using aad v2 app, registered at: https://apps.dev.microsoft.com
	#Connect-PnPMicrosoftGraph depricated, moved to Connect-PnPOnline
	if (![string]::IsNullOrEmpty($AppId) -and ![string]::IsNullOrEmpty($AppSecret) -and ![string]::IsNullOrEmpty($AADDomain)) {
		Connect-PnPOnline -AppId $AppId -AppSecret $AppSecret -AADDomain $AADDomain
	}
	else {
		Connect-PnPOnline -scopes "Group.ReadWrite.All","User.Read.All"
	}

	Write-Host -ForegroundColor Green "Connected to graph"

	#get a list of groups if so desired
	#Get-PnPUnifiedGroup

	#set up the command we will use to create a group
	$exp = '$group = New-PnPUnifiedGroup -DisplayName $($DisplayName) -Description $($Description) -MailNickname $($MailNickname)'

	#owners
	if ($Owners -ne $null)
	{
		Write-Host "Setting Owners" $Owners
		$exp += " -Owners "
		for ($i=0; $i -lt $Owners.length; $i++) {
			if ($i -gt 0) {
				$exp += ', '
			}
			$exp += '"' + $Owners[$i] + '"'
		}
	}

	#members
	if ($Members -ne $null)
	{
		Write-Host "Setting Members" $Members
		$exp += " -Members "
		for ($i=0; $i -lt $Members.length; $i++) {
			if ($i -gt 0) {
				$exp += ', '
			}
			$exp += '"' + $Members[$i] + '"'
		}
	}

	#logo
	<#
	if ($GroupLogoPath -ne "")
	{
		Write-Host "Setting Group Logo: " $GroupLogoPath
		$exp += ' -GroupLogoPath "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts\templates\SiteAssets\PnP.png"'
	}
	#>

	#private group?
	if($IsPrivate)
	{
		Write-Host "Make group private"
		$exp += " -IsPrivate"
	}

	#Force?
	if($Force)
	{
		Write-Host "Force create the group"
		$exp += " -Force"
	}

	Write-Host ""
	Write-Host -ForegroundColor White "Going to create group using the following command:"
	Write-Host $exp

	Invoke-Expression $exp

	if ($group -ne $null)
	{
		Write-Host "Group provisioned " $group.GroupId
		$group

		<#
		if ($GroupLogoPath -ne "")
		{
			Write-Host "Setting Group Logo: " $GroupLogoPath
			#Set-PnPUnifiedGroup -Identity $group -GroupLogoPath $GroupLogoPath
			Set-PnPUnifiedGroup -Identity $group.GroupId -GroupLogoPath "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts\templates\SiteAssets\PnP.png"
			Write-Host "Setting Group Logo: complete"
		}
		#>
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Group Provisioned"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}