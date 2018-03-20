<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Get a list of Office 365 groups

.EXAMPLE
PS C:\> .\Get-Group.ps1 -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com

.EXAMPLE
PS C:\> .\Get-Group.ps1 -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Get-Group.ps1 -GroupId "8fbe9d2e-7f5d-4eb0-91f7-6394abc482ed" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Get-Group.ps1 -DisplayName "Modern Team Site 21" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
PS C:\> .\Get-Group.ps1 -Url "https://pixelmilldev1.sharepoint.com/sites/collabcomm120" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = $false, HelpMessage="Filter Group Id")]
    [String]
    $GroupId,
	
	[Parameter(Mandatory = $false, HelpMessage="Filter Url")]
    [String]
    $Url,

	[Parameter(Mandatory = $false, HelpMessage="Filter Display Name")]
    [String]
    $DisplayName,

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

Write-Host -ForegroundColor White "---------------------------------------------"
Write-Host -ForegroundColor White "|                 Get Groups                |"
Write-Host -ForegroundColor White "---------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Getting list of groups"
Write-Host ""

try
{
	#connent to graph to get groups
	Write-Host -ForegroundColor White "Connecting to graph"
	#using aad v2 app, registered at: https://apps.dev.microsoft.com
	if (![string]::IsNullOrEmpty($AppId) -and ![string]::IsNullOrEmpty($AppSecret) -and ![string]::IsNullOrEmpty($AADDomain)) {
		Connect-PnPOnline -AppId $AppId -AppSecret $AppSecret -AADDomain $AADDomain
	}
	else {
		Connect-PnPOnline -scopes "Group.ReadWrite.All","User.Read.All"
	}
	Write-Host -ForegroundColor Green "Connected to graph"


	if(![string]::IsNullOrEmpty($GroupId))
	{
		$groups = Get-PnPUnifiedGroup -Identity $GroupId
	}
	else {
		$groups = Get-PnPUnifiedGroup
	}

	if ($groups -ne $null)
	{
		Write-Host ""
		Write-Host -ForegroundColor Green "Groups retrieved"

		if(![string]::IsNullOrEmpty($DisplayName))
		{
			$groups = $groups | Where-Object {$_.DisplayName -eq $DisplayName}
		}

		if(![string]::IsNullOrEmpty($Url))
		{
			$groups = $groups | Where-Object {$_.SiteUrl -eq $Url}
		}

		if ($groups -ne $null)
		{
			$groups
		}
		else {
			Write-Host -ForegroundColor Yellow "After filtering, no valid groups found to match your request"
		}

    }
	else {
		Write-Host ""
		Write-Host -ForegroundColor Yellow "Unable to find any groups, ignoring request"
	}
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}