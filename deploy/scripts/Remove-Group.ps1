<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1712.2 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

Must create an AAD App per:
https://github.com/SharePoint/PnP-PowerShell/tree/master/Samples/Graph.ConnectUsingAppPermissions

Once an app is created, it must be concented:
https://login.microsoftonline.com/<tenant>/adminconsent?client_id=<clientid>&state=<something>

https://login.microsoftonline.com/pixelmilldev1.onmicrosoft.com/adminconsent?client_id=8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e&state=CA

Based on PnP: Remove-PnPUnifiedGroup commandlet
	Remove-PnPUnifiedGroup -Identity $group

.SYNOPSIS
Remove a given group based on display name

.EXAMPLE
PS C:\> .\Remove-Group.ps1 -DisplayName "CollabComm1"

.EXAMPLE
PS C:\> .\Remove-Group.ps1 -DisplayName "CollabComm20" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $true, HelpMessage="Group Display Name")]
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

Write-Host -ForegroundColor White "------------------------------------------------"
Write-Host -ForegroundColor White "|                 Remove Group                 |"
Write-Host -ForegroundColor White "------------------------------------------------"
Write-Host ""
Write-Host -ForegroundColor Green "Remove a Group: $($DisplayName)"
Write-Host ""

try
{
	Write-Host -ForegroundColor White "Connecting to graph"
	if (![string]::IsNullOrEmpty($AppId) -and ![string]::IsNullOrEmpty($AppSecret) -and ![string]::IsNullOrEmpty($AADDomain)) {
		Connect-PnPOnline -AppId $AppId -AppSecret $AppSecret -AADDomain $AADDomain
	}
	else {
		Connect-PnPOnline -scopes "Group.ReadWrite.All","User.Read.All"
	}
	
	Write-Host -ForegroundColor Green "Connected to graph"

	Write-Host ""
	Write-Host -ForegroundColor Green "Looking for group to remove"
	$groups = Get-PnPUnifiedGroup
	$group = $groups | Where-Object {$_.DisplayName -eq $DisplayName}
	if ($group -ne $null)
	{
		Write-Host ""
		Write-Host -ForegroundColor White "Group found, removing"
    	Remove-PnPUnifiedGroup -Identity $group
		Write-Host -ForegroundColor Green "Group removed"
    }
	else {
		Write-Host -ForegroundColor Yellow "Unable to find group to remove."
		
		exit
	}

	Write-Host ""
	Write-Host -ForegroundColor Green "Group Removed"
}
catch
{
    Write-Host -ForegroundColor Red "Exception occurred!" 
    Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"
}