<#
.REQUIREMENTS
Requires PnP-PowerShell SharePointPnPPowerShellOnline version 2.20.1711.1 or later
https://github.com/OfficeDev/PnP-PowerShell/releasess

.SYNOPSIS
Utility functions
#>

<#
Verify credentials to return either valid credentials provided, or credentials retrieved by PS.
If null will be returned, then will actually exit script

$Credentials = [string] if a credential manager "name" or [PSCredential]
$AppId = [string] if credentials are not being supplied, the verify an appId and appsecret are being supplied. Will have $credentials return $null
$AppSecret = [string] App secret for App Id
$Url (optional) = [string] the url we wish to connect to. if provided, and valid PSCredentials not yet found, will use to check if credential manager has url available in stored credentials

#Returns
[PSCredential]

#Example
Verify-Credentials -Credentials (Get-Credential)
Verify-Credentials -Credentials "yourname"
Verify-Credentials -Url "https://yourdomain.sharepoint.com/sites/test-site"
Verify-Credentials -AppId "your app id" -AppSecret "your app secret"
#>
function Verify-Credentials {
	Param (
        [Parameter(Mandatory=$false)]$Credentials,
		[Parameter(Mandatory=$false)]$AppId,
		[Parameter(Mandatory=$false)]$AppSecret,
        [Parameter(Mandatory=$false)][string]$Url
    )

	try
	{
		#first check to see if credentials provided
		if (![string]::IsNullOrEmpty($Credentials))
		{
			#if not PSCredential, then attempt to get via stored credentials
			if($Credentials.GetType().Name -ne "PSCredential") {
				$Credentials = Get-PnPStoredCredential -Name $Credentials -Type PSCredential
			}
		}

		#if still null, check to see if url provided, if so, we will want to check stored credentials for that url
		if ([string]::IsNullOrEmpty($Credentials))
		{
			if (![string]::IsNullOrEmpty($Url))
			{
				$Credentials = Get-PnPStoredCredential -Name $Url -Type PSCredential
			}	
		}

		#check app id now if credentials still not available
		if ([string]::IsNullOrEmpty($Credentials))
		{
			if (![string]::IsNullOrEmpty($AppId) -and ![string]::IsNullOrEmpty($AppSecret))
			{
				#we are returning null and will have to check when connecting to SharePoint
				return $null
			}
		}

		#if still null, then assume we have an issue so request
		if ([string]::IsNullOrEmpty($Credentials))
		{
			$Credentials = Get-Credential -Message "Enter Admin Credentials"
		}

		#if still null, stop!
		if ([string]::IsNullOrEmpty($Credentials))
		{
			Write-Host -ForegroundColor Red "Unable to get credentials" 
			exit
		}

		return $Credentials
	}
	catch
	{
		Write-Host -ForegroundColor Red "Exception occurred!" 
		Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"

		$Credentials = Get-Credential -Message "Enter Admin Credentials"
		return $Credentials
	}
}

<#
Verify SPO tenant level credentials to return either valid credentials provided, or credentials retrieved by PS.
If null will be returned, then will actually exit script

$Credentials = [string] if a credential manager "name" or [PSCredential]
$Url (optional) = [string] the url we wish to connect to. if provided, and valid PSCredentials not yet found, will use to check if credential manager has url available in stored credentials

#Returns
[PSCredential]

#Example
Verify-Credentials -Credentials (Get-Credential)
Verify-Credentials -Credentials "yourname"
Verify-Credentials -Url "https://yourdomain.sharepoint.com/sites/test-site"
#>
function Verify-SPOCredentials {
	Param (
        [Parameter(Mandatory=$false)]$Credentials,
        [Parameter(Mandatory=$false)][string]$Url
    )

	try
	{
		#first check to see if credentials provided
		if($Credentials -ne $null) {
			
			#if not PSCredential, then attempt to get via stored credentials
			if($Credentials.GetType().Name -ne "PSCredential") {
				$Credentials = Get-PnPStoredCredential -Name $Credentials -Type PSCredential
			}
		}

		#if still null, check to see if url provided, if so, we will want to check stored credentials for that url
		if($Credentials -eq $null)
		{
			if($Url -ne $null) {
				$Credentials = Get-PnPStoredCredential -Name $Url -Type PSCredential
			}	
		}

		#if still null, then assume we have an issue so request
		if($Credentials -eq $null)
		{
			$Credentials = Get-Credential -Message "Enter Admin Credentials"
		}

		#if still null, stop!
		if($Credentials -eq $null)
		{
			Write-Host -ForegroundColor Red "Unable to get credentials" 
			exit
		}

		return $Credentials
	}
	catch
	{
		Write-Host -ForegroundColor Red "Exception occurred!" 
		Write-Host -ForegroundColor Red "Exception Type: $($_.Exception.GetType().FullName)"
		Write-Host -ForegroundColor Red "Exception Message: $($_.Exception.Message)"

		$Credentials = Get-Credential -Message "Enter Admin Credentials"
		return $Credentials
	}
}

<#
Convert a Hast Table to a Dictionary, normally for SPO tenant themeing
https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-powershell

$ht = [Hashtable] the input hashtable

#Returns
[Dictionary]

#Example
HashToDictionary @{}
#>
function HashToDictionary {
	Param ([Hashtable]$ht)
	
	$dictionary = New-Object "System.Collections.Generic.Dictionary``2[System.String,System.String]"

	foreach ($entry in $ht.GetEnumerator()) {
		$dictionary.Add($entry.Name, $entry.Value)
	}

	return $dictionary
}