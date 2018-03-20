cd "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts"

$creds = Get-Credential
$creds = Get-PnPStoredCredential -Name "pmdev1eo" -Type PSCredential
$appId = "ed561a1f-79d7-4a90-bf82-c849764a3949"
$appSecret = "iaJq/B+bDZ0wCu351DDd+H9J6sI2RO4kRVqeKAUFl98="

#Provisioning SPFx solution to tenant app catalog
.\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials $creds
.\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials "pmdev1eo"
.\Provision-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -AppId $appId -AppSecret $appSecret

#Remove SPFx solution from tenant app catalog
.\Remove-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials $creds
.\Remove-SPFx.ps1 -TargetCatalogWebUrl "https://pixelmilldev1.sharepoint.com/sites/apps" -Credentials "pmdev1eo"

#provision assets to CDN
.\Provision-SPFx-Assets.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN" -Credentials $creds
.\Provision-SPFx-Assets.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN" -Credentials "pmdev1eo"

#Provision a given site to use extension and more
.\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials $creds -ProvisionIA -ProvisionAssets -ProvisionPages -ActivateExtensions -ActivateSiteSettings -ProvisionTheme -ProvisionLogo
.\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials "pmdeveo1"
.\Provision-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/collabcomm194" -Credentials $creds -ProvisionIA -ProvisionAssets -ProvisionPages -ActivateExtensions -ActivateSiteSettings -ProvisionTheme -ProvisionLogo

#Remove setup of given site of extension and more
.\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials $creds -DeactivateExtensions -DeactivateSiteSettings -ResetLogo
.\Remove-Site-Setup.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/DevCollabCommPlayground1" -Credentials "pmdeveo1" -DeactivateExtensions -DeactivateSiteSettings -ResetLogo

#Theming administration
$credsSPO = Get-Credential
$creds = Get-PnPStoredCredential -Name "pmdev1eo" -Type PSCredential
$creds = Get-PnPStoredCredential -Name "pmdev1admin" -Type PSCredential
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -AddTheme
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -HideDefaultThemes
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -ShowDefaultThemes
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddTheme
.\Tenant-Theming.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RemoveTheme

#Site Design Administration
$credsSPO = Get-Credential
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GetSiteScript -Id b01992de-eacb-4e1d-9ca6-af4c9c054076
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteScript -Title "site script 1" -Description "a desc" -Script ".\templates\sitescript1.json"
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RemoveSiteScript -Id 47d678fa-d4a5-4813-b0cf-0a87f2144df4

.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GetSiteDesign
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteDesign -Title "Site Design 1" -WebTemplate "64" -SiteScripts "b01992de-eacb-4e1d-9ca6-af4c9c054076"
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -AddSiteDesign -Title "Site Design 1" -WebTemplate "64" -SiteScripts "" -PreviewImageUrl "" -PreviewImageAltText "" -IsDefault
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RemoveSiteDesign -Id c69499b7-37e2-480c-98f4-035974645887

.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GetSiteDesignRights -Id c69499b7-37e2-480c-98f4-035974645887
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -GrantSiteDesignRights -Id c69499b7-37e2-480c-98f4-035974645887 -Principals "eoverfield@pixelmilldev1.onmicrosoft.com", "admin@pixelmilldev1.pixelmill.com" -Rights "view"
.\Tenant-SiteDesigns.ps1 -OrgName "pixelmilldev1" -Credentials $credsSPO -RevokeSiteDesignRights -Id c69499b7-37e2-480c-98f4-035974645887 -Principals "eoverfield@pixelmilldev1.onmicrosoft.com"

#Group Provisioning
.\Provision-Group.ps1 -DisplayName "CollabComm120" -Description "Collab Comm 20 Group" -MailNickname "collabcomm120" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com","admin@pixelmilldev1.pixelmill.com" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Provision-Group.ps1 -DisplayName "CollabComm121" -Description "Collab Comm 1 Group" -MailNickname "collabcomm121" -GroupLogoPath ".\templates\SiteAssets\PnP.png" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Provision-Group.ps1 -DisplayName "CollabComm122" -Description "Collab Comm 1 Group" -MailNickname "collabcomm122" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com","admin@pixelmilldev1.pixelmill.com" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Provision-Group.ps1 -DisplayName "CollabComm123" -Description "Collab Comm 1 Group" -MailNickname "collabcomm123" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com" -GroupLogoPath "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts\templates\SiteAssets\PnP.png" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Provision-Group.ps1 -DisplayName "CollabComm124" -Description "Collab Comm 1 Group" -MailNickname "collabcomm124" -Owners "eoverfield@pixelmilldev1.onmicrosoft.com" -Members "admin@pixelmilldev1.pixelmill.com" -GroupLogoPath "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts\templates\SiteAssets\PnP.png" -IsPrivate -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com

#Group Removal
.\Remove-Group.ps1 -DisplayName "CollabComm24" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com

#Provision Communication Sites
.\Provision-Communication.ps1 -Tenant "pixelmilldev1" -Title "Comm Site 20" -Alias "CommSite20" -Description "A comm site" -Credentials $creds
.\Provision-Communication.ps1 -Tenant "pixelmilldev1" -Title "Comm Site 21" -Alias "CommSite21" -Description "A comm site" -SiteDesign Topic -Classification "HB1" -AllowFileSharingForGuestUsers -Credentials $creds

#Provision Modern Team Sites
.\Provision-Modern-Team.ps1 -Tenant "pixelmilldev1" -Title "Modern Team Site 22" -Alias "ModernTeam22" -Description "A modern team site" -Credentials $creds
.\Provision-Modern-Team.ps1 -Tenant "pixelmilldev1" -Title "Modern Team Site 21" -Alias "ModernTeam21" -Description "A modern team site" -Classification "HB1" -IsPublic -Credentials $creds

#Provision Classic Sites
.\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Team Site 1" -Alias "ClassicTeam20" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0" -StorageQuota 9 -ResourceQuota 19 -Force -Credentials $creds
.\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Team Site 1" -Alias "ClassicTeam20" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0" -Force -Credentials $creds
.\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Team Site 1" -Alias "ClassicTeam22" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "STS#0" -Force -Wait -Credentials $creds
.\Provision-Classic.ps1 -Tenant "pixelmilldev1" -Title "Publishing Site 1" -Alias "ClassicPub20" -Owner "eoverfield@pixelmilldev1.onmicrosoft.com" -TimeZone 13 -Template "BLANKINTERNETCONTAINER#0" -Force -Credentials $creds

#Remove Classic Sites
.\Remove-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam31" -Force -Credentials $creds
.\Remove-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam32" -Force -SkipRecycleBin -Credentials $creds

.\Remove-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/collabcomm12" -Force -SkipRecycleBin -Credentials $creds

#Clear Classic Site from RecycleBin
.\ClearRecycle-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam31" -Force -Credentials $creds
.\ClearRecycle-Classic.ps1 -Tenant "pixelmilldev1" -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam22" -Force -Wait -Credentials $creds

#Groupify a Classic Site
.\Groupify-Classic.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam20" -Alias "ClassicTeam20" -DisplayName "Classic Team 20" -Description "ClassicTeam20 desc" -IsPublic -Credentials $creds
.\Groupify-Classic.ps1 -TargetWebUrl "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam21" -Alias "ClassicTeam21" -DisplayName "Classic Team 2" -Description "ClassicTeam2 desc" -IsPublic -Credentials $creds

#Get Tenant level sites
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Credentials $creds
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Url "https://pixelmilldev1.sharepoint.com/sites/ClassicTeam21" -Credentials $creds
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "STS#0" -Credentials $creds
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "BLANKINTERNETCONTAINER#0" -Credentials $creds
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "DEV#0" -Credentials $creds

.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "GROUP#0" -Credentials $creds
.\Get-TenantSites.ps1 -Tenant "pixelmilldev1" -Template "SITEPAGEPUBLISHING#0" -Credentials $creds

#Get GRoups
.\Get-Group.ps1 -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Get-Group.ps1 -GroupId "8fbe9d2e-7f5d-4eb0-91f7-6394abc482ed" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Get-Group.ps1 -DisplayName "Modern Team Site 21" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
.\Get-Group.ps1 -Url "https://pixelmilldev1.sharepoint.com/sites/collabcomm120" -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com



Disconnect-PnPOnline
Get-PnPAccessToken
$group = Get-PnPUnifiedGroup -Identity "8fbe9d2e-7f5d-4eb0-91f7-6394abc482ed"
$group.SiteUrl

Get-PnPTenantSite -IncludeOneDriveSites

Connect-PnPOnline -Url https://pixelmilldev1-admin.sharepoint.com -Credential $credential

$tenantSites = Get-PnPListItem -List DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS -Fields ID,Title,TemplateTitle,SiteUrl,IsGroupConnected
$tenantSites | format-table @{Expression = {$_['ID']};Label="ID"},@{Expression = {$_['Title']};Label="Title"},@{Expression = {$_['TemplateTitle']};Label="TemplateTitle"},@{Expression = {$_['SiteUrl']};Label="SiteURL"},@{Expression = {$_['IsGroupConnected']};Label="IsGroupConnected"}



Connect-PnPOnline -Url https://pixelmilldev1.sharepoint.com/sites/DevCollabCommCDN -AppId ed561a1f-79d7-4a90-bf82-c849764a3949 -AppSecret "iaJq/B+bDZ0wCu351DDd+H9J6sI2RO4kRVqeKAUFl98="
get-pnplist