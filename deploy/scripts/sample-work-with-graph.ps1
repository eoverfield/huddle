cd "E:\Datastore\Git Projects\PixelMill\CollabCommPlay\deploy\scripts"

Connect-PnPMicrosoftGraph -AppId 8a6ec2a0-6092-4b02-8a81-17f0d41a1c3e -AppSecret 'mypFI6?gusbRTGHY3410(~@' -AADDomain pixelmilldev1.onmicrosoft.com
connect-pnpmicrosoftgraph -scopes "Group.ReadWrite.All","User.Read.All"

$access_token = Get-PnPAccessToken;
$access_token

Get-PnPUnifiedGroup
$group = Get-PnPUnifiedGroup -Identity 04e5ddcd-ccda-4388-9191-da2c4b6cc300
$group.Description
$group.SiteUrl

$error[0].Exception.Stacktrace

Set-PnPUnifiedGroup -Identity 04e5ddcd-ccda-4388-9191-da2c4b6cc300 -GroupLogoPath "E:\Datastore\PnP.png" -Verbose:$true
Set-PnPUnifiedGroup -Identity 04e5ddcd-ccda-4388-9191-da2c4b6cc300 -Description "new"
