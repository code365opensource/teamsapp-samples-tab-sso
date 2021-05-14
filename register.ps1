[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 1)]
    [string]
    $customer,
    [Parameter(Mandatory = $false, Position = 2)]
    [string[]]
    $admins,
    [Parameter(Mandatory = $false, Position = 3)]
    [string[]]
    $validdomains
)

$ErrorActionPreference = "Stop"
Import-Module AzureAD

# create the aad application in azure(global) environment (Microsoft.com)

Connect-AzureAD
$app = New-AzureADApplication -DisplayName "Teams Workhub -$customer" -Homepage "https://workhub.microsoft.com" -AvailableToOtherTenants $true -Oauth2AllowImplicitFlow $true -ReplyUrls "https://$web_app_name.chinacloudsites.cn/auth", "https://$web_app_name.chinacloudsites.cn/home/notsupported?message=Done"
$clientId = $app.AppId
$clientSecret = (New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -EndDate "2099-1-1").Value
Start-Sleep -Seconds 10
$owner = Get-AzureADApplicationOwner -ObjectId $app.ObjectId
if ($owner.ObjectId -ne "f3b94dd3-20cc-49a3-98ce-b1287658e8cf") {
    Add-AzureADMSApplicationOwner -ObjectId $app.ObjectId -RefObjectId f3b94dd3-20cc-49a3-98ce-b1287658e8cf 
}

#graph 00000003-0000-0000-c000-000000000000
#permission f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb,205e70e5-aba6-4c52-a976-6d2d46c48043,465a38f9-76ea-45b9-9f34-9e8b0d4b0b42,e1fe6dd8-ba31-4d61-89e7-88639da4683d
$req = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$acc1 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "Scope"
$acc2 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb", "Scope"
$acc3 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "205e70e5-aba6-4c52-a976-6d2d46c48043", "Scope"
$acc4 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42", "Scope"
$req.ResourceAccess = $acc1, $acc2, $acc3, $acc4
$req.ResourceAppId = "00000003-0000-0000-c000-000000000000"
Set-AzureADApplication -ObjectId $app.ObjectId -RequiredResourceAccess $req


#expose an API
Set-AzureADApplication -ObjectId $app.ObjectId -IdentifierUris "api://$web_app_name.chinacloudsites.cn/$($app.AppId)"
#preauthorize application
$apiApp = Get-AzureADMSApplication -ObjectId $app.ObjectId
$permissionId = $apiApp.Api.Oauth2PermissionScopes[0].Id
$preAuthorizedApplication1 = New-Object 'Microsoft.Open.MSGraph.Model.PreAuthorizedApplication'
$preAuthorizedApplication1.AppId = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"
$preAuthorizedApplication1.DelegatedPermissionIds = @($permissionId)
$preAuthorizedApplication2 = New-Object 'Microsoft.Open.MSGraph.Model.PreAuthorizedApplication'
$preAuthorizedApplication2.AppId = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
$preAuthorizedApplication2.DelegatedPermissionIds = @($permissionId)
$apiApp.Api.PreAuthorizedApplications = New-Object 'System.Collections.Generic.List[Microsoft.Open.MSGraph.Model.PreAuthorizedApplication]'
$apiApp.Api.PreAuthorizedApplications.Add($preAuthorizedApplication1)
$apiApp.Api.PreAuthorizedApplications.Add($preAuthorizedApplication2)
Set-AzureADMSApplication -ObjectId $app.ObjectId -Api $apiApp.Api
