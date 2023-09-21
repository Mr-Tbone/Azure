#Change for your environment
$TenantID = "11111-08a2-4ade-9a68-0db7586d80ad"
$ManagedIdentity = "Tbone-Automate"

#Scope permissions needed by Contact Manager
$Permissions = @("Group.Read.All","User.Read.All","Contacts.ReadWrite","Mail.Send")
$GraphAppId = "00000003-0000-0000-c000-000000000000"

Connect-AzureAD -TenantId $TenantID
$ManagedIdentityServicePrincipal = (Get-AzureADServicePrincipal -Filter "displayName eq '$ManagedIdentity'")
$GraphServicePrincipal = Get-AzureADServicePrincipal -Filter "appId eq '$GraphAppId'"

foreach ($Permission in $Permissions)
    {
    $AppRole = $GraphServicePrincipal.AppRoles | Where-Object {$_.Value -eq $Permission -and $_.AllowedMemberTypes -contains "Application"}
    New-AzureAdServiceAppRoleAssignment -ObjectId $ManagedIdentityServicePrincipal.ObjectId -PrincipalId $ManagedIdentityServicePrincipal.ObjectId -ResourceId $GraphServicePrincipal.ObjectId -Id $AppRole.Id
    }
 