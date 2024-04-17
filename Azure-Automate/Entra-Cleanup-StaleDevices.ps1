<#PSScriptInfo
.SYNOPSIS
    Script for Entra ID to cleanup stale device objects
 
.DESCRIPTION
    This script will get the Entra device objects 
    The script then compare the ApproximateLastSignInDateTime with the cleanup threshold and remove the device if it is older than the threshold 
    The script uses Ms Graph with MGGraph modules
        
.EXAMPLE
   .\Entra-Cleanup-StaleDevices.ps1
    Will cleanup stale devices 

.NOTES
    Written by Mr-Tbone (Tbone Granheden) Coligo AB
    torbjorn.granheden@coligo.se

.VERSION
    1.0

.RELEASENOTES
    1.0 2023-02-14 Initial Build

.AUTHOR
    Tbone Granheden 
    @MrTbone_se

.COMPANYNAME 
    Coligo AB

.GUID 
    00000000-0000-0000-0000-000000000000

.COPYRIGHT
    Feel free to use this, But would be grateful if My name is mentioned in Notes 

.CHANGELOG
    1.0.2202.1 - Initial Version
#>

#region ---------------------------------------------------[Set script requirements]-----------------------------------------------
#
#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules Microsoft.Graph.identity.DirectoryManagement
#
#endregion

#region ---------------------------------------------------[Script Parameters]-----------------------------------------------
#endregion

#region ---------------------------------------------------[Modifiable Parameters and defaults]------------------------------------
# Customizations
[int]$DeviceDisableThreshold = 60        # Number of inactive days to determine a stale device to disable
[int]$DeviceDeleteThreshold  = 90        # Number of inactive days to determine a stale device to delete
[Bool]$TestMode             = $true    # $True = No devices will be deleted, $False = Stale devices will be deleted
[Bool]$Verboselogging       = $True     # $True = Enable verbose logging for t-shoot. $False = Disable Verbose Logging
#endregion

#region ---------------------------------------------------[Set global script settings]--------------------------------------------
Set-StrictMode -Version Latest
#endregion

#region ---------------------------------------------------[Import Modules and Extensions]-----------------------------------------
import-Module Microsoft.Graph.Authentication
import-Module Microsoft.Graph.identity.DirectoryManagement
#endregion

#region ---------------------------------------------------[Static Variables]------------------------------------------------------
[System.Collections.ArrayList]$RequiredScopes   = @("Device.ReadWrite.All")
[datetime]$scriptStartTime                      = Get-Date
[string]$disableDate = "$(($scriptStartTime).AddDays(-$DeviceDisableThreshold).ToString("yyyy-MM-dd"))T00:00:00z"
[string]$deleteDate = "$(($scriptStartTime).AddDays(-$DeviceDeleteThreshold).ToString("yyyy-MM-dd"))T00:00:00z"
if ($Verboselogging){$VerbosePreference         = "Continue"}
else{$VerbosePreference                         = "SilentlyContinue"}
#endregion

#region ---------------------------------------------------[Functions]------------------------------------------------------------
function ConnectTo-MgGraph {
    param (
        [System.Collections.ArrayList]$RequiredScopes
    )
    Begin {
        $ErrorActionPreference = 'stop'
        [String]$resourceURL = "https://graph.microsoft.com/"
        $GraphAccessToken = $null
        if ($env:AUTOMATION_ASSET_ACCOUNTID) {  [Bool]$ManagedIdentity = $true}  # Check if running in Azure Automation
        else {                                  [Bool]$ManagedIdentity = $false} # Otherwise running in Local PowerShell
        }
    Process {
        if ($ManagedIdentity){ #Connect to the Microsoft Graph using the ManagedIdentity and get the AccessToken
            Try{$response = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
                $GraphAccessToken = $response.access_token
                Write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to get an Access Token to Graph for managed identity"
                }
            Catch{Write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to get an Access Token to Graph for managed identity, with error: $_"}
            $GraphVersion = ($GraphVersion = (Get-Module -Name 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue).Version | Sort-Object -Desc | Select-Object -First 1)
            if ('2.0.0' -le $GraphVersion) {
                Try{Connect-MgGraph -Identity -Nowelcome
                    $GraphAccessToken = convertto-securestring($response.access_token) -AsPlainText -Force
                    Write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to connect to Graph with module 2.x and Managedidentity"}
                Catch{Write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to connect to Graph with module 2.x and Managedidentity, with error: $_"}
                }
            else {#Connect to the Microsoft Graph using the AccessToken
                Try{Connect-mgGraph -AccessToken $GraphAccessToken -NoWelcome
	                Write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to connect to Graph with module 1.x and Managedidentity"}
                Catch{Write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to connect to Graph with module 1.x and Managedidentity, with error: $_"}
                }
            }
        else{
            Try{Connect-MgGraph -Scope $RequiredScopes
                Write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to connect to Graph manually"}
            Catch{Write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to connect to Graph manually, with error: $_"}
            }
        #Checking if all permissions are granted to the script identity in Graph and exit if not
        [System.Collections.ArrayList]$CurrentPermissions  = (Get-MgContext).Scopes
        foreach ($RequiredScope in $RequiredScopes) {
            if (Compare-Object $currentpermissions $RequiredScope -IncludeEqual | Where-Object -FilterScript {$_.SideIndicator -eq '=='}){
                Write-Verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success, Script identity has a scope permission: $RequiredScope"
                }
            else {Write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed, Script identity is missing a scope permission: $RequiredScope"}
            }
        #Return the access token if available and cleanup memory after connecting to Graph
        return $GraphAccessToken
        }
    End {$MemoryUsage = [System.GC]::GetTotalMemory($true)
        Write-Verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to cleanup Memory usage after connect to Graph to: $(($MemoryUsage/1024/1024).ToString('N2')) MB"
        }   
}
#endregion

#region ---------------------------------------------------[[Script Execution]------------------------------------------------------
$StartTime = Get-Date
$MgGraphAccessToken = ConnectTo-MgGraph -RequiredScopes $RequiredScopes

#Get Pending Devices to disable
try{$pendingdevices = Get-MgDevice -All -Filter "ApproximateLastSignInDateTime le $($disableDate) AND ApproximateLastSignInDateTime ge $($deleteDate)"
    write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to get $($pendingdevices.count) Pending Devices to disable"}
catch{write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to get Pending Devices with error: $_"}

#Get Stale Devices to delete
try{$staledevices = Get-MgDevice -All -Filter "ApproximateLastSignInDateTime le $($deleteDate)"
    write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to get $($staledevices.count) Stale Devices to delete"}
catch{write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to get Stale Devices with error: $_"}

#Disable Pending Devices
foreach ($device in $pendingdevices) {
    Write-Output "Device $($device.DisplayName) is pending to be disabled"
    if ($TestMode -eq $False) {
        try{Update-MgDevice -DeviceId $device.Id -AccountEnabled:$False
            write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to disable Device $($device.DisplayName)"}
        catch{write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to disable Device $($device.DisplayName) with error: $_"}
    }
}

#Delete Stale Devices
foreach ($device in $staledevices) {
    Write-Output "Device $($device.DisplayName) is stale and will be removed"
    if ($TestMode -eq $False) {
        try{Remove-MgDevice -DeviceId $device.Id
            write-verbose "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Success to remove Device $($device.DisplayName)"}
        catch{write-Error "$(Get-Date -Format 'yyyy-MM-dd'),$(Get-Date -format 'HH:mm:ss'),Failed to remove Device $($device.DisplayName) with error: $_"}
    }
}

[datetime]$scriptEndTime    = Get-Date
write-Output "Script execution time: $(($scriptEndTime-$scriptStartTime).ToString('hh\:mm\:ss'))"
$VerbosePreference = "SilentlyContinue"
