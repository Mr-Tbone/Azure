<#PSScriptInfo
.SYNOPSIS
    Script to set Default Password

.DESCRIPTION
    This script will set the default MFA method for a specific user of for all users in the tenant to the preferred method defined in the script.
        
.EXAMPLE
    .\set-MFADefaultMethod.ps1
    Will set the default MFA method for all users in the tenant to the preferred method defined in the script.
    .\set-MFADefaultMethod.ps1 -UserPrincipalName mr.tbone@tbone.se
    Will set the default MFA method for the specified user

.NOTES
    Written by Mr-Tbone (Tbone Granheden) Coligo AB
    torbjorn.granheden@coligo.se

.VERSION
    1.0

.RELEASENOTES
    1.0 2023-09-18 Initial Build
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
    1.0.2309.1 - Initial Version   
#>

#region ---------------------------------------------------[Set script requirements]-----------------------------------------------
#Requires -Version 7.0
#endregion

#region ---------------------------------------------------[Script Parameters]-----------------------------------------------
param(
  [Parameter(
    Mandatory = $false,
    ParameterSetName  = "UserPrincipalName",
    HelpMessage = "Enter a single UserPrincipalName",
    Position = 0
    )]
  [string[]]$UserPrincipalName
)
#endregion

#region ---------------------------------------------------[Modifiable Parameters and defaults]------------------------------------
$UserDefaultMethod = "push" # Set the preferred MFA method here
$SystemPreferredMethodEnabled = "true" # Set to true if you want to enable the system preferred authentication method
#endregion

#region ---------------------------------------------------[Set global script settings]--------------------------------------------
Set-StrictMode -Version Latest
#endregion

#region ---------------------------------------------------[Static Variables]------------------------------------------------------
$AllMgUsers       = @()

#Log File Info
#endregion

#region ---------------------------------------------------[Import Modules and Extensions]-----------------------------------------
Import-Module Microsoft.Graph.users
Import-Module Microsoft.Graph.Authentication
#endregion

#region ---------------------------------------------------[Functions]------------------------------------------------------------
#endregion

#region ---------------------------------------------------[[Script Execution]------------------------------------------------------

# Connect to Graph
Connect-MgGraph -scopes "User.Read.All,UserAuthenticationMethod.ReadWrite.All" -NoWelcome

if ($UserPrincipalName) {$AllMgUsers += Get-MgUser -UserId $UserPrincipalName}
else{$AllMgUsers = Get-MgUser -Filter "userType eq 'Member'"  -all}
$TotalMgUsers = $AllMgUsers.Count

# Get all the user's authentication methods. Batches is faster than getting the user and then the methods
$starttime = get-date
$AllAuthMethods = @()
for($i=0;$i -lt $TotalMgUsers;$i+=20){
    $req = @{}                
    # Use select to create hashtables of id, method and url for each call                                     
    if($i + 19 -lt $TotalMgUsers){
        $req['requests'] = ($AllMgUsers[$i..($i+19)] 
            | select @{n='id';e={$_.id}},@{n='method';e={'GET'}},`
            @{n='url';e={"/users/$($_.id)/authentication/methods"}})
    } else {
        $req['requests'] = ($AllMgUsers[$i..($TotalMgUsers-1)] 
            | select @{n='id';e={$_.id}},@{n='method';e={'GET'}},`
            @{n='url';e={"/users/$($_.id)/authentication/methods"}})
    }
    $response = invoke-mggraphrequest -Method POST `
        -URI "https://graph.microsoft.com/beta/`$batch" `
        -body ($req | convertto-json)
    $CurrentMgUser = $i  
    $response.responses | foreach {
        if($_.status -eq 200 -and $_.body.value -ne $null){
            $AuthMethod  = [PSCustomObject]@{
                "userprincipalname" = $AllMgUsers[$CurrentMgUser].userPrincipalName
                "AdditionalProperties" = $_.body
            }
            $AllAuthMethods += $AuthMethod
        } else {
            # request failed
            #write-host "the request for signInPreference failed"
        }
        $CurrentMgUser++
    }
#progressbar
    $Elapsedtime = (get-date) - $starttime
    $timeLeft = [TimeSpan]::FromMilliseconds((($ElapsedTime.TotalMilliseconds / $CurrentMgUser) * ($TotalMgUsers - $CurrentMgUser)))
    Write-Progress -Activity "Getting Authentication Methods for users $($CurrentMgUser) of $($TotalMgUsers)" `
        -Status "Est Time Left: $($timeLeft.Hours) Hours, $($timeLeft.Minutes) Minutes, $($timeLeft.Seconds) Seconds" `
        -PercentComplete $([math]::ceiling($($CurrentMgUser / $TotalMgUsers) * 100))
}

# Build the JSON Body request
$body = @'
{
    "isSystemPreferredAuthenticationMethodEnabled": SystemPreferredMethodEnabled,
    "userPreferredMethodForSecondaryAuthentication": "UserDefaultMethod"
}
'@
$body = $body -replace 'SystemPreferredMethodEnabled', $SystemPreferredMethodEnabled
$body = $body -replace 'UserDefaultMethod', $UserDefaultMethod

#Start the loop to set the default method for each user
$CurrentMgUser = 0
$starttime = get-date
foreach($MgUser in $AllMgUsers){
#Reset all the variables 
    $Authenticator = $false
    $Phone = $false
    $PasswordLess = $false
    $Fido2 = $false
    $TAP = $false
    $WHFB = $false
    $3part = $false
    $Email = $false

    $uri = "https://graph.microsoft.com/beta/users/$($Mguser.id)/authentication/signInPreferences"
#Get the current setting, if it matches the preferred method, do nothing, else set the preferred method
    $CurrentDefaults = Invoke-MgGraphRequest -uri $uri -Method GET -OutputType PSObject
    If ($CurrentDefaults.userPreferredMethodForSecondaryAuthentication -eq $UserDefaultMethod){}
    else{
        #Get the user's authentication methods
        $mfaData = $AllAuthMethods | where userprincipalname -eq $mguser.UserPrincipalName | select -ExpandProperty AdditionalProperties
        # Populate $userobject with MFA methods from $mfaData
        if ($mfaData) {
            ForEach ($method in $mfaData.value) {
                Switch ($method['@odata.type']) {
                    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                        # Microsoft Authenticator App
                        $Authenticator = $true
                    }
                    "#microsoft.graph.phoneAuthenticationMethod" {
                        # Phone authentication
                        $Phone = $true
                    }
                    "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                        # Passwordless
                        $PasswordLess = $true
                    }
                    "#microsoft.graph.fido2AuthenticationMethod" {
                        # FIDO2 key
                        $Fido2 = $true
                    }
                    "microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                        # Temporary Access pass
                        $TAP = $true
                    }
                    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                        # Windows Hello
                        $WHFB = $true
                    }
                    "#microsoft.graph.softwareOathAuthenticationMethod" {
                        # ThirdPartyAuthenticator
                        $3part = $true
                    }
                    "#microsoft.graph.emailAuthenticationMethod" {
                        # Email Authentication
                        $Email = $true
                    }
                }
            }
        }
        #Set the default method if the user has a registered method that matches the preferred method
        If ((($UserDefaultMethod -eq "push") -or ($UserDefaultMethod -eq "oath"))-and ($Authenticator -eq $true))
            {
            Invoke-MgGraphRequest -uri $uri -Body $body -Method PATCH
        }
        elseIf ((($UserDefaultMethod -eq "phone") -or ($UserDefaultMethod -eq "voiceMobile") -or ($UserDefaultMethod -eq "voiceAlternateMobile") -or ($UserDefaultMethod -eq "voiceOffice") -or ($UserDefaultMethod -eq "sms")) -and ($Phone -eq $true))
            {
            Invoke-MgGraphRequest -uri $uri -Body $body -Method PATCH
        }
        else{
            #User has no registered method that matches the preferred method
        }
    }
    $CurrentMgUser++
    
    #progressbar
    if($CurrentMgUser % 100 -eq 0){
        $Elapsedtime = (get-date) - $starttime
        $timeLeft = [TimeSpan]::FromMilliseconds((($ElapsedTime.TotalMilliseconds / $CurrentMgUser) * ($TotalMgUsers - $CurrentMgUser)))
        Write-Progress -Activity "Getting Authentication Methods for users $($CurrentMgUser) of $($TotalMgUsers)" `
            -Status "Est Time Left: $($timeLeft.Hours) Hours, $($timeLeft.Minutes) Minutes, $($timeLeft.Seconds) Seconds" `
            -PercentComplete $([math]::ceiling($($CurrentMgUser / $TotalMgUsers) * 100))
        }
}
Disconnect-Graph
#endregion