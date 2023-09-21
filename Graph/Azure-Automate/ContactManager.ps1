<#PSScriptInfo
.SYNOPSIS
    ContactManager script for Azure Automation

.DESCRIPTION
    This script will add a contact folder to users in the users group and add contacts from contacts group to that contact folder
    The script will also send a notification message to users in Nofify group with a change report and an attached log file
    The script need six powershell modules:
    Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.PersonalContacts, Microsoft.Graph.Users.Actions and PsLogLite
    The Managed Account needs 4 permissions in Azure AD: Group.Read.All, User.Read.All, Contacts.ReadWrite and Mail.Send
        
.EXAMPLE
    .\ContactManager.ps1

.NOTES

.VERSION
    1.0

.RELEASENOTES
    1.0 2022-02-18 Initial Build
 
 .AUTHOR
    Tbone Granheden @ Coligo AB
    @MrTbone_se

.COMPANYNAME 
    Coligo AB

.GUID 
    00000000-0000-0000-0000-000000000000

.COPYRIGHT
    Feel free to use this, but would be grateful if my name is mentioned in notes 

.CHANGELOG
      
#>

#region ---------------------------------------------------[Set script requirements]-----------------------------------------------
#
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.PersonalContacts, Microsoft.Graph.Users.Actions, PsLogLite
#endregion

#region ---------------------------------------------------[Script Parameters]-----------------------------------------------
#endregion

#region ---------------------------------------------------[Set global script settings]--------------------------------------------
#Use replacement module PsLogLite for logging and output to get a logfile and output to Azure Automation
Using Module PsLogLite
Set-StrictMode -Version Latest
#endregion

#region ---------------------------------------------------[Modifiable Parameters and defaults]------------------------------------

#Variables for testing script execution
$GlobalRunMode      = "Prod"        #"Test" = No changes will be made,"Prod" = Script will run for production
$TestOneTarget      = $false        #"$true" = Will only process one single user in $TestTargetUser, "$false" = Production mode
$TestTargetUser     = "me@coligo.se"#Test Target user email when running in $TestOneTarget mode"
$TestSendReport     = $false        #"$true" = Will only send report one single user in to $testReportTarget, "$false" = Production mode
$TestReportRecipient= "me@coligo.se"#Test Report Recipipent when running $TestSendReport mode

#Script Variables
$CMUsersID      = "459ba4fb-0ecc-4cf5-bfa5-02ed6a17be31"    #GroupId of group with users that will get contacts
$CMContactsID   = "c1d405f2-f6ca-42c1-86db-eb97ab6d7d4a"    #GroupId of group with contacts that will be added to users
$CMNotifysID    = "4c92c297-9775-4f28-b768-459fb567b640"    #GroupId of group with users that will get the report email
$Contactlist    = "Coligo Employees"                        #Name of the contact folder that will be added to users
$ReferenceUser  = "mr.tbone@coligo.se"                      #Reference user that will be used to compare changes
$MsgFrom        = "contactmanager@coligo.se"                #From address of notification message
#endregion

#region ---------------------------------------------------[Static Variables]------------------------------------------------------
#Log settings
Set-LogLevel    Output # Can be set to "Output"= default, "Verbose" = Verbose logging, "Debug" = Debug logging 
$logfile        = "$($env:TEMP)\$(get-date -format yyMMdd).txt"
set-logpath     "$($logfile)"

#Initialize variables
$CMUsersGroupMembers    = @()
$CMContactsGroupMembers = @()
$CMNotifysGroupMembers  = @()
$added                  = 0
$updated                = 0
$removed                = 0
$errors                 = 0
$skipped                = 0
$skippedUsers           = 0
$warnings               = 0
$Totalwarnings          = 0
$TotalAdded             = 0
$TotalUpdated           = 0
$TotalRemoved           = 0
$Totalerrors             = 0
$userReports            = @()
[datetime]$startTime     = (Get-Date -format g) 
#endregion

#region ---------------------------------------------------[Import Modules and Extensions]-----------------------------------------
import-module Microsoft.Graph.Authentication
import-module Microsoft.Graph.Groups
import-module Microsoft.Graph.Users
import-module Microsoft.Graph.PersonalContacts
import-module Microsoft.Graph.Users.Actions
import-module PsLogLite
#endregion

#region ---------------------------------------------------[Functions]------------------------------------------------------------
function GetGraphToken {
#Get an AccessToken to Microsoft Graph
    Try{
        $resourceURL    = "https://graph.microsoft.com/" 
        $response       = [System.Text.Encoding]::Default.GetString((Invoke-WebRequest -UseBasicParsing -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" -Method 'GET' -Headers @{'X-IDENTITY-HEADER' = "$env:IDENTITY_HEADER"; 'Metadata' = 'True'}).RawContentStream.ToArray()) | ConvertFrom-Json 
        $accessToken    = $response.access_token
        Write-verbose "$($GlobalRunMode) - Success getting an Access Token to Graph"
        }
    catch{$errors++;Write-Error "$($GlobalRunMode) - Failed getting an Access Token to Graph, with error: $_"}
    return $accessToken
}

function GetGroup {
#Get group from Azure AD
Param(
    [parameter(Mandatory = $true)]
    [String]
    $GroupID
)
    try{$Group = get-mggroup -GroupId $GroupID
        write-verbose "$($GlobalRunMode) - Success to get Group with name $($Group.displayname)"}
    catch{$errors++;write-error "$($GlobalRunMode) - Failed to get Group with id $($GroupID) with error: $_"}
    if($Group){write-verbose "$($GlobalRunMode) - Success to get Group with name $($Group.displayname)"}
    else {write-error "$($GlobalRunMode) - Failed to get Group with id $($GroupID)"}
return $Group
}

Function Populate-MessageRecipient {
    [cmdletbinding()]
    Param(
        [array]$ListOfAddresses )
    ForEach ($SMTPAddress in $ListOfAddresses) {
        @{
            emailAddress = @{address = $SMTPAddress}
        }    
    }    
}
Function Populate-Attachments {
    [cmdletbinding()]
    Param(
        [array]$ListOfAttachments )
	
    ForEach ($Attachment in $ListOfAttachments) {
     Write-Host "Processing" $Attachment
     $EncodedAttachmentFile = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Attachment))
        @{
            "@odata.type"= "#microsoft.graph.fileAttachment"
            name = ($Attachment -split '\\')[-1]
            contentBytes = $EncodedAttachmentFile
            contentType = "text/plain"
          }
        }    
    }

#endregion

#region ---------------------------------------------------[[Script Execution]------------------------------------------------------

#Connect to Graph
$GraphToken = GetGraphToken
Try{
	Connect-mgGraph -AccessToken $GraphToken
	Write-output "$($GlobalRunMode) - Success to connect to Graph"}
catch{$errors++;Write-Error "$($GlobalRunMode) - Failed to connect to Graph, with error: $_"}

#To get all required user properties, it needs to use Beta profile
Select-MgProfile beta

#get Contactmanager Users group
$CMUsersGroup = GetGroup $CMUsersID

#get Contactmanager Users Members
$GroupID = $CMUsersGroup.id
$Groupmembers = @()
try{[array]$Groupmembers = get-mggroupmember -GroupId $GroupID -all
    write-output "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) - Success to get $($Groupmembers.count) Group members from group with id $($GroupID)"}
catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) - Failed to get Group members from group with id $($GroupID) with error: $_"}
if($Groupmembers){write-verbose "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) -Success to get Group members from group with id $($GroupID)"}
else {write-error "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) - Failed to get Group members from group with id $($GroupID)"}
if($TestOneTarget){
    $TestTargetuserinfo = get-mguser -UserId $TestTargetUser
    $CMUsersGroupMembers += $TestTargetuserinfo}
else{
    foreach ($user in $Groupmembers) {
        $userobject = New-Object System.Object
        try{$userobject = Get-MgUser -UserId $user.Id | Select-Object id,userprincipalname,DisplayName,GivenName,Surname,middlename,Mail,MobilePhone,BusinessPhones,OfficeLocation,JobTitle,Photo,Companyname,Department 
            write-verbose "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) - Success to get info for user $($User.Id)"}
        catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMUsersGroup.displayname) - Failed to get info for user $($User.Id) with error: $_"}
    #Add additional info to userobject
    if (!($license = Get-MgUserlicensedetail -UserId $user.Id)){
        $userReport ="<tr style='color:red'><td>$($userobject.displayname)</td><td>$($userobject.mail)</td><td>Skipped,</td><td>User</td></td><td>have</td></td><td>no</td></td><td>mailbox</td><td>or license</td></tr>";$userReports += $userReport
        $errors++;write-error "$($GlobalRunMode) - User: $($userobject.displayname) - Skipped due to no license"
        $skippedUsers++;continue}
    else{if($photo = Get-MgUserPhoto -UserId $user.Id -erroraction silentlycontinue){$PhotoExist = $true}
        else{$photoExist = $false}
        }
        $userObject | Add-Member -Name "PhotoExist" -MemberType NoteProperty -Value $PhotoExist
        $userObject | Add-Member -Name "BusinessPhone" -MemberType NoteProperty -Value ($userobject.businessphones | select -First 1)
        $CMUsersGroupMembers += $userobject
        }
    }

#get Contactmanager Contacts group
$CMContactsGroup = GetGroup $CMContactsID

#get Contactmanager Contacts Members
$GroupID = $CMContactsGroup.id
$Groupmembers = @()
try{[array]$Groupmembers = get-mggroupmember -GroupId $GroupID -all
    write-output "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Success to get $($Groupmembers.count) Group members from group with id $($GroupID)"}
catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Failed to get Group members from group with id $($GroupID) with error: $_"}
if($Groupmembers){write-verbose "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Success to get Group members from group with id $($GroupID)"}
else {write-error "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Failed to get Group members from group with id $($GroupID)"}
foreach ($user in $Groupmembers) {
    $userobject = New-Object System.Object
    try{$userobject = Get-MgUser -UserId $user.Id | Select-Object id,userprincipalname,DisplayName,GivenName,Surname,middlename,Mail,MobilePhone,BusinessPhones,OfficeLocation,JobTitle,Photo,Companyname,Department 
        write-verbose "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Success to get info for user $($User.Id)"}
    catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) - Failed to get info for user $($User.Id) with error: $_"}
    #Add additional info to userobject
    if (!($license = Get-MgUserlicensedetail -UserId $user.Id)){
        $errors++;write-error "$($GlobalRunMode) - User: $($userobject.displayname) - Contact has no license, Skipped Contact"
        $skipped++;continue}
    else{if($photo = Get-MgUserPhoto -UserId $user.Id -erroraction silentlycontinue){
        $PhotoExist = $true
        try{Get-MgUserPhotoContent -userid $userobject.mail -outfile  "$($env:TEMP)\$($userobject.mail).jpg" |Out-Null
            write-verbose "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) -  Success to get photo of user $($userobject.displayName)"}
        catch{$warnings++;write-warning "$($GlobalRunMode) - Group: $($CMContactsGroup.displayname) -  Failed to get photo of user $($userobject.displayName) with error: $_"}
        }
        else{$photoExist = $false}
        }
    $userObject | Add-Member -Name "PhotoExist" -MemberType NoteProperty -Value $PhotoExist
    $userObject | Add-Member -Name "BusinessPhone" -MemberType NoteProperty -Value ($userobject.businessphones | select -First 1)
    $CMContactsGroupMembers += $userobject
    }

#get Contactmanager Notifys group
$CMNotifysGroup = GetGroup $CMNotifysID

#get Contactmanager Notifys Members
$GroupID = $CMNotifysGroup.id
$Groupmembers = @()
try{[array]$Groupmembers = get-mggroupmember -GroupId $GroupID -all
    write-output "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Success to get $($Groupmembers.count) Group members from group with id $($GroupID)"}
catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Failed to get Group members from group with id $($GroupID) with error: $_"}
if($Groupmembers){write-verbose "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Success to get Group members from group with id $($GroupID)"}
else {write-error "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Failed to get Group members from group with id $($GroupID)"}
if($TestSendReport){
    $userObject = New-Object System.Object
    $userObject | Add-Member -Name "mail" -MemberType NoteProperty -Value $TestReportRecipient
    $CMNotifysGroupMembers += $userObject}
else{
    #Add additional info to userobject
    foreach ($user in $Groupmembers) {
        $userobject = New-Object System.Object
        try{$userobject = Get-MgUser -UserId $user.Id | Select-Object id,userprincipalname,DisplayName,GivenName,Surname,middlename,Mail,MobilePhone,BusinessPhones,OfficeLocation,JobTitle,Photo,Companyname,Department 
            write-verbose "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Success to get info for user $($User.Id)"}
        catch{$errors++;write-error "$($GlobalRunMode) - Group: $($CMNotifysGroup.displayname) - Failed to get info for user $($User.Id) with error: $_"}
        if (!($license = Get-MgUserlicensedetail -UserId $user.Id)){
            $errors++;write-error "$($GlobalRunMode) - User: $($userobject.displayname) - Notify user has no license, Skipped notify user"
            continue}
        else{if($photo = Get-MgUserPhoto -UserId $user.Id -erroraction silentlycontinue){$PhotoExist = $true}
            else{$photoExist = $false}
            }
        $userObject | Add-Member -Name "photoExist" -MemberType NoteProperty -Value $PhotoExist
        $userObject | Add-Member -Name "businessPhone" -MemberType NoteProperty -Value ($userobject.businessphones | select -First 1)
        $CMNotifysGroupMembers += $userobject
        }
    }
$Totalerrors = $errors;$errors = 0
$Totalwarnings = $warnings;$warnings=0

#switch to release v1.0 to get contact info correctly
Select-MgProfile "v1.0"

#loop through all users in Contactmanager users group
$i=1
foreach ($CMUser in $CMUsersGroupMembers){
    write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Start processing $($i) of $($CMUsersGroupMembers.count)"
    $i++
    #get contact folder for the user, create if not exist 
    $ContactFolder = $null
    if(!($ContactFolder = Get-MgUserContactFolder -UserId $CMUser.id -Filter "displayname eq '$Contactlist'")){
        $warnings++;write-warning "$($GlobalRunMode) - User: $($CMUser.displayname) - Contactfolder $($Contactlist) does not exist, needs to be created"
        try{if($GlobalRunMode -ne "Test"){New-MgUserContactFolder -userid $CMUser.id -displayname $contactlist}
            write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to create contactfolder"}
        catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to create contactfolder with error: $_"}
        if(!($ContactFolder = Get-MgUserContactFolder -UserId $CMUser.id -Filter "displayname eq '$Contactlist'"))
        {write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to create contactfolder, continue next user"
        continue}
        }

    #get contacts in contacts folder for the user, create if not exist 
    [array]$PersonalContacts = @()
    try{[array]$PersonalContacts = Get-MgUserContactFolderContact -UserId "$($CMUser.id)" -ContactFolderId $ContactFolder.id -all 
        write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to get Personal contacts"}
    catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to get Personal contacts with error: $_"
        continue}
    if ([array]$PersonalContacts) {write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to get $([array]$PersonalContacts.count) Personal contacts"}

    #Add attributes to personal contacts object
    foreach ($PersonalContact in $PersonalContacts){
        $PersonalContact | Add-Member -Name "businessPhone" -MemberType NoteProperty -Value (($PersonalContact | select -ExpandProperty businessphones)|select -First 1)
        $PersonalContact | Add-Member -Name "mail" -MemberType NoteProperty -Value (($PersonalContact.emailaddresses | Select-Object -ExpandProperty address) |select -First 1)
        if ($photo = Get-MgUserContactFolderContactPhoto -UserId $CMUser.id -ContactFolderId $ContactFolder.id -ContactId $PersonalContact.id -erroraction silentlycontinue){$PhotoExist = $true}
        else{$photoExist = $false} 
        $PersonalContact | Add-Member -Name "photoExist" -MemberType NoteProperty -Value $PhotoExist
        }

    #loop throuh all contactmanager contacts and verify if they exist in personal contacts
    foreach ($CMContact in $CMContactsGroupMembers){
        write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Checking existance of Contact: $($cmcontact.displayName) with email: $($CMContact.mail)"

        #check if contact exist in current personal contacts
        $Contactmatch = $null
        if ($Contactmatch = $PersonalContacts | where-object {$_.mail -eq $CMContact.mail}){
            write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Found existing contact: $($contactmatch.displayName) with email: $($contactmatch.mail)"
            if ($Contactmatch -is [array]){
                $isFirst = $true
                foreach ($Contact in $Contactmatch){
                    if ($isFirst) {
                        $isFirst = $false
                        $Contactmatches = $null
                        $Contactmatches = $Contact
                        }
                        else{  
                        try{if($GlobalRunMode -ne "Test"){Remove-MgUserContactFolderContact -userid $CMUser.id -ContactFolderId $ContactFolder.id -ContactID $Contact.id}
                            $removed++
                            write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to remove duplicate contact $($Contact.displayName)"}
                        catch{$warnings++;write-warning "$($GlobalRunMode) - User: $($CMUser.displayname) - Failes to remove duplicate contact $($Contact.displayName) with error: $_"}
                        }
                    }
                    $Contactmatch = $null
                    $Contactmatch = $Contactmatches 
                }
            $changed = $false
            $Attribute = ""
            if (((!([string]::IsNullOrWhitespace($CMContact.givenName))) -or (!([string]::IsNullOrWhitespace($Contactmatch.givenName)))) -and ($CMContact.givenName -ne $Contactmatch.givenName)){
                $changed = $true;$Attribute=$Attribute+"givenname,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New Givenname: $($CMContact.givenName) old givenname $($Contactmatch.givenname)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.middleName))) -or (!([string]::IsNullOrWhitespace($Contactmatch.middleName)))) -and ($CMContact.middleName -ne $Contactmatch.middleName)){
                $changed = $true;$Attribute=$Attribute+"middleName,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New middleName: $($CMContact.middleName) old middleName $($Contactmatch.middleName)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.surname))) -or (!([string]::IsNullOrWhitespace($Contactmatch.surname)))) -and ($CMContact.surname -ne $Contactmatch.surname)){
                $changed = $true;$Attribute=$Attribute+"surname,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New surname: $($CMContact.surname) old surname $($Contactmatch.surname)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.jobTitle))) -or (!([string]::IsNullOrWhitespace($Contactmatch.jobTitle)))) -and ($CMContact.jobTitle -ne $Contactmatch.jobTitle)){
                $changed = $true;$Attribute=$Attribute+"jobTitle,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New jobTitle: $($CMContact.jobTitle) old jobTitle $($Contactmatch.jobTitle)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.department))) -or (!([string]::IsNullOrWhitespace($Contactmatch.department)))) -and ($CMContact.department -ne $Contactmatch.department)){
                $changed = $true;$Attribute=$Attribute+"department,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New department: $($CMContact.department) old department $($Contactmatch.department)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.companyName))) -or (!([string]::IsNullOrWhitespace($Contactmatch.companyName)))) -and ($CMContact.companyName -ne $Contactmatch.companyName)){
                $changed = $true;$Attribute=$Attribute+"companyName,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New companyName: $($CMContact.companyName) old companyName $($Contactmatch.companyName)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.OfficeLocation))) -or (!([string]::IsNullOrWhitespace($Contactmatch.OfficeLocation)))) -and ($CMContact.OfficeLocation -ne $Contactmatch.OfficeLocation)){
                $changed = $true;$Attribute=$Attribute+"OfficeLocation,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New OfficeLocation: $($CMContact.OfficeLocation) old OfficeLocation $($Contactmatch.OfficeLocation)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.mobilePhone))) -or (!([string]::IsNullOrWhitespace($Contactmatch.mobilePhone)))) -and ($CMContact.mobilePhone -ne $Contactmatch.mobilePhone)){
                $changed = $true;$Attribute=$Attribute+"mobilePhone,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New mobilePhone: $($CMContact.mobilePhone) old mobilePhone $($Contactmatch.mobilePhone)"}
            if (((!([string]::IsNullOrWhitespace($CMContact.businessPhone))) -or (!([string]::IsNullOrWhitespace($Contactmatch.businessPhone)))) -and ($CMContact.businessPhone -ne $Contactmatch.businessPhone)){
                $changed = $true;$Attribute=$Attribute+"businessPhone,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - New businessPhone: $($CMContact.businessPhone) old businessPhone $($Contactmatch.businessPhone)"}
            if (($CMContact.PhotoExist) -and (!($Contactmatch.PhotoExist))){
                $changed = $true;$Attribute=$Attribute+"Photo,"; write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Photo added in AAD, and missing on contact"}

            #Contact has changed and needs an update
            if ($changed){
                write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Contact $($Contactmatch.displayName) has changed $($Attribute)"
                try{if($GlobalRunMode -ne "Test"){Update-MgUserContactFolderContact -userid $CMUser.id -ContactId $Contactmatch.id -contactFolderId $ContactFolder.id -fileas "$($CMContact.displayName)" -GivenName "$($CMContact.givenname)" -MiddleName "$($CMContact.middlename)" -surname "$($CMContact.surname)" -CompanyName "$($CMContact.companyname)" -Department "$($CMContact.department)" -JobTitle "$($CMContact.jobtitle)" -EmailAddresses @{name="$($CMContact.displayname)";address="$($CMContact.mail)"} -MobilePhone "$($CMContact.MobilePhone)" -OfficeLocation "$($CMContact.OfficeLocation)" -BusinessPhones "$($CMContact.businessPhone)"}
                    $updated++
                    write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to Update contact $($Contactmatch.displayName)"}
                catch{$warnings++;write-Warning "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to Update contact $($Contactmatch.displayName) with error: $_"}
                if ($CMContact.PhotoExist){
                    try{if($GlobalRunMode -ne "Test"){Set-MgUserContactFolderContactPhotoContent -UserId $CMUser.id -ContactFolderId $ContactFolder.id -ContactId $Contactmatch.id -infile "$($env:TEMP)\$($CMContact.mail).jpg" |Out-Null}
                        write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to Update photo on contact $($CMContact.displayName)"}
                    catch{$warnings++;write-Warning "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to Update photo on contact $($CMContact.displayName) with error: $_"}
                    }
                }
            else{write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - No Need to update existing contact: $($contactmatch.displayname) with email: $($contactmatch.mail)"}
            }

            #Contact does not exist and needs to be created            
            else{write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Contact not found, needs to create new contact: $($cmcontact.displayname) with email: $($CMContact.mail)"
                try{if($GlobalRunMode -ne "Test"){$newcontact = New-MgUserContactFolderContact -userid $CMUser.id -contactFolderId $ContactFolder.id -fileas "$($CMContact.displayName)" -GivenName "$($CMContact.givenname)" -MiddleName "$($CMContact.middlename)" -surname "$($CMContact.surname)" -CompanyName "$($CMContact.companyname)" -Department "$($CMContact.department)" -JobTitle "$($CMContact.jobtitle)" -EmailAddresses @{name="$($CMContact.displayname)";address="$($CMContact.mail)"} -MobilePhone "$($CMContact.MobilePhone)" -OfficeLocation "$($CMContact.OfficeLocation)" -BusinessPhones "$($CMContact.businessPhone)"}
                    $added++
                    write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to create contact $($CMContact.displayName)"}
                catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to create contact $($CMContact.displayName) with error: $_"}
                if ($CMContact.PhotoExist){
                    try{if($GlobalRunMode -ne "Test"){Set-MgUserContactFolderContactPhotoContent -UserId $CMUser.id -ContactFolderId $ContactFolder.id -ContactId $newcontact.id -infile "$($env:TEMP)\$($CMContact.mail).jpg" |Out-Null}
                        write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to Update photo on contact $($CMContact.displayName)"}
                    catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to Update photo on contact $($CMContact.displayName) with error: $_"}
                    }
            }
        }   
    #loop through all existing personal contacts and verify if they exist in contactmanager group, remove if not
        [array]$PersonalContacts = @()
        try{[array]$PersonalContacts = Get-MgUserContactFolderContact -UserId "$($CMUser.id)" -ContactFolderId $ContactFolder.id -all 
            write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to get Personal contacts"}
        catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Failed to get Personal contacts with error: $_"
            continue}
        if ([array]$PersonalContacts) {write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to get $([array]$PersonalContacts.count) Personal contacts"}

            foreach ($PersonalContact in $PersonalContacts){
                $PersonalContact | Add-Member -Name "mail" -MemberType NoteProperty -Value (($PersonalContact.emailaddresses | Select-Object -ExpandProperty address) |select -First 1)
                write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Checking existance of Contact: $($PersonalContact.displayName) with email: $($PersonalContact.mail)"
                $Contactmatch = $null
                if ($Contactmatch = $CMContactsGroupMembers | where-object {$_.mail -eq $PersonalContact.mail}){
                    write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Keep existing contact: $($contactmatch.displayName) with email: $($contactmatch.mail)"}
                else{
                    write-output "$($GlobalRunMode) - User: $($CMUser.displayname) - Need to remove old contact: $($PersonalContact.displayName) with email: $($PersonalContact.mail)"
                    try{if($GlobalRunMode -ne "Test"){Remove-MgUserContactFolderContact -userid $CMUser.id -ContactFolderId $ContactFolder.id -ContactID $PersonalContact.id}
                        $removed++
                        write-verbose "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to remove contact $($PersonalContact.displayName)"}
                    catch{$errors++;write-error "$($GlobalRunMode) - User: $($CMUser.displayname) - Success to remove contact $($PersonalContact.displayName) with error: $_"}
            }
        }
        if ($warnings -gt 0){$UserReport     = "<tr style='color:red'><td>$($CMUser.displayname)</td><td>($($CMUser.mail))</td><td>$($added)</td><td>$($updated)</td><td>$($removed)</td><td>$($skipped)</td><td>$($warnings)</td><td>$($errors)</td></tr>"}
        elseif($errors -gt 0){$UserReport     = "<tr style='color:yellow'><td>$($CMUser.displayname)</td><td>($($CMUser.mail))</td><td>$($added)</td><td>$($updated)</td><td>$($removed)</td><td>$($skipped)</td><td>$($warnings)</td><td>$($errors)</td></tr>"}
        else{$UserReport     = "<tr><td>$($CMUser.displayname)</td><td>($($CMUser.mail))</td><td>$($added)</td><td>$($updated)</td><td>$($removed)</td><td>$($skipped)</td><td>$($warnings)</td><td>$($errors)</td></tr>"}
    $userReports    += $UserReport
    $TotalAdded     = $TotalAdded   + $added;   $added   = 0
    $TotalUpdated   = $TotalUpdated + $updated; $updated = 0
    $TotalRemoved   = $TotalRemoved + $removed; $removed = 0
    $Totalerrors    = $Totalerrors  + $errors;  $errors  = 0
    $Totalwarnings  = $Totalwarnings+ $warnings;$warnings= 0
    }
[datetime]$endTime  = (Get-Date -format g) 

#Send report with E-mail

#Add logfile as attachment
[array]$AttachmentsList = "$($logfile)"
[array]$MsgAttachments = Populate-Attachments -ListOfAttachments $AttachmentsList

#Modify some Report text
if ($Totalerrors -gt 0){$status = "<span style='color:red;'>Failed</span> - Se attached log for more info"}
else {$status = "<span style='color:green;'>successful</span>"}
$referenceReport = $userReports | where {$_ -like "*$($ReferenceUser)*"}
$totaltime = ([datetime]$endTime - [datetime]$startTime)

#Create the HTML header
$htmlhead="<html>
     <style>
      BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
    TD {text-align: center; vertical-align: middle;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"

#Create the HTML body
$HtmlBody = "<body>
     <h1>ContactManager Report $($startTime.AddHours(2))</h1>
     <p><strong>Started:</strong> $($startTime.AddHours(2)) <strong>Ended:</strong> $($endTime.AddHours(2))</p>
     <p><strong>Total Runtime:</strong> $($totaltime.hours) hours $($totaltime.minutes) minutes</p>
     <p><strong>Processed $($CMUsersGroupMembers.count) Users and $($CMContactsGroupMembers.count) Contacts 
     <h2>ContactManager Result was $($Status)</h2>
     <br>
     <p><u><b>Report for a reference User</b></u></p>
     <table><tr><th>User</th><th>email</th><th>Added</th><th>Updated</th><th>Removed</th><th>Skipped</th><th>Warnings</th><th>Errors</th></tr>
     $($referenceReport)
     </table>
     <br>
     <p><u><b>Total for all users</b></u></p>
     <table><tr><th>User</th><th>Added</th><th>Updated</th><th>Removed</th><th>Skipped Contacts</th><th>Skipped Users</th><th>Warnings</th><th>Errors</th></tr>
     <tr><td>All Users</td><td>$($TotalAdded)</td><td>$($TotalUpdated)</td><td>$($TotalRemoved)</td><td>$($Skipped)</td><td>$($skippedUsers)</td><td>$($Totalwarnings)</td><td>$($Totalerrors)</td></tr>
     </table>
     <br>
     <p><u><b>Individual user reports</b></u></p>
     <table><tr><th>User</th><th>email</th><th>Added</th><th>Updated</th><th>Removed</th><th>Skipped</th><th>Warnings</th><th>Errors</th></tr>
     $($userReports)
     </table>"

#Set the message subject
$MsgSubject = "ContactManager Report $($startTime)"

#Build and send E-mail with report to all members of the Notification group
ForEach ($User in $CMNotifysGroupMembers) {
    $ToRecipientList   = @( $User.mail )
    [array]$MsgToRecipients = Populate-MessageRecipient -ListOfAddresses $ToRecipientList
  $HtmlMsg = "</body></html>" + $HtmlHead + $htmlbody + "<p>"
  $MsgBody = @{
     Content = "$($HtmlMsg)"
     ContentType = 'html'   }

  $Message =  @{subject           = $MsgSubject}
  $Message += @{toRecipients      = $MsgToRecipients}  
  $Message += @{attachments       = $MsgAttachments}
  $Message += @{body              = $MsgBody}
  $Params   = @{'message'         = $Message}
  $Params  += @{'saveToSentItems' = $False}
  $Params  += @{'isDeliveryReceiptRequested' = $False}

  try{Send-MgUserMail -UserId $MsgFrom -BodyParameter $Params
  write-verbose "$($GlobalRunMode) - Sendmail - Success to send report to $($user.mail)"}
  catch{write-error "$($GlobalRunMode) - Sendmail - Failed to send report to $($user.mail) with error: $_"}
}

#Cleanup Temporary storage
remove-item "$($env:TEMP)\*.jpg" -force
remove-item "$($env:TEMP)\*.txt" -force

disconnect-mgGraph
#endregion