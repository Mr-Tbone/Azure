<#PSScriptInfo
.SYNOPSIS
    Script for Intune Remediation to remediate and add Shortcuts (Example calculator & notepad)

.DESCRIPTION
    This script will Remediate and change or add shortcuts and verify the correct targets
    It will log the transcript and write it to eventlog, The transcript text will also be returned to Intune

.NOTES
    .AUTHOR         Mr Tbone Granheden @MrTbone_se 
    .COMPANYNAME    Coligo AB @coligoAB
    .COPYRIGHT      Feel free to use this, but would be grateful if my name is mentioned in notes

.RELESENOTES
    1.0 Initial version
#>

#region ------------------------------------------------[Set script requirements]------------------------------------------------
#Requires -Version 4.0
#Requires -RunAsAdministrator
#endregion

#region -------------------------------------------------[Modifiable Parameters]-------------------------------------------------
$RemediationName    = "Shortcut-Apps"    # Used for Eventlog
$Logpath            = "$($env:TEMP)"    # Path to log transcript
$Shortcuts = @(
    @{
        Name        = "Notepad"                      # Name of the shortcut
        TargetPath  = "c:\Windows\System32\notepad.exe" # Path to the target
        Location    = "Desktop"                         # Valid options is Desktop, StartMenu, Startup, Favorites, CommonDesktop, CommonStartmenu, CommonStartup       
        IconPath    = ""    # if no icon is specified, the default icon is used
        IconIndex   = 0     # if 0 use default first icon, if 1 use second icon etc
    },
    @{
        Name        = "Calculator"    
        TargetPath  = "c:\Windows\System32\calc.exe"
        Location    = "CommonDesktop"
        IconPath    = "C:\Windows\System32\calc.exe"
        IconIndex   = 0
    }
)
#endregion

#region --------------------------------------------------[Static Variables]-----------------------------------------------------
#Declare variables
[string]$Transcript = $null
[string]$EventType  = $null
[int32]$eventID     = $null
#set Eventsource and Logfile depending on remediation or detection mode
[string]$Logfile    = "$($Logpath)\Detect-$($RemediationName).log"
[string]$eventsource="Detect-$($RemediationName)"
#endregion

#region --------------------------------------------[Import Modules and Extensions]----------------------------------------------
#endregion

#region ------------------------------------------------------[Functions]--------------------------------------------------------
Function Remediate-Shortcuts {
    Param(
        [array]$Shortcuts
    )
    Begin {}
    Process {
        Foreach ($Shortcut in $Shortcuts) {
            $ShortcutPath = "$([Environment]::GetFolderPath("$($Shortcut.ShortcutLocation)"))\$($Shortcut.ShortcutName).lnk"
            $Shell = New-Object -ComObject WScript.Shell
            if (Test-Path -Path $ShortcutPath -ErrorAction SilentlyContinue) {
                $ShortcutObject = $Shell.CreateShortcut($ShortcutPath)
                if (($null -ne $ShortcutObject.TargetPath) -and ($error.count -eq 0)) {
                    if ("$([string]$ShortcutObject.TargetPath)" -replace '\s' -ceq "$($Shortcut.TargetPath)" -replace '\s') {
                        Write-Verbose "The target path of shortcut $($ShortcutPath) is correct." -Verbose
                    }
                    else {
                        $ShortcutObject.TargetPath = $Shortcut.TargetPath
                        $ShortcutObject.Save()
                        Write-Verbose "The target path of shortcut $($ShortcutPath) has been remediated." -Verbose
                    }
                }
                else {
                    $ShortcutObject.TargetPath = $Shortcut.TargetPath
                    $ShortcutObject.Save()
                    Write-Verbose "The target path of shortcut $($ShortcutPath) has been remediated." -Verbose
                }
                if (($null -ne $ShortcutObject.IconLocation) -and ($error.count -eq 0)) {
                    if ($ShortcutObject.IconLocation -replace '\s' -ceq "$($Shortcut.IconPath)" -replace '\s') {
                        if ($ShortcutObject.IconIndex -eq $Shortcut.IconIndex) {
                            Write-Verbose "The icon path and index of shortcut $($ShortcutPath) are correct." -Verbose
                        }
                        else {
                            $ShortcutObject.IconIndex = $Shortcut.IconIndex
                            $ShortcutObject.Save()
                            Write-Verbose "The icon index of shortcut $($ShortcutPath) has been remediated." -Verbose
                        }
                    }
                    else {
                        $ShortcutObject.IconLocation = $Shortcut.IconPath
                        $ShortcutObject.IconIndex = $Shortcut.IconIndex
                        $ShortcutObject.Save()
                        Write-Verbose "The icon path and index of shortcut $($ShortcutPath) have been remediated." -Verbose
                    }
                }
                else {
                    $ShortcutObject.IconLocation = $Shortcut.IconPath
                    $ShortcutObject.IconIndex = $Shortcut.IconIndex
                    $ShortcutObject.Save()
                    Write-Verbose "The icon path and index of shortcut $($ShortcutPath) have been remediated." -Verbose
                }
            }
            else {
                Write-Warning "The shortcut $($ShortcutPath) does not exist."
                return $false
            }
        }
    }
    End {}
}

function Write-ToEventlog {
    Param(
        [string]$Logtext,
        [string]$EventSource,
        [int]$EventID,
        [validateset("Information", "Warning", "Error")]$EventType = "Information"
    )
    Begin {}
    Process {
    if (!([System.Diagnostics.EventLog]::SourceExists($EventSource))) {
        New-EventLog -LogName 'Application' -Source $EventSource -ErrorAction ignore | Out-Null
        }
    Write-EventLog -LogName 'Application' -Source $EventSource -EntryType $EventType -EventId $EventID -Message $Logtext -ErrorAction ignore | Out-Null
    }
    End {}
}
#endregion
#region --------------------------------------------------[Script Execution]-----------------------------------------------------
#Clear all Errors before start
$Error.Clear()
#Start logging
start-transcript -Path $Logfile

#Detect
$Detected = Remediate-Shortcuts $Shortcuts

#Get transcript and cleanup the text
Stop-Transcript |out-null
$Transcript = ((Get-Content $LogFile -Raw) -split ([regex]::Escape("**********************")))[-3]
#endregion

#region -----------------------------------------------------[Detection]---------------------------------------------------------
#Compliant
if ($Detected){
    $EventText =  "Compliant - No need to Remediate `n$($Transcript)";$eventID=10;$EventType="Information"
    Write-ToEventlog $EventText $EventSource $eventID $EventType
    Write-output "$($EventText -replace "`n",", " -replace "`r",", ")" #with no line breaks
    Exit 0
}
#Non Compliant
Else{
    $EventText =  "NON Compliant - Need to Remediate `n$($Transcript)";$eventID=11;$EventType="Warning"
    Write-ToEventlog $EventText $EventSource $eventID $EventType
    Write-output "$($EventText -replace "`n",", " -replace "`r",", ")" #with no line breaks
    Exit 1
}
#endregion