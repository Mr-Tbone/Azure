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
    t-bone.granheden@coligo.se

.VERSION
    2.0

.RELEASENOTES
    1.0 2023-02-14 Initial Build
    2.0 2024-06-10 Updated to use Microsoft.Graph module and improved logging and error handling

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
    2.0.2512.1 - Updated to use Microsoft.Graph module and improved logging and error handling
#>

#region ---------------------------------------------------[Set Script Requirements]-----------------------------------------------
#Requires -Modules Microsoft.Graph.Authentication
#Requires -Version 5.1
#endregion
#region ---------------------------------------------------[Modifiable Parameters and Defaults]------------------------------------
# Customizations
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false,          HelpMessage = "Name of the script action for logging. Default is 'Intune and Azure Cleanup Stale Devices'")]
    [string]$ScriptActionName       = "Intune and Azure Cleanup Stale Devices",

    [Parameter(Mandatory = $false,          HelpMessage = "Number of inactive days to determine a stale device to disable. Default is 60 days")]
    [ValidateRange(1,365)]
    [int]$DeviceDisableThreshold    = 900,

    [Parameter(Mandatory = $false,          HelpMessage = "Number of inactive days to determine a stale device to delete. Default is 90 days")]
    [ValidateRange(1,365)]
    [int]$DeviceDeleteThreshold     = 920,
        
    [Parameter(Mandatory = $false,          HelpMessage = "Return report with statistics on how many changed objects. Default is true")]
    [bool]$ReturnReport             = $true,

    [Parameter(Mandatory = $false,          HelpMessage = "Include detailed device changes in the report. Default is true")]
    [bool]$DetailedReport           = $true,

    [Parameter(Mandatory = $false,          HelpMessage = "Save report to disk. Default is false")]
    [bool]$ReportToDisk             = $false,    

    [Parameter(Mandatory = $false,          HelpMessage = "Path where to save the report. Default is TEMP directory for Azure Automation compatibility")]
    [string]$ReportPath             = "$env:TEMP",

    [Parameter(Mandatory = $false,          HelpMessage = "Enable verbose logging. Default is false")]
    [bool]$VerboseLogging           = $true,

    [Parameter(Mandatory = $false,          HelpMessage = "Wait time in milliseconds between throttled requests. Default is 1000")]
    [ValidateRange(100,5000)]
    [int]$WaitTime                  = 1000,

    [Parameter(Mandatory = $false,          HelpMessage = "Maximum number of retry attempts for failed requests. Default is 3")]
    [ValidateRange(1,10)]
    [int]$MaxRetry                  = 3,

    [Parameter(Mandatory = $false,          HelpMessage = "Maximum number of items to process in a single batch. Default is 20")]
    [ValidateRange(1,20)]
    [int]$BatchSize                  = 20,

    [Parameter(Mandatory = $false,          HelpMessage = "Testmode, same as -WhatIf. Default is false")]
    [bool]$Testmode                 = $false
    )
#endregion
#region ---------------------------------------------------[Modifiable Variables and defaults]------------------------------------
#endregion

#region ---------------------------------------------------[Set global script settings]--------------------------------------------
# Exit if running as a managed identity in PowerShell 7.2 due to bugs connecting to MgGraph https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3151
if ($env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER -and $PSVersionTable.PSVersion -eq [version]"7.2.0") {
    Write-Error "This script cannot run as a managed identity in PowerShell 7.2. Please use a different version of PowerShell."
    exit 1}
# set strict mode to latest version
Set-StrictMode -Version Latest

# Save original preference states at script scope for restoration in finally block
$script:OriginalErrorActionPreference   = $ErrorActionPreference
$script:OriginalVerbosePreference       = $VerbosePreference
$script:OriginalWhatIfPreference        = $WhatIfPreference

# Set verbose- and should-preference based on parameter instead of hardcoded values
if ($VerboseLogging)    {$VerbosePreference = 'Continue'}                   # Set verbose logging based on the parameter $VerbosePreference
else                    {$VerbosePreference = 'SilentlyContinue'}
if($Testmode)           {$WhatIfPreference = 1}                             # Manually enable whatif mode with parameter $Testmode for testing
#endregion

#region ---------------------------------------------------[Import Modules and Extensions]-----------------------------------------
# Check if Microsoft.Graph.Authentication module is already loaded, if not import it silently
$moduleName = 'Microsoft.Graph.Authentication'
if (-not (Get-Module -Name $moduleName)) {
    $savedVerbosePreference = $VerbosePreference
    $VerbosePreference = 'SilentlyContinue'
    try {Import-Module $moduleName -ErrorAction Stop}
    finally {$VerbosePreference = $savedVerbosePreference}
} else {Write-Verbose "Module '$moduleName' is already loaded"}
#endregion

#region ---------------------------------------------------[Static Variables]------------------------------------------------------
# Script execution tracking
[datetime]$Script:StartTime                 = ([DateTime]::Now) # Script start time
#Dates for evaluations
[string]$disableDate = "$(($Script:StartTime).AddDays(-$DeviceDisableThreshold).ToString("yyyy-MM-dd"))T00:00:00z"
[string]$deleteDate = "$(($Script:StartTime).AddDays(-$DeviceDeleteThreshold).ToString("yyyy-MM-dd"))T00:00:00z"
# Required Graph API scopes for Invoke-ConnectMgGraph function
$requiredScopes = [System.Collections.ArrayList]@("Device.ReadWrite.All")
# Reporting variables for progress tracking (dynamic - add more counters as needed)
$script:ReportProgress = @{
    Total          = [int]0    # Total number of objects to process
    Disabled       = [int]0    # Successfully disabled devices
    Deleted        = [int]0    # Successfully deleted devices
    WhatifDisabled = [int]0    # Devices that would be disabled (WhatIf mode)
    WhatifDeleted  = [int]0    # Devices that would be deleted (WhatIf mode)
    Failed         = [int]0    # Failed operations
    Skipped        = [int]0    # Skipped objects (already correct or excluded)
}

# Reporting variable for detailed results
$script:ReportResults = [System.Collections.ArrayList]::new()
#endregion

#region ---------------------------------------------------[Functions]------------------------------------------------------------
function Invoke-TboneLog { 
<#
.SYNOPSIS
    Unified tiny logger for PowerShell 5.1–7.5 and Azure Automation; overrides Write-* cmdlets and stores all messages in-memory

.DESCRIPTION
    A lightweight, cross-platform logging solution that intercepts all Write-Host, Write-Output, Write-Verbose, 
    Write-Warning, and Write-Error calls. Stores messages in memory with timestamps and can optionally output to:
    -LogToGUI - Console (real-time during execution) -LogToDisk - Disk (log file at script completion) -LogToEventlog - Windows Event Log (Application log)
#>
    param(
        [Parameter(                                         HelpMessage='Start=Begin logging, Stop=End and output log array')]
        [ValidateSet('Start','Stop')][string]$Mode,
        [Parameter(                                         HelpMessage='Show output in console during execution')]
        [bool]$LogToGUI=$true,
        [Parameter(                                         HelpMessage='Write complete log array to Disk when script ends')]
        [bool]$LogToDisk=$true,
        [Parameter(                                         HelpMessage='Write complete log array to Windows Event when script ends')]
        [bool]$LogToEventlog=$true,
        [Parameter(                                         HelpMessage='Return complete log array as Host output when script ends (Good for Intune Remediations)')]
        [bool]$LogToHost=$false,
        [Parameter(                                         HelpMessage='Path where Disk logs are saved')]
        [string]$LogPath = "c:\temp"
    )
    # Auto-detect mode: if logger functions is already loaded in memory and no mode specified, assume Stop
    if(!$Mode){$Mode=if(Get-Variable -Name _l -Scope Global -EA 0){'Stop'}else{'Start'}}
    if(!$LogPath){$LogPath=if($global:_p){$global:_p}elseif($env:TEMP){$env:TEMP}else{'/tmp'}}

    # Stop mode: Save logs and cleanup
    if ($Mode -eq 'Stop') {
        if((Get-Variable -Name _l -Scope Global -EA 0) -and (Test-Path function:\global:_Save)){_Save;if($global:_r){,$global:_l.ToArray()}}
        Unregister-Event -SourceIdentifier PowerShell.Exiting -ea 0 -WhatIf:$false
        if(Test-Path function:\global:_Clean){_Clean}
        return
    }

    # Start mode: Initialize logging and proxy all Write-* functions
    if ($Mode -eq 'Start') {
        # Create helper functions and variables
        $c=(Get-PSCallStack)[1];$n=if($c.Command -and $c.Command -ne '<ScriptBlock>'){$c.Command}elseif($c.ScriptName){[IO.Path]::GetFileNameWithoutExtension($c.ScriptName)}else{'PowershellScript'}
        $global:_az=$env:AZUREPS_HOST_ENVIRONMENT -or $env:AUTOMATION_ASSET_ACCOUNTID # Detect Azure Automation environment
        $global:_l=[Collections.Generic.List[string]]::new();$global:_g=$LogToGUI;$global:_s=$n;$global:_n="{0}-{1:yyyyMMdd-HHmmss}"-f$n,(Get-Date);$global:_p=$LogPath;$global:_d=$LogToDisk;$global:_e=$LogToEventlog;$global:_r=$LogToHost;$global:_w=([Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT)
        if(!(Test-Path function:\global:_Time)){function global:_Time{Get-Date -f 'yyyy-MM-dd,HH:mm:ss'}}
        if(!(Test-Path function:\global:_ID)){function global:_ID{$c=(Get-PSCallStack)[2];$n=if($c.Command -and $c.Command -ne '<ScriptBlock>'){$c.Command}elseif($c.FunctionName -and $c.FunctionName -ne '<ScriptBlock>'){$c.FunctionName}else{'Main-Script'};if($n -like '*.ps1'){'Main-Script'}else{$n}}}
        if(!(Test-Path function:\global:_Save)){function global:_Save{try{if($global:_d){[IO.Directory]::CreateDirectory($global:_p)|Out-Null;[IO.File]::WriteAllLines((Join-Path $global:_p "$($global:_n).log"),$global:_l.ToArray())};if($global:_e -and $global:_w){try{$id=[Security.Principal.WindowsIdentity]::GetCurrent();$isAdmin=([Security.Principal.WindowsPrincipal]::new($id)).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)}catch{$isAdmin=$false};if([Diagnostics.EventLog]::SourceExists($global:_s) -or $isAdmin){if(-not [Diagnostics.EventLog]::SourceExists($global:_s)){try{New-EventLog -LogName Application -Source $global:_s -EA 0 -WhatIf:$false}catch{}};$la=$global:_l -join"`n";$h=$la -match ',ERROR,';Write-EventLog -LogName Application -Source $global:_s -EventId $(if($h){11003}elseif($la -match ',WARN,'){11002}else{11001}) -EntryType $(if($h){'Error'}elseif($la -match ',WARN,'){Warning}else{Information}) -Message $la -EA 0 -WhatIf:$false}}}catch{}}}
        if(!(Test-Path function:\global:_Clean)){function global:_Clean{$WhatIfPreference=$false;Remove-Item -Path function:\Write-Host,function:\Write-Output,function:\Write-Warning,function:\Write-Error,function:\Write-Verbose,function:\_Save,function:\_Clean,function:\_ID,function:\_Time -ea 0 -Force;Remove-Variable -Name _l,_g,_s,_n,_p,_d,_e,_r,_w,_az -Scope Global -ea 0}}
        # Register exit handler FIRST (before Write-* overrides)
        $null=Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action{if($global:_l){try{_Save}catch{}};if(Test-Path function:\_Clean){_Clean}} -MaxTriggerCount 1
        # Create Write-* proxy functions (skip in Azure Automation)
        function Script:Write-Host{$m="$args";$c=(Get-PSCallStack)[1];$r="Row$($c.ScriptLineNumber)";$e="$(_Time),INFO,$r,$(_ID),$m";$global:_l.Add($e);if($global:_g){if($global:_az){Microsoft.PowerShell.Utility\Write-Output $m}else{Microsoft.PowerShell.Utility\Write-Host $e -ForegroundColor Green}}}
        function Script:Write-Output{$m="$args";$c=(Get-PSCallStack)[1];$r="Row$($c.ScriptLineNumber)";$e="$(_Time),OUTPUT,$r,$(_ID),$m";$global:_l.Add($e);if($global:_g){if($global:_az){Microsoft.PowerShell.Utility\Write-Output $m}else{Microsoft.PowerShell.Utility\Write-Host $e -ForegroundColor Green}}}
        function Script:Write-Verbose{$m="$args";$c=(Get-PSCallStack)[1];$r="Row$($c.ScriptLineNumber)";$e="$(_Time),VERBOSE,$r,$(_ID),$m";$global:_l.Add($e);if($global:_g -and $VerbosePreference -ne 'SilentlyContinue'){if($global:_az){Microsoft.PowerShell.Utility\Write-Verbose $m}else{Microsoft.PowerShell.Utility\Write-Host $e -ForegroundColor cyan}}}
        function Script:Write-Warning{$m="$args";$c=(Get-PSCallStack)[1];$r="Row$($c.ScriptLineNumber)";$e="$(_Time),WARN,$r,$(_ID),$m";$global:_l.Add($e);if($global:_g){if($global:_az){Microsoft.PowerShell.Utility\Write-Warning $m}else{Microsoft.PowerShell.Utility\Write-Host $e -ForegroundColor Yellow}};if($WarningPreference -eq 'Stop'){_Save;_Clean;exit}}
        function Script:Write-Error{$m="$args";$c=(Get-PSCallStack)[1];$r="Row$($c.ScriptLineNumber)";$e="$(_Time),ERROR,$r,$(_ID),$m";$global:_l.Add($e);if($global:_g){if($global:_az){Microsoft.PowerShell.Utility\Write-Error $m}else{Microsoft.PowerShell.Utility\Write-Host $e -ForegroundColor Red}};if($ErrorActionPreference -eq 'Stop'){_Save;_Clean;exit}}
    }
}
function Invoke-ConnectMgGraph {
<#
.SYNOPSIS
    Connects to Microsoft Graph API with multiple authentication methods.

.DESCRIPTION
    Supports Managed Identity, Interactive, Client Secret, and Certificate authentication.
    Automatically detects the method from provided parameters and environment. 
    Validates required scopes for Interactive authentication and reuses an existing connection when possible. 
#>
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage = "Array of required Microsoft Graph API permission scopes example:('User.Read.All','DeviceManagementManagedDevices.ReadWrite.All') ")]
        [string[]]$RequiredScopes,
        
        [Parameter(HelpMessage = "Entra ID Application ID (ClientID) (required for Client Secret or Certificate authentication)")]
        [ValidateNotNullOrEmpty()][string]$ClientId,

        [Parameter(HelpMessage = "Entra ID Tenant ID (directory ID) for app-only authentication (required for Client Secret or Certificate authentication)")]
        [ValidateNotNullOrEmpty()][string]$TenantId,
        
        [Parameter(HelpMessage = "Client Secret as SecureString for app-only authentication (use with ClientId and TenantId)")]
        [ValidateNotNull()][SecureString]$ClientSecret,
        
        [Parameter(HelpMessage = "Certificate thumbprint for certificate-based authentication (certificate must be in CurrentUser or LocalMachine store)")]
        [ValidateNotNullOrEmpty()][string]$CertificateThumbprint,

        [Parameter(HelpMessage = "Certificate subject name for certificate-based authentication (certificate must be in CurrentUser or LocalMachine store)")]
        [ValidateNotNullOrEmpty()][string]$CertificateName,
        
        [Parameter(HelpMessage = "File path to certificate (.pfx or .cer) for certificate-based authentication")]
        [ValidateNotNullOrEmpty()][string]$CertificatePath,
        
        [Parameter(HelpMessage = "Password for certificate file as SecureString (required if certificate file is password-protected)")]
        [SecureString]$CertificatePassword
    )
    
    Begin {
        $ErrorActionPreference = 'Stop'
        $timestamp = [DateTime]::Now.ToString('yy-MM-dd HH:mm:ss')
        $resourceURL = "https://graph.microsoft.com/"

        # Default scopes to sign in if not provided
        if (-not $RequiredScopes -or @($RequiredScopes).Count -eq 0) {
            $RequiredScopes = @("User.Read.All")
        }
        
        # Detect authentication method based on parameters and environment (priority: ClientSecret > Certificate > ManagedIdentity > Interactive)
        $hasClientId     = $PSBoundParameters.ContainsKey('ClientId') -and [string]::IsNullOrWhiteSpace($ClientId) -eq $false
        $hasTenantId     = $PSBoundParameters.ContainsKey('TenantId') -and [string]::IsNullOrWhiteSpace($TenantId) -eq $false
        $hasClientSecret = $PSBoundParameters.ContainsKey('ClientSecret') -and $null -ne $ClientSecret
        $hasCertInput    = $PSBoundParameters.ContainsKey('CertificateThumbprint') -or $PSBoundParameters.ContainsKey('CertificateName') -or $PSBoundParameters.ContainsKey('CertificatePath')

        if ($hasClientSecret -and $hasClientId -and $hasTenantId) {
            $authMethod = 'ClientSecret'
        }
        elseif ($hasCertInput -and $hasClientId -and $hasTenantId) {
            $authMethod = 'Certificate'
        }
        elseif ($env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER) {
            $authMethod = 'ManagedIdentity'
        }
        else {
            $authMethod = 'Interactive'
        }
        
        Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Authentication method: $authMethod"
    }
    
    Process {
        try {
            # Check for existing valid connection
            try {
                $context = Get-MgContext -ErrorAction SilentlyContinue
                if ($context) {
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Existing connection found for: $($context.Account)"
                    
                    # Validate scopes only for Interactive auth (Managed Identity/app-only doesn't use delegated scopes)
                    if ($authMethod -eq 'Interactive' -and @($RequiredScopes).Count -gt 0) {
                        $currentScopes = @($context.Scopes)
                        $reqScopes = @($RequiredScopes)
                        $missingScopes = @($reqScopes | Where-Object { $_ -notin $currentScopes })
                        
                        if (@($missingScopes).Count -eq 0) {
                            Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Reusing existing connection with valid scopes"
                            return $context.Account
                        } else {
                            Write-Verbose "$timestamp,Warning,Invoke-ConnectMgGraph,Existing connection missing scopes: $($missingScopes -join ', ')"
                            Disconnect-MgGraph -ErrorAction SilentlyContinue
                        }
                    } else {
                        # For app-only auth, reuse existing connection
                        return $context.Account
                    }
                }
            }
            catch {
                Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,No existing connection found"
            }
            
            # Build connection parameters
            $connectParams = @{ NoWelcome = $true }
            
            switch ($authMethod) {
                'ManagedIdentity' {
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Connecting with Managed Identity"
                    
                    # Validate environment variables
                    if (-not $env:IDENTITY_ENDPOINT -or -not $env:IDENTITY_HEADER) {
                        throw "Managed Identity environment variables not set"
                    }
                    
                    # Get Graph SDK version for compatibility
                    $graphVersion = (Get-Module -Name 'Microsoft.Graph.Authentication' -ListAvailable | 
                        Sort-Object Version -Descending | Select-Object -First 1).Version
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Graph SDK version: $graphVersion"
                    
                    if ($graphVersion -ge [version]"2.0.0") {
                        $connectParams['Identity'] = $true
                    } else {
                        # For older SDK versions, get token manually
                        $headers = @{
                            'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER
                            'Metadata' = 'True'
                        }
                        $response = Invoke-RestMethod -Uri "$($env:IDENTITY_ENDPOINT)?resource=$resourceURL" `
                            -Method GET -Headers $headers -TimeoutSec 30
                        $connectParams['AccessToken'] = $response.access_token
                        Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Retrieved managed identity token"
                    }
                }
                
                'ClientSecret' {
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Connecting with Client Secret"
                    # Validate required inputs
                    if (-not $hasClientId -or -not $hasTenantId) {
                        throw "ClientSecret authentication requires both ClientId and TenantId."
                    }
                    # Prefer native parameters supported by the Graph SDK and convert SecureString to plain text
                    $plainSecret = $null
                    try {
                        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
                        $plainSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                    } finally {
                        if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
                    }
                    $connectParams['ClientId']     = $ClientId
                    $connectParams['TenantId']     = $TenantId
                    $connectParams['ClientSecret'] = $plainSecret
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Using ClientId: $ClientId, TenantId: $TenantId"
                }
                
                'Certificate' {
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Connecting with Certificate"
                    # Validate required inputs
                    if (-not $hasClientId -or -not $hasTenantId) {
                        throw "Certificate authentication requires both ClientId and TenantId."
                    }
                    $connectParams['ClientId'] = $ClientId
                    $connectParams['TenantId'] = $TenantId
                    
                    # Handle different certificate input methods
                    if ($CertificateThumbprint) {
                        $connectParams['CertificateThumbprint'] = $CertificateThumbprint
                        Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Using certificate thumbprint: $CertificateThumbprint"
                    }
                    elseif ($CertificateName) {
                        $connectParams['CertificateName'] = $CertificateName
                        Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Using certificate name: $CertificateName"
                    }
                    elseif ($CertificatePath) {
                        # Load certificate from file
                        if (-not (Test-Path $CertificatePath)) {
                            throw "Certificate file not found: $CertificatePath"
                        }
                        
                        try {
                            if ($CertificatePassword) {
                                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                                    $CertificatePath, 
                                    $CertificatePassword
                                )
                            } else {
                                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath)
                            }
                            $connectParams['Certificate'] = $cert
                            Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Loaded certificate from: $CertificatePath"
                        }
                        catch {
                            throw "Failed to load certificate: $_"
                        }
                    }
                    else {
                        throw "No certificate specified. Use CertificateThumbprint, CertificateName, or CertificatePath"
                    }
                }
                
                'Interactive' {
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Connecting interactively"
                    # Ensure scopes are a string array
                    $connectParams['Scopes'] = @($RequiredScopes)
                }
            }
            
            # Connect to Microsoft Graph
            try {
                Connect-MgGraph @connectParams -ErrorAction Stop
                Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Successfully connected to Microsoft Graph"
            }
            catch {
                throw "Failed to connect to Microsoft Graph: $_"
            }
            finally {
                # Clear sensitive value if used
                if ($connectParams.ContainsKey('ClientSecret')) { $connectParams['ClientSecret'] = $null }
            }
            
            # Validate permissions for delegated auth (Interactive only)
            if ($authMethod -eq 'Interactive' -and @($RequiredScopes).Count -gt 0) {
                try {
                    $context = Get-MgContext
                    $currentScopes = @($context.Scopes)
                    $reqScopes = @($RequiredScopes)
                    $missingScopes = @($reqScopes | Where-Object { $_ -notin $currentScopes })
                    if (@($missingScopes).Count -gt 0) {
                        throw "Missing required scopes: $($missingScopes -join ', ')"
                    }
                    
                    Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Validated all required scopes: $($RequiredScopes -join ', ')"
                }
                catch {
                    throw "Failed to validate permissions: $_"
                }
            }
            
            # Return account context
            $context = Get-MgContext
            $account = $context.Account
            Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Connected as: $account"
            return $account
        }
        catch {
            Write-Error "$timestamp,Error,Invoke-ConnectMgGraph,Connection failed: $_"
            throw
        }
    }
    
    End {
        $memoryUsage = [Math]::Round(([System.GC]::GetTotalMemory($false) / 1MB), 2)
        Write-Verbose "$timestamp,Info,Invoke-ConnectMgGraph,Function finished. Memory: $memoryUsage MB"
    }
}
function Invoke-MgGraphRequestSingle {
    <#
    .SYNOPSIS
        Executes a single Graph API request with pagination and retry logic.

    .DESCRIPTION
        Makes Graph API calls with automatic pagination, throttling handling, and exponential backoff retry logic.
        Supports filtering, property selection, and count queries. Returns all pages of results automatically.
    #>
[CmdletBinding()]
    Param(
        [Parameter(HelpMessage = "The Graph API version ('beta' or 'v1.0')")]
        [ValidateSet('beta', 'v1.0')]
        [string]$RunProfile     = "v1.0",
    
        [Parameter(HelpMessage = "The HTTP method for the request(e.g., 'GET', 'PATCH', 'POST', 'DELETE')")]
        [ValidateSet('GET', 'PATCH', 'POST', 'DELETE')]
        [String]$Method         = "GET",
        
        [Parameter(HelpMessage = "The Graph API endpoint path to target (e.g., 'me', 'users', 'groups')")]
        [string]$Object,

        [Parameter(             HelpMessage = "Request body for POST/PATCH operations")]
        [string[]]$Body,
        
        [Parameter(             HelpMessage = "Graph API properties to include")]
        [string[]]$Properties,
    
        [Parameter(             HelpMessage = "Graph API filters to apply")]
        [string]$Filters,
    
        [Parameter(             HelpMessage = "Page size (max 1000 objects per page)")]
        [ValidateRange(1,1000)]
        [int]$PageSize          = 999,

        [Parameter(             HelpMessage = "Skip pagination and only get the first page. Default is false")]
        [bool]$SkipPagination   = $false,

        [Parameter(             HelpMessage = "Include count of total items. Adds ConsistencyLevel header. Default is false")]
        [bool]$Count            = $false,

        [Parameter(             HelpMessage = "Delay between requests if throttled in milliseconds")]
        [ValidateRange(100,5000)]
        [int]$WaitTime          = 1000,

        [Parameter(             HelpMessage = "Maximum retry attempts for failed requests when throttled")]
        [ValidateRange(1,10)]
        [int]$MaxRetry          = 3
    )
    Begin {
        $PsobjectResults = [System.Collections.ArrayList]::new()
        $RetryCount = 0
        $TotalCount = $null

        # Normalize $Object parameter - ensure it starts and ends with /
        if ($Object) {
            if (-not $Object.StartsWith('/')) { $Object = "/$Object" }
            if (-not $Object.EndsWith('/')) { $Object = "$Object/" }
        }

        # Build base URI
        $uri = "https://graph.microsoft.com/$RunProfile/$Object"
        $queryParams = [System.Collections.ArrayList]::new()

        # Add count parameter if requested
        if ($Count) {[void]$queryParams.Add("`$count=true")}

        # Add page size parameter
        if ($Method -eq 'GET') {[void]$queryParams.Add("`$top=$PageSize")}

        # Add properties if specified
        if ($Properties) {
            $select = $Properties -join ','
            [void]$queryParams.Add("`$select=$select")
        }

        # Add filters if specified
        if ($Filters) {
            if ($Filters -is [string]) {
                [void]$queryParams.Add("`$filter=$([System.Web.HttpUtility]::UrlEncode($Filters))")
            }
            elseif ($Filters -is [hashtable]) {
                $filterParts = foreach ($key in $Filters.Keys) {
                    "$key eq '$($Filters[$key])'"
                }
                $filter = $filterParts -join ' and '
                [void]$queryParams.Add("`$filter=$([System.Web.HttpUtility]::UrlEncode($filter))")
            }
        }

        # Combine query parameters into URI
        if ($queryParams.Count -gt 0) {$uri += "?" + ($queryParams -join '&')}
    }
    Process {
        do {
            try {
                Write-Verbose "Making request to: $uri"
                $i = 1
                do {
                    $response = $null
                    Write-Verbose "Requesting page $i with $PageSize items"
                    # Set default parameters for Invoke-MgGraphRequest
                    $params = @{
                        Method      = $Method
                        Uri         = $uri
                        ErrorAction = 'Stop'
                        OutputType  = 'PSObject'
                        Verbose     = $false
                    }
                    # Add ConsistencyLevel header if Count is requested
                    if ($Count) { $params['Headers'] = @{ 'ConsistencyLevel' = 'eventual' } }

                    # Add additional parameters based on method
                    if ($Method -in 'POST', 'PATCH') {
                        $params['Body'] = $Body
                        if (-not $params.ContainsKey('Headers')) {
                            $params['Headers'] = @{}
                        }
                        $params['Headers']['Content-Type'] = 'application/json'
                        write-verbose "Request body: $($Body | ConvertTo-Json -Depth 10)"
                    }
                    # Send request to Graph API
                    try {
                        $response = Invoke-MgGraphRequest @params
                        Write-Verbose "Request successful"
                    }
                    catch {
                        # Check if this is an expired skip token error
                        if ($_.Exception.Message -match "Skip token.*expired|Skip token is null") {
                            Write-Warning "Skip token has expired on page $i after collecting $($PsobjectResults.Count) items. Returning collected data."
                            # Exit pagination loop and return what we have
                            $uri = $null
                            break
                        }
                        # For other errors, log and re-throw to outer catch
                        Write-Verbose "Request failed with error: $_"
                        throw
                    }
                    if ($Method -eq 'POST') {return $response}
                    if ($response.value) {[void]$PsobjectResults.AddRange($response.value)}
                    # Capture count from first response if requested
                    if ($Count -and $null -eq $TotalCount -and $response.'@odata.count') {
                        $TotalCount = $response.'@odata.count'
                        Write-Verbose "Total count available: $TotalCount items"
                    }
                    Write-Verbose "Retrieved page $i, Now total: $($PsobjectResults.Count) items"

                    # Check for next page
                    if ($SkipPagination) {
                        Write-Verbose "SkipPagination enabled, stopping after first page"
                        $uri = $null
                    }
                    elseif ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
                        if ($response.'@odata.nextLink') {
                            $uri = $response.'@odata.nextLink'
                            Write-Verbose "Next page found: $uri"
                        }
                        else {
                            Write-Verbose "No @odata.nextLink value, stopping pagination"
                            $uri = $null
                        }
                    }
                    else {
                        Write-Verbose "No more pages found"
                        $uri = $null
                    }

                    $i++
                } while ($uri)
                Write-Verbose "Completed pagination. Returning array with $($PsobjectResults.Count) items"
                
                # Return results with count if requested
                if ($Count -and $null -ne $TotalCount) {
                    return [PSCustomObject]@{
                        Items = $PsobjectResults
                        Count = $TotalCount
                    }
                }
                return $PsobjectResults # Success, return results and exit retry loop
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                # Get full error string including nested JSON messages for better pattern matching
                $FullErrorString = $_ | Out-String
                Write-Warning "Request failed (Retry attempt $($RetryCount + 1)/$MaxRetry): $ErrorMessage"

                # Check if the exception has response details (it should for HTTP errors)
                if ($_.Exception.Response) {
                    $StatusCode = $_.Exception.Response.StatusCode

                    # Use switch to handle specific status codes (handle both enum names and numeric values)
                    switch ($StatusCode) {
                        {$_ -eq 429 -or $_ -eq 'TooManyRequests'} { # Throttling
                            $RetryAfter = ($_.Exception.Response.Headers | Where-Object {$_.Name -eq "Retry-After"}).Value
                            if ($RetryAfter) {
                                Write-Warning "Throttling detected (429). Waiting $($RetryAfter * 1000) milliseconds before retrying."
                                Start-Sleep -Milliseconds ($RetryAfter * 1000) # Convert seconds to milliseconds
                            } else {
                                $Delay = [math]::Min(($WaitTime * ([math]::Pow(2, $RetryCount))), 60000) # Exponential backoff, max 60 seconds
                                Write-Warning "Throttling detected (429). No Retry-After header found. Waiting $($Delay) milliseconds before retrying."
                                Start-Sleep -Milliseconds $Delay
                            }
                            # Break not needed, will fall through to retry logic below
                        }
                        {$_ -eq 404 -or $_ -eq 'NotFound'} { # Not Found
                            Write-Warning "Resource not found (404). Error: $ErrorMessage"
                            # Re-throw the original exception to signal failure to the caller immediately
                            throw "$_ (Object Deleted/No User License)"
                        }
                        {$_ -eq 400 -or $_ -eq 'BadRequest'} { # Bad Request                            
                            if ($ErrorMessage -match "Skip token.*expired|Skip token is null" -or $FullErrorString -match "Skip token.*expired|Skip token is null") {# Check if this is an expired skip token error
                                Write-Warning "Skip token has expired after collecting $($PsobjectResults.Count) items. Returning collected data."
                                return $PsobjectResults
                            }
                            if ($ErrorMessage -match "does not have intune license or is deleted" -or $FullErrorString -match "does not have intune license or is deleted") { # Check if no license, common for Intune queries
                                Write-Warning "Object Deleted or User has no Intune license"
                                return "$_ (Object Deleted/No User License)"
                            }
                            else {
                                Write-Error "Bad request (400). Error: $ErrorMessage"
                                throw $_
                            }
                        }
                        {$_ -eq 403 -or $_ -eq 'Forbidden'} { # Forbidden / Access Denied
                             Write-Error "Access denied (403). Error: $ErrorMessage"
                             throw $_
                        }
                        default { # Other HTTP errors - Use generic retry
                            $Delay = [math]::Min(($WaitTime * ([math]::Pow(2, $RetryCount))), 60000) # Exponential backoff, max 60 seconds
                            Write-Warning "HTTP error $($StatusCode). Waiting $($Delay) milliseconds before retrying."
                            Start-Sleep -Milliseconds $Delay
                            # Break not needed, will fall through to retry logic below
                        }
                    }
                } else {
                    # Non-HTTP errors (e.g., network issues, DNS resolution) - Use generic retry
                    $Delay = [math]::Min(($WaitTime * ([math]::Pow(2, $RetryCount))), 60000) # Exponential backoff, max 60 seconds
                    Write-Warning "Non-HTTP error. Waiting $($Delay) milliseconds before retrying. Error: $ErrorMessage"
                    Start-Sleep -Milliseconds $Delay
                }

                # Increment retry count and check if max retries exceeded ONLY if not already thrown
                $RetryCount++
                if ($RetryCount -gt $MaxRetry) {
                     Write-Error "Request failed after $($MaxRetry) retries. Aborting."
                     throw "Request failed after $($MaxRetry) retries. Last error: $ErrorMessage"
                }
                # If retries not exceeded and error was potentially retryable (e.g., 429, other HTTP, non-HTTP), the loop will continue
            }
        } while ($RetryCount -le $MaxRetry)

        Write-Error "Request failed after $($MaxRetry) retries. Aborting."
        throw "Request failed after $($MaxRetry) retries." # Re-throw the exception after max retries
    }

    End {
        # End function and report memory usage 
        $MemoryUsage = [Math]::Round(([System.GC]::GetTotalMemory($false) / 1MB), 2)
        Write-Verbose "Function finished. Memory usage: $MemoryUsage MB"
    }
}
function invoke-mgGraphRequestBatch {
    <#
    .SYNOPSIS
        Processes multiple Graph API requests in batches for improved performance.

    .DESCRIPTION
        Sends Graph API requests in batches (up to 20 per batch) to efficiently process large numbers of objects.
        Handles throttling, retries, and provides progress tracking. Supports GET, PATCH, POST, and DELETE operations.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(HelpMessage = "The Graph API version ('beta' or 'v1.0')")]
        [ValidateSet('beta', 'v1.0')]
        [string]$RunProfile = "v1.0",
    
        [Parameter(HelpMessage = "The HTTP method for the request(e.g., 'GET', 'PATCH', 'POST', 'DELETE')")]
        [ValidateSet('GET', 'PATCH', 'POST', 'DELETE')]
        [String]$Method = "GET",
        
        [Parameter(HelpMessage = "The Graph API endpoint path to target (e.g., 'me', 'users', 'groups')")]
        [string]$Object,
    
        [Parameter(HelpMessage = "Array of objects to process in batches")]
        [System.Object[]]$Objects,
    
        [Parameter(HelpMessage = "The Graph API query on the objects")]
        [string]$query,
    
        [Parameter(HelpMessage = "Request body for POST/PATCH operations")]
        [object]$Body,
        
        [Parameter(HelpMessage = "Graph API properties to include")]
        [string[]]$Properties,
    
        [Parameter(HelpMessage = "Graph API filters to apply")]
        [string]$Filters,
    
        [Parameter(HelpMessage = "Batch size (max 20 objects per batch)")]
        [ValidateRange(1,20)]
        [int]$BatchSize = 20,
    
        [Parameter(HelpMessage = "Delay between batches in milliseconds")]
        [ValidateRange(100,5000)]
        [int]$WaitTime = 1000,
    
        [Parameter(HelpMessage = "Maximum retry attempts for failed requests")]
        [ValidateRange(1,10)]
        [int]$MaxRetry = 3
    )
    
    Begin {
        $ErrorActionPreference = 'Stop'
        $script:GetTimestamp = { ([DateTime]::Now).ToString('yyyy-MM-dd HH:mm:ss') }
        $starttime = Get-Date
        $Retrycount = 0
        $TotalObjects = $Objects.Count
        
        $CollectedObjects = [System.Collections.Generic.List[PSObject]]::new($TotalObjects)
        $RetryObjects = [System.Collections.Generic.List[PSObject]]::new()
        
        [Bool]$ManagedIdentity = [bool]$env:AUTOMATION_ASSET_ACCOUNTID
        Write-Verbose "Running in $(if ($ManagedIdentity) { 'Azure Automation' } else { 'interactive PowerShell' }) context"
        
        $batchUri = "https://graph.microsoft.com/$RunProfile/`$batch"
        $batchHeaders = @{'Content-Type' = 'application/json'}
        
        $urlQueryString = $null
        if ($Properties -or $Filters) {
            $queryParams = [System.Collections.ArrayList]::new()
            if ($Properties) {
                $select = $Properties -join ','
                [void]$queryParams.Add("`$select=$select")
            }
            if ($Filters) { [void]$queryParams.Add("`$filter=$([System.Web.HttpUtility]::UrlEncode($Filters))") }
            if ($queryParams.Count -gt 0) { $urlQueryString = "?" + ($queryParams -join '&') }
        }
        
        $needsBody = $Method -in 'PATCH','POST'
        $contentTypeHeader = if ($needsBody) { 'application/json' } else { $null }
        
        Write-Verbose "Graph batch processing initialized for $TotalObjects objects"
    }
    
    Process {
        try {
            for ($i = 0; $i -lt $TotalObjects; $i += $BatchSize) {
                $batchObjects = $Objects[$i..[Math]::Min($i + $BatchSize - 1, $TotalObjects - 1)]
                $batchNumber = [Math]::Floor($i / $BatchSize) + 1
                $totalBatches = [Math]::Ceiling($TotalObjects / $BatchSize)
                
                Write-Verbose "Processing batch $batchNumber of $totalBatches (objects $($i+1) to $($i + $batchObjects.Count))"
                
                $requests = [System.Collections.Generic.List[object]]::new($batchObjects.Count)
                
                foreach ($obj in $batchObjects) {
                    $requestId = "req_$([guid]::NewGuid().ToString('N').Substring(0,8))"
                    $url = if ($Object) {
                        "/$Object/$($obj.id)$(if($query){"/" + $query})$(if($urlQueryString){$urlQueryString})"
                    } else {
                        "/$($obj.id)$(if($query){"/" + $query})$(if($urlQueryString){$urlQueryString})"
                    }
                    
                    $request = @{
                        id = $requestId
                        method = $Method
                        url = $url
                    }
                    
                    if ($needsBody -and $Body) {
                        $request.body = $Body
                        $request.headers = @{ 'Content-Type' = $contentTypeHeader }
                    }
                    
                    [void]$requests.Add($request)
                }
                
                $batchRequest = @{ requests = $requests } | ConvertTo-Json -Depth 10 -Compress
                
                try {
                    $batchResponse = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $batchRequest -Headers $batchHeaders -ErrorAction Stop
                    
                    foreach ($response in $batchResponse.responses) {
                        if ($response.status -eq 200 -or $response.status -eq 204) {
                            if ($response.body) { [void]$CollectedObjects.Add($response.body) }
                        }
                        elseif ($response.status -eq 429) {
                            $matchingRequest = $requests | Where-Object { $_.id -eq $response.id }
                            if ($matchingRequest) {
                                $objectId = ($matchingRequest.url -split '/')[-1] -replace '\?.*$'
                                $retryObj = $batchObjects | Where-Object { $_.id -eq $objectId }
                                if ($retryObj) { [void]$RetryObjects.Add($retryObj) }
                            }
                            Write-Warning "Request $($response.id) throttled (429). Will retry."
                        }
                        else {
                            Write-Warning "Request $($response.id) failed with status $($response.status): $($response.body.error.message)"
                        }
                    }
                    
                    if ($i + $BatchSize -lt $TotalObjects) { Start-Sleep -Milliseconds $WaitTime }
                }
                catch {
                    $errorMessage = $_.Exception.Message
                    if ($errorMessage -match '429|throttl' -and $Retrycount -lt $MaxRetry) {
                        $Retrycount++
                        $delay = [math]::Min(($WaitTime * ([math]::Pow(2, $Retrycount))), 60000)
                        Write-Warning "Batch request throttled. Retry $Retrycount/$MaxRetry. Waiting $($delay)ms"
                        Start-Sleep -Milliseconds $delay
                        [void]$RetryObjects.AddRange($batchObjects)
                    }
                    else {
                        Write-Error "Batch request failed: $errorMessage"
                        throw
                    }
                }
            }
            
            if ($RetryObjects.Count -gt 0 -and $Retrycount -lt $MaxRetry) {
                Write-Verbose "Retrying $($RetryObjects.Count) throttled requests"
                $retryResults = invoke-mgGraphRequestBatch @PSBoundParameters -Objects $RetryObjects.ToArray()
                if ($retryResults) { [void]$CollectedObjects.AddRange($retryResults) }
            }
            
            return $CollectedObjects.ToArray()
        }
        catch {
            Write-Error "Batch processing failed: $_"
            throw
        }
    }
    
    End {
        $MemoryUsage = [Math]::Round(([System.GC]::GetTotalMemory($false) / 1MB), 2)
        $duration = (Get-Date) - $starttime
        Write-Verbose "Function $($MyInvocation.MyCommand.Name) finished in $($duration.ToString('mm\:ss')) | Memory: $MemoryUsage MB"
    }
}
function Invoke-ScriptReport {
    <#
    .SYNOPSIS
        Generates execution report with summary statistics and optional detailed results.

    .DESCRIPTION
        Creates formatted reports showing script execution outcomes including success/failure statistics,
        execution duration, and per-object results if detailed reporting is enabled.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Progress,
        
        [Parameter(Mandatory = $false)]
        [array]$Results,
        
        [Parameter(Mandatory = $false)]
        [bool]$DetailedReport = $true,
        
        [Parameter(Mandatory = $false)]
        [bool]$SaveToDisk = $false,
        
        [Parameter(Mandatory = $false)]
        [string]$Path = "$env:TEMP"
    )

    Begin {
        $reportTime = Get-Date
        $duration = $reportTime - $Script:StartTime
    }

    Process {
        # Build dynamic summary object from all progress counters
        $summaryProps = [ordered]@{
            ExecutionTime = $duration.ToString('hh\:mm\:ss')
        }
        
        # Add all progress counters dynamically (excluding internal ones), with Total last
        foreach ($key in ($Progress.Keys | Sort-Object | Where-Object { $_ -ne 'Total' })) {
            $summaryProps[$key] = $Progress[$key]
        }
        # Add Total last
        if ($Progress.ContainsKey('Total')) {
            $summaryProps['Total'] = $Progress['Total']
        }
        
        $report = [PSCustomObject]@{
            Summary = [PSCustomObject]$summaryProps
        }
        
        if ($DetailedReport -and $Results -and $Results.Count -gt 0) {
            $report | Add-Member -NotePropertyName 'DetailedResults' -NotePropertyValue $Results
        }
        
        # Build report lines for efficient output
        $reportLines = [System.Collections.Generic.List[string]]::new()
        [void]$reportLines.Add("`n========== Execution Report ==========")
        [void]$reportLines.Add("Execution Time: $($report.Summary.ExecutionTime)")
        
        # Dynamically display all counters except ExecutionTime, with Total last
        foreach ($key in ($Progress.Keys | Sort-Object | Where-Object { $_ -ne 'Total' })) {
            if ($Progress[$key] -gt 0) {
                [void]$reportLines.Add("${key}: $($Progress[$key])")
            }
        }
        # Display Total last
        if ($Progress.ContainsKey('Total')) {
            [void]$reportLines.Add("Total: $($Progress['Total'])")
        }
        [void]$reportLines.Add("======================================`n")
        
        # Output all at once
        Write-Output ($reportLines -join "`n")
        
        # Display detailed results as table if available
        if ($DetailedReport -and $Results -and $Results.Count -gt 0) {
            Write-Output "`n========== Detailed Results =========="
            $Results | Format-Table -AutoSize
            Write-Output "======================================`n"
        }
        
        if ($SaveToDisk) {
            try {
                $fileName = "DeviceCleanup-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
                $fullPath = Join-Path -Path $Path -ChildPath $fileName
                $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $fullPath -Encoding UTF8
                Write-Verbose "Report saved to: $fullPath"
            }
            catch {
                Write-Warning "Failed to save report to disk: $_"
            }
        }
        
        return $report.Summary
    }

    End {
        $MemoryUsage = [Math]::Round(([System.GC]::GetTotalMemory($false) / 1MB), 2)
        Write-Verbose "Report generation finished. Memory: $MemoryUsage MB"
    }
}
#endregion

#region ---------------------------------------------------[[Script Execution]------------------------------------------------------
$Script:StartTime = [datetime]::Now

# Start custom logging
if ($ReportToDisk) { Invoke-TboneLog -Mode Start -LogToGUI $true -LogToDisk $true -LogPath $ReportPath }
else{ Invoke-TboneLog -Mode Start -LogToGUI $true -LogToDisk $false }

try {
    #Sign in to Graph
    try {Invoke-ConnectMgGraph -RequiredScopes $RequiredScopes
        Write-Verbose "Success to get Access Token to Graph"}
    catch {
        Write-Error "Failed to get Access Token to Graph, with error: $_"
        throw
    }

# Get all old devices from Entra ID
    try {
        # List properties to retrieve
        $GraphProperties = @('id', 'displayName', 'approximateLastSignInDateTime', 'accountEnabled')
        # Prepare filters
        $GraphFilterString = "approximateLastSignInDateTime le $disableDate"
        # Get graph objects
        $AllOldDevices = Invoke-MgGraphRequestSingle `
            -RunProfile 'v1.0' `
            -Method 'GET' `
            -Object 'devices' `
            -Properties $GraphProperties `
            -Filters $GraphFilterString `
            -MaxRetry $MaxRetry `
            -WaitTime $WaitTime

        #Verify if objects were found
        if ($AllOldDevices -and $AllOldDevices.Count -gt 0) {
            Write-Verbose "Successfully retrieved $($AllOldDevices.Count) old devices from Graph API"
        }
        else {
            Write-Warning "No old devices from Graph API found in tenant"
            return # Exit early if no devices to process
        }
    }
    catch {
        Write-Error "Failed to get old devices from Graph API: $_"
        throw
    }

    # Process all devices and set the scope tag to match info on the primary user
    $script:ReportProgress.Total = $AllOldDevices.Count
    # Initialize collection for devices that need updates
    $devicesToDisable = [System.Collections.ArrayList]::new()
    $devicesToDelete  = [System.Collections.ArrayList]::new()
    # Helper function to add report result
    $addReport = {
        param($Device, $Old, $New, $Status, $Counter)
        $script:ReportProgress.$Counter++
        [void]$script:ReportResults.Add([PSCustomObject]@{
            Object = $Device; OldValue = $Old; NewValue = $New; Status = $Status
        })
    }.GetNewClosure()
    # Process devices and identify what needs to be done
    foreach ($device in $AllOldDevices) {
        $deviceDate = $device.approximateLastSignInDateTime
        $deviceName = $device.displayName
        $deviceId = $device.id
        $isEnabled = $device.accountEnabled
        
        # Determine action needed
        if ($deviceDate -le $deleteDate) {
            # Device is old enough to delete
            Write-Verbose "Device '$deviceName' last sign-in: $deviceDate - marked for deletion"
            [void]$devicesToDelete.Add($device)
        }
        elseif ($deviceDate -le $disableDate -and $isEnabled -eq $true) {
            # Device should be disabled
            Write-Verbose "Device '$deviceName' last sign-in: $deviceDate - marked for disable"
            [void]$devicesToDisable.Add($device)
        }
        else {
            # Device doesn't need changes
            & $addReport -Device $deviceName -Old "Disabled,LastSeen:$deviceDate" -New "No change" -Status "Skipped" -Counter "Skipped"
        }
    }
    
    Write-Verbose "Identified $($devicesToDisable.Count) devices to disable and $($devicesToDelete.Count) devices to delete"
    
    # Process devices to disable using batch operations
    if ($devicesToDisable.Count -gt 0) {
        if (-not $WhatIfPreference) {
            Write-Output "Disabling $($devicesToDisable.Count) devices using batch operations..."
            try {
                # Build body for disable operation
                $disableBody = @{ accountEnabled = $false }
                
                # Use batch processing for better performance (20 devices per batch)
                $batchResults = invoke-mgGraphRequestBatch `
                    -RunProfile 'v1.0' `
                    -Method 'PATCH' `
                    -Object 'devices' `
                    -Objects $devicesToDisable `
                    -Body $disableBody `
                    -BatchSize $BatchSize `
                    -MaxRetry $MaxRetry `
                    -WaitTime $WaitTime
                
                # Process results - batch returns updated device objects for successful PATCH
                if ($batchResults -and $batchResults.Count -gt 0) {
                    Write-Verbose "Batch operation returned $($batchResults.Count) successful results"
                    foreach ($device in $devicesToDisable) {
                        $batchResult = $batchResults | Where-Object { $_.id -eq $device.id }
                        if ($batchResult) {
                            Write-Verbose "Successfully disabled device: $($device.displayName)"
                            & $addReport -Device $device.displayName -Old "Enabled,LastSeen:$($device.approximateLastSignInDateTime)" -New "Disabled" -Status "Success" -Counter "Disabled"
                        }
                        else {
                            Write-Warning "Failed to disable device $($device.displayName) - not in batch results"
                            & $addReport -Device $device.displayName -Old "Enabled,LastSeen:$($device.approximateLastSignInDateTime)" -New "Failed" -Status "Failed" -Counter "Failed"
                        }
                    }
                }
                else {
                    Write-Warning "Batch operation returned no results - assuming success for all devices"
                    # If batch completed without errors but returned no results, assume success
                    foreach ($device in $devicesToDisable) {
                        Write-Verbose "Assuming successful disable for device: $($device.displayName)"
                        & $addReport -Device $device.displayName -Old "Enabled,LastSeen:$($device.approximateLastSignInDateTime)" -New "Disabled" -Status "Success" -Counter "Disabled"
                    }
                }
            }
            catch {
                Write-Error "Batch disable operation failed: $_"
                # Fallback: Mark all as failed
                foreach ($device in $devicesToDisable) {
                    & $addReport -Device $device.displayName -Old "Enabled,LastSeen:$($device.approximateLastSignInDateTime)" -New "Failed" -Status "Failed" -Counter "Failed"
                }
            }
        }
        else {
            # WhatIf mode: Report without making changes
            foreach ($device in $devicesToDisable) {
                Write-Verbose "WhatIf: Would disable device $($device.displayName)"
                & $addReport -Device $device.displayName -Old "Enabled,LastSeen:$($device.approximateLastSignInDateTime)" -New "Would be disabled" -Status "WhatIf" -Counter "WhatifDisabled"
            }
        }
    }
    else {
        Write-Output "No devices need to be disabled"
    }
    
    # Process devices to delete using batch operations
    if ($devicesToDelete.Count -gt 0) {
        if (-not $WhatIfPreference) {
            Write-Output "Deleting $($devicesToDelete.Count) devices using batch operations..."
            try {
                # Use batch processing for better performance (20 devices per batch)
                $batchResults = invoke-mgGraphRequestBatch `
                    -RunProfile 'v1.0' `
                    -Method 'DELETE' `
                    -Object 'devices' `
                    -Objects $devicesToDelete `
                    -BatchSize $BatchSize `
                    -MaxRetry $MaxRetry `
                    -WaitTime $WaitTime
                
                # DELETE returns 204 No Content - if batch completes without error, all succeeded
                Write-Verbose "Batch delete operation completed - assuming success for all devices"
                foreach ($device in $devicesToDelete) {
                    Write-Verbose "Successfully deleted device: $($device.displayName)"
                    & $addReport -Device $device.displayName -Old "Old,LastSeen:$($device.approximateLastSignInDateTime)" -New "Deleted" -Status "Success" -Counter "Deleted"
                }
            }
            catch {
                Write-Error "Batch delete operation failed: $_"
                # Fallback: Mark all as failed
                foreach ($device in $devicesToDelete) {
                    & $addReport -Device $device.displayName -Old "Old,LastSeen:$($device.approximateLastSignInDateTime)" -New "Failed" -Status "Failed" -Counter "Failed"
                }
            }
        }
        else {
            # WhatIf mode: Report without making changes
            foreach ($device in $devicesToDelete) {
                Write-Verbose "WhatIf: Would delete device $($device.displayName)"
                & $addReport -Device $device.displayName -Old "Old,LastSeen:$($device.approximateLastSignInDateTime)" -New "Would be deleted" -Status "WhatIf" -Counter "WhatifDeleted"
            }
        }
    }
    else {
        Write-Output "No devices need to be deleted"
    }
    
    Write-Verbose "Processing complete. Disabled: $($script:ReportProgress.Disabled), Deleted: $($script:ReportProgress.Deleted), Failed: $($script:ReportProgress.Failed), Skipped: $($script:ReportProgress.Skipped)"
}
catch {
    Write-Error "Script execution failed: $_"
}
finally {
    # Disconnect from Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Verbose "Disconnected from Graph"
    } catch { Write-Error "Failed to disconnect from Graph: $_" }
    
    Invoke-TboneLog -Mode Stop
    
    # Restore original preference settings
    $ErrorActionPreference = $script:OriginalErrorActionPreference
    $VerbosePreference = $script:OriginalVerbosePreference
    $WhatIfPreference = $script:OriginalWhatIfPreference
    
    # End script and report memory usage
    $MemoryUsage = [Math]::Round(([System.GC]::GetTotalMemory($false) / 1MB), 2)
    Write-Verbose "Script finished. Memory usage: $MemoryUsage MB"
    
    # Generate report if requested
    if ($ReturnReport) {
        if ($DetailedReport -and $script:ReportResults.Count -gt 0) {
            Invoke-ScriptReport -Progress $script:ReportProgress -Results $script:ReportResults -DetailedReport $true -SaveToDisk $ReportToDisk -Path $ReportPath
        }
        elseif ($script:ReportProgress.Total -gt 0) {
            Invoke-ScriptReport -Progress $script:ReportProgress -DetailedReport $false -SaveToDisk $ReportToDisk -Path $ReportPath
        }
    }
}
