# Import the REST module so that the EXO* cmdlets are present before Connect-ExchangeOnline in the powershell instance.
$RestModule = "Microsoft.Exchange.Management.RestApiClient.dll"
$RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestModule)
Import-Module $RestModulePath 

$ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll"
$ExoPowershellModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule)
Import-Module $ExoPowershellModulePath

#Keep track of Execution status of last cmdlet
$global:EXO_LastExecutionStatus = $true;

############# Helper Functions Begin #############

    <#
    Get the ExchangeOnlineManagement module version.
    Same function is present in the autogen module. Both the codes should be kept in sync.
    #>
    function Get-ModuleVersion
    {
        try
        {
            # Return the already computed version info if available.
            if ($script:ModuleVersion -ne $null -and $script:ModuleVersion -ne '')
            {
                Write-Verbose "Returning precomputed version info: $script:ModuleVersion"
                return $script:ModuleVersion;
            }

            $exoModule = Get-Module ExchangeOnlineManagement
            
            # Check for ExchangeOnlineManagementBeta in case the psm1 is loaded directly
            if ($exoModule -eq $null)
            {
               $exoModule = (Get-Command -Name Connect-ExchangeOnline).Module
            }

            # Get the module version from the loaded module info.
            $script:ModuleVersion = $exoModule.Version.ToString()

            # Look for prerelease information from the corresponding module manifest.
            $exoModuleRoot = (Get-Item $exoModule.Path).Directory.Parent.FullName

            $exoModuleManifestPath = Join-Path -Path $exoModuleRoot -ChildPath ExchangeOnlineManagement.psd1
            $isExoModuleManifestPathValid = Test-Path -Path $exoModuleManifestPath
            if ($isExoModuleManifestPathValid -ne $true)
            {
                # Could be a local debug build import for testing. Skip extracting prerelease info for those.
                Write-Verbose "Module manifest path invalid, path: $exoModuleManifestPath, skipping extracting prerelease info"
                return $script:ModuleVersion
            }

            $exoModuleManifestContent = Get-Content -Path $exoModuleManifestPath
            $preReleaseInfo = $exoModuleManifestContent -match "Prerelease = '(.*)'"
            if ($preReleaseInfo -ne $null)
            {
                $script:ModuleVersion = "{0}-{1}" -f $exoModule.Version.ToString(),$preReleaseInfo[0].Split('=')[1].Trim().Trim("'")
            }

            Write-Verbose "Computed version info: $script:ModuleVersion"
            return $script:ModuleVersion
        }
        catch
        {
            return [string]::Empty
        }
    }

    <#
    .Synopsis Validates a given Uri
    #>
    function Test-Uri
    {
        [CmdletBinding()]
        [OutputType([bool])]
        Param
        (
            # Uri to be validated
            [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
            [string]
            $UriString
        )

        [Uri]$uri = $UriString -as [Uri]

        $uri.AbsoluteUri -ne $null -and $uri.Scheme -eq 'https'
    }

    <#
    .Synopsis Is Cloud Shell Environment
    #>
    function global:IsCloudShellEnvironment()
    {
        return [Microsoft.Exchange.Management.AdminApiProvider.Utility]::IsCloudShellEnvironment();
    }

    <#
    .Synopsis Override Get-PSImplicitRemotingSession function for reconnection
    #>
    function global:UpdateImplicitRemotingHandler()
    {
        # Remote Powershell Sessions created by the ExchangeOnlineManagement module are given a name that starts with "ExchangeOnlineInternalSession".
        # Only modules from such sessions should be modified here, to prevent modfification of RPS tmp_* modules created by running the New-PSSession cmdlet directly, or when connecting to exchange on-prem tenants.
        $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}

        if ($existingPSSession.count -gt 0) 
        {
            foreach ($session in $existingPSSession)
            {
                $module = Get-Module $session.CurrentModuleName
                if ($module -eq $null)
                {
                    continue
                }

                [bool]$moduleProcessed = $false
                [string] $moduleUrl = $module.Description
                [int] $queryStringIndex = $moduleUrl.IndexOf("?")

                if ($queryStringIndex -gt 0)
                {
                    $moduleUrl = $moduleUrl.SubString(0,$queryStringIndex)
                }

                if ($moduleUrl.EndsWith("/PowerShell-LiveId", [StringComparison]::OrdinalIgnoreCase) -or $moduleUrl.EndsWith("/PowerShell", [StringComparison]::OrdinalIgnoreCase))
                {
                    & $module { ${function:Get-PSImplicitRemotingSession} = `
                    {
                        param(
                            [Parameter(Mandatory = $true, Position = 0)]
                            [string]
                            $commandName
                        )

                        $shouldRemoveCurrentSession = $false;
                        # Clear any left over PS tmp modules
                        if (($script:PSSession -ne $null) -and ($script:PSSession.PreviousModuleName -ne $null) -and ($script:PSSession.PreviousModuleName -ne $script:MyModule.Name))
                        {
                            $null = Remove-Module -Name $script:PSSession.PreviousModuleName -ErrorAction SilentlyContinue
                            $script:PSSession.PreviousModuleName = $null
                        }

                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                        {
                            Set-PSImplicitRemotingSession `
                                (& $script:GetPSSession `
                                    -InstanceId $script:PSSession.InstanceId.Guid `
                                    -ErrorAction SilentlyContinue )
                        }
                        if ($script:PSSession -ne $null)
                        {
                            if ($script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Disconnected')
                            {
                                # If we are handed a disconnected session, try re-connecting it before creating a new session.
                                Set-PSImplicitRemotingSession `
                                    (& $script:ConnectPSSession `
                                        -Session $script:PSSession `
                                        -ErrorAction SilentlyContinue)
                            }
                            else
                            {
                                # Import the module once more to ensure that Test-ActiveToken is present
                                Import-Module $global:_EXO_ModulePath -Cmdlet Test-ActiveToken;

                                # If there is no active token run the new session flow
                                $hasActiveToken = Test-ActiveToken -TokenExpiryTime $script:PSSession.TokenExpiryTime
                                $sessionIsOpened = $script:PSSession.Runspace.RunspaceStateInfo.State -eq 'Opened'
                                if (($hasActiveToken -eq $false) -or ($sessionIsOpened -ne $true))
                                {
                                    #If there is no active user token or opened session then ensure that we remove the old session
                                    $shouldRemoveCurrentSession = $true;
                                }
                            }
                        }
                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened') -or ($shouldRemoveCurrentSession -eq $true))
                        {
                            # Import the module once more to ensure that New-ExoPSSession is present
                            Import-Module $global:_EXO_ModulePath -Cmdlet New-ExoPSSession;

                            Write-PSImplicitRemotingMessage ('Creating a new Remote PowerShell session using Modern Authentication for implicit remoting of "{0}" command ...' -f $commandName)
                            $session = New-ExoPSSession -PreviousSession $script:PSSession

                            if ($session -ne $null)
                            {
                                if ($shouldRemoveCurrentSession -eq $true)
                                {
                                    Remove-PSSession $script:PSSession
                                }

                                # Import the latest session to ensure that the next cmdlet call would occur on the new PSSession instance.
                                if ([string]::IsNullOrEmpty($script:MyModule.ModulePrefix))
                                {
                                    $PSSessionModuleInfo = Import-PSSession $session -AllowClobber -DisableNameChecking -CommandName $script:MyModule.CommandName -FormatTypeName $script:MyModule.FormatTypeName
                                }
                                else
                                {
                                    $PSSessionModuleInfo = Import-PSSession $session -AllowClobber -DisableNameChecking -CommandName $script:MyModule.CommandName -FormatTypeName $script:MyModule.FormatTypeName -Prefix $script:MyModule.ModulePrefix
                                }

                                # Add the name of the module to clean up in case of removing the broken session
                                $session | Add-Member -NotePropertyName "CurrentModuleName" -NotePropertyValue $PSSessionModuleInfo.Name

                                $CurrentModule = Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $script:MyModule.ModulePrefix -PassThru
                                $CurrentModule | Add-Member -NotePropertyName "ModulePrefix" -NotePropertyValue $script:MyModule.ModulePrefix
                                $CurrentModule | Add-Member -NotePropertyName "CommandName" -NotePropertyValue $script:MyModule.CommandName
                                $CurrentModule | Add-Member -NotePropertyName "FormatTypeName" -NotePropertyValue $script:MyModule.FormatTypeName

                                $session | Add-Member -NotePropertyName "PreviousModuleName" -NotePropertyValue $script:MyModule.Name

                                UpdateImplicitRemotingHandler
                                $script:PSSession = $session
                            }
                        }
                        if (($script:PSSession -eq $null) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened'))
                        {
                            throw 'No session has been associated with this implicit remoting module'
                        }

                        return [Management.Automation.Runspaces.PSSession]$script:PSSession
                    }}
                }
            }
        }
    }

    <#
    .SYNOPSIS Extract organization name from UserPrincipalName
    #>
    function Get-OrgNameFromUPN
    {
        param([string] $UPN)
        $fields = $UPN -split '@'
        return $fields[-1]
    }

    <#
    .SYNOPSIS Get the command from the given module
    #>
    function global:Get-WrappedCommand
    {
        param(
        [string] $CommandName,
        [string] $ModuleName,
        [string] $CommandType)

        $cmd = (Get-Module $moduleName).ExportedFunctions[$CommandName]
        return $cmd
    }

    <#
    .Synopsis Writes a message to the console in Yellow.
    Call this method with caution since it uses Write-Host internally which cannot be suppressed by the user.
    #>
    function Write-Message
    {
        param([string] $message)
        Write-Host $message -ForegroundColor Yellow
    }

############# Helper Functions End #############

###### Begin Main ######

$EOPConnectionInProgress = $false
function Connect-ExchangeOnline 
{
    [CmdletBinding()]
    param(

        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = '',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = '',

        # Exchange Environment name
        [Microsoft.Exchange.Management.RestApiClient.ExchangeEnvironment] $ExchangeEnvironmentName = 'O365Default',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # Prefix 
        [string] $Prefix = '',

        # Show Banner of Exchange cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true,

        #Cmdlets to Import for rps cmdlets , by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*"),

        # Use Remote PowerShell Session based connection
        [switch] $UseRPSSession = $false,

        # Use to Skip Exchange Format file loading into current shell.
        [switch] $SkipLoadingFormatData = $false,

        # Use to skip downloading and loading the help files into the current connection.
        [switch] $SkipLoadingCmdletHelp = $true,

        # Use to enable downloading and loading the help files into the current connection.
        [switch] $LoadCmdletHelp = $false,

        # Externally provided access token
        [string] $AccessToken = '',

        # Client certificate to sign the temp module
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $SigningCertificate = $null,

        # switch to disable WAM
        [switch] $DisableWAM = $false
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and valiate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            # Switch to use Managed Identity flow. 
            $ManagedIdentity = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentity', [switch], $attributeCollection)
            $ManagedIdentity.Value = $false

            # ManagedIdentityAccountId to be used in case of User Assigned Managed Identity flow
            $ManagedIdentityAccountId = New-Object System.Management.Automation.RuntimeDefinedParameter('ManagedIdentityAccountId', [string], $attributeCollection)
            $ManagedIdentityAccountId.Value = ''

# EXO params start

            # Switch to track perfomance 
            $TrackPerformance = New-Object System.Management.Automation.RuntimeDefinedParameter('TrackPerformance', [bool], $attributeCollection)
            $TrackPerformance.Value = $false

            # Flag to enable or disable showing the number of objects written
            $ShowProgress = New-Object System.Management.Automation.RuntimeDefinedParameter('ShowProgress', [bool], $attributeCollection)
            $ShowProgress.Value = $false

            # Switch to enable/disable Multi-threading in the EXO cmdlets
            $UseMultithreading = New-Object System.Management.Automation.RuntimeDefinedParameter('UseMultithreading', [bool], $attributeCollection)
            $UseMultithreading.Value = $true

            # Pagesize Param
            $PageSize = New-Object System.Management.Automation.RuntimeDefinedParameter('PageSize', [uint32], $attributeCollection)
            $PageSize.Value = 1000

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

# EXO params end
            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('TrackPerformance', $TrackPerformance)
            $paramDictionary.Add('ShowProgress', $ShowProgress)
            $paramDictionary.Add('UseMultithreading', $UseMultithreading)
            $paramDictionary.Add('PageSize', $PageSize)
            $paramDictionary.Add('ManagedIdentity', $ManagedIdentity)
            $paramDictionary.Add('ManagedIdentityAccountId', $ManagedIdentityAccountId)
            if($PSEdition -eq 'Core')
            {
                $paramDictionary.Add('Device', $Device)
                $paramDictionary.Add('InlineCredential', $InlineCredential);
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            # Switch to collect telemetry on command execution. 
            $EnableErrorReporting = New-Object System.Management.Automation.RuntimeDefinedParameter('EnableErrorReporting', [switch], $attributeCollection)
            $EnableErrorReporting.Value = $false
            
            # Where to store EXO command telemetry data. By default telemetry is stored in the directory "%TEMP%/EXOTelemetry" in the file : EXOCmdletTelemetry-yyyymmdd-hhmmss.csv.
            $LogDirectoryPath = New-Object System.Management.Automation.RuntimeDefinedParameter('LogDirectoryPath', [string], $attributeCollection)
            $LogDirectoryPath.Value = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EXOCmdletTelemetry")

            # Create a new attribute and validate set against the LogLevel
            $LogLevelAttribute = New-Object System.Management.Automation.ParameterAttribute
            $LogLevelAttribute.Mandatory = $false
            $LogLevelAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $LogLevelAttributeCollection.Add($LogLevelAttribute)
            $LogLevelList = @([Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::Default, [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel]::All)
            $ValidateSet = New-Object System.Management.Automation.ValidateSetAttribute($LogLevelList)
            $LogLevel = New-Object System.Management.Automation.RuntimeDefinedParameter('LogLevel', [Microsoft.Online.CSE.RestApiPowerShellModule.Instrumentation.LogLevel], $LogLevelAttributeCollection)
            $LogLevel.Attributes.Add($ValidateSet)

            # Switch to CmdInline parameters
            $InlineCredential = New-Object System.Management.Automation.RuntimeDefinedParameter('InlineCredential', [switch], $attributeCollection)
            $InlineCredential.Value = $false

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            $paramDictionary.Add('EnableErrorReporting', $EnableErrorReporting)
            $paramDictionary.Add('LogDirectoryPath', $LogDirectoryPath)
            $paramDictionary.Add('LogLevel', $LogLevel)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('InlineCredential', $InlineCredential)
            return $paramDictionary
        }
    }
    process {
        $global:EXO_LastExecutionStatus = $true;
        $startTime = Get-Date

        # Validate parameters
        if (($ConnectionUri -ne '') -and (-not (Test-Uri $ConnectionUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        if (($AzureADAuthorizationEndpointUri -ne '') -and (-not (Test-Uri $AzureADAuthorizationEndpointUri)))
        {
            $global:EXO_LastExecutionStatus = $false;
            throw "Invalid AzureADAuthorizationEndpointUri parameter '$AzureADAuthorizationEndpointUri'"
        }
        if (($Prefix -ne ''))
        {
            if ($Prefix -notmatch '^[a-z0-9]+$') 
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Use of any special characters in the Prefix string is not supported."
            }
            if ($Prefix -eq 'EXO') 
            {
                $global:EXO_LastExecutionStatus = $false;
                throw "Prefix 'EXO' is a reserved Prefix, please use a different prefix."
            }
        }

        # Keep track of error count at beginning.
        $errorCountAtStart = $global:Error.Count;
        try
        {
            $moduleVersion = Get-ModuleVersion
            Write-Verbose "ModuleVersion: $moduleVersion"

            # Generate a ConnectionId to use in all logs and to send in all server calls.
            $connectionContextID = [System.Guid]::NewGuid()

            $cmdletLogger = New-CmdletLogger -ExoModuleVersion $moduleVersion -LogDirectoryPath $LogDirectoryPath.Value -EnableErrorReporting:$EnableErrorReporting.Value -ConnectionId $connectionContextID -IsRpsSession:$UseRPSSession.IsPresent
            $logFilePath = $cmdletLogger.GetCurrentLogFilePath()
            
            if ($EnableErrorReporting.Value -eq $true -and $UseRPSSession -eq $false)
            {
                Write-Message ("Writing cmdlet logs to " + $logFilePath)
            }

            $cmdletLogger.InitLog($connectionContextID)
            $cmdletLogger.LogStartTime($connectionContextID, $startTime)
            $cmdletLogger.LogCmdletName($connectionContextID, "Connect-ExchangeOnline");
            $cmdletLogger.LogCmdletParameters($connectionContextID, $PSBoundParameters);

            if ($isCloudShell -eq $false)
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value `
                -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring `
                -DelegatedOrg $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value `
                -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $CertificateThumbprint.Value -AppId $AppId.Value `
                -Organization $Organization.Value -Device:$Device.Value -InlineCredential:$InlineCredential.Value -CommandName $CommandName `
                -FormatTypeName $FormatTypeName -Prefix $Prefix -PageSize $PageSize.Value -ExoModuleVersion:$moduleVersion -Logger $cmdletLogger `
                -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value `
                -ManagedIdentity:$ManagedIdentity.Value -ManagedIdentityAccountId $ManagedIdentityAccountId.Value -AccessToken $AccessToken -DisableWAM:$DisableWAM -LogDirectoryPath $LogDirectoryPath.Value
            }
            else
            {
                $ConnectionContext = Get-ConnectionContext -ExchangeEnvironmentName $ExchangeEnvironmentName -ConnectionUri $ConnectionUri `
                -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -Credential $Credential.Value -PSSessionOption $PSSessionOption `
                -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -InlineCredential:$InlineCredential.Value `
                -DelegatedOrg $DelegatedOrganization -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $prefix -ExoModuleVersion:$moduleVersion `
                -Logger $cmdletLogger -ConnectionId $connectionContextID -IsRpsSession $UseRPSSession.IsPresent -EnableErrorReporting:$EnableErrorReporting.Value -AccessToken $AccessToken -DisableWAM:$DisableWAM -LogDirectoryPath $LogDirectoryPath.Value
            }

            if ($isCloudShell -eq $false)
            {
                $global:_EXO_EnableErrorReporting = $EnableErrorReporting.Value;
            }

            if ($ShowBanner -eq $true)
            {
                try
                {
                    $BannerContent = Get-EXOBanner -ConnectionContext:$ConnectionContext -IsRPSSession:$UseRPSSession.IsPresent
                    Write-Host -ForegroundColor Yellow $BannerContent
                }
                catch
                {
                    Write-Verbose "Failed to fetch banner content from server. Reason: $_"
                    $cmdletLogger.LogGenericError($connectionContextID, $_);
                }
            }

            if (($ConnectionUri -ne '') -and ($AzureADAuthorizationEndpointUri -eq ''))
            {
                Write-Information "Using ConnectionUri:'$ConnectionUri', in the environment:'$ExchangeEnvironmentName'."
            }
            if (($AzureADAuthorizationEndpointUri -ne '') -and ($ConnectionUri -eq ''))
            {
                Write-Information "Using AzureADAuthorizationEndpointUri:'$AzureADAuthorizationEndpointUri', in the environment:'$ExchangeEnvironmentName'."
            }

            $ImportedModuleName = '';
            $LogModuleDirectoryPath = [System.IO.Path]::GetTempPath();

            if ($UseRPSSession -eq $true)
            {
                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
                $ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);

                Import-Module $ModulePath;

                $global:_EXO_ModulePath = $ModulePath;

                $PSSession = New-ExoPSSession -ConnectionContext $ConnectionContext

                if ($PSSession -ne $null)
                {
                    if ([string]::IsNullOrEmpty($Prefix))
                    {
                        $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking -CommandName $CommandName -FormatTypeName $FormatTypeName
                    }
                    else
                    {
                        $PSSessionModuleInfo = Import-PSSession $PSSession -AllowClobber -DisableNameChecking -CommandName $CommandName -FormatTypeName $FormatTypeName -Prefix $Prefix
                    }
                    # Add the name of the module to clean up in case of removing the broken session
                    $PSSession | Add-Member -NotePropertyName "CurrentModuleName" -NotePropertyValue $PSSessionModuleInfo.Name

                    # Import the above module globally. This is needed as with using psm1 files, 
                    # any module which is dynamically loaded in the nested module does not reflect globally.
                    $CurrentModule = Import-Module $PSSessionModuleInfo.Path -Global -DisableNameChecking -Prefix $Prefix -PassThru
                    $CurrentModule | Add-Member -NotePropertyName "ModulePrefix" -NotePropertyValue $Prefix
                    $CurrentModule | Add-Member -NotePropertyName "CommandName" -NotePropertyValue $CommandName
                    $CurrentModule | Add-Member -NotePropertyName "FormatTypeName" -NotePropertyValue $FormatTypeName

                    UpdateImplicitRemotingHandler

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                    $ImportedModuleName = $PSSessionModuleInfo.Name;
                }
            }
            else
            {
                # Download the new web based EXOModule
                if ($SigningCertificate -ne $null)
                {
                    $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData -SigningCertificate $SigningCertificate;
                }
                else
                {
                    $ImportedModule = New-EXOModule -ConnectionContext $ConnectionContext -SkipLoadingFormatData:$SkipLoadingFormatData;
                }
                if ($null -ne $ImportedModule)
                {
                    $ImportedModuleName = $ImportedModule.Name;
                    $LogModuleDirectoryPath = $ImportedModule.ModuleBase

                    Write-Verbose "AutoGen EXOModule created at  $($ImportedModule.ModuleBase)"

                    if ($LoadCmdletHelp -eq $true -and $HelpFileNames -ne $null -and $HelpFileNames -is [array] -and $HelpFileNames.Count -gt 0)
                    {
                        Get-HelpFiles -HelpFileNames $HelpFileNames -ConnectionContext $ConnectionContext -ImportedModule $ImportedModule -EnableErrorReporting:$EnableErrorReporting.Value
                    }
                    else
                    {
                        $cmdletLogger.LogGenericInfo($connectionContextID, "Skipping cmdlet help data");
                    }
                }
                else
                {
                    $global:EXO_LastExecutionStatus = $false;
                    throw "Module could not be correctly formed. Please run Connect-ExchangeOnline again."
                }
            }

            # If we are configured to collect telemetry, add telemetry wrappers in case of an RPS connection
            if ($EnableErrorReporting.Value -eq $true)
            {
                if ($UseRPSSession -eq $true)
                {
                    $FilePath = Add-EXOClientTelemetryWrapper -Organization (Get-OrgNameFromUPN -UPN $UserPrincipalName.Value) -PSSessionModuleName $ImportedModuleName -LogDirectoryPath $LogDirectoryPath.Value -LogModuleDirectoryPath $LogModuleDirectoryPath
                    Write-Message("Writing telemetry records for this session to " + $FilePath[0] );
                    $global:_EXO_TelemetryFilePath = $FilePath[0]
                    Import-Module $FilePath[1] -DisableNameChecking -Global

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }

                $endTime = Get-Date
                $cmdletLogger.LogEndTime($connectionContextID, $endTime);
                $cmdletLogger.CommitLog($connectionContextID);

                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }
            }
            else 
            {
                if ($EOPConnectionInProgress -eq $false)
                {
                    # Set the AppSettings disabling the logging
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -EnableErrorReporting $false
                }
            }
        }
        catch
        {
            # If telemetry is enabled, log errors generated from this cmdlet also.
            # If telemetry is not enabled, calls to cmdletLogger will be a no-op.
            $errorCountAtProcessEnd = $global:Error.Count 
            $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
            for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
            {
                $cmdletLogger.LogGenericError($connectionContextID, $global:Error[$i]);
            }

            $cmdletLogger.CommitLog($connectionContextID);

            if ($EnableErrorReporting.Value -eq $true -and $UseRPSSession -eq $true)
            {
                if ($global:_EXO_TelemetryFilePath -eq $null)
                {
                    $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath -LogDirectoryPath $LogDirectoryPath.Value

                    # Import the REST module
                    $RestPowershellModule = "Microsoft.Exchange.Management.RestApiClient.dll";
                    $RestModulePath = [System.IO.Path]::Combine($PSScriptRoot, $RestPowershellModule);
                    Import-Module $RestModulePath -Cmdlet Set-ExoAppSettings;

                    # Set the AppSettings
                    Set-ExoAppSettings -ShowProgress $ShowProgress.Value -PageSize $PageSize.Value -UseMultithreading $UseMultithreading.Value -TrackPerformance $TrackPerformance.Value -ConnectionUri $ConnectionUri -EnableErrorReporting $true -LogDirectoryPath $LogDirectoryPath.Value -LogLevel $LogLevel.Value
                }

                # Log errors which are encountered during Connect-ExchangeOnline execution. 
                Write-Message ("Writing Connect-ExchangeOnline error log to " + $global:_EXO_TelemetryFilePath)
                Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Connect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
            }

            $global:EXO_LastExecutionStatus = $false;

            if ($_.Exception -ne $null)
            {
                # Connect-ExchangeOnline Failed, Remove ConnectionContext from Map.
                if ([Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveConnectionContextUsingConnectionId($connectionContextID))
                {
                    Write-Verbose "ConnectionContext Removed"
                }

                if ($_.Exception.InnerException -ne $null)
                {
                    throw $_.Exception.InnerException;
                }
                else
                {
                    throw $_.Exception;
                }
            }
            else
            {
                throw $_;
            }
        }
    }
}

function Connect-IPPSSession
{
    [CmdletBinding()]
    param(
        # Connection Uri for the Remote PowerShell endpoint
        [string] $ConnectionUri = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId',

        # Azure AD Authorization endpoint Uri that can issue the OAuth2 access tokens
        [string] $AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations',

        # Delegated Organization Name
        [string] $DelegatedOrganization = '',

        # PowerShell session options to be used when opening the Remote PowerShell session
        [System.Management.Automation.Remoting.PSSessionOption] $PSSessionOption = $null,

        # Switch to bypass use of mailbox anchoring hint.
        [switch] $BypassMailboxAnchoring = $false,

        # Prefix 
        [string] $Prefix = '',

        #Cmdlets to Import, by default it would bring all
        [string[]] $CommandName = @("*"),

        #The way the output objects would be printed on the console
        [string[]] $FormatTypeName = @("*"),

        # Use Remote PowerShell Session based connection
        [switch] $UseRPSSession = $false,

        # Show Banner of scc cmdlets Mapping and recent updates
        [switch] $ShowBanner = $true,

        # switch to disable WAM
        [switch] $DisableWAM = $false,

        # Externally provided access token
        [string] $AccessToken = ''
    )
    DynamicParam
    {
        if (($isCloudShell = IsCloudShellEnvironment) -eq $false)
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # User Principal Name or email address of the user
            $UserPrincipalName = New-Object System.Management.Automation.RuntimeDefinedParameter('UserPrincipalName', [string], $attributeCollection)
            $UserPrincipalName.Value = ''

            # User Credential to Logon
            $Credential = New-Object System.Management.Automation.RuntimeDefinedParameter('Credential', [System.Management.Automation.PSCredential], $attributeCollection)
            $Credential.Value = $null

            # Certificate
            $Certificate = New-Object System.Management.Automation.RuntimeDefinedParameter('Certificate', [System.Security.Cryptography.X509Certificates.X509Certificate2], $attributeCollection)
            $Certificate.Value = $null

            # Certificate Path 
            $CertificateFilePath = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateFilePath', [string], $attributeCollection)
            $CertificateFilePath.Value = ''

            # Certificate Password
            $CertificatePassword = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificatePassword', [System.Security.SecureString], $attributeCollection)
            $CertificatePassword.Value = $null

            # Certificate Thumbprint
            $CertificateThumbprint = New-Object System.Management.Automation.RuntimeDefinedParameter('CertificateThumbprint', [string], $attributeCollection)
            $CertificateThumbprint.Value = ''

            # Application Id
            $AppId = New-Object System.Management.Automation.RuntimeDefinedParameter('AppId', [string], $attributeCollection)
            $AppId.Value = ''

            # Organization
            $Organization = New-Object System.Management.Automation.RuntimeDefinedParameter('Organization', [string], $attributeCollection)
            $Organization.Value = ''

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('UserPrincipalName', $UserPrincipalName)
            $paramDictionary.Add('Credential', $Credential)
            $paramDictionary.Add('Certificate', $Certificate)
            $paramDictionary.Add('CertificateFilePath', $CertificateFilePath)
            $paramDictionary.Add('CertificatePassword', $CertificatePassword)
            $paramDictionary.Add('AppId', $AppId)
            $paramDictionary.Add('Organization', $Organization)
            if($PSEdition -eq 'Core')
            {
                # We do not want to expose certificate thumprint in Linux as it is not feasible there.
                if($IsWindows)
                {
                    $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
                }
            }
            else 
            {
                $paramDictionary.Add('CertificateThumbprint', $CertificateThumbprint);
            }

            return $paramDictionary
        }
        else
        {
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.Mandatory = $false

            $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)

            # Switch to MSI auth 
            $Device = New-Object System.Management.Automation.RuntimeDefinedParameter('Device', [switch], $attributeCollection)
            $Device.Value = $false

            $paramDictionary = New-object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Device', $Device)
            return $paramDictionary
        }
    }
    process 
    {
        try
        {
            $EOPConnectionInProgress = $true
            if ($isCloudShell -eq $false)
            {
                $certThumbprint = $CertificateThumbprint.Value
                # Will pass CertificateThumbprint if it is not null or not empty
                if($certThumbprint)
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -CertificateThumbprint $certThumbprint -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession -DisableWAM:$DisableWAM
                }
                else
                {
                    Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -UserPrincipalName $UserPrincipalName.Value -PSSessionOption $PSSessionOption -Credential $Credential.Value -BypassMailboxAnchoring:$BypassMailboxAnchoring -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Certificate $Certificate.Value -CertificateFilePath $CertificateFilePath.Value -CertificatePassword $CertificatePassword.Value -AppId $AppId.Value -Organization $Organization.Value -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession -DisableWAM:$DisableWAM -AccessToken $AccessToken
                }
            }
            else
            {
                Connect-ExchangeOnline -ConnectionUri $ConnectionUri -AzureADAuthorizationEndpointUri $AzureADAuthorizationEndpointUri -PSSessionOption $PSSessionOption -BypassMailboxAnchoring:$BypassMailboxAnchoring -Device:$Device.Value -ShowBanner:$ShowBanner -DelegatedOrganization $DelegatedOrganization -Prefix $Prefix -CommandName $CommandName -FormatTypeName $FormatTypeName -UseRPSSession:$UseRPSSession -DisableWAM:$DisableWAM -AccessToken $AccessToken
            }
        }
        finally
        {
            $EOPConnectionInProgress = $false
        }
    }
}

function Disconnect-ExchangeOnline
{
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact='High', DefaultParameterSetName='DefaultParameterSet')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='ConnectionId', ValueFromPipelineByPropertyName=$true)]
        [string[]] $ConnectionId,
        [Parameter(Mandatory=$true, ParameterSetName='ModulePrefix')]
        [string[]] $ModulePrefix
    )

    process
    {
        $global:EXO_LastExecutionStatus = $true;
        $disconnectConfirmationMessage = ""
        Switch ($PSCmdlet.ParameterSetName)
        {
            'ConnectionId'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByConnectionId($ConnectionId)
                break
            }
            'ModulePrefix'
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageByModulePrefix($ModulePrefix)
                break
            }
            Default
            {
                $disconnectConfirmationMessage = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetDisconnectConfirmationMessageWithInbuilt()
            }
        }
	
        if ($PSCmdlet.ShouldProcess(
            $disconnectConfirmationMessage,
            "Press(Y/y/A/a) if you want to continue.",
            $disconnectConfirmationMessage))
        {

            # Keep track of error count at beginning.
            $errorCountAtStart = $global:Error.Count;
            $startTime = Get-Date

            try
            {
                # Get all the connection contexts so that the logger can be initialized.
                $connectionContexts = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetAllConnectionContexts()
                $disconnectCmdletId = [System.Guid]::NewGuid().ToString()

                # denotes if any of the connections is an RPS session.
                # This is used to Push-EXOTelemetryRecord in case any RPS connections are present.
                $rpsConnectionWithErrorReportingExists = $false
                
                foreach ($context in $connectionContexts)
                {
                    $context.Logger.InitLog($disconnectCmdletId);
                    $context.Logger.LogStartTime($disconnectCmdletId, $startTime);
                    $context.Logger.LogCmdletName($disconnectCmdletId, "Disconnect-ExchangeOnline");
                    if ($context.IsRpsSession -and $context.ErrorReportingEnabled)
                    {
                        $rpsConnectionWithErrorReportingExists = $true
                    }
                }

                # Import the module once more to ensure that Test-ActiveToken is present
                $ExoPowershellModule = "Microsoft.Exchange.Management.ExoPowershellGalleryModule.dll";
                $ModulePath = [System.IO.Path]::Combine($PSScriptRoot, $ExoPowershellModule);
                Import-Module $ModulePath -Cmdlet Clear-ActiveToken;

                $existingPSSession = Get-PSSession | Where-Object {$_.ConfigurationName -like "Microsoft.Exchange" -and $_.Name -like "ExchangeOnlineInternalSession*"}

                if ($existingPSSession.count -gt 0) 
                {
                    for ($index = 0; $index -lt $existingPSSession.count; $index++)
                    {
                        $session = $existingPSSession[$index]
                        Remove-PSSession -session $session

                        Write-Information "Removed the PSSession $($session.Name) connected to $($session.ComputerName)"

                        # Remove any active access token from the cache
                        # If the connectionId of the session being cleared is equal to AppSettings.ConnectionId, this means connection to EXO cmdlets will break.
                        if ($session.ConnectionContext.ConnectionId -ieq [Microsoft.Exchange.Management.AdminApiProvider.AppSettings]::ConnectionId)
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$true
                        }
                        else
                        {
                            Clear-ActiveToken -TokenProvider $session.TokenProvider -IsSessionUsedByInbuiltCmdlets:$false
                        }

                        # Remove any previous modules loaded because of the current PSSession
                        if ($session.PreviousModuleName -ne $null)
                        {
                            if ((Get-Module $session.PreviousModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.PreviousModuleName -ErrorAction SilentlyContinue
                            }

                            $session.PreviousModuleName = $null
                        }

                        # Remove any leaked module in case of removal of broken session object
                        if ($session.CurrentModuleName -ne $null)
                        {
                            if ((Get-Module $session.CurrentModuleName).Count -ne 0)
                            {
                                $null = Remove-Module -Name $session.CurrentModuleName -ErrorAction SilentlyContinue
                            }
                        }
                    }
                }

                $modulesToRemove = $null
                Switch ($PSCmdlet.ParameterSetName)
                {
                    'ConnectionId'
                    {
                        # Call GetModulesToRemoveByConnectionId in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByConnectionId($ConnectionId)
                        break
                    }
                    'ModulePrefix'
                    {
                        # Call GetModulesToRemoveByModulePrefix in this scenario
                        $modulesToRemove = [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::GetModulesToRemoveByModulePrefix($ModulePrefix)
                        break
                    }
                    Default
                    {
                        # Remove all the AutoREST modules from this instance of powershell if created
                        $existingAutoRESTModules = Get-Module "tmpEXO_*"
                        foreach ($module in $existingAutoRESTModules)
                        {  
                            $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                        }

                        # The below call to remove all connection contexts could be removed as we already have an OnRemove event hooked to the module. Work Item 3461604 to investigate this
                        # Remove all ConnectionContexts
                        # this internally clears all the active tokens in ConnectionContexts
                        [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionContextFactory]::RemoveAllConnectionContexts()
                    }
                }

                if ($modulesToRemove -ne $null -and $modulesToRemove.Count -gt 0)
                {
                    $null = Remove-Module $modulesToRemove -ErrorAction SilentlyContinue
                }

                Write-Information "Disconnected successfully !"

                # Remove all the Wrapped modules from this instance of powershell if created
                $existingWrappedModules = Get-Module "EXOCmdletWrapper-*"
                foreach ($module in $existingWrappedModules)
                {
                    $null = Remove-Module -Name $module -ErrorAction SilentlyContinue
                }

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid
                }
            }
            catch
            {
                # If telemetry is enabled for any of the connections, log errors generated from this cmdlet also. 
                $errorCountAtProcessEnd = $global:Error.Count
                $global:EXO_LastExecutionStatus = $false;

                $endTime = Get-Date
                foreach ($context in $connectionContexts)
                {
                    $numErrorRecordsToConsider = $errorCountAtProcessEnd - $errorCountAtStart
                    for ($i = 0 ; $i -lt $numErrorRecordsToConsider ; $i++)
                    {
                        $context.Logger.LogGenericError($disconnectCmdletId, $global:Error[$i]);
                    }

                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                    $logFilePath = $context.Logger.GetCurrentLogFilePath();
                }

                if ($rpsConnectionWithErrorReportingExists)
                {
                    if ($global:_EXO_TelemetryFilePath -eq $null)
                    {
                        $global:_EXO_TelemetryFilePath = New-EXOClientTelemetryFilePath
                    }

                    # Log errors which are encountered during Disconnect-ExchangeOnline execution. 
                    Write-Message ("Writing Disconnect-ExchangeOnline errors to " + $global:_EXO_TelemetryFilePath)

                    Push-EXOTelemetryRecord -TelemetryFilePath $global:_EXO_TelemetryFilePath -CommandName Disconnect-ExchangeOnline -CommandParams $PSCmdlet.MyInvocation.BoundParameters -OrganizationName  $global:_EXO_ExPSTelemetryOrganization -ScriptName $global:_EXO_ExPSTelemetryScriptName  -ScriptExecutionGuid $global:_EXO_ExPSTelemetryScriptExecutionGuid -ErrorObject $global:Error -ErrorRecordsToConsider ($errorCountAtProcessEnd - $errorCountAtStart) 
                }

                throw $_
            }

            $endTime = Get-Date
            foreach ($context in $connectionContexts)
            {
                if ($context.Logger -ne $null)
                {
                    $context.Logger.LogEndTime($disconnectCmdletId, $endTime);
                    $context.Logger.CommitLog($disconnectCmdletId);
                }
            }
        }
    }
}

# SIG # Begin signature block
# MIIoKAYJKoZIhvcNAQcCoIIoGTCCKBUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDS7PPw9W6auhLX
# UZO1Rh/NNAhk/yT9m0z9IxgpK6RhNKCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGggwghoEAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggZAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJKoZIhvcNAQkEMSIE
# IAUQ0tIUmoj57fHOlyKSx5Tc5xwH2arrscmojIjkaCsmMEIGCisGAQQBgjcCAQwx
# NDAyoBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20wDQYJKoZIhvcNAQEBBQAEggEASyuTEoiPjU/2So5SZJaZQTb/WufIvEFF
# mSC7mVBHzmIvZP+/gp8AwCM6NhrVS8iYXKYcFWrrnn5ILQBcYGHoyckn71Jwu/6d
# SPzhOmy0LdI6ptM1wBQyQSI4KMdaV3FQAMgvO7HaKLcxOcPheLH/l36B+Drrj40S
# /Nae+FAy5LZw8ngthHeCe8gdNUCPXIufa66aVu9vWXHCiX1sD9oNNFtwy9FAnzs0
# JU0IzUJ6btlVonrl+ksZgXhk9qDD8AGvnJHnOhoj1eM19SijrikRc6yvk7tig502
# TXJpI82PNkyaBT7nA5uPtofkcwEuGZqAu2XCKKFapoBfvwnldQAmjaGCF7Awghes
# BgorBgEEAYI3AwMBMYIXnDCCF5gGCSqGSIb3DQEHAqCCF4kwgheFAgEDMQ8wDQYJ
# YIZIAWUDBAIBBQAwggFaBgsqhkiG9w0BCRABBKCCAUkEggFFMIIBQQIBAQYKKwYB
# BAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDy+wCQ7UNpt994ClylcjhHZrvXvtk+
# FID+Er2pLshY8wIGZ+0kBdH4GBMyMDI1MDQxNzAxMzEyOC4xMjlaMASAAgH0oIHZ
# pIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYD
# VQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNV
# BAsTHm5TaGllbGQgVFNTIEVTTjo0QzFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEf4wggcoMIIFEKADAgECAhMzAAAB
# /xI4fPfBZdahAAEAAAH/MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFBDQSAyMDEwMB4XDTI0MDcyNTE4MzExOVoXDTI1MTAyMjE4MzExOVowgdMx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1p
# Y3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNo
# aWVsZCBUU1MgRVNOOjRDMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQg
# VGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAyeiV0pB7bg8/qc/mkiDdJXnzJWPYgk9mTGeI3pzQpsyrRJREWcKYHd/9db+g
# 3z4dU4VCkAZEXqvkxP5QNTtBG5Ipexpph4PhbiJKwvX+US4KkSFhf1wflDAY1tu9
# CQqhhxfHFV7vhtmqHLCCmDxhZPmCBh9/XfFJQIUwVZR8RtUkgzmN9bmWiYgfX0R+
# bDAnncUdtp1xjGmCpdBMygk/K0h3bUTUzQHb4kPf2ylkKPoWFYn2GNYgWw8PGBUO
# 0vTMKjYD6pLeBP0hZDh5P3f4xhGLm6x98xuIQp/RFnzBbgthySXGl+NT1cZAqGyE
# hT7L0SdR7qQlv5pwDNerbK3YSEDKk3sDh9S60hLJNqP71iHKkG175HAyg6zmE5p3
# fONr9/fIEpPAlC8YisxXaGX4RpDBYVKpGj0FCZwisiZsxm0X9w6ZSk8OOXf8JxTY
# WIqfRuWzdUir0Z3jiOOtaDq7XdypB4gZrhr90KcPTDRwvy60zrQca/1D1J7PQJAJ
# Obbiaboi12usV8axtlT/dCePC4ndcFcar1v+fnClhs9u3Fn6LkHDRZfNzhXgLDEw
# b6dA4y3s6G+gQ35o90j2i6amaa8JsV/cCF+iDSGzAxZY1sQ1mrdMmzxfWzXN6sPJ
# My49tdsWTIgZWVOSS9uUHhSYkbgMxnLeiKXeB5MB9QMcOScCAwEAAaOCAUkwggFF
# MB0GA1UdDgQWBBTD+pXk/rT/d7E/0QE7hH0wz+6UYTAfBgNVHSMEGDAWgBSfpxVd
# AF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIw
# UENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsGAQUFBzAChlBo
# dHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMBAf8EAjAAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0B
# AQsFAAOCAgEAOSNN5MpLiyunm866frWIi0hdazKNLgRp3WZPfhYgPC3K/DNMzLli
# YQUAp6WtgolIrativXjOG1lIjayG9r6ew4H1n5XZdDfJ12DLjopap5e1iU/Yk0eu
# tPyfOievfbsIzTk/G51+uiUJk772nVzau6hI2KGyGBJOvAbAVFR0g8ppZwLghT4z
# 3mkGZjq/O4Z/PcmVGtjGps2TCtI4rZjPNW8O4c/4aJRmYQ/NdW91JRrOXRpyXrTK
# UPe3kN8N56jpl9kotLhdvd89RbOsJNf2XzqbAV7XjV4caCglA2btzDxcyffwXhLu
# 9HMU3dLYTAI91gTNUF7BA9q1EvSlCKKlN8N10Y4iU0nyIkfpRxYyAbRyq5QPYPJH
# GA0Ty0PD83aCt79Ra0IdDIMSuwXlpUnyIyxwrDylgfOGyysWBwQ/js249bqQOYPd
# pyOdgRe8tXdGrgDoBeuVOK+cRClXpimNYwr61oZ2/kPMzVrzRUYMkBXe9WqdSezh
# 8tytuulYYcRK95qihF0irQs6/WOQJltQX79lzFXE9FFln9Mix0as+C4HPzd+S0bB
# N3A3XRROwAv016ICuT8hY1InyW7jwVmN+OkQ1zei66LrU5RtAz0nTxx5OePyjnTa
# ItTSY4OGuGU1SXaH49JSP3t8yGYA/vorbW4VneeD721FgwaJToHFkOIwggdxMIIF
# WaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNy
# b3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0yMTA5MzAx
# ODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
# MDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0ciELeaLL
# 1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/XE/HZveVU3Fa4n5K
# Wv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxRMTeg
# Cjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62AW36MEBydUv62
# 6GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SH
# JMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCGMFxPLOJiss25
# 4o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/Nme
# Rd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP8BDyt0cY7afo
# mXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLi
# Mxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHzNgY1GIRH29wb
# 0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV2xo3xwgVGD94q0W2
# 9R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsGAQQBgjcVAQQF
# AgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUpzxD/LwTuMB0GA1Ud
# DgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQBgjdM
# g30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJ
# KwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8w
# TTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVj
# dHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBK
# BggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9N
# aWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcNAQELBQADggIBAJ1V
# ffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/2mRsfNB1
# OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYAA7AFvonoaeC6Ce57
# 32pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveVtihV
# J9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn8AtqgcKBGUIZ
# UnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW
# 9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyxTkctwRQEcb9k
# +SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pF
# EUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9y8FBSX5+k77L
# +DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1
# ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhwRNGQ8cirOoo6
# CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDWTCCAkECAQEwggEBoYHZ
# pIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYD
# VQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJzAlBgNV
# BAsTHm5TaGllbGQgVFNTIEVTTjo0QzFBLTA1RTAtRDk0NzElMCMGA1UEAxMcTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAqROMbMS8
# JcUlcnPkwRLFRPXFspmggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMDANBgkqhkiG9w0BAQsFAAIFAOuqv1EwIhgPMjAyNTA0MTYyMzQ0MTdaGA8y
# MDI1MDQxNzIzNDQxN1owdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA66q/UQIBADAK
# AgEAAgIt+wIB/zAHAgEAAgISoDAKAgUA66wQ0QIBADA2BgorBgEEAYRZCgQCMSgw
# JjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3
# DQEBCwUAA4IBAQCTirUyWiijvaLhtaKibIyCMdzexUQTGHhyDROQga134/2KOEBn
# RkqvnPR6ZMIq3ssWYzpKj/3HXoP+x/nAAvbsjojMWZ4C5tG/PjS2rDaxwK95fW7/
# Yf1FrppZBEL2bntM+YCmGUgZefxDR54MrYN+VNLyJ3Cxovr4WQtZAASYAhdq5dH2
# FYGV1mzp3hksJnInv4Jvuy2vFJ5TPKmdzIL2jJdxUCZ+BiStJY08nkdnbtChow6I
# EtlQJtOL65puxA+wEP5wNj4+F+HjwH41Vg8qN815suY7PoAFYCfcHijyUX+oZ9ZB
# mFuTVygKhu80ZBsuyDmz/nz/46ERIV+cictTMYIEDTCCBAkCAQEwgZMwfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAH/Ejh898Fl1qEAAQAAAf8wDQYJ
# YIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkq
# hkiG9w0BCQQxIgQgkt7py3oz1/JJYtH4iJj25goaRO9iTw2MC7N5gujxkHMwgfoG
# CyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCDkMu++yQJ3aaycIuMT6vA7JNuMaVOI
# 3qDjSEV8upyn/TCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# AhMzAAAB/xI4fPfBZdahAAEAAAH/MCIEIOr88Aj91ZsvEnhd7gLhRjzmDbjuhnsj
# SmPRa1n1PC2MMA0GCSqGSIb3DQEBCwUABIICAHQndEiLG0hJdCOFyG8kmmnoKrvk
# Oe1nRWO8Dvf2SvmigBUGk6lfMekPzGOYicFiofTEOYn0tMA+e3qPOHiPYrSYHjYp
# PnZfQMSBWho9bLfd5z1aEn7soZYuPhwZLmIirIJwOPWm377oAFF1ER3UpLh3cGWr
# POv9DLXLEUj9VyvvQnVqs8P0AlPvaaU+4ooWVh4R/IcoJG8401ml/z4qJbTWg5c9
# KSBWjjQr4/juTXOdLgYfA27Xn3hdeVBQ0x61XiykyiQxuVeee4vbo4D52mDUov1Q
# CffhyXFKC0YcJI6iwlFrWZGr8WU6NIXrblvlnc1JAuta9i+1hT88ugzF0l5eMaiI
# 27j3CTI5GvU2qGwCA7zB4mijhZT7laDna3pIZSBRgp1oEpqCQD2I1jvVqVfg8903
# 2gB3q4bipWf5WbHqZ4WA1tyuH+aypb5mrHKYvQiCHlLepwpdUFBPQa3ghsIWw5G3
# 2RWccjLoAzp6OooQAKk2rTDVVdWsmfx501r0Eh77Vod7z1fq82PXzCu+wMJVD1qF
# Ztn8/R5FOy+kDXTqXEpQ043cj0iyJ5KP937INJp6shR7slozmYL3ay9gR+3uBmsU
# qaYbQKZlPOjMo6h8r2RjdQpO2HwGRGAaGX4JvNqWCUJeVw5v1PtJ7tAZdUbe8wHL
# EhFyuDYHy767qyf/
# SIG # End signature block
