<#
.SYNOPSIS
    Manages SCCM software on client machines, providing status checks, difference analysis,
    policy synchronization, and options for application installation.

.DESCRIPTION
    This script facilitates the management of SCCM deployed applications on Windows clients.
    It performs the following key functions:
    1.  Retrieves a detailed list of locally installed software from the client's registry.
    2.  Fetches comprehensive status information about SCCM applications known to the client,
        including their deployment state, target type (machine/user), and availability for installation.
    3.  Introduces a robust "server-intended vs. client-actual" difference analysis. This compares
        what the SCCM server *intends* for the client (via application assignments) with what
        is *actually* installed on the client, highlighting discrepancies.
    4.  Offers options to:
        - Trigger an immediate Machine Policy Retrieval & Evaluation Cycle on the client (to refresh policy).
        - Initiate other specific SCCM client cycles (e.g., Hardware Inventory, App Deployment Evaluation).
        - Automatically install all identified pending/relevant SCCM applications.
        - Prompt for individual installation of pending/relevant SCCM applications.
    5.  Ensures full compatibility with PowerShell 5.1 in on-premise Windows environments.
    6.  Implements German locale settings for CSV output and date/time formatting.
    7.  Provides extensive verbose logging and detailed execution traceability.
    8.  Features interactive prompts for essential missing parameters, ensuring flexible execution.

.PARAMETER ComputerName
    Specifies the target computer name or IP address. If not provided, it defaults to the local
    computer and will prompt interactively if run without this parameter.

.PARAMETER InstallAll
    Use this switch to automatically trigger the installation of all identified pending or
    otherwise relevant SCCM applications without further user interaction.
    Cannot be used with -PromptInstall.

.PARAMETER PromptInstall
    Use this switch to be prompted for individual installation of each identified pending or
    otherwise relevant SCCM application. This is the default installation behavior if
    -InstallAll is not specified and no other installation choice is made interactively.

.PARAMETER RefreshPolicy
    Use this switch to trigger an immediate "Machine Policy Retrieval & Evaluation Cycle"
    on the target client before processing application statuses. This ensures the client
    has the latest policies from the SCCM server.

.PARAMETER ShowDifferences
    Use this switch to display a detailed comparison between what the SCCM server has
    assigned to the client (server's intent, as seen by client) and the actual
    installation status on the client. Differences will be highlighted. This display occurs
    *before* any policy refresh or installation actions.

.PARAMETER AdditionalCycle
    Specifies an additional SCCM client cycle to trigger after a potential policy refresh.
    Valid options include: "MachinePolicy", "UserPolicy", "HardwareInventory",
    "SoftwareInventory", "AppDeploymentEval", "SoftwareUpdateScan", "DiscoveryData".

.PARAMETER DisableVerbose
    Use this switch to disable the verbose output logging. By default, verbose output is enabled.

.EXAMPLE
    # Interactive mode (prompts for ComputerName, then installation choice, then individual apps)
    .\Invoke-SCCMClientAppMgmt.ps1

.EXAMPLE
    # Check status and show differences for a remote PC, then prompt for individual installs.
    .\Invoke-SCCMClientAppMgmt.ps1 -ComputerName "L26PBZD70080021" -ShowDifferences -PromptInstall

.EXAMPLE
    # Trigger a policy refresh and automatically install all pending apps on a remote PC.
    .\Invoke-SCCMClientAppMgmt.ps1 -ComputerName "L26PBZD70080021" -RefreshPolicy -InstallAll

.EXAMPLE
    # Just show the status and differences for the local PC, no installation prompts.
    .\Invoke-SCCMClientAppMgmt.ps1 -ComputerName $env:COMPUTERNAME -ShowDifferences -PromptInstall:$false

.EXAMPLE
    # Trigger a Hardware Inventory Cycle on a remote PC and then display status.
    .\Invoke-SCCMClientAppMgmt.ps1 -ComputerName "L26PBZD70080021" -AdditionalCycle "HardwareInventory" -PromptInstall:$false

.NOTES
    Author: Coding-Assistent (Optimized based on User Feedback)
    Date: 2025-05-23
    Version: 2.0.0
    Compatibility: PowerShell 5.1+, Windows 10, Windows Server 2019 (On-Premise)

    Required Permissions:
        - Local Administrator on the target computer for WMI and Invoke-Command.
        - WinRM (Windows Remote Management) enabled on the target computer.

    Known Limitations:
        - "Server-intended" status is derived from the client's cached assignments, not live SCCM server data.
        - Fuzzy matching for application names (Convert-AppNameToNormalized) is heuristic and may not be perfect.

    References:
        - AI-Discovered:
            - Microsoft Docs: PowerShell Cmdlet Naming Conventions (https://learn.microsoft.com/en-us/powershell/scripting/developer/cmdlet/approved-verbs-for-windows-powershell-commands)
            - Microsoft Docs: About CommonParameters (-Verbose) (https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_commonparameters#-verbose)
            - Microsoft Docs: WMI Class CCM_Application (https://learn.microsoft.com/en-us/previous-versions/system-center/developer/ccm-application)
            - Microsoft Docs: WMI Class SMS_Client (https://learn.microsoft.com/en-us/previous-versions/system-center/developer/sms-client)
            - Microsoft Docs: WMI Class SMS_ApplicationAssignment (https://learn.microsoft.com/en-us/previous-versions/system-center/developer/sms-applicationassignment)
            - Stack Overflow: PowerShell DateTime format to German (https://stackoverflow.com/questions/21356265/powershell-datetime-format-to-german-culture)
        - User-Provided:
            - Feedback on SCCM Client COM Objects and WMI classes.
#>

# --- Global Configuration and Helper Definitions ---

# Region-specific settings for Germany
$script:GermanCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
$script:CsvSeparator = ";" # German CSV separator

# Dictionary with the GUIDs for common SCCM client actions
# This makes the script more flexible and readable.
$global:SCCMCycleGUIDs = @{
    "MachinePolicy" = "{00000000-0000-0000-0000-000000000021}"; # Machine Policy Retrieval & Evaluation Cycle
    "UserPolicy" = "{00000000-0000-0000-0000-000000000027}"; # User Policy Retrieval & Evaluation Cycle
    "HardwareInventory" = "{00000000-0000-0000-0000-000000000001}"; # Hardware Inventory Cycle
    "SoftwareInventory" = "{00000000-0000-0000-0000-000000000002}"; # Software Inventory Cycle
    "AppDeploymentEval" = "{00000000-0000-0000-0000-000000000121}"; # Application Deployment Evaluation Cycle
    "SoftwareUpdateScan" = "{00000000-0000-0000-0000-000000000113}"; # Software Updates Scan Cycle
    "DiscoveryData" = "{00000000-0000-0000-0000-000000000003}"; # Data Discovery Record (DDR) Cycle
}

# --- Helper Functions ---

<#
.SYNOPSIS
    Normalizes a software name to facilitate consistent comparison.

.DESCRIPTION
    This function takes a raw software name (e.g., from installed applications or SCCM)
    and attempts to normalize it by removing common prefixes (like manufacturers),
    architecture suffixes (x64), version numbers, and other common variations.
    This helps in robustly comparing names from different sources that might use
    slightly different naming conventions.

.PARAMETER Name
    The software name string to be normalized.

.RETURNS
    The normalized string.
#>
function Convert-AppNameToNormalized {
    [CmdletBinding(SupportsShouldProcess=$false)]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Name
    )
    Process {
        Write-Verbose "Normalizing name: '$Name'"
        $normalizedName = $Name.ToLower().Trim()

        # Remove known prefixes and suffixes (case-insensitive)
        $normalizedName = $normalizedName -replace "(?i)^(microsoft|adobe|google|mozilla|ibm|cisco|vmware|nvidia|intel|hp|dell)\s+"
        $normalizedName = $normalizedName -replace "(?i)\s+(x64|x86|64-bit|32-bit|arm)$"
        # More robust version number removal
        $normalizedName = $normalizedName -replace "(?i)\s+v?(\d+(\.\d+){0,4}(\.\d+)?(-[a-z0-9]+)?(\s+\d+)?(\s+\(.*\))?)$"
        $normalizedName = $normalizedName -replace "(?i)\s+(setup|installer|client|agent|runtime|package|redistributable|sdk|framework)$"
        $normalizedName = $normalizedName -replace "(?i)\s+(freeware|shareware|community edition|professional|enterprise|standard|plus|viewer|reader|free)$"
        
        # Remove any non-alphanumeric characters (except spaces)
        $normalizedName = $normalizedName -replace "[^a-z0-9\s]", ""
        # Replace multiple spaces with a single space
        $normalizedName = $normalizedName -replace "\s+", " "
        
        $normalizedName = $normalizedName.Trim()
        Write-Verbose "Normalized result: '$normalizedName'"
        return $normalizedName
    }
}

# --- Core Data Retrieval Functions ---

<#
.SYNOPSIS
    Retrieves a list of installed software from a specified computer's registry.

.DESCRIPTION
    This function connects to the target computer and queries the standard
    Uninstall registry paths for both 32-bit and 64-bit applications.
    It returns key details such as DisplayName, DisplayVersion, and Publisher.
    Error handling is included for remote connection issues.

.PARAMETER ComputerName
    The name or IP address of the computer to query. Defaults to the local computer.

.RETURNS
    An array of objects, each representing an installed application.
.NOTES
    This method is preferred over Win32_Product for performance and to avoid triggering
    MSI repair actions.
#>
function Get-InstalledApp {
    [CmdletBinding(DefaultParameterSetName='Default', SupportsShouldProcess=$false, ConfirmImpact='None')]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ComputerName = $env:COMPUTERNAME
    )

    Write-Verbose "Attempting to retrieve installed software from registry on '${ComputerName}'."
    $installedSoftware = @()
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    try {
        foreach ($path in $regPaths) {
            Write-Verbose "Querying registry path: '$path' on '${ComputerName}'."
            # Use Invoke-Command for remote registry access
            $softwareEntries = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                param($regPath)
                Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -ne $null -and $_.SystemComponent -ne 1 -and $_.WindowsInstaller -ne 1 } |
                    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation
            } -ArgumentList $path

            if ($softwareEntries) {
                $installedSoftware += $softwareEntries
                Write-Verbose "Found $($softwareEntries.Count) entries from '$path'."
            } else {
                Write-Verbose "No entries found from '$path'."
            }
        }
        Write-Verbose "Successfully retrieved total installed software entries: $($installedSoftware.Count)."
    }
    catch {
        Write-Error "Failed to retrieve installed software on '${ComputerName}': ${_.Exception.Message}"
        Write-Verbose "Error details for Get-InstalledApp: $($_.Exception.ToString())"
        Write-Error "Ensure SCCM client is installed and WinRM is enabled on '${ComputerName}'."
        return $null
    }
    # Sort and remove duplicates based on DisplayName (case-insensitive for uniqueness)
    return $installedSoftware | Sort-Object DisplayName -Unique
}

<#
.SYNOPSIS
    Retrieves detailed status information about SCCM applications from the client.

.DESCRIPTION
    This function queries the CCM_Application WMI class on the target client to get
    comprehensive details about applications known to the SCCM client. It includes
    information such as deployment state, target type (machine/user), and availability for installation.
    It also converts WMI dates to readable DateTime objects and status codes to descriptive names.

.PARAMETER ComputerName
    The name or IP address of the computer to query. Defaults to the local computer.

.RETURNS
    An array of custom objects, each representing an SCCM application with its status.
#>
function Get-SCCMAppStatus {
    [CmdletBinding(DefaultParameterSetName='Default', SupportsShouldProcess=$false, ConfirmImpact='None')]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ComputerName = $env:COMPUTERNAME
    )

    Write-Verbose "Attempting to retrieve SCCM application status from '${ComputerName}'."
    $sccmApplications = @()
    try {
        $installStates = @{
            0 = "Unknown"; 1 = "Installed"; 2 = "Available"; 3 = "Downloading"; 4 = "Pending";
            5 = "Installing"; 6 = "Failed"; 7 = "Superseded"; 8 = "NoContent"; 9 = "NotApplicable";
            10 = "PendingRestart"; 11 = "PendingInstall"; 12 = "PendingUserLogoff"; 13 = "DownloadingContent";
            14 = "WaitingForContent"; 15 = "WaitingForServiceWindow"; 16 = "WaitingForDependencies";
            17 = "WaitingForRestart"; 18 = "WaitingForUserLogon"; 19 = "WaitingForSystemReset";
            20 = "WaitingForConfiguration"; 21 = "WaitingForCompletion"; 22 = "WaitingForRetry"
        }

        # Query CCM_Application class in root\ccm\ClientSDK namespace
        $rawSCCMApps = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-WmiObject -Namespace "root\ccm\ClientSDK" -Class CCM_Application -ErrorAction Stop |
                Select-Object Name, InstallState, IsMachineTarget, UserInstallAvailable, IsUpdate, Id, PolicyID,
                            InstallSource, LastUpdateTime, EstimatedInstallTime, Deadline, RequestedExecutionLevel
        }
        Write-Verbose "Retrieved $($rawSCCMApps.Count) raw SCCM application entries."

        foreach ($app in $rawSCCMApps) {
            $isInstalled = ($app.InstallState -eq 1)
            # IsUpdatePending: App is marked as an update AND is not currently installed.
            $isUpdatePending = ($app.IsUpdate -and -not $isInstalled)
            # IsPendingAction: App is in a state requiring action (not installed, unknown, or not applicable)
            $isPendingAction = ($app.InstallState -ne 1 -and $app.InstallState -ne 0 -and $app.InstallState -ne 9)

            $sccmApplications += [PSCustomObject]@{
                Name = $app.Name
                Status = $installStates[$app.InstallState]
                IsInstalled = $isInstalled
                IsUpdate = $app.IsUpdate
                IsUpdatePending = $isUpdatePending
                IsPendingAction = $isPendingAction
                IsMachineTarget = $app.IsMachineTarget
                UserInstallable = $app.UserInstallable
                AppId = $app.Id
                PolicyId = $app.PolicyID
                InstallSource = $app.InstallSource
                LastUpdated = if ($app.LastUpdateTime) { [DateTime]::ParseExact($app.LastUpdateTime.Substring(0,14), "yyyyMMddHHmmss", $null) } else { $null }
                EstimatedTime = $app.EstimatedInstallTime
                Deadline = if ($app.Deadline) { [DateTime]::ParseExact($app.Deadline.Substring(0,14), "yyyyMMddHHmmss", $null) } else { $null }
                ExecutionLevel = $app.RequestedExecutionLevel
            }
        }
        Write-Verbose "Processed $($sccmApplications.Count) SCCM application details."
    }
    catch {
        Write-Error "Failed to get SCCM application status on '${ComputerName}': ${_.Exception.Message}"
        Write-Verbose "Error details for Get-SCCMAppStatus: $($_.Exception.ToString())"
        Write-Error "Ensure SCCM client is installed and WinRM is enabled on '${ComputerName}'."
        return $null
    }
    return $sccmApplications | Sort-Object Name -Unique
}

<#
.SYNOPSIS
    Retrieves all application assignments (deployments) from the SCCM client.

.DESCRIPTION
    This function queries the SMS_ApplicationAssignment WMI class on the target client.
    These assignments represent the applications that the SCCM server has explicitly
    deployed or made available to this specific client. This is crucial for comparing
    "server intent" (as perceived by the client) with the actual installed software.

.PARAMETER ComputerName
    The name or IP address of the computer to query for assignments. Defaults to the local computer.

.RETURNS
    An array of custom objects, each representing an SCCM application assignment.
#>
function Get-SCCMAppAssignment {
    [CmdletBinding(DefaultParameterSetName='Default', SupportsShouldProcess=$false, ConfirmImpact='None')]
    param(
        [Parameter(Mandatory=$false)]
        [string]$ComputerName = $env:COMPUTERNAME
    )

    Write-Verbose "Attempting to retrieve SCCM application assignments from '${ComputerName}'."
    $assignments = @()
    try {
        $rawAssignments = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-WmiObject -Namespace "root\ccm\ClientSDK" -Class SMS_ApplicationAssignment -ErrorAction Stop |
                Select-Object ApplicationName, AssignmentName, DesiredState, CurrentState, LastEvaluationTime, IsEnabled
        }
        Write-Verbose "Retrieved $($rawAssignments.Count) raw SCCM application assignments."

        # Map DesiredState and CurrentState to readable names if needed
        # DesiredState: 0=Unknown, 1=Installed, 2=RequiresUninstall, 3=Available, 4=Mandatory
        # CurrentState: 0=Unknown, 1=Installed, 2=NotInstalled, 3=Pending, 4=Failed
        $stateMap = @{
            0 = "Unknown"; 1 = "Installed"; 2 = "RequiresUninstall"; 3 = "Available"; 4 = "Mandatory";
            "Current_0" = "Unknown"; "Current_1" = "Installed"; "Current_2" = "NotInstalled";
            "Current_3" = "Pending"; "Current_4" = "Failed"
        }

        foreach ($assignment in $rawAssignments) {
            $assignments += [PSCustomObject]@{
                ApplicationName = $assignment.ApplicationName
                AssignmentName = $assignment.AssignmentName
                DesiredState = $stateMap[$assignment.DesiredState]
                CurrentState = $stateMap["Current_$($assignment.CurrentState)"]
                LastEvaluationTime = if ($assignment.LastEvaluationTime) { [DateTime]::ParseExact($assignment.LastEvaluationTime.Substring(0,14), "yyyyMMddHHmmss", $null) } else { $null }
                IsEnabled = $assignment.IsEnabled
            }
        }
        Write-Verbose "Processed $($assignments.Count) SCCM application assignments."
    }
    catch {
        Write-Error "Failed to get SCCM application assignments on '${ComputerName}': ${_.Exception.Message}"
        Write-Verbose "Error details for Get-SCCMAppAssignment: $($_.Exception.ToString())"
        Write-Error "Ensure SCCM client is installed and WinRM is enabled on '${ComputerName}'."
        return $null
    }
    return $assignments | Sort-Object ApplicationName -Unique
}

# --- Action Functions ---

<#
.SYNOPSIS
    Triggers the installation of a specified SCCM application on a client.

.DESCRIPTION
    This function uses the SCCM Client COM object's ApplicationManagementInterface
    to initiate the installation of an application. It requires the application's
    ID and PolicyID, which can be obtained from Get-SCCMAppStatus.

.PARAMETER ComputerName
    The name or IP address of the computer where the installation should be triggered.

.PARAMETER AppId
    The unique ID of the SCCM application to install.

.PARAMETER PolicyId
    The policy ID associated with the application's deployment.

.NOTES
    This operation simulates clicking the "Install" button in the Software Center.
    Requires local administrator privileges on the target machine.
#>
function Invoke-SCCMAppInstall {
    [CmdletBinding(DefaultParameterSetName='Default', SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        [Parameter(Mandatory=$true)]
        [string]$PolicyId
    )

    if ($PSCmdlet.ShouldProcess("Application ID '$AppId' on '$ComputerName'", "Trigger SCCM Installation")) {
        Write-Verbose "Attempting to trigger installation for App ID '${AppId}' on '${ComputerName}'."
        try {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                param($appId, $policyId)
                $SMSClient = New-Object -ComObject 'Microsoft.SMS.Client'
                $AppMgmt = $SMSClient.GetClientSettings().GetApplicationManagementInterface()
                
                if ($AppMgmt) {
                    $AppMgmt.InstallApplication($appId, $policyId)
                    Write-Host "Installation for App ID '${appId}' successfully triggered." -ForegroundColor Green
                    Write-Verbose "SCCM InstallApplication method called for App ID '$appId'."
                } else {
                    Write-Error "Could not retrieve Application Management Interface. Installation not triggered."
                    Write-Verbose "Failed to get Application Management Interface on '$ComputerName'."
                }
            } -ArgumentList $AppId, $PolicyId
        }
        catch {
            Write-Error "Failed to trigger installation on '${ComputerName}': ${_.Exception.Message}"
            Write-Verbose "Error details for Invoke-SCCMAppInstall: $($_.Exception.ToString())"
            Write-Error "Ensure SCCM client is running and permissions are sufficient on '${ComputerName}'."
        }
    }
}

<#
.SYNOPSIS
    Triggers a specific SCCM client action cycle.

.DESCRIPTION
    This function allows initiating various SCCM client management cycles,
    such as Machine Policy Retrieval, Hardware Inventory, or Application Deployment Evaluation.
    It uses the `SMS_Client` WMI class and its `TriggerSchedule` method with the
    appropriate GUID for the desired cycle.

.PARAMETER ComputerName
    The name or IP address of the computer where the cycle should be triggered.

.PARAMETER CycleName
    The name of the SCCM client cycle to trigger.
    Valid options are defined in the global $SCCMCycleGUIDs dictionary.

.NOTES
    Requires local administrator privileges on the target machine.
#>
function Invoke-SCCMClientCycle {
    [CmdletBinding(DefaultParameterSetName='Default', SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ComputerName,
        
        [Parameter(Mandlandatory=$true)]
        [ValidateSet("MachinePolicy", "UserPolicy", "HardwareInventory", "SoftwareInventory", "AppDeploymentEval", "SoftwareUpdateScan", "DiscoveryData")]
        [string]$CycleName
    )

    $cycleGuid = $global:SCCMCycleGUIDs[$CycleName]

    if (-not $cycleGuid) {
        Write-Error "Unknown SCCM client cycle name: '$CycleName'. Valid options are: $($global:SCCMCycleGUIDs.Keys -join ', ')"
        Write-Verbose "Attempted to trigger unknown cycle: '$CycleName'."
        return
    }

    if ($PSCmdlet.ShouldProcess("SCCM '$CycleName' cycle on '$ComputerName'", "Trigger SCCM Cycle")) {
        Write-Verbose "Attempting to initiate SCCM '$CycleName' cycle (${cycleGuid}) on '${ComputerName}'."
        try {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                param($guid, $cycleNameForOutput)
                $SMSClient = New-Object -ComObject 'Microsoft.SMS.Client'
                $SMSClient.TriggerSchedule($guid)
                Write-Host "SCCM '$cycleNameForOutput' cycle successfully triggered." -ForegroundColor Green
                Write-Verbose "SCCM TriggerSchedule method called for GUID '$guid'."
            } -ArgumentList $cycleGuid, $CycleName
        }
        catch {
            Write-Error "Failed to trigger '$CycleName' cycle on '${ComputerName}': ${_.Exception.Message}"
            Write-Verbose "Error details for Invoke-SCCMClientCycle: $($_.Exception.ToString())"
            Write-Error "Ensure SCCM client is running and permissions are sufficient on '${ComputerName}'."
        }
    }
}

# --- Main Script Function (Cmdlet) ---

<#
.SYNOPSIS
    Manages SCCM applications on client computers.

.DESCRIPTION
    This is the main function of the script. It orchestrates the retrieval of
    installed software, SCCM application status, and SCCM assignments.
    It provides detailed difference analysis, allows triggering SCCM client cycles,
    and offers options for interactive or automatic application installations.

.PARAMETER ComputerName
    Specifies the target computer name or IP address. If not provided, it defaults
    to the local computer and will prompt interactively.

.PARAMETER InstallAll
    Automatically install all identified pending SCCM applications.

.PARAMETER PromptInstall
    Prompt for individual installation of pending SCCM applications.

.PARAMETER RefreshPolicy
    Trigger an immediate Machine Policy Retrieval & Evaluation Cycle on the client.

.PARAMETER ShowDifferences
    Show detailed differences between server-intended packages (client view) and installed status.

.PARAMETER AdditionalCycle
    Specify an additional SCCM client cycle to trigger. Valid options are:
    "MachinePolicy", "UserPolicy", "HardwareInventory", "SoftwareInventory",
    "AppDeploymentEval", "SoftwareUpdateScan", "DiscoveryData".

.PARAMETER DisableVerbose
    Disable verbose output logging. (Default: Verbose output is enabled)

.EXAMPLE
    Invoke-SCCMClientAppMgmt
    # Prompts for ComputerName, then installation choice, then individual app installs.

.EXAMPLE
    Invoke-SCCMClientAppMgmt -ComputerName "Server01" -ShowDifferences -PromptInstall
    # Checks status and differences on Server01, then prompts for each relevant app to install.

.EXAMPLE
    Invoke-SCCMClientAppMgmt -ComputerName "Client01" -RefreshPolicy -InstallAll
    # Refreshes policy on Client01, then automatically installs all pending apps.

.EXAMPLE
    Invoke-SCCMClientAppMgmt -ComputerName $env:COMPUTERNAME -ShowDifferences -PromptInstall:$false
    # Shows status and differences for the local PC, no prompts, verbose output disabled.

.EXAMPLE
    Invoke-SCCMClientAppMgmt -ComputerName "TestPC" -AdditionalCycle "HardwareInventory" -PromptInstall:$false
    # Triggers a Hardware Inventory Cycle on TestPC, then shows status.
#>
function Invoke-SCCMClientAppMgmt {
    [CmdletBinding(DefaultParameterSetName='StatusAndPrompt',
                   SupportsShouldProcess=$false, # ShouldProcess is handled in sub-functions
                   ConfirmImpact='Medium')]
    param(
        [Parameter(Mandatory=$false,
                   HelpMessage="Specify the computer name to manage. Defaults to local computer.")]
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter(ParameterSetName='InstallAll', Mandatory=$true,
                   HelpMessage="Automatically install all identified pending SCCM applications.")]
        [switch]$InstallAll,

        [Parameter(ParameterSetName='StatusAndPrompt', Mandatory=$false,
                   HelpMessage="Prompt for individual installation of pending SCCM applications.")]
        [switch]$PromptInstall, 

        [Parameter(Mandatory=$false,
                   HelpMessage="Trigger an immediate Machine Policy Retrieval & Evaluation Cycle on the client.")]
        [switch]$RefreshPolicy,

        [Parameter(Mandatory=$false,
                   HelpMessage="Show differences between server-intended packages (client view) and installed status.")]
        [switch]$ShowDifferences,
        
        [Parameter(Mandatory=$false,
                   HelpMessage="Specify an additional SCCM client cycle to trigger (e.g., 'HardwareInventory', 'AppDeploymentEval').")]
        [ValidateSet("MachinePolicy", "UserPolicy", "HardwareInventory", "SoftwareInventory", "AppDeploymentEval", "SoftwareUpdateScan", "DiscoveryData")]
        [string]$AdditionalCycle,

        [Parameter(Mandatory=$false,
                   HelpMessage="Disable verbose output logging.")]
        [switch]$DisableVerbose
    )

    # Enable verbose logging if not disabled
    if (-not $DisableVerbose) {
        $PSCmdlet.WriteVerbose("Verbose logging enabled.")
        $VerbosePreference = "Continue"
    } else {
        $PSCmdlet.WriteVerbose("Verbose logging disabled by parameter.")
        $VerbosePreference = "SilentlyContinue"
    }

    Write-Verbose "Starting Invoke-SCCMClientAppMgmt function."
    # Corrected: Use the parameter $ComputerName directly.
    Write-Verbose "Parameters received: ComputerName='$ComputerName', InstallAll=$InstallAll, PromptInstall=$PromptInstall, RefreshPolicy=$RefreshPolicy, ShowDifferences=$ShowDifferences, AdditionalCycle='$AdditionalCycle', DisableVerbose=$DisableVerbose."

    # --- Interactive Parameter Prompts ---
    if (-not $PSBoundParameters.ContainsKey('ComputerName')) {
        $ComputerName = Read-Host "Enter the computer name to manage (default: $env:COMPUTERNAME)"
        if ([string]::IsNullOrWhiteSpace($ComputerName)) { $ComputerName = $env:COMPUTERNAME }
        Write-Verbose "ComputerName set interactively to '$ComputerName'."
    }

    # Interactive choice for installation mode if not explicitly set
    if (-not $PSBoundParameters.ContainsKey('InstallAll') -and -not $PSBoundParameters.ContainsKey('PromptInstall')) {
        Write-Host ""
        $choice = Read-Host "Choose installation mode: (1) Prompt for each, (2) Install all, (3) Just show status [1]"
        switch ($choice) {
            "1" { $PromptInstall = $true; Write-Verbose "Installation mode: Prompt for each." }
            "2" { $InstallAll = $true; Write-Verbose "Installation mode: Install all." }
            "3" { $PromptInstall = $false; Write-Verbose "Installation mode: Just show status." }
            default { $PromptInstall = $true; Write-Verbose "Installation mode defaulted to: Prompt for each." }
        }
    }
    Write-Verbose "Final installation mode: InstallAll=$InstallAll, PromptInstall=$PromptInstall."
    # --- End Interactive Parameter Prompts ---

    Write-Host "`n--- Starting SCCM Client Application Management for ${ComputerName} ---" -ForegroundColor Green

    # --- 1. Display Server-Intended vs. Client-Actual Differences (before policy refresh) ---
    if ($ShowDifferences) {
        Write-Host "`n--- Current Policy Adherence & Software Installation Status (before sync) ---" -ForegroundColor Yellow
        Write-Host "This view compares applications assigned by the server (client's knowledge) with actual installations." -ForegroundColor Yellow
        Write-Verbose "Executing Get-SCCMAppAssignment for server intent comparison."
        $sccmAssignments = Get-SCCMAppAssignment -ComputerName $ComputerName
        Write-Verbose "Executing Get-InstalledApp for local installation comparison."
        $installedApps = Get-InstalledApp -ComputerName $ComputerName
        
        if (-not $sccmAssignments) {
            Write-Warning "Could not retrieve SCCM application assignments on '${ComputerName}'. Cannot perform server-client difference analysis."
        } elseif (-not $installedApps) {
            Write-Warning "Could not retrieve installed software on '${ComputerName}'. Cannot perform server-client difference analysis."
        } else {
            Write-Verbose "Starting server-client difference analysis."
            $assignedNormalizedNames = $sccmAssignments | Select-Object @{N='NName'; E={Convert-AppNameToNormalized $_.ApplicationName}}, ApplicationName, DesiredState, CurrentState
            $installedNormalizedNames = $installedApps | Select-Object @{N='NName'; E={Convert-AppNameToNormalized $_.DisplayName}}, DisplayName, DisplayVersion

            $diffs = @()

            # --- Compare: What is assigned by SCCM but not (yet) installed? ---
            Write-Verbose "Checking for SCCM assigned applications not found installed locally."
            foreach ($assignment in $assignedNormalizedNames) {
                $match = $installedNormalizedNames | Where-Object { $_.NName -eq $assignment.NName }
                
                # Condition 1: Assigned, but not found locally AND not explicitly installed according to assignment's CurrentState
                if (-not $match -and $assignment.CurrentState -ne "Installed") {
                    $diffs += [PSCustomObject]@{
                        Category = "SCCM Assigned - Not Installed"
                        Application = $assignment.ApplicationName
                        AssignedState = $assignment.DesiredState # Server's intent
                        ClientActual = $assignment.CurrentState  # Client's reported current state for assignment
                        InstalledVersion = "N/A"
                        Notes = "Assigned by server, but client's CurrentState is '$($assignment.CurrentState)' (not 'Installed')."
                    }
                    Write-Verbose "Difference found: SCCM assigned '$($assignment.ApplicationName)' (Desired: $($assignment.DesiredState), Current: $($assignment.CurrentState)) but no local installation match."
                } 
                # Condition 2: Assigned, found locally, but assignment's CurrentState is NOT 'Installed'
                elseif ($match -and $assignment.CurrentState -ne "Installed") {
                     $diffs += [PSCustomObject]@{
                        Category = "SCCM Assignment Status Discrepancy"
                        Application = $assignment.ApplicationName
                        AssignedState = $assignment.DesiredState
                        ClientActual = $assignment.CurrentState
                        InstalledVersion = $match.DisplayVersion
                        Notes = "Installed locally, but client's SCCM assignment status is still '$($assignment.CurrentState)' (not 'Installed')."
                    }
                    Write-Verbose "Discrepancy found: '$($assignment.ApplicationName)' installed locally, but SCCM CurrentState is '$($assignment.CurrentState)'."
                }
            }

            # --- Compare: What is installed but not (clearly) assigned by SCCM? ---
            # (Potential manual installs or old versions not managed by SCCM)
            Write-Verbose "Checking for locally installed applications not matching SCCM assignments."
            foreach ($installedApp in $installedNormalizedNames) {
                $match = $assignedNormalizedNames | Where-Object { $_.NName -eq $installedApp.NName }
                if (-not $match) {
                    $diffs += [PSCustomObject]@{
                        Category = "Locally Installed - Not SCCM Assigned"
                        Application = $installedApp.DisplayName
                        AssignedState = "N/A"
                        ClientActual = "Installed (Local)"
                        InstalledVersion = $installedApp.DisplayVersion
                        Notes = "Installed locally, but no matching SCCM assignment found."
                    }
                    Write-Verbose "Difference found: Installed '$($installedApp.DisplayName)' but no SCCM assignment match."
                }
            }

            if ($diffs.Count -gt 0) {
                Write-Host "`nDetailed Differences Found (Server Intent vs. Client Actual):" -ForegroundColor Red
                # Format output with German culture for dates/numbers if applicable - not directly for this table
                $diffs | Format-Table Category, Application, AssignedState, ClientActual, InstalledVersion, Notes -AutoSize
            } else {
                Write-Host "No significant differences found between SCCM assignments and local installations." -ForegroundColor Green
            }
        }
        Write-Host "--- End of Policy Adherence & Software Installation Status Display ---" -ForegroundColor Green
    }

    # --- 2. Trigger SCCM Client Cycles ---
    # Trigger Machine Policy Retrieval & Evaluation Cycle, if requested
    if ($RefreshPolicy) {
        Write-Verbose "Initiating RefreshPolicy: MachinePolicy cycle."
        Invoke-SCCMClientCycle -ComputerName $ComputerName -CycleName "MachinePolicy"
        Write-Host "Please wait a moment for the client to process the new policy..." -ForegroundColor DarkYellow
        Start-Sleep -Seconds 30
    }
    
    # Trigger additional specified client cycle, if requested
    if (-not [string]::IsNullOrWhiteSpace($AdditionalCycle)) {
        Write-Verbose "Initiating additional cycle: '$AdditionalCycle'."
        Invoke-SCCMClientCycle -ComputerName $ComputerName -CycleName $AdditionalCycle
        Write-Host "Please wait a moment for the client to process the '$AdditionalCycle' cycle..." -ForegroundColor DarkYellow
        Start-Sleep -Seconds 15
    }

    # --- 3. Retrieve Latest SCCM App Status and Offer Installation Options ---
    Write-Host "`n--- Current SCCM Application Status & Pending Actions (after sync) ---" -ForegroundColor Yellow
    $sccmAppsAfterRefresh = Get-SCCMAppStatus -ComputerName $ComputerName

    if (-not $sccmAppsAfterRefresh) {
        Write-Warning "Could not retrieve latest SCCM application details. Skipping installation options."
        return
    }

    $allRelevantSCCMApps = @()
    foreach ($app in $sccmAppsAfterRefresh) {
        # An application is relevant if it's pending action OR it's an update that's not installed
        if ($app.IsPendingAction -or $app.IsUpdatePending) {
            $allRelevantSCCMApps += $app
            Write-Verbose "Relevant SCCM app found: $($app.Name) (Status: $($app.Status), IsUpdatePending: $($app.IsUpdatePending))."
        }
    }

    if ($allRelevantSCCMApps.Count -gt 0) {
        Write-Host "Following SCCM applications require attention (installations/updates):" -ForegroundColor Red
        # Format output with German culture for dates/numbers
        $allRelevantSCCMApps | Format-Table Name, Status, IsUpdatePending, IsMachineTarget, UserInstallable, Deadline, EstimatedTime -AutoSize

        # --- Offer Installation Options ---
        # The interactive prompt logic should come here, *after* showing relevant apps.
        # This handles the case where -InstallAll or -PromptInstall was not specified,
        # or where the user chose '1' or '2' from the interactive prompt.
        if ($InstallAll) {
            Write-Host "`nAutomatically installing all pending/relevant applications..." -ForegroundColor Magenta
            foreach ($appToInstall in $allRelevantSCCMApps) {
                if ($appToInstall.AppId -and $appToInstall.PolicyId) {
                    Write-Host "Starting installation for: $($appToInstall.Name)" -ForegroundColor DarkYellow
                    Invoke-SCCMAppInstall -ComputerName $ComputerName -AppId $appToInstall.AppId -PolicyId $appToInstall.PolicyId
                } else {
                    Write-Warning "Application '${appToInstall.Name}' has no valid ID/PolicyId for installation. Skipping."
                    Write-Verbose "Skipping installation of '$($appToInstall.Name)' due to missing AppId/PolicyId."
                }
            }
        } elseif ($PromptInstall) {
            Write-Host "`nPrompting for individual pending/relevant applications..." -ForegroundColor Magenta
            foreach ($appToInstall in $allRelevantSCCMApps) {
                if ($appToInstall.AppId -and $appToInstall.PolicyId) {
                    $confirm = Read-Host "Install '${appToInstall.Name}' (Status: $($appToInstall.Status), Update Pending: $($appToInstall.IsUpdatePending))? (J/N)"
                    if ($confirm -eq "J") {
                        Write-Host "Starting installation for: $($appToInstall.Name)" -ForegroundColor DarkYellow
                        Invoke-SCCMAppInstall -ComputerName $ComputerName -AppId $appToInstall.AppId -PolicyId $appToInstall.PolicyId
                    } else {
                        Write-Host "Installation of '${appToInstall.Name}' skipped." -ForegroundColor Yellow
                        Write-Verbose "User chose to skip installation for '$($appToInstall.Name)'."
                    }
                } else {
                    Write-Warning "Application '${appToInstall.Name}' has no valid ID/PolicyId for installation. Skipping prompt for this."
                    Write-Verbose "Skipping prompt for '$($appToInstall.Name)' due to missing AppId/PolicyId."
                }
            }
        } else {
            Write-Host "`nNo installation option selected by user/parameters. To install, re-run with -InstallAll or -PromptInstall." -ForegroundColor Green
            Write-Verbose "No installation option selected by user/parameters."
        }
    } else {
        Write-Host "No pending/relevant installations or updates found in SCCM after sync." -ForegroundColor Green
        Write-Verbose "No relevant SCCM applications found to process for installation."
    }
    
    # --- 4. Final Detailed SCCM Application Status List ---
    Write-Host "`n--- Detailed list of all SCCM applications and their final status ---" -ForegroundColor DarkCyan
    # Format output with German culture for dates/numbers
    $sccmAppsAfterRefresh | Format-Table Name, Status, IsInstalled, IsUpdate, IsUpdatePending, IsPendingAction, IsMachineTarget, UserInstallable, InstallSource, Deadline, EstimatedTime, @{N='LastUpdated (DE)'; E={$_.LastUpdated.ToString($script:GermanCulture)}}, @{N='Deadline (DE)'; E={$_.Deadline.ToString($script:GermanCulture)}}, ExecutionLevel -AutoSize

    Write-Host "`n--- End of SCCM Client Application Management for ${ComputerName} ---" -ForegroundColor Green
    Write-Verbose "Invoke-SCCMClientAppMgmt function completed."
}

# --- Skriptaufruf ---
# WICHTIG: Um das Skript auszuführen, wähle EINE der folgenden Zeilen und entferne
# das Kommentarzeichen (#) davor. Passe den Computernamen bei Bedarf an.

# 1. Interaktiver Modus (Standard): Fragt nach Computername, dann nach Installationsoption.
#    Verbose ist standardmäßig AN.
# Invoke-SCCMClientAppMgmt

# 2. Status prüfen, Detaillierte Differenzen zeigen, dann für jede App fragen (Remote-PC):
#    Ersetze "L26PBZE70080022" durch den tatsächlichen Computernamen.
# Invoke-SCCMClientAppMgmt -ComputerName "L26PBZE70080022" -ShowDifferences -PromptInstall

# 3. Richtlinie aktualisieren und dann alle ausstehenden Apps automatisch installieren (Remote-PC):
#    Ersetze "L26PBZE70080022" durch den tatsächlichen Computernamen.
# Invoke-SCCMClientAppMgmt -ComputerName "L26PBZE70080022" -RefreshPolicy -InstallAll

# 4. Nur Status und detaillierte Differenzen anzeigen (Remote-PC), OHNE Installationsoptionen:
#    Hier wird PromptInstall explizit auf $false gesetzt, um die interaktive Abfrage zu überspringen.
#    Ersetze "L26PBZE70080022" durch den tatsächlichen Computernamen.
Invoke-SCCMClientAppMgmt -ComputerName "L26PBZE70080022" -ShowDifferences -PromptInstall:$false

# 5. Nur Status anzeigen für den lokalen PC, OHNE Installationsoptionen:
# Invoke-SCCMClientAppMgmt -ComputerName $env:COMPUTERNAME -PromptInstall:$false

# 6. Einen spezifischen zusätzlichen Client-Zyklus auslösen (z.B. Hardware-Inventur)
#    und dann den Status anzeigen (keine Installationsoptionen):
# Invoke-SCCMClientAppMgmt -ComputerName "L26PBZE70080022" -AdditionalCycle "HardwareInventory" -PromptInstall:$false

# 7. Nur Status anzeigen, Verbose Output deaktivieren:
# Invoke-SCCMClientAppMgmt -ComputerName "L26PBZE70080022" -PromptInstall:$false -DisableVerbose
