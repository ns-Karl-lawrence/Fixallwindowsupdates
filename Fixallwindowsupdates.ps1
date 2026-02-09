<#
.SYNOPSIS
    Repairs Windows Update issues and installs available updates.

.DESCRIPTION
    Comprehensive Windows Update repair and management script that:
    - Runs pre-flight checks (disk space, network, pending reboots, WUA health)
    - Creates system restore point
    - Repairs Windows Update components (services, folders, DLLs, DISM, SFC)
    - Scans for and installs available updates with retry logic
    - Provides detailed logging and recommendations

    The script is designed to run with Administrator privileges and handles
    both standard and aggressive repair scenarios.

.PARAMETER SkipFix
    Skip the repair phase and go directly to update scanning.

.PARAMETER SkipPreflightChecks
    Skip pre-flight checks (disk space, network, pending reboots, WUA health).

.PARAMETER SkipSystemRestore
    Skip creating a system restore point before repairs.

.PARAMETER AutoInstall
    Automatically install all available updates without user interaction. Default: $true

.PARAMETER ForceAutoInstall
    Force installation even if warnings exist (overrides safety prompts).

.PARAMETER AggressiveRepair
    Run aggressive repair including DISM /RestoreHealth and SFC /scannow.

.PARAMETER TestConnectivity
    Run extended Microsoft connectivity tests (Windows Update endpoints).

.PARAMETER MaxRetries
    Maximum retry attempts for failed updates. Default: 3

.PARAMETER LogPath
    Path to save the log file. Default: C:\Logs\WindowsUpdate_<timestamp>.log

.PARAMETER TranscriptPath
    Path to save full PowerShell transcript.

.PARAMETER RecommendationsCsvPath
    Path to save update recommendations CSV.

.EXAMPLE
    PS C:\> .\1.Fixallwindowsupdates.ps1
    Run with defaults - repairs WU and auto-installs updates.

.EXAMPLE
    PS C:\> .\1.Fixallwindowsupdates.ps1 -SkipFix -AutoInstall
    Skip repairs, just scan and install updates.

.EXAMPLE
    PS C:\> .\1.Fixallwindowsupdates.ps1 -AggressiveRepair -MaxRetries 5
    Run aggressive repair with 5 retry attempts.

.EXAMPLE
    PS C:\> .\1.Fixallwindowsupdates.ps1 -SkipSystemRestore -TestConnectivity
    Skip restore point and run connectivity tests.

.RISKS
    ┌──────────────────────────────┬──────────┬─────────────────────────────────────────────┐
    │ Operation                    │ Risk     │ Description                                 │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Service Stop/Start           │ Medium   │ Stops wuauserv, cryptsvc, bits, msiserver   │
    │                              │          │ - may interrupt downloads/installs          │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Folder Rename/Delete         │ High     │ Renames SoftwareDistribution & catroot2     │
    │                              │          │ - loses update history, may fail if stuck │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ DLL Re-registration          │ Low      │ Runs regsvr32 on WU-related DLLs            │
    │                              │          │ - generally safe, may fail if files locked  │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ DISM /RestoreHealth          │ Medium   │ Repairs component store from Windows Update │
    │                              │          │ - 30+ min runtime, requires internet        │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ SFC /scannow                 │ Low      │ Scans and repairs system files              │
    │                              │          │ - read-only scan, safe but time-consuming   │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Security Descriptor Reset    │ Medium   │ Resets BITS/WU service permissions          │
    │                              │          │ - restores defaults, may affect custom ACLs │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Take Ownership (Recursive)   │ High     │ takeown/icacls on system folders            │
    │                              │          │ - required for locked files, security risk  │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Registry Modifications       │ Medium   │ Deletes SusClientId, resets WSUS config     │
    │                              │          │ - forces re-registration with WU/WSUS       │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ BITS Queue Clear             │ Low      │ Removes pending BITS transfers              │
    │                              │          │ - may interrupt other application downloads │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ System Restore Point         │ Low      │ Creates restore point before changes        │
    │                              │          │ - requires System Protection enabled        │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Update Installation          │ Medium   │ Downloads and installs Windows Updates      │
    │                              │          │ - may require reboot, drivers may change    │
    ├──────────────────────────────┼──────────┼─────────────────────────────────────────────┤
    │ Auto-Install (No Prompt)     │ High     │ Installs all updates automatically          │
    │                              │          │ - default behavior, review updates first    │
    └──────────────────────────────┴──────────┴─────────────────────────────────────────────┘

.NOTES
    Version:        1.0
    Author:         Karl Lawrence
    Requirements:   PowerShell 5.1+, Administrator privileges
    Compatibility:  Windows 10, Windows 11, Windows Server 2016+

    WARNING: This script makes significant system changes. Review the RISKS table
    before running. Use -SkipSystemRestore only if you have other backups.

.LINK
    https://docs.microsoft.com/en-us/windows/deployment/update/windows-update-troubleshooting

#powershell -file "c:\temp\1.Fixallwindowsupdates.ps1" -MaxRetries 5 -SkipSystemRestore -AggressiveRepair -AutoInstall
#>

[CmdletBinding()]
param(
    [switch]$SkipFix,
    [switch]$SkipPreflightChecks,
    [switch]$SkipSystemRestore,
    [switch]$AutoInstall,
    [switch]$ForceAutoInstall,
    [switch]$AggressiveRepair,
    [switch]$TestConnectivity,
    [int]$MaxRetries = 3,
    [string]$LogPath = "C:\Logs\WindowsUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    [string]$TranscriptPath = "",
    [string]$RecommendationsCsvPath = "C:\Logs\WindowsUpdate_Recommendations_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Script-level variables for tracking
$script:FailedUpdates = @()
$script:SuccessfulUpdates = @()
$script:StartTime = Get-Date
$script:RebootRequired = $false
$script:RecommendationsCsvPath = $RecommendationsCsvPath
$script:CorrelationId = [System.Guid]::NewGuid().ToString("N")
$script:PhaseTimings = @{}

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug", "Verbose", "PhaseStart", "PhaseEnd")]
        [string]$Level = "Info",
        [string]$ErrorCode = "",
        [string]$PhaseName = "",
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorSuffix = if ($ErrorCode) { " [ErrorCode: $ErrorCode]" } else { "" }
    $phaseSuffix = if ($PhaseName) { " [Phase: $PhaseName]" } else { "" }
    $corrIdShort = $script:CorrelationId.Substring(0, 8)
    $logMessage = "[$timestamp] [$corrIdShort] [$Level]$phaseSuffix $Message$errorSuffix"
    
    # Console output with colors (unless suppressed)
    if (-not $NoConsole) {
        switch ($Level) {
            "Info"         { Write-Host $logMessage -ForegroundColor Cyan }
            "Warning"      { Write-Host $logMessage -ForegroundColor Yellow }
            "Error"        { Write-Host $logMessage -ForegroundColor Red }
            "Success"      { Write-Host $logMessage -ForegroundColor Green }
            "Debug"        { Write-Host $logMessage -ForegroundColor Magenta }
            "Verbose"      { Write-Host $logMessage -ForegroundColor Gray }
            "PhaseStart"   { Write-Host $logMessage -ForegroundColor White -BackgroundColor DarkGreen }
            "PhaseEnd"     { Write-Host $logMessage -ForegroundColor White -BackgroundColor DarkCyan }
        }
    }
    
    # File logging with rotation check
    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        # Log rotation - if file exceeds 10MB, rotate
        if (Test-Path $LogPath) {
            $logFile = Get-Item $LogPath -ErrorAction SilentlyContinue
            if ($logFile -and $logFile.Length -gt 10MB) {
                $rotatedPath = $LogPath -replace '\.log$', "_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                Move-Item -Path $LogPath -Destination $rotatedPath -Force -ErrorAction SilentlyContinue
            }
        }
        
        Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
        
        # Write to Windows Event Log for errors
        if ($Level -eq "Error") {
            try {
                $source = "WindowsUpdateScript"
                if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
                    [System.Diagnostics.EventLog]::CreateEventSource($source, "Application")
                }
                Write-EventLog -LogName "Application" -Source $source -EventId 1001 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
            }
            catch {
                # Silent fail for event log issues
            }
        }
    }
    catch {
        # Silent fail for logging issues
    }
}

function Start-Phase {
    param(
        [Parameter(Mandatory)]
        [string]$PhaseName,
        [string]$Description = ""
    )
    
    $script:PhaseTimings[$PhaseName] = @{
        StartTime = Get-Date
        Description = $Description
    }
    
    $msg = if ($Description) { "Starting: $Description" } else { "Starting phase: $PhaseName" }
    Write-Log -Message $msg -Level "PhaseStart" -PhaseName $PhaseName
}

function Complete-Phase {
    param(
        [Parameter(Mandatory)]
        [string]$PhaseName,
        [string]$Status = "Complete"
    )
    
    $phaseInfo = $script:PhaseTimings[$PhaseName]
    if ($phaseInfo) {
        $elapsed = (Get-Date) - $phaseInfo.StartTime
        $elapsedStr = "{0:mm\:ss\.fff}" -f $elapsed
        Write-Log -Message "Phase completed in $elapsedStr - Status: $Status" -Level "PhaseEnd" -PhaseName $PhaseName
        $script:PhaseTimings[$PhaseName].EndTime = Get-Date
        $script:PhaseTimings[$PhaseName].Elapsed = $elapsed
        $script:PhaseTimings[$PhaseName].Status = $Status
    }
    else {
        Write-Log -Message "Phase ended (timing not tracked) - Status: $Status" -Level "PhaseEnd" -PhaseName $PhaseName
    }
}

function New-WUComObject {
    param(
        [Parameter(Mandatory)]
        [string]$ProgID,
        [string]$Purpose = "COM operation"
    )
    
    try {
        $obj = New-Object -ComObject $ProgID -ErrorAction Stop
        return $obj
    }
    catch [System.Runtime.InteropServices.COMException] {
        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "COM Error creating '$ProgID' for $Purpose`: $_ (Error: 0x$($errorCode.ToString('X8')))" -Level Error
        
        # Provide specific guidance based on error
        if ($_.Exception.Message -match "80040154|class not registered") {
            Write-Log "Windows Update components not registered. Try running: Register-WindowsUpdateDLLs" -Level Warning
        }
        elseif ($_.Exception.Message -match "80070005|access denied") {
            Write-Log "Access denied creating COM object. Ensure running as Administrator." -Level Warning
        }
        elseif ($_.Exception.Message -match "80080005|server execution failed") {
            Write-Log "WUA service may be stopped or disabled. Try restarting services." -Level Warning
        }
        
        return $null
    }
    catch {
        Write-Log "Failed to create COM object '$ProgID' for $Purpose`: $_" -Level Error
        return $null
    }
}

# Windows Update Error Code Translation (Comprehensive)
$script:WUErrorCodes = @{
    # Windows Update Agent errors (0x8024xxxx)
    0x80240001 = "WU_E_NO_SERVICE - Windows Update Agent was unable to provide the service"
    0x80240002 = "WU_E_MAX_CAPACITY_REACHED - Maximum capacity of the service was exceeded"
    0x80240003 = "WU_E_UNKNOWN_ID - An ID cannot be found"
    0x80240004 = "WU_E_NOT_INITIALIZED - The object could not be initialized"
    0x80240005 = "WU_E_RANGEOVERLAP - The update handler requested a byte range overlapping a previously requested range"
    0x80240006 = "WU_E_TOOMANYRANGES - The requested number of byte ranges exceeds the maximum number"
    0x80240007 = "WU_E_INVALIDINDEX - The index to a collection was invalid"
    0x80240008 = "WU_E_ITEMNOTFOUND - The key for the item queried could not be found"
    0x80240009 = "WU_E_OPERATIONINPROGRESS - Another conflicting operation was in progress"
    0x8024000A = "WU_E_COULDNOTCANCEL - Cancellation of the operation was not allowed"
    0x8024000B = "WU_E_CALL_CANCELLED - Operation was cancelled"
    0x8024000C = "WU_E_NOOP - No operation was required"
    0x8024000D = "WU_E_XML_MISSINGDATA - Windows Update Agent could not find required information in the update's XML data"
    0x8024000E = "WU_E_XML_INVALID - Invalid XML found in the update's XML data"
    0x8024000F = "WU_E_CYCLE_DETECTED - Circular update relationships were detected in the metadata"
    0x80240010 = "WU_E_TOO_DEEP_RELATION - Update relationships too deep to evaluate"
    0x80240011 = "WU_E_INVALID_RELATIONSHIP - An invalid update relationship was detected"
    0x80240012 = "WU_E_REG_VALUE_INVALID - An invalid registry value was read"
    0x80240013 = "WU_E_DUPLICATE_ITEM - Operation tried to add a duplicate item to a list"
    0x80240014 = "WU_E_INVALID_INSTALL_REQUESTED - Cannot install exclusive update with other updates"
    0x80240015 = "WU_E_INSTALL_NOT_ALLOWED - Cannot install while other installations in progress"
    0x80240016 = "WU_E_INSTALL_NOT_ALLOWED - Operation tried to install while another installation was in progress"
    0x80240017 = "WU_E_NOT_APPLICABLE - Operation was not performed because there are no applicable updates"
    0x80240018 = "WU_E_NO_USERTOKEN - Operation failed because a required user token is missing"
    0x80240019 = "WU_E_EXCLUSIVE_INSTALL_CONFLICT - An exclusive update cannot be installed with other updates at the same time"
    0x8024001A = "WU_E_POLICY_NOT_SET - A policy value was not set"
    0x8024001B = "WU_E_SELFUPDATE_IN_PROGRESS - The operation could not be performed because the Windows Update Agent is self-updating"
    0x8024001D = "WU_E_INVALID_UPDATE - An update contains invalid metadata"
    0x8024001E = "WU_E_SERVICE_STOP - Operation did not complete because the service or system was being shut down"
    0x8024001F = "WU_E_NO_CONNECTION - No network connection was available"
    0x80240020 = "WU_E_NO_INTERACTIVE_USER - Operation did not complete because there is no logged-on interactive user"
    0x80240021 = "WU_E_TIME_OUT - Operation did not complete because it timed out"
    0x80240022 = "WU_E_ALL_UPDATES_FAILED - Operation failed for all the updates"
    0x80240023 = "WU_E_EULAS_DECLINED - The license terms for all updates were declined"
    0x80240024 = "WU_E_NO_UPDATE - There are no updates"
    0x80240025 = "WU_E_USER_ACCESS_DISABLED - Group Policy settings prevented access to Windows Update"
    0x80240026 = "WU_E_INVALID_UPDATE_TYPE - The type of update is invalid"
    0x80240027 = "WU_E_URL_TOO_LONG - The URL exceeded the maximum length"
    0x80240028 = "WU_E_UNINSTALL_NOT_ALLOWED - The update could not be uninstalled because the request did not originate from a WSUS server"
    0x80240029 = "WU_E_INVALID_PRODUCT_LICENSE - Search may have missed some updates before there is an unlicensed application on the system"
    0x8024002A = "WU_E_MISSING_HANDLER - A component required to detect applicable updates was missing"
    0x8024002B = "WU_E_LEGACYSERVER - An operation did not complete because it requires a newer version of server"
    0x8024002C = "WU_E_BIN_SOURCE_ABSENT - A delta-compressed update could not be installed because it required the source"
    0x8024002D = "WU_E_SOURCE_ABSENT - A full-file update could not be installed because it required the source"
    0x8024002E = "WU_E_WU_DISABLED - Access to an unmanaged server is not allowed"
    0x8024002F = "WU_E_CALL_CANCELLED_BY_POLICY - Operation did not complete because the DisableWindowsUpdateAccess policy was set"
    0x80240030 = "WU_E_INVALID_PROXY_SERVER - The format of the proxy list was invalid"
    0x80240031 = "WU_E_INVALID_FILE - The file is in the wrong format"
    0x80240032 = "WU_E_INVALID_CRITERIA - The search criteria string was invalid"
    0x80240033 = "WU_E_EULA_UNAVAILABLE - License terms could not be downloaded"
    0x80240034 = "WU_E_DOWNLOAD_FAILED - Update failed to download"
    0x80240035 = "WU_E_UPDATE_NOT_PROCESSED - The update was not processed"
    0x80240036 = "WU_E_INVALID_OPERATION - The object's current state did not allow the operation"
    0x80240037 = "WU_E_NOT_SUPPORTED - The functionality for the operation is not supported"
    0x80240038 = "WU_E_WINHTTP_INVALID_FILE - The downloaded file has an unexpected content type"
    0x80240039 = "WU_E_TOO_MANY_RESYNC - Agent is asked by server to resync too many times"
    0x80240040 = "WU_E_NO_SERVER_CORE_SUPPORT - WUA API method does not run on Server Core installation"
    0x80240041 = "WU_E_SYSPREP_IN_PROGRESS - Service is not available while sysprep is running"
    0x80240042 = "WU_E_UNKNOWN_SERVICE - The update service is no longer registered with AU"
    0x80240043 = "WU_E_NO_UI_SUPPORT - No support for WUA UI"
    0x80240044 = "WU_E_PER_MACHINE_UPDATE_ACCESS_DENIED - Only administrators can perform this operation on per-machine updates"
    0x80240045 = "WU_E_UNSUPPORTED_SEARCHSCOPE - A search was attempted with a scope that is not currently supported"
    0x80240046 = "WU_E_BAD_FILE_URL - The URL does not point to a file"
    0x80240047 = "WU_E_NOTSUPPORTED - The operation requested is not supported"
    0x80240048 = "WU_E_INVALID_NOTIFICATION_INFO - The featured update notification info returned by the server is invalid"
    0x80240049 = "WU_E_OUTOFRANGE - The data is out of range"
    0x8024004A = "WU_E_SETUP_IN_PROGRESS - Windows Update Agent could not be updated because setup is in progress"
    0x80240FFF = "WU_E_UNEXPECTED - An operation failed due to reasons not covered by another error code"
    
    # Download Manager errors (0x80246xxx)
    0x80246001 = "WU_E_DM_URLNOTAVAILABLE - A download manager operation could not be completed because the requested file does not have a URL"
    0x80246002 = "WU_E_DM_INCORRECTFILEHASH - A download manager operation could not be completed because the file digest was not recognized"
    0x80246003 = "WU_E_DM_UNKNOWNALGORITHM - A download manager operation could not be completed because the file metadata requested an unrecognized hash algorithm"
    0x80246004 = "WU_E_DM_NEEDDOWNLOADREQUEST - An operation could not be completed because a download request is required from the download handler"
    0x80246005 = "WU_E_DM_NONETWORK - A download manager operation could not be completed because the network connection was unavailable"
    0x80246006 = "WU_E_DM_WRONGBITSVERSION - A download manager operation could not be completed because the version of BITS is incompatible"
    0x80246007 = "WU_E_DM_NOTDOWNLOADED - The update has not been downloaded"
    0x80246008 = "WU_E_DM_FAILTOCONNECTTOBITS - A download manager operation failed because the download manager was unable to connect the BITS service"
    0x80246009 = "WU_E_DM_BITSTRANSFERERROR - A download manager operation failed because there was an unspecified BITS transfer error"
    0x8024600A = "WU_E_DM_DOWNLOADLOCATIONCHANGED - A download must be restarted because the location of the source of the download has changed"
    0x8024600B = "WU_E_DM_CONTENTCHANGED - A download must be restarted because the update content changed in a new revision"
    0x80246FFF = "WU_E_DM_UNEXPECTED - There was a download manager error not covered by another WU_E_DM_* error code"
    
    # Update Handler errors (0x80242xxx)
    0x80242000 = "WU_E_UH_REMOTEUNAVAILABLE - A request for a remote update handler could not be completed because no remote process is available"
    0x80242001 = "WU_E_UH_LOCALONLY - A request for a remote update handler could not be completed because the handler is local only"
    0x80242002 = "WU_E_UH_UNKNOWNHANDLER - A request for an update handler could not be completed because the handler could not be recognized"
    0x80242003 = "WU_E_UH_REMOTEALREADYACTIVE - A remote update handler could not be created because one already exists"
    0x80242004 = "WU_E_UH_DOESNOTSUPPORTACTION - A request for the handler to install (uninstall) an update could not be completed because the update does not support install (uninstall)"
    0x80242005 = "WU_E_UH_WRONGHANDLER - An operation did not complete because the wrong handler was specified"
    0x80242006 = "WU_E_UH_INVALIDMETADATA - A handler operation could not be completed because the update contains invalid metadata"
    0x80242007 = "WU_E_UH_INSTALLERHUNG - An operation could not be completed because the installer exceeded the time limit"
    0x80242008 = "WU_E_UH_OPERATIONCANCELLED - An operation being done by the update handler was cancelled"
    0x80242009 = "WU_E_UH_BADHANDLERXML - An operation could not be completed because the handler-specific metadata is invalid"
    0x8024200A = "WU_E_UH_CANREQUIREINPUT - A request to the handler to install an update could not be completed because the update requires user input"
    0x8024200B = "WU_E_UH_INSTALLERFAILURE - The installer failed to install (uninstall) one or more updates"
    0x8024200C = "WU_E_UH_FALLBACKTOSELFCONTAINED - The update handler should download self-contained content rather than delta-compressed content for the update"
    0x8024200D = "WU_E_UH_NEEDANOTHERDOWNLOAD - The update handler did not install the update because it needs to be downloaded again"
    0x8024200E = "WU_E_UH_NOTIFYFAILURE - The update handler failed to send notification of the status of the install (uninstall) operation"
    0x8024200F = "WU_E_UH_INCONSISTENT_FILE_NAMES - The file names contained in the update metadata and in the update package are inconsistent"
    0x80242010 = "WU_E_UH_FALLBACKERROR - The update handler failed to fall back to the self-contained content"
    0x80242011 = "WU_E_UH_TOOMANYDOWNLOADREQUESTS - The update handler has exceeded the maximum number of download requests"
    0x80242012 = "WU_E_UH_UNEXPECTEDCBSRESPONSE - The update handler has received an unexpected response from CBS"
    0x80242013 = "WU_E_UH_BADCBSPACKAGEID - The update metadata contains an invalid CBS package identifier"
    0x80242014 = "WU_E_UH_POSTREBOOTSTILLPENDING - The post-reboot operation for the update is still in progress"
    0x80242015 = "WU_E_UH_POSTREBOOTRESULTUNKNOWN - The result of the post-reboot operation for the update could not be determined"
    0x80242016 = "WU_E_UH_POSTREBOOTUNEXPECTEDSTATE - The state of the update after its post-reboot operation has completed is unexpected"
    0x80242017 = "WU_E_UH_NEW_SERVICING_STACK_REQUIRED - The OS servicing stack must be updated before this update is downloaded or installed"
    0x80242FFF = "WU_E_UH_UNEXPECTED - An update handler error not covered by another WU_E_UH_* code"
    
    # Data Store errors (0x80248xxx)
    0x80248000 = "WU_E_DS_SHUTDOWN - An operation failed because Windows Update Agent is shutting down"
    0x80248001 = "WU_E_DS_INUSE - An operation failed because the data store was in use"
    0x80248002 = "WU_E_DS_INVALID - The current and expected states of the data store do not match"
    0x80248003 = "WU_E_DS_TABLEMISSING - The data store is missing a table"
    0x80248004 = "WU_E_DS_TABLEINCORRECT - The data store contains a table with unexpected columns"
    0x80248005 = "WU_E_DS_INVALIDTABLENAME - A table could not be opened because the table is not in the data store"
    0x80248006 = "WU_E_DS_BADVERSION - The current and expected versions of the data store do not match"
    0x80248007 = "WU_E_DS_NODATA - The information requested is not in the data store"
    0x80248008 = "WU_E_DS_MISSINGDATA - The data store is missing required information or has a NULL in a table column that requires a non-null value"
    0x80248009 = "WU_E_DS_MISSINGREF - The data store is missing required information or has a reference to missing license terms, file, localized property or linked row"
    0x8024800A = "WU_E_DS_UNKNOWNHANDLER - The update was not processed because its update handler could not be recognized"
    0x8024800B = "WU_E_DS_CANTDELETE - The update was not deleted because it is still referenced by one or more services"
    0x8024800C = "WU_E_DS_LOCKTIMEOUTEXPIRED - The data store section could not be locked within the allotted time"
    0x8024800D = "WU_E_DS_NOCATEGORIES - The category was not added because it contains no parent categories and is not a top-level category itself"
    0x8024800E = "WU_E_DS_ROWEXISTS - The row was not added because an existing row has the same primary key"
    0x8024800F = "WU_E_DS_STOREFILELOCKED - The data store could not be initialized because it was locked by another process"
    0x80248010 = "WU_E_DS_CANNOTREGISTER - The data store is not allowed to be registered with COM in the current process"
    0x80248011 = "WU_E_DS_UNABLETOSTART - Could not create a data store object in another process"
    0x80248013 = "WU_E_DS_DUPLICATEUPDATEID - The server sent the same update to the client with two different revision IDs"
    0x80248014 = "WU_E_DS_UNKNOWNSERVICE - An operation did not complete because the service is not in the data store"
    0x80248015 = "WU_E_DS_SERVICEEXPIRED - An operation did not complete because the registration of the service has expired"
    0x80248016 = "WU_E_DS_DECLINENOTALLOWED - A request to hide an update was declined because it is a mandatory update or because it was deployed with a deadline"
    0x80248017 = "WU_E_DS_TABLESESSIONMISMATCH - A table was not closed because it is not associated with the session"
    0x80248018 = "WU_E_DS_SESSIONLOCKMISMATCH - A table was not closed because it is not associated with the session"
    0x80248019 = "WU_E_DS_NEEDWINDOWSSERVICE - A request to remove the Windows Update service or to unregister it with Automatic Updates was declined"
    0x8024801A = "WU_E_DS_INVALIDOPERATION - A request was declined because the operation is not allowed"
    0x8024801B = "WU_E_DS_SCHEMAMISMATCH - The schema of the current data store and the schema of a table in a backup XML document do not match"
    0x8024801C = "WU_E_DS_RESETREQUIRED - The data store requires a session reset; release the session and retry with a new session"
    0x8024801D = "WU_E_DS_IMPERSONATED - A data store operation did not complete because it was requested with an impersonated identity"
    0x80248FFF = "WU_E_DS_UNEXPECTED - A data store error not covered by another WU_E_DS_* code"
    
    # Inventory errors (0x80249xxx)
    0x80249001 = "WU_E_INVENTORY_PARSEFAILED - Parsing of the rule file failed"
    0x80249002 = "WU_E_INVENTORY_GET_INVENTORY_TYPE_FAILED - Failed to get the requested inventory type from the server"
    0x80249003 = "WU_E_INVENTORY_RESULT_UPLOAD_FAILED - Failed to upload inventory result to the server"
    0x80249004 = "WU_E_INVENTORY_UNEXPECTED - There was an inventory error not covered by another error code"
    0x80249005 = "WU_E_INVENTORY_WMI_ERROR - A WMI error occurred when enumerating the instances for a particular class"
    
    # AU (Automatic Updates) errors (0x8024Axxx)
    0x8024A000 = "WU_E_AU_NOSERVICE - Automatic Updates was unable to service incoming requests"
    0x8024A002 = "WU_E_AU_NONLEGACYSERVER - The old version of the Automatic Updates client has stopped because the WSUS server has been upgraded"
    0x8024A003 = "WU_E_AU_LEGACYCLIENTDISABLED - The old version of the Automatic Updates client was disabled"
    0x8024A004 = "WU_E_AU_PAUSED - Automatic Updates was unable to process incoming requests because it was paused"
    0x8024A005 = "WU_E_AU_NO_REGISTERED_SERVICE - No unmanaged service is registered with AU"
    0x8024A006 = "WU_E_AU_DETECT_SVCID_MISMATCH - The default service registered with AU changed during the search"
    0x8024A007 = "WU_E_REBOOT_IN_PROGRESS - A reboot is in progress"
    0x8024AFFF = "WU_E_AU_UNEXPECTED - An Automatic Updates error not covered by another WU_E_AU* code"
    
    # Reporter errors (0x8024Cxxx)
    0x8024C001 = "WU_E_DRV_PRUNED - A driver was skipped"
    0x8024C002 = "WU_E_DRV_NOPROP_OR_LEGACY - A property for the driver could not be found. It may not conform with required specifications"
    0x8024C003 = "WU_E_DRV_REG_MISMATCH - The registry type read for the driver does not match the expected type"
    0x8024C004 = "WU_E_DRV_NO_METADATA - The driver update is missing metadata"
    0x8024C005 = "WU_E_DRV_MISSING_ATTRIBUTE - The driver update is missing a required attribute"
    0x8024C006 = "WU_E_DRV_SYNC_FAILED - Driver synchronization failed"
    0x8024C007 = "WU_E_DRV_NO_PRINTER_CONTENT - Information required for the synchronization of applicable printers is missing"
    0x8024CFFF = "WU_E_DRV_UNEXPECTED - A driver error not covered by another WU_E_DRV_* code"
    
    # Windows Installer errors (0x8024Dxxx)
    0x8024D001 = "WU_E_SETUP_INVALID_INFDATA - Windows Update Agent could not be updated because an INF file contains invalid information"
    0x8024D002 = "WU_E_SETUP_INVALID_IDENTDATA - Windows Update Agent could not be updated because the wuident.cab file contains invalid information"
    0x8024D003 = "WU_E_SETUP_ALREADY_INITIALIZED - Windows Update Agent could not be updated because of an internal error that caused setup initialization to be performed twice"
    0x8024D004 = "WU_E_SETUP_NOT_INITIALIZED - Windows Update Agent could not be updated because setup initialization never completed successfully"
    0x8024D005 = "WU_E_SETUP_SOURCE_VERSION_MISMATCH - Windows Update Agent could not be updated because the versions specified in the INF do not match the actual source file versions"
    0x8024D006 = "WU_E_SETUP_TARGET_VERSION_GREATER - Windows Update Agent could not be updated because a WUA file on the target system is newer than the corresponding source file"
    0x8024D007 = "WU_E_SETUP_REGISTRATION_FAILED - Windows Update Agent could not be updated because regsvr32.exe returned an error"
    0x8024D008 = "WU_E_SELFUPDATE_SKIP_ON_FAILURE - An update to the Windows Update Agent was skipped because previous attempts to update have failed"
    0x8024D009 = "WU_E_SETUP_SKIP_UPDATE - An update to the Windows Update Agent was skipped due to a directive in the wuident.cab file"
    0x8024D00A = "WU_E_SETUP_UNSUPPORTED_CONFIGURATION - Windows Update Agent could not be updated because the current system configuration is not supported"
    0x8024D00B = "WU_E_SETUP_BLOCKED_CONFIGURATION - Windows Update Agent could not be updated because the system is configured to block the update"
    0x8024D00C = "WU_E_SETUP_REBOOT_TO_FIX - Windows Update Agent could not be updated because a restart of the system is required"
    0x8024D00D = "WU_E_SETUP_ALREADYRUNNING - Windows Update Agent setup is already running"
    0x8024D00E = "WU_E_SETUP_REBOOTREQUIRED - Windows Update Agent setup package requires a reboot to complete installation"
    0x8024D00F = "WU_E_SETUP_HANDLER_EXEC_FAILURE - Windows Update Agent could not be updated because the setup handler failed during execution"
    0x8024D010 = "WU_E_SETUP_INVALID_REGISTRY_DATA - Windows Update Agent could not be updated because the registry contains invalid information"
    0x8024D011 = "WU_E_SELFUPDATE_REQUIRED - Windows Update Agent must be updated before search can continue"
    0x8024D012 = "WU_E_SELFUPDATE_REQUIRED_ADMIN - Windows Update Agent must be updated before search can continue. An administrator is required to perform the operation"
    0x8024D013 = "WU_E_SETUP_WRONG_SERVER_VERSION - Windows Update Agent could not be updated because the server does not contain update information for this version"
    0x8024DFFF = "WU_E_SETUP_UNEXPECTED - Windows Update Agent could not be updated because of an error not covered by another WU_E_SETUP_* error code"
    
    # Expression Evaluator errors (0x8024Exxx)
    0x8024E001 = "WU_E_EE_UNKNOWN_EXPRESSION - An expression evaluator operation could not be completed because an expression was unrecognized"
    0x8024E002 = "WU_E_EE_INVALID_EXPRESSION - An expression evaluator operation could not be completed because an expression was invalid"
    0x8024E003 = "WU_E_EE_MISSING_METADATA - An expression evaluator operation could not be completed because an expression contains an incorrect number of metadata nodes"
    0x8024E004 = "WU_E_EE_INVALID_VERSION - An expression evaluator operation could not be completed because the version of the serialized expression data is invalid"
    0x8024E005 = "WU_E_EE_NOT_INITIALIZED - The expression evaluator could not be initialized"
    0x8024E006 = "WU_E_EE_INVALID_ATTRIBUTEDATA - An expression evaluator operation could not be completed because there was an invalid attribute"
    0x8024E007 = "WU_E_EE_CLUSTER_ERROR - An expression evaluator operation could not be completed because the cluster state of the computer could not be determined"
    0x8024EFFF = "WU_E_EE_UNEXPECTED - There was an expression evaluator error not covered by another WU_E_EE_* error code"
    
    # Common Windows/System errors
    0x80070005 = "E_ACCESSDENIED - Access denied. Run as Administrator"
    0x8007000E = "E_OUTOFMEMORY - Not enough memory to complete the operation"
    0x80070057 = "E_INVALIDARG - One or more arguments are invalid"
    0x80070070 = "ERROR_DISK_FULL - Not enough disk space to complete the operation"
    0x800700A1 = "ERROR_BAD_PATHNAME - The specified path is invalid"
    0x80070422 = "ERROR_SERVICE_DISABLED - The service cannot be started, either because it is disabled or no enabled devices are associated with it"
    0x80070426 = "ERROR_SERVICE_NOT_ACTIVE - The service has not been started"
    0x8007043B = "ERROR_SERVICE_NOT_IN_EXE - The executable program that this service is configured to run in does not implement the service"
    0x80070490 = "ERROR_NOT_FOUND - Element not found"
    0x800704C7 = "ERROR_CANCELLED - The operation was canceled by the user"
    0x800706BA = "RPC_S_SERVER_UNAVAILABLE - The RPC server is unavailable"
    0x800706BE = "RPC_S_CALL_FAILED - The remote procedure call failed"
    0x80071A90 = "ERROR_TRANSACTIONAL_CONFLICT - The function attempted to use a name that is reserved for use by another transaction"
    
    # Network/WinHTTP errors
    0x80072EE2 = "WININET_E_TIMEOUT - The operation timed out"
    0x80072EE7 = "WININET_E_NAME_NOT_RESOLVED - The server name or address could not be resolved"
    0x80072EFD = "WININET_E_CANNOT_CONNECT - Cannot connect to server"
    0x80072EFE = "WININET_E_CONNECTION_ABORTED - The connection with the server was terminated abnormally"
    0x80072F05 = "WININET_E_SEC_INVALID_CERT - SSL certificate is invalid"
    0x80072F06 = "WININET_E_SEC_CERT_DATE_INVALID - SSL certificate date is invalid"
    0x80072F07 = "WININET_E_SEC_CERT_CN_INVALID - SSL certificate common name is incorrect"
    0x80072F0C = "WININET_E_SEC_CERT_ERRORS - SSL certificate contains errors"
    0x80072F0D = "WININET_E_SEC_CERT_REV_FAILED - SSL certificate revocation check failed"
    0x80072F8F = "WININET_E_DECODING_FAILED - SSL certificate error"
    
    # Certificate errors
    0x800B0100 = "TRUST_E_NOSIGNATURE - No signature was present in the subject"
    0x800B0101 = "CERT_E_EXPIRED - A required certificate is not within its validity period"
    0x800B0109 = "CERT_E_UNTRUSTEDROOT - A certificate chain processed, but terminated in a root certificate which is not trusted"
    0x800B010A = "CERT_E_UNTRUSTEDTESTROOT - The certification path terminates with the test root which is not trusted"
    0x800B010B = "CERT_E_CHAINING - A certificate chain could not be built to a trusted root authority"
    0x800B010E = "CERT_E_REVOCATION_FAILURE - The revocation process could not continue, certificate revocation check failed"
    
    # Component Store/CBS errors
    0x80073701 = "ERROR_SXS_ASSEMBLY_MISSING - The referenced assembly is not installed on your system"
    0x80073702 = "ERROR_SXS_CANT_GEN_ACTCTX - The referenced assembly manifest could not be found"
    0x80073712 = "ERROR_SXS_COMPONENT_STORE_CORRUPT - The component store is in an inconsistent state"
    0x80073713 = "ERROR_SXS_FILE_HASH_MISMATCH - A component's file does not match the verification information present in the component manifest"
    0x80073D01 = "ERROR_DEPLOYMENT_BLOCKED - The application cannot be installed because it has been blocked"
    
    # CBS (Component-Based Servicing) errors
    0x800F0801 = "CBS_E_NOT_INITIALIZED - CBS session not initialized"
    0x800F0802 = "CBS_E_ALREADY_INITIALIZED - CBS session already initialized"
    0x800F0803 = "CBS_E_INVALID_PACKAGE - The update package is not valid"
    0x800F0805 = "CBS_E_INVALID_PACKAGE_FORMAT - The update package format is invalid"
    0x800F0806 = "CBS_E_INVALID_INSTALL_STATE - Cannot install to the specified install state"
    0x800F0816 = "CBS_E_PENDING - A system restart is required to complete installation"
    0x800F0818 = "CBS_E_IMAGE_NOT_SERVICED - This package cannot be installed to the specified image"
    0x800F081E = "CBS_E_NOT_APPLICABLE_STATE - The package is not applicable to this image because of the current state"
    0x800F081F = "CBS_E_SOURCE_MISSING - The source for the package or file not found"
    0x800F0820 = "CBS_E_CANCEL - The operation was cancelled"
    0x800F0821 = "CBS_E_ABORT - The operation was aborted"
    0x800F0822 = "CBS_E_ILLEGAL_OPERATION - Illegal operation attempted"
    0x800F0825 = "CBS_E_INVALID_SELF_UPDATE - Cannot perform self-update in this state"
    0x800F0826 = "CBS_E_CORRUPTED_IMAGE - The image is corrupted"
    0x800F0900 = "CBS_E_XML_PARSER_FAILURE - Unexpected XML parser failure"
    0x800F0902 = "CBS_E_INVALID_MANIFEST - The update manifest is invalid"
    0x800F0905 = "CBS_E_MISSING_MANIFEST - The update manifest is missing"
    0x800F0906 = "CBS_E_DOWNLOAD_FAILURE - The source files could not be downloaded"
    0x800F0907 = "CBS_E_GROUPPOLICY_DISALLOWED - WSUS server group policy does not permit servicing"
    0x800F0908 = "CBS_E_MISSING_REQUIRED - Some required update components are missing"
    0x800F0922 = "CBS_E_INSTALLERS_FAILED - Processing advanced installers and generic commands failed"
    0x800F0923 = "CBS_E_NOT_ENOUGH_DISK_SPACE - Not enough disk space to install package"
    0x800F0950 = "CBS_E_NOT_APPLICABLE - The package is not applicable to this image"
    0x800F0954 = "CBS_E_REBOOT_PENDING - A reboot is required to complete the operation"
    0x800F0955 = "CBS_E_STORE_CORRUPTION - The component store is corrupt"
    0x800F0982 = "CBS_E_STORE_STACK_OVERFLOW - The servicing stack encountered a serious error"
    0x800F0984 = "CBS_E_STORE_MERGE_FAILURE - CBS store failed to merge during session finalize"
    0x800F0988 = "CBS_E_INSUFFICIENT_SYS_RESOURCES - Insufficient system resources to complete the operation"
    
    # BITS errors
    0x80190190 = "BG_E_HTTP_ERROR_400 - Bad Request"
    0x80190191 = "BG_E_HTTP_ERROR_401 - Unauthorized"
    0x80190193 = "BG_E_HTTP_ERROR_403 - Forbidden"
    0x80190194 = "BG_E_HTTP_ERROR_404 - Not Found"
    0x80190197 = "BG_E_HTTP_ERROR_407 - Proxy Authentication Required"
    0x801901F4 = "BG_E_HTTP_ERROR_500 - Internal Server Error"
    0x801901F7 = "BG_E_HTTP_ERROR_503 - Service Unavailable"
}

function Get-WUErrorDescription {
    param([int64]$ErrorCode)
    
    $hexCode = "0x{0:X8}" -f $ErrorCode
    if ($script:WUErrorCodes.ContainsKey($ErrorCode)) {
        return "$hexCode - $($script:WUErrorCodes[$ErrorCode])"
    }
    return "$hexCode - Unknown error code. Check CBS.log for details."
}

function Get-SystemLogIssues {
    param(
        [int]$HoursBack = 24,
        [int]$MaxEntries = 100
    )
    
    Write-Log "========================================" -Level Info
    Write-Log "Scanning System Logs for Issues (Last $HoursBack hours)" -Level Info
    Write-Log "========================================" -Level Info
    
    $issues = @{
        CBS = @()
        DISM = @()
        WindowsUpdate = @()
        System = @()
        Application = @()
    }
    
    $cutoffTime = (Get-Date).AddHours(-$HoursBack)
    
    # 1. Parse CBS.log for errors
    $cbsLogPath = "$env:SystemRoot\Logs\CBS\CBS.log"
    if (Test-Path $cbsLogPath) {
        Write-Log "Parsing CBS.log..." -Level Info
        try {
            $cbsContent = Get-Content $cbsLogPath -Tail 5000 -ErrorAction SilentlyContinue
            $errorPatterns = @(
                'Error\s+',
                'FAIL',
                'failed',
                'ERROR_',
                '0x80[0-9A-Fa-f]{6}',
                'STATUS_',
                'corrupt',
                'missing',
                'not found',
                'access denied',
                'cannot',
                'unable to'
            )
            $combinedPattern = ($errorPatterns -join '|')
            
            foreach ($line in $cbsContent) {
                if ($line -match $combinedPattern -and $line -notmatch 'Success|Successful|succeeded') {
                    # Extract timestamp if present
                    if ($line -match '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}') {
                        $issues.CBS += $line.Trim()
                    }
                    elseif ($line -match 'Error|FAIL|0x80') {
                        $issues.CBS += $line.Trim()
                    }
                }
            }
            
            # Deduplicate and limit
            $issues.CBS = $issues.CBS | Select-Object -Unique | Select-Object -Last $MaxEntries
            
            if ($issues.CBS.Count -gt 0) {
                Write-Log "Found $($issues.CBS.Count) CBS log issues" -Level Warning
                foreach ($issue in ($issues.CBS | Select-Object -Last 10)) {
                    Write-Log "  CBS: $issue" -Level Warning -NoConsole
                }
            }
            else {
                Write-Log "No CBS errors found in recent log entries" -Level Success
            }
        }
        catch {
            Write-Log "Failed to parse CBS.log: $_" -Level Warning
        }
    }
    else {
        Write-Log "CBS.log not found at $cbsLogPath" -Level Warning
    }
    
    # 2. Parse DISM.log for errors
    $dismLogPath = "$env:SystemRoot\Logs\DISM\dism.log"
    if (Test-Path $dismLogPath) {
        Write-Log "Parsing DISM.log..." -Level Info
        try {
            $dismContent = Get-Content $dismLogPath -Tail 2000 -ErrorAction SilentlyContinue
            foreach ($line in $dismContent) {
                if ($line -match 'Error|FAIL|failed|0x80[0-9A-Fa-f]{6}' -and $line -notmatch 'Success') {
                    $issues.DISM += $line.Trim()
                }
            }
            $issues.DISM = $issues.DISM | Select-Object -Unique | Select-Object -Last $MaxEntries
            
            if ($issues.DISM.Count -gt 0) {
                Write-Log "Found $($issues.DISM.Count) DISM log issues" -Level Warning
                foreach ($issue in ($issues.DISM | Select-Object -Last 5)) {
                    Write-Log "  DISM: $issue" -Level Warning -NoConsole
                }
            }
            else {
                Write-Log "No DISM errors found in recent log entries" -Level Success
            }
        }
        catch {
            Write-Log "Failed to parse DISM.log: $_" -Level Warning
        }
    }
    
    # 3. Parse WindowsUpdate.log (Windows 10+)
    $wuLogPath = "$env:SystemRoot\Logs\WindowsUpdate\WindowsUpdate.log"
    if (-not (Test-Path $wuLogPath)) {
        # Try to generate it on Windows 10+
        try {
            Write-Log "Generating WindowsUpdate.log..." -Level Info
            $null = Get-WindowsUpdateLog -LogPath "$env:TEMP\WindowsUpdate.log" -ErrorAction SilentlyContinue
            $wuLogPath = "$env:TEMP\WindowsUpdate.log"
        }
        catch {
            # Fallback to legacy location
            $wuLogPath = "$env:SystemRoot\WindowsUpdate.log"
        }
    }
    
    if (Test-Path $wuLogPath) {
        Write-Log "Parsing WindowsUpdate.log..." -Level Info
        try {
            $wuContent = Get-Content $wuLogPath -Tail 2000 -ErrorAction SilentlyContinue
            foreach ($line in $wuContent) {
                if ($line -match 'Error|FAIL|failed|WARNING|0x80[0-9A-Fa-f]{6}' -and $line -notmatch 'Success') {
                    $issues.WindowsUpdate += $line.Trim()
                }
            }
            $issues.WindowsUpdate = $issues.WindowsUpdate | Select-Object -Unique | Select-Object -Last $MaxEntries
            
            if ($issues.WindowsUpdate.Count -gt 0) {
                Write-Log "Found $($issues.WindowsUpdate.Count) WindowsUpdate log issues" -Level Warning
            }
            else {
                Write-Log "No WindowsUpdate errors found in recent log entries" -Level Success
            }
        }
        catch {
            Write-Log "Failed to parse WindowsUpdate.log: $_" -Level Warning
        }
    }
    
    # 4. Check Windows Event Logs for critical system issues
    Write-Log "Checking Windows Event Logs..." -Level Info
    try {
        # System Event Log - Critical and Error events
        $systemEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 1,2  # Critical, Error
            StartTime = $cutoffTime
        } -MaxEvents 50 -ErrorAction SilentlyContinue
        
        foreach ($evt in $systemEvents) {
            $issues.System += "[EventID:$($evt.Id)] [$($evt.TimeCreated)] $($evt.Message -replace '\r?\n', ' ' | Select-Object -First 200)"
        }
        
        if ($issues.System.Count -gt 0) {
            Write-Log "Found $($issues.System.Count) System event log errors" -Level Warning
            # Log notable ones
            $notableEvents = $systemEvents | Where-Object { 
                $_.ProviderName -match 'Microsoft-Windows-WindowsUpdateClient|DCOM|Service Control Manager|Kernel|WHEA|Ntfs|Disk|volsnap' 
            } | Select-Object -First 10
            foreach ($evt in $notableEvents) {
                $shortMsg = ($evt.Message -split '\r?\n')[0]
                if ($shortMsg.Length -gt 150) { $shortMsg = $shortMsg.Substring(0, 150) + "..." }
                Write-Log "  System [$($evt.ProviderName)]: $shortMsg" -Level Warning -NoConsole
            }
        }
        else {
            Write-Log "No critical System events in the last $HoursBack hours" -Level Success
        }
    }
    catch {
        Write-Log "Failed to query System event log: $_" -Level Warning
    }
    
    try {
        # Application Event Log - Critical and Error events
        $appEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            Level = 1,2  # Critical, Error
            StartTime = $cutoffTime
        } -MaxEvents 50 -ErrorAction SilentlyContinue
        
        foreach ($evt in $appEvents) {
            $issues.Application += "[EventID:$($evt.Id)] [$($evt.TimeCreated)] $($evt.Message -replace '\r?\n', ' ' | Select-Object -First 200)"
        }
        
        if ($issues.Application.Count -gt 0) {
            Write-Log "Found $($issues.Application.Count) Application event log errors" -Level Warning
            # Log notable application errors
            $notableAppEvents = $appEvents | Where-Object { 
                $_.ProviderName -match 'Application Error|Windows Error Reporting|.NET Runtime|VSS|MSIINSTALLER|SideBySide' 
            } | Select-Object -First 10
            foreach ($evt in $notableAppEvents) {
                $shortMsg = ($evt.Message -split '\r?\n')[0]
                if ($shortMsg.Length -gt 150) { $shortMsg = $shortMsg.Substring(0, 150) + "..." }
                Write-Log "  App [$($evt.ProviderName)]: $shortMsg" -Level Warning -NoConsole
            }
        }
        else {
            Write-Log "No critical Application events in the last $HoursBack hours" -Level Success
        }
    }
    catch {
        Write-Log "Failed to query Application event log: $_" -Level Warning
    }
    
    # 5. Check for specific problematic patterns and provide recommendations
    Write-Log "Analyzing issues for recommendations..." -Level Info
    
    $recommendations = @()
    
    # Helper function to find matching entries and extract context
    function Find-MatchingEntries {
        param(
            [array]$Entries,
            [string]$Pattern,
            [int]$MaxResults = 5
        )
        $matched = @()
        foreach ($entry in $Entries) {
            if ($entry -match $Pattern) {
                $matched += $entry
                if ($matched.Count -ge $MaxResults) { break }
            }
        }
        return $matched
    }
    
    # Helper function to extract file paths from text
    function Get-FilePathsFromText {
        param([string]$Text)
        $paths = @()
        # Match Windows file paths
        $pathMatches = [regex]::Matches($Text, '[A-Za-z]:\\[^\s\,\;\"\<\>\|]+\.[a-zA-Z0-9]+')
        foreach ($match in $pathMatches) {
            $paths += $match.Value
        }
        # Also match paths starting with \Windows or similar
        $pathMatches2 = [regex]::Matches($Text, '\\Windows\\[^\s\,\;\"\<\>\|]+')
        foreach ($match in $pathMatches2) {
            $paths += $match.Value
        }
        return ($paths | Select-Object -Unique)
    }
    
    # Helper function to extract error codes from text
    function Get-ErrorCodesFromText {
        param([string]$Text)
        $codes = @()
        $codeMatches = [regex]::Matches($Text, '0x[0-9A-Fa-f]{8}|0x[0-9A-Fa-f]{4}')
        foreach ($match in $codeMatches) {
            $codes += $match.Value.ToUpper()
        }
        return ($codes | Select-Object -Unique)
    }
    
    # Helper function to add recommendation with context
    function Add-RecommendationWithContext {
        param(
            [string]$Title,
            [string]$Description,
            [array]$MatchedEntries,
            [ref]$RecommendationsList
        )
        
        $rec = @{
            Title = $Title
            Description = $Description
            Files = @()
            ErrorCodes = @()
            Context = @()
        }
        
        foreach ($entry in $MatchedEntries) {
            $rec.Files += Get-FilePathsFromText -Text $entry
            $rec.ErrorCodes += Get-ErrorCodesFromText -Text $entry
            # Truncate long entries for context
            $truncated = if ($entry.Length -gt 200) { $entry.Substring(0, 200) + "..." } else { $entry }
            $rec.Context += $truncated
        }
        
        $rec.Files = $rec.Files | Select-Object -Unique | Select-Object -First 5
        $rec.ErrorCodes = $rec.ErrorCodes | Select-Object -Unique | Select-Object -First 5
        $rec.Context = $rec.Context | Select-Object -Unique | Select-Object -First 3
        
        $RecommendationsList.Value += $rec
    }
    
    # Store all issue arrays for searching
    $allCBSDISM = $issues.CBS + $issues.DISM
    $allWU = $issues.WindowsUpdate
    $allSystem = $issues.System
    $allApp = $issues.Application
    $allEntries = $allCBSDISM + $allWU + $allSystem + $allApp
    
    # Combine all text-based issues for quick pattern matching
    $allIssues = ($issues.CBS + $issues.DISM) -join " "
    $allWUIssues = $issues.WindowsUpdate -join " "
    $allSystemIssues = $issues.System -join " "
    $allAppIssues = $issues.Application -join " "
    $combinedIssues = "$allIssues $allWUIssues $allSystemIssues $allAppIssues"
    
    # ============================================
    # COMPONENT STORE / CBS ISSUES
    # ============================================
    
    $pattern = 'component store|corrupt|0x80073712|0x800F0955|store corruption'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "COMPONENT STORE CORRUPTION" -Description "Run 'DISM /Online /Cleanup-Image /RestoreHealth' then 'sfc /scannow'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800F081F|CBS_E_SOURCE_MISSING|source.*missing|payload.*missing'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "MISSING SOURCE FILES" -Description "Mount Windows ISO and run 'DISM /Online /Cleanup-Image /RestoreHealth /Source:D:\sources\install.wim'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'manifest|0x800F0902|0x800F0905|invalid manifest'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "MANIFEST CORRUPTION" -Description "CBS manifests are damaged - Run DISM RestoreHealth with Windows installation media" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'servicing stack|0x800F0982|stack overflow'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "SERVICING STACK ERROR" -Description "Install the latest Servicing Stack Update (SSU) manually from Microsoft Update Catalog" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'merge failure|0x800F0984|session finalize'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "CBS MERGE FAILURE" -Description "Component store failed to merge - May require in-place upgrade repair" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800F0988|insufficient.*resources|system resources'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "INSUFFICIENT RESOURCES" -Description "Close applications, increase virtual memory, or add RAM" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # PENDING OPERATIONS / REBOOT ISSUES
    # ============================================
    
    $pattern = 'pending|0x800F0954|0x800F0816|reboot.*pending|restart.*required'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "PENDING REBOOT" -Description "System has pending operations - Reboot and run script again" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'pending\.xml|poqexec|TrustedInstaller.*pending'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "PENDING TRANSACTIONS" -Description "Incomplete servicing transactions - Reboot in Safe Mode, then normal mode" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'sessions\.xml|0x800F0922|advanced installers'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "INSTALLER FAILURE" -Description "Advanced installers failed - Check for pending .NET or VC++ installations" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # DISK / STORAGE ISSUES
    # ============================================
    
    $pattern = 'disk.*full|0x80070070|0x800F0923|not enough.*space|insufficient.*disk'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "DISK SPACE" -Description "Free up space - Run Disk Cleanup, clear Windows Update cache (C:\Windows\SoftwareDistribution)" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'Ntfs|chkdsk|file system|0x80070570|corrupt.*file'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "FILE SYSTEM ERRORS" -Description "Run 'chkdsk C: /scan' then 'chkdsk C: /spotfix' (requires reboot)" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'Disk.*error|bad.*sector|I/O.*error|0x8007045D|read.*fault'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "DISK I/O ERRORS" -Description "Physical disk may be failing - Run 'wmic diskdrive get status' and check SMART data" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'volsnap|shadow.*copy|VSS|0x8004230F'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "VSS/SHADOW COPY" -Description "Volume Shadow Copy issues - Run 'vssadmin list writers' and restart VSS service" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'WHEA|hardware.*error|machine.*check|0x124'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "HARDWARE ERROR (WHEA)" -Description "CPU/Memory hardware errors detected - Check RAM with memtest86, check CPU temps" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # NETWORK / CONNECTIVITY ISSUES
    # ============================================
    
    $pattern = '0x80072EE7|name.*not.*resolved|DNS|could not.*resolve'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "DNS RESOLUTION" -Description "Cannot resolve server names - Run 'ipconfig /flushdns' and check DNS settings" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80072EFD|cannot.*connect|connection.*refused|0x80072EE2|timeout'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "CONNECTION FAILED" -Description "Cannot connect to update servers - Check firewall, proxy settings, and internet connectivity" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80072F8F|SSL|certificate.*error|secure.*channel|TLS'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "SSL/TLS ERROR" -Description "Certificate validation failed - Check system date/time, update root certificates" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'proxy|0x80240030|407|authentication.*required'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "PROXY ISSUES" -Description "Proxy authentication or configuration problem - Check 'netsh winhttp show proxy'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'BITS.*error|0x80200010|background.*transfer|BG_E'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "BITS ERROR" -Description "Background transfer failed - Run 'bitsadmin /reset /allusers' and restart BITS service" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'firewall|blocked|0x80070422|access.*denied.*network'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "FIREWALL/NETWORK BLOCK" -Description "Check Windows Firewall rules and third-party security software" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # CERTIFICATE / TRUST ISSUES
    # ============================================
    
    $pattern = '0x800B0109|untrusted.*root|root.*certificate'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "UNTRUSTED ROOT CERT" -Description "Update root certificates - Run 'certutil -generateSSTFromWU roots.sst' then import" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800B0101|certificate.*expired|validity.*period'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "EXPIRED CERTIFICATE" -Description "A certificate has expired - Check system date/time, update root certs" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800B010B|certificate.*chain|chaining'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "CERTIFICATE CHAIN" -Description "Cannot build certificate chain - May need intermediate certificates" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800B010E|revocation|CRL|OCSP'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "CERT REVOCATION CHECK FAILED" -Description "Cannot verify certificate revocation - Check internet connectivity" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x800B0100|no.*signature|TRUST_E_NOSIGNATURE'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "MISSING SIGNATURE" -Description "File signature missing or invalid - File may be corrupted or tampered" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # SERVICE ISSUES
    # ============================================
    
    $pattern = 'Service Control Manager.*failed|service.*failed.*start|0x80070422'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "SERVICE FAILURE" -Description "Critical service failed to start - Check Services.msc for disabled/failed services" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'wuauserv|Windows Update.*service|0x80070426'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "WU SERVICE ERROR" -Description "Windows Update service issue - Run 'sc config wuauserv start=demand' then 'net start wuauserv'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'TrustedInstaller|0x80070005.*TrustedInstaller'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "TRUSTEDINSTALLER ERROR" -Description "Restart 'Windows Modules Installer' service or run from Safe Mode" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'cryptsvc|Cryptographic.*Services'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "CRYPTO SERVICE ERROR" -Description "Restart Cryptographic Services - Delete C:\Windows\System32\catroot2 contents" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'DCOM.*10016|DistributedCOM'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "DCOM PERMISSIONS" -Description "Usually safe to ignore - If causing issues, fix via Component Services (dcomcnfg)" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'RPC.*unavailable|0x800706BA|endpoint.*mapper'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "RPC ERROR" -Description "Remote Procedure Call failed - Restart 'Remote Procedure Call' and 'DCOM Server Process Launcher'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'msiserver|Windows Installer|0x80070643|1603'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "MSI INSTALLER ERROR" -Description "Windows Installer issue - Re-register: 'msiexec /unregister' then 'msiexec /regserver'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # .NET FRAMEWORK ISSUES
    # ============================================
    
    $pattern = '\.NET Runtime|CLR|mscorlib|System\.'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title ".NET RUNTIME ERROR" -Description "Run .NET Framework Repair Tool from Microsoft or 'DISM /Online /Enable-Feature /FeatureName:NetFx3'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'fusion|assembly.*load|0x80131040|GAC'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title ".NET ASSEMBLY ERROR" -Description "Global Assembly Cache issue - Run 'gacutil /l' to check, consider .NET repair" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'ngen|native.*image|0x80131F06'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "NGEN ERROR" -Description "Run 'ngen update' from Developer Command Prompt to regenerate native images" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # VISUAL C++ / SIDEBYSIDE ISSUES
    # ============================================
    
    $pattern = 'SideBySide|0x80073701|0x80073702|activation.*context'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title "SIDEBYSIDE ERROR" -Description "Missing or corrupted Visual C++ Redistributable - Reinstall all VC++ versions from Microsoft" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'VCRUNTIME|msvcr|msvcp|api-ms-win'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "VC++ RUNTIME MISSING" -Description "Install Visual C++ Redistributable 2015-2022 (both x86 and x64)" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80073713|file.*hash.*mismatch|hash.*verification'
    if ($allIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allCBSDISM -Pattern $pattern
        Add-RecommendationWithContext -Title "FILE HASH MISMATCH" -Description "Component file corrupted - Run 'sfc /scannow' then 'DISM /RestoreHealth'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # WINDOWS UPDATE SPECIFIC ISSUES
    # ============================================
    
    $pattern = '0x80240017|not.*applicable|WU_E_NOT_APPLICABLE'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "UPDATE NOT APPLICABLE" -Description "Update doesn't apply to this system - May need prerequisite updates first" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x8024402C|0x80244022|WU_E_PT|protocol.*error'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "WU PROTOCOL ERROR" -Description "Reset Windows Update components - Stop services, rename SoftwareDistribution, restart services" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80248007|WU_E_DS_NODATA|data.*store'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "WU DATA STORE CORRUPT" -Description "Delete C:\Windows\SoftwareDistribution\DataStore folder and restart WU service" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80242006|0x80242007|handler.*hung|installer.*timeout'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "UPDATE HANDLER TIMEOUT" -Description "Update installation timed out - Try installing updates one at a time" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80240020|no.*interactive.*user'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "NO INTERACTIVE USER" -Description "Update requires user interaction - Log in to console session and run manually" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x8024A000|AU.*unable|automatic.*updates'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "AUTO UPDATE ERROR" -Description "Automatic Updates service issue - Reset WU components and check Group Policy settings" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'WSUS|0x8024401C|0x80244017|same.*client.*id'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "WSUS CLIENT ID" -Description "Reset WSUS client ID - Delete SusClientId registry keys and run 'wuauclt /resetauthorization'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x8024D001|0x8024D007|WUA.*setup|selfupdate'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "WU AGENT UPDATE NEEDED" -Description "Windows Update Agent needs updating - Download latest WU Agent from Microsoft" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80246008|0x80246007|download.*not.*available'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "DOWNLOAD ERROR" -Description "Update download failed - Clear download cache, check connectivity, try Microsoft Update Catalog" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # REGISTRY ISSUES
    # ============================================
    
    $pattern = 'registry|0x80070002.*registry|HKLM.*missing|reg.*corrupt'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "REGISTRY ISSUE" -Description "Registry key missing or corrupted - Run 'sfc /scannow', consider registry backup restore" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80240012|WU_E_REG_VALUE_INVALID|invalid.*registry'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "INVALID REGISTRY VALUE" -Description "Windows Update registry corrupted - Reset WU components or edit registry manually" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # MEMORY / RESOURCE ISSUES
    # ============================================
    
    $pattern = '0x8007000E|out.*memory|insufficient.*memory|E_OUTOFMEMORY'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "OUT OF MEMORY" -Description "Close applications, increase page file size, check for memory leaks with Task Manager" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'pool.*depletion|nonpaged.*pool|paged.*pool'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "KERNEL POOL DEPLETION" -Description "System running low on kernel memory - Reboot, check for driver leaks" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'handle.*leak|handle.*count|GDI.*objects'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "HANDLE/GDI LEAK" -Description "Process leaking handles - Identify with Process Explorer and restart affected application" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # DRIVER ISSUES
    # ============================================
    
    $pattern = '0x8024C001|driver.*pruned|WU_E_DRV'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "DRIVER UPDATE SKIPPED" -Description "Driver was skipped - May need manual driver update from manufacturer" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'driver.*failed|0x8007F0DA|driver.*blocked'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "DRIVER FAILURE" -Description "Driver installation failed - Update driver manually or roll back to previous version" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'bugcheck|blue.*screen|BSOD|0x0000'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "BSOD/BUGCHECK" -Description "System crash detected - Check minidump files in C:\Windows\Minidump for driver info" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # SECURITY / POLICY ISSUES
    # ============================================
    
    $pattern = '0x80070005|access.*denied|E_ACCESSDENIED'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "ACCESS DENIED" -Description "Permissions issue - Run as Administrator, check folder permissions, disable antivirus temporarily" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80240025|group.*policy|WU_E_USER_ACCESS_DISABLED'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "GROUP POLICY BLOCK" -Description "Windows Update blocked by policy - Check 'gpedit.msc' > Computer Config > Admin Templates > Windows Update" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x8024002F|DisableWindowsUpdateAccess|policy.*disabled'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "WU DISABLED BY POLICY" -Description "Run 'gpupdate /force', check for conflicting GPOs or third-party management tools" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'defender|antimalware|security.*software|0x80070643.*defender'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "SECURITY SOFTWARE CONFLICT" -Description "Temporarily disable antivirus/antimalware during updates, then re-enable" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'SecureBoot|0x80070032|secure.*boot'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "SECURE BOOT ISSUE" -Description "Check Secure Boot status in BIOS, ensure proper signing of boot components" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # APPLICATION ERRORS
    # ============================================
    
    $pattern = 'Application Error|faulting.*module|exception.*code'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title "APPLICATION CRASH" -Description "Check for application updates, reinstall affected application, run as Administrator" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'Windows Error Reporting|WER|crash.*dump'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title "WER REPORTS" -Description "Check Action Center > Reliability History for crash details and solutions" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'MSIINSTALLER.*1603|1618|1619'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title "MSI ERROR" -Description "Installation failed - Check temp folder space, run as Admin, try MSI in Safe Mode" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'Office|0x80070BC9|0x80070057.*Office'
    if ($allAppIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allApp -Pattern $pattern
        Add-RecommendationWithContext -Title "OFFICE ERROR" -Description "Run Office Online Repair from Control Panel > Programs > Microsoft Office > Change" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # STORE / UWP ISSUES
    # ============================================
    
    $pattern = 'AppXDeployment|Store.*error|0x80073D05|0x80073CF9'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "STORE/UWP ERROR" -Description "Run 'wsreset.exe', or PowerShell: Get-AppXPackage | Foreach {Add-AppxPackage -Register}" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80073D01|deployment.*blocked'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "APP DEPLOYMENT BLOCKED" -Description "Reset Microsoft Store: Settings > Apps > Microsoft Store > Advanced > Reset" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # BOOT / STARTUP ISSUES
    # ============================================
    
    $pattern = 'boot.*fail|BCD|bootmgr|0xc000000e'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "BOOT ISSUE" -Description "Boot configuration problem - Run 'bootrec /fixmbr', 'bootrec /fixboot', 'bootrec /rebuildbcd'" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'Kernel-Power|unexpected.*shutdown|0x41'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "UNEXPECTED SHUTDOWN" -Description "Power loss or system hang - Check power supply, thermals, and Event ID 41 details" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = 'boot.*configuration|0xc0000034|winload'
    if ($allSystemIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allSystem -Pattern $pattern
        Add-RecommendationWithContext -Title "BOOT LOADER ERROR" -Description "Run 'bcdboot C:\Windows /s C:' from recovery environment" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # TIME / DATE ISSUES
    # ============================================
    
    $pattern = 'time.*sync|W32Time|0x80072F8F.*date'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "TIME SYNC ERROR" -Description "Run 'w32tm /resync', check CMOS battery, verify time zone settings" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # ============================================
    # GENERAL / CATCH-ALL
    # ============================================
    
    $pattern = '0x80004005|E_FAIL|unspecified.*error'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "UNSPECIFIED ERROR (0x80004005)" -Description "Generic failure - Check CBS.log for specific details, run SFC and DISM" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80070002|0x80070003|file.*not.*found|path.*not.*found'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "FILE NOT FOUND" -Description "Required file missing - Run 'sfc /scannow' to restore system files" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    $pattern = '0x80070057|E_INVALIDARG|invalid.*parameter'
    if ($combinedIssues -match $pattern) {
        $matched = Find-MatchingEntries -Entries $allEntries -Pattern $pattern
        Add-RecommendationWithContext -Title "INVALID PARAMETER" -Description "Check for corrupted settings - Reset component, clear cache, check for third-party interference" -MatchedEntries $matched -RecommendationsList ([ref]$recommendations)
    }
    
    # Output recommendations with detailed context
    if ($recommendations.Count -gt 0) {
        Write-Log "========================================" -Level Warning
        Write-Log "RECOMMENDATIONS BASED ON LOG ANALYSIS:" -Level Warning
        Write-Log "========================================" -Level Warning
        
        foreach ($rec in $recommendations) {
            Write-Log "" -Level Info
            Write-Log "[$($rec.Title)]" -Level Warning
            Write-Log "  Action: $($rec.Description)" -Level Warning
            
            if ($rec.ErrorCodes.Count -gt 0) {
                Write-Log "  Error Codes: $($rec.ErrorCodes -join ', ')" -Level Warning
            }
            
            if ($rec.Files.Count -gt 0) {
                Write-Log "  Related Files:" -Level Warning
                foreach ($file in $rec.Files) {
                    Write-Log "    - $file" -Level Warning
                }
            }
            
            if ($rec.Context.Count -gt 0) {
                Write-Log "  Log Entries:" -Level Info
                foreach ($ctx in $rec.Context) {
                    Write-Log "    > $ctx" -Level Info
                }
            }
        }
        
        Write-Log "" -Level Info
        Write-Log "Note: Multiple issues may be related. Address them in order of severity." -Level Info
        Write-Log "For persistent issues, consider an in-place upgrade repair (run Windows Setup keeping files/apps)." -Level Info
    }
    else {
        Write-Log "No specific recommendations - System logs appear healthy." -Level Success
    }
    
    # Export recommendations to CSV with structured output
    if ($recommendations.Count -gt 0) {
        try {
            $hostname = $env:COMPUTERNAME
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $csvData = @()
            
            # Build structured CSV data with more columns
            foreach ($rec in $recommendations) {
                $errorCodes = if ($rec.ErrorCodes.Count -gt 0) { $rec.ErrorCodes -join '; ' } else { 'N/A' }
                $files = if ($rec.Files.Count -gt 0) { $rec.Files -join '; ' } else { 'N/A' }
                $context = if ($rec.Context.Count -gt 0) { ($rec.Context -join ' | ') -replace '"', '""' } else { 'N/A' }
                
                # Categorize the issue type
                $issueCategory = switch -Regex ($rec.Title) {
                    "CBS|Corruption|corrupt" { "Corruption" }
                    "DISM" { "DISM" }
                    "Windows Update|WUA|AU" { "WindowsUpdate" }
                    "Disk|Space|Storage" { "Storage" }
                    "Permission|Access|Denied" { "Permissions" }
                    "Network|Connectivity|Proxy" { "Network" }
                    "Service|wuauserv|bits" { "Services" }
                    default { "Other" }
                }
                
                $csvData += [PSCustomObject]@{
                    Hostname = $hostname
                    Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                    CorrelationId = $script:CorrelationId
                    IssueCategory = $issueCategory
                    IssueType = $rec.Title
                    Severity = $rec.Severity
                    ErrorCodes = $errorCodes
                    AffectedFiles = $files
                    LogContext = $context
                    RecommendedFix = $rec.Description
                    AutomatedFixAvailable = $rec.CanAutoFix
                    ScriptVersion = "2.0"
                    OSVersion = [System.Environment]::OSVersion.VersionString
                }
            }
            
            # Create per-host subdirectory for CSV files
            $csvBaseDir = Split-Path -Parent $script:RecommendationsCsvPath
            $csvDir = Join-Path $csvBaseDir $hostname
            if (-not (Test-Path $csvDir)) {
                New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
            }
            
            # Create filename with hostname and timestamp
            $csvFilename = "WU_Recommendations_$hostname`_$timestamp.csv"
            $csvFullPath = Join-Path $csvDir $csvFilename
            
            # Export to CSV (no append for structured output - new file each run)
            $csvData | Export-Csv -Path $csvFullPath -NoTypeInformation
            
            Write-Log "Recommendations exported to CSV: $csvFullPath" -Level Success
            Write-Log "  Exported $($csvData.Count) recommendation(s)" -Level Info
        }
        catch {
            Write-Log "Failed to export recommendations to CSV: $_" -Level Warning
        }
    }
    
    # Summary
    $totalIssues = $issues.CBS.Count + $issues.DISM.Count + $issues.WindowsUpdate.Count + $issues.System.Count + $issues.Application.Count
    Write-Log "Log analysis complete. Total issues found: $totalIssues" -Level $(if ($totalIssues -gt 0) { "Warning" } else { "Success" })
    
    return @{
        Issues = $issues
        Recommendations = $recommendations
        TotalCount = $totalIssues
    }
}

function Test-AdminPrivileges {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PendingReboot {
    $rebootPending = $false
    $reasons = @()
    
    # Check Component Based Servicing
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        $rebootPending = $true
        $reasons += "Component Based Servicing"
    }
    
    # Check Windows Update
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
        $rebootPending = $true
        $reasons += "Windows Update"
    }
    
    # Check Pending File Rename Operations
    $pfro = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
    if ($pfro.PendingFileRenameOperations) {
        $rebootPending = $true
        $reasons += "Pending File Rename Operations"
    }
    
    # Check SCCM Client
    try {
        $sccmReboot = Invoke-CimMethod -Namespace "root\ccm\ClientSDK" -ClassName "CCM_ClientUtilities" -MethodName "DetermineIfRebootPending" -ErrorAction SilentlyContinue
        if ($sccmReboot -and ($sccmReboot.RebootPending -or $sccmReboot.IsHardRebootPending)) {
            $rebootPending = $true
            $reasons += "SCCM Client"
        }
    }
    catch {
        # SCCM not installed, ignore
    }
    
    return @{
        RebootPending = $rebootPending
        Reasons = $reasons
    }
}

function Test-DiskSpace {
    param(
        [int]$MinimumGB = 10
    )
    
    try {
        $systemDrive = $env:SystemDrive
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop
        $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        
        return @{
            Success = $freeSpaceGB -ge $MinimumGB
            FreeSpaceGB = $freeSpaceGB
            RequiredGB = $MinimumGB
        }
    }
    catch {
        Write-Log "Failed to check disk space: $_" -Level Warning
        return @{
            Success = $true  # Assume OK if we can't check
            FreeSpaceGB = "Unknown"
            RequiredGB = $MinimumGB
        }
    }
}

function Test-NetworkConnectivity {
    # First check if WSUS is configured - if so, test WSUS server instead of Microsoft
    $wsusConfig = Test-WSUSConfiguration
    
    $results = @()
    $networkWorking = $false
    
    # Test basic internet connectivity first (DNS resolution)
    try {
        $dns = [System.Net.Dns]::GetHostAddresses("www.microsoft.com")
        if ($dns.Count -gt 0) {
            $networkWorking = $true
            $results += @{ Name = "DNS Resolution"; Reachable = $true; Status = "OK - Resolved www.microsoft.com" }
        }
    }
    catch {
        $results += @{ Name = "DNS Resolution"; Reachable = $false; Status = "Failed to resolve www.microsoft.com" }
    }
    
    # If using WSUS, test the WSUS server instead
    if ($wsusConfig.UsingWSUS -and $wsusConfig.WSUSServer) {
        try {
            $request = [System.Net.WebRequest]::Create($wsusConfig.WSUSServer)
            $request.Timeout = 10000
            $request.Method = "HEAD"
            try {
                $response = $request.GetResponse()
                $response.Close()
                $results += @{ Name = "WSUS Server"; Reachable = $true; Status = "OK - $($wsusConfig.WSUSServer)" }
                $networkWorking = $true
            }
            catch [System.Net.WebException] {
                $webEx = $_.Exception
                $statusCode = $null
                if ($webEx.Response) {
                    $statusCode = [int]$webEx.Response.StatusCode
                }
                # 403, 401, 400 mean the server IS reachable, just blocking/requiring auth
                if ($statusCode -in @(400, 401, 403, 404, 405, 503)) {
                    $results += @{ Name = "WSUS Server"; Reachable = $true; Status = "Reachable (HTTP $statusCode) - $($wsusConfig.WSUSServer)" }
                    $networkWorking = $true
                }
                else {
                    $results += @{ Name = "WSUS Server"; Reachable = $false; Status = $_.Exception.Message }
                }
            }
        }
        catch {
            $results += @{ Name = "WSUS Server"; Reachable = $false; Status = "Failed: $($_.Exception.Message)" }
        }
        
        # Return early for WSUS - don't need to check Microsoft servers
        return @{
            Success = $networkWorking
            Results = $results
            UsingWSUS = $true
        }
    }
    
    # Test Microsoft endpoints (for non-WSUS environments)
    $endpoints = @(
        @{ Name = "Microsoft Update"; URL = "https://update.microsoft.com" },
        @{ Name = "Windows Update"; URL = "https://windowsupdate.microsoft.com" },
        @{ Name = "Microsoft Download"; URL = "https://download.microsoft.com" }
    )
    
    foreach ($endpoint in $endpoints) {
        try {
            $request = [System.Net.WebRequest]::Create($endpoint.URL)
            $request.Timeout = 10000
            $request.Method = "HEAD"
            try {
                $response = $request.GetResponse()
                $response.Close()
                $results += @{ Name = $endpoint.Name; Reachable = $true; Status = "OK" }
                $networkWorking = $true
            }
            catch [System.Net.WebException] {
                $webEx = $_.Exception
                $statusCode = $null
                if ($webEx.Response) {
                    $statusCode = [int]$webEx.Response.StatusCode
                }
                # 403, 401, 400, 503 mean server IS reachable (proxy/firewall blocking, or service unavailable)
                # This indicates network connectivity works, just policy/access issues
                if ($statusCode -in @(400, 401, 403, 404, 405, 500, 503)) {
                    $results += @{ Name = $endpoint.Name; Reachable = $true; Status = "Reachable (HTTP $statusCode - may be proxy/firewall)" }
                    $networkWorking = $true
                }
                else {
                    $results += @{ Name = $endpoint.Name; Reachable = $false; Status = $_.Exception.Message }
                }
            }
        }
        catch {
            $results += @{ Name = $endpoint.Name; Reachable = $false; Status = $_.Exception.Message }
        }
    }
    
    return @{
        Success = $networkWorking
        Results = $results
        UsingWSUS = $false
    }
}

function Test-WSUSConfiguration {
    $result = @{
        UsingWSUS = $false
        WSUSServer = $null
        WUStatusServer = $null
        DualScanEnabled = $false
        DualScanSource = $null
        Policies = @()
        IsManaged = $false
    }
    
    try {
        # Primary WSUS settings
        $wuServer = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "WUServer" -ErrorAction SilentlyContinue
        $wuStatusServer = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "WUStatusServer" -ErrorAction SilentlyContinue
        $useWUServer = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -ErrorAction SilentlyContinue
        
        if ($wuServer -and $useWUServer.UseWUServer -eq 1) {
            $result.UsingWSUS = $true
            $result.WSUSServer = $wuServer.WUServer
            $result.WUStatusServer = $wuStatusServer.WUStatusServer
            $result.Policies += "WSUS Server configured"
        }
        
        # Check for Dual Scan (WUfB + WSUS)
        $doPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        $disableDualScan = Get-ItemProperty -Path $doPath -Name "DisableDualScan" -ErrorAction SilentlyContinue
        $setPolicySource = Get-ItemProperty -Path $doPath -Name "SetPolicyDrivenUpdateSourceForOtherUpdates" -ErrorAction SilentlyContinue
        $setPolicySourceForDrivers = Get-ItemProperty -Path $doPath -Name "SetPolicyDrivenUpdateSourceForDriverUpdates" -ErrorAction SilentlyContinue
        
        if ($disableDualScan.DisableDualScan -ne 1 -and $result.UsingWSUS) {
            # Dual Scan is not disabled - could be in effect
            if ($setPolicySource.SetPolicyDrivenUpdateSourceForOtherUpdates -eq 1 -or 
                $setPolicySourceForDrivers.SetPolicyDrivenUpdateSourceForDriverUpdates -eq 1) {
                $result.DualScanEnabled = $true
                $result.Policies += "Dual Scan enabled"
                
                if ($setPolicySource.SetPolicyDrivenUpdateSourceForOtherUpdates -eq 1) {
                    $result.DualScanSource = "Windows Update for Business (WUfB)"
                }
            }
        }
        
        # Check for Windows Update for Business (WUfB) deferral policies
        $deferFeature = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DeferFeatureUpdatesPeriodInDays" -ErrorAction SilentlyContinue
        $deferQuality = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DeferQualityUpdatesPeriodInDays" -ErrorAction SilentlyContinue
        
        if ($deferFeature.DeferFeatureUpdatesPeriodInDays -gt 0) {
            $result.Policies += "Feature updates deferred: $($deferFeature.DeferFeatureUpdatesPeriodInDays) days"
            $result.IsManaged = $true
        }
        if ($deferQuality.DeferQualityUpdatesPeriodInDays -gt 0) {
            $result.Policies += "Quality updates deferred: $($deferQuality.DeferQualityUpdatesPeriodInDays) days"
            $result.IsManaged = $true
        }
        
        # Check for Microsoft Update (Office updates)
        $muEnabled = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "AllowMUUpdateService" -ErrorAction SilentlyContinue
        if ($muEnabled.AllowMUUpdateService -eq 1) {
            $result.Policies += "Microsoft Update enabled (Office updates)"
        }
        
        # Check if any group policies are configured
        if ($result.Policies.Count -gt 0) {
            $result.IsManaged = $true
        }
    }
    catch {
        # Not using WSUS or error reading registry
    }
    
    return $result
}

function Test-WUAgentHealth {
    Write-Log "Checking Windows Update Agent health..." -Level Info
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        
        # Get WUA version from registry
        $wuaPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate"
        $wuaVersion = $null
        try {
            $wuaVersion = (Get-ItemProperty -Path $wuaPath -Name "SusClientVersion" -ErrorAction SilentlyContinue).SusClientVersion
        }
        catch {
            # Version not in registry, try COM
        }
        
        # Test basic WUA functionality with a quick search
        $testResult = $updateSearcher.Search("IsInstalled=1")
        if (-not $testResult) {
            throw "WUA search returned null result"
        }

        if ($testResult.ResultCode -ne 2) {  # 2 = SearchResult_Succeeded
            throw "WUA search failed with result code: $($testResult.ResultCode)"
        }
        Write-Log "Windows Update Agent version: $wuaVersion" -Level Info
        Write-Log "WUA basic functionality test: PASSED" -Level Success
        Write-Log "WUA basic functionality test: PASSED (found $($testResult.Updates.Count) installed updates)" -Level Success
        return @{
            Success = $true
            Version = $wuaVersion
            CanSearch = $true
            Error = $null
        }
    }
    catch {
        Write-Log "Windows Update Agent health check failed: $_" -Level Error
        return @{
            Success = $false
            Version = $null
            CanSearch = $false
            Error = $_.Exception.Message
        }
    }
}

function Test-NetworkConfiguration {
    Write-Log "Checking network configuration..." -Level Info
    
    $issues = @()
    $isMetered = $false
    $hasProxy = $false
    $proxySettings = $null
    
    # Check for metered connection
    try {
        $networkCost = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($networkCost) {
            # Windows 8+ metered connection detection via registry
            $meteredRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\DefaultMediaCost"
            if (Test-Path $meteredRegPath) {
                $costValue = (Get-ItemProperty -Path $meteredRegPath -Name "WiFi" -ErrorAction SilentlyContinue).WiFi
                if ($costValue -eq 2) { $isMetered = $true }
            }
        }
    }
    catch {
        # Unable to detect metered status
    }
    
    # Check WinHTTP proxy
    try {
        $proxyOutput = netsh winhttp show proxy 2>&1
        if ($proxyOutput -match "Proxy Server\(s\):\s*(\S+)" -and $matches[1] -ne "Direct") {
            $hasProxy = $true
            $proxySettings = $matches[1]
        }
    }
    catch {
        # Unable to check proxy
    }
    
    # Check network category (public vs private)
    $networkCategory = "Unknown"
    try {
        $netProfile = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($netProfile) {
            $networkCategory = $netProfile.NetworkCategory
        }
    }
    catch {
        # Unable to check category
    }
    
    # Check for VPN connections
    $hasVPN = $false
    try {
        $vpnConnections = Get-VpnConnection -ErrorAction SilentlyContinue
        if ($vpnConnections | Where-Object { $_.ConnectionStatus -eq "Connected" }) {
            $hasVPN = $true
            $issues += "VPN connection detected - may affect update downloads"
        }
    }
    catch {
        # Unable to check VPN
    }
    
    # Check Delivery Optimization settings
    try {
        $doPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config"
        $doEnabled = (Get-ItemProperty -Path $doPath -Name "DODownloadMode" -ErrorAction SilentlyContinue).DODownloadMode
        if ($doEnabled -eq 0) {
            $issues += "Delivery Optimization disabled - updates will download directly from Microsoft"
        }
    }
    catch {
        # Unable to check DO settings
    }
    
    # Log findings
    if ($isMetered) {
        Write-Log "WARNING: Metered connection detected - downloads may be blocked" -Level Warning
        $issues += "Metered connection - Windows may block update downloads"
    }
    
    if ($hasProxy) {
        Write-Log "Proxy configured: $proxySettings" -Level Info
    }
    else {
        Write-Log "No proxy configured (direct connection)" -Level Info
    }
    
    Write-Log "Network category: $networkCategory" -Level Info
    
    if ($hasVPN) {
        Write-Log "VPN connection active" -Level Warning
    }
    
    return @{
        IsMetered = $isMetered
        HasProxy = $hasProxy
        ProxySettings = $proxySettings
        NetworkCategory = $networkCategory
        HasVPN = $hasVPN
        Issues = $issues
    }
}

function Test-MicrosoftConnectivity {
    param(
        [int[]]$Ports = @(443),
        [switch]$IncludeTraceroute
    )
    
    Write-Log "========================================" -Level Info
    Write-Log "Running Microsoft Connectivity Tests" -Level Info
    Write-Log "========================================" -Level Info
    
    # Windows Update focused endpoints from MS_connecttest.ps1
    $targets = @(
        @{ Display = 'Windows Update (Generic)'; Host = 'windowsupdate.com' },
        @{ Display = 'Windows Update Microsoft'; Host = 'update.microsoft.com' },
        @{ Display = 'Windows Update Download'; Host = 'download.windowsupdate.com' },
        @{ Display = 'Windows Update Auth'; Host = 'wustat.windows.com' },
        @{ Display = 'Windows Update CTLDL'; Host = 'ctldl.windowsupdate.com' },
        @{ Display = 'Windows Update SLS'; Host = 'sls.update.microsoft.com' },
        @{ Display = 'Delivery Optimization (MP)'; Host = 'delivery.mp.microsoft.com' },
        @{ Display = 'Windows Update Download (AU)'; Host = 'au.download.windowsupdate.com' },
        @{ Display = 'Microsoft Store Licensing'; Host = 'licensing.mp.microsoft.com' },
        @{ Display = 'Office CDN'; Host = 'officecdn.microsoft.com' },
        @{ Display = 'Login Microsoft Online (AAD)'; Host = 'login.microsoftonline.com' },
        @{ Display = 'Microsoft CRL'; Host = 'mscrl.microsoft.com' },
        @{ Display = 'DigiCert OCSP'; Host = 'ocsp.digicert.com' },
        @{ Display = 'Windows Time Service'; Host = 'time.windows.com' },
        @{ Display = 'NCSI Test'; Host = 'www.msftconnecttest.com' }
    )
    
    $results = @()
    $passed = 0
    $failed = 0
    
    foreach ($target in $targets) {
        Write-Log "Testing $($target.Display) ($($target.Host))..." -Level Info
        
        $dnsResolved = $false
        $resolvedAddress = $null
        $tcpResults = @()
        
        # DNS Resolution
        try {
            $dnsResolution = Resolve-DnsName -Name $target.Host -ErrorAction Stop
            $dnsResult = $dnsResolution | Where-Object { $_.Type -eq 'A' } | Select-Object -First 1
            if (-not $dnsResult) {
                $dnsResult = $dnsResolution | Where-Object { $_.Type -eq 'AAAA' } | Select-Object -First 1
            }
            
            if ($dnsResult) {
                $dnsResolved = $true
                $resolvedAddress = $dnsResult.IPAddress
                Write-Log "  DNS: Resolved to $resolvedAddress" -Level Success
            }
            else {
                Write-Log "  DNS: No IP address found" -Level Warning
                $failed++
            }
        }
        catch {
            Write-Log "  DNS: Failed - $($_.Exception.Message)" -Level Warning
            $failed++
        }
        
        # TCP Port Tests
        if ($dnsResolved -and $resolvedAddress) {
            foreach ($port in $Ports) {
                try {
                    $tcpClient = New-Object System.Net.Sockets.TcpClient
                    $connection = $tcpClient.BeginConnect($resolvedAddress, $port, $null, $null)
                    $success = $connection.AsyncWaitHandle.WaitOne(5000, $false)
                    
                    if ($success -and $tcpClient.Connected) {
                        $tcpClient.Close()
                        Write-Log "  TCP Port $port`: OPEN" -Level Success
                        $tcpResults += @{ Port = $port; Reachable = $true }
                        $passed++
                    }
                    else {
                        $tcpClient.Close()
                        Write-Log "  TCP Port $port`: CLOSED/FAILED" -Level Warning
                        $tcpResults += @{ Port = $port; Reachable = $false }
                        $failed++
                        
                        # Optional traceroute on failure
                        if ($IncludeTraceroute) {
                            Write-Log "  Running traceroute to $resolvedAddress..." -Level Info
                            try {
                                $traceResult = Test-NetConnection -ComputerName $resolvedAddress -TraceRoute -InformationLevel Quiet -ErrorAction SilentlyContinue
                                if ($traceResult.TraceRoute) {
                                    $hops = $traceResult.TraceRoute.Count
                                    Write-Log "    Traceroute: $hops hops completed" -Level Info
                                }
                            }
                            catch {
                                Write-Log "    Traceroute failed" -Level Warning
                            }
                        }
                    }
                }
                catch {
                    Write-Log "  TCP Port $port`: ERROR - $($_.Exception.Message)" -Level Warning
                    $tcpResults += @{ Port = $port; Reachable = $false; Error = $_.Exception.Message }
                    $failed++
                }
            }
        }
        else {
            Write-Log "  TCP: SKIPPED (DNS resolution failed)" -Level Warning
            foreach ($port in $Ports) {
                $tcpResults += @{ Port = $port; Reachable = $false; Error = 'DNS failed' }
                $failed++
            }
        }
        
        $results += [PSCustomObject]@{
            DisplayName = $target.Display
            Host = $target.Host
            DnsResolved = $dnsResolved
            ResolvedAddress = $resolvedAddress
            TcpResults = $tcpResults
        }
    }
    
    Write-Log "Connectivity test complete: $passed passed, $failed failed" -Level $(if ($failed -eq 0) { 'Success' } elseif ($failed -lt ($passed + $failed) / 2) { 'Warning' } else { 'Error' })
    
    return @{
        Results = $results
        Passed = $passed
        Failed = $failed
        Total = $passed + $failed
        Success = ($failed -eq 0)
    }
}

function Invoke-PreflightChecks {
    Write-Log "========================================" -Level Info
    Write-Log "Running Pre-flight Checks" -Level Info
    Write-Log "========================================" -Level Info
    
    $allPassed = $true
    $warnings = @()
    
    # Check 1: Pending Reboot
    Write-Log "Checking for pending reboots..." -Level Info
    $rebootCheck = Test-PendingReboot
    if ($rebootCheck.RebootPending) {
        Write-Log "WARNING: System has pending reboot (Reasons: $($rebootCheck.Reasons -join ', '))" -Level Warning
        $warnings += "Pending reboot detected. Some updates may fail until system is restarted."
    }
    else {
        Write-Log "No pending reboot detected" -Level Success
    }
    
    # Check 2: Disk Space
    Write-Log "Checking disk space..." -Level Info
    $diskCheck = Test-DiskSpace -MinimumGB 10
    if (-not $diskCheck.Success) {
        Write-Log "ERROR: Insufficient disk space. Free: $($diskCheck.FreeSpaceGB) GB, Required: $($diskCheck.RequiredGB) GB" -Level Error
        $allPassed = $false
    }
    else {
        Write-Log "Disk space OK: $($diskCheck.FreeSpaceGB) GB free" -Level Success
    }
    
    # Check 3: Network Connectivity
    Write-Log "Checking network connectivity..." -Level Info
    $networkCheck = Test-NetworkConnectivity
    if (-not $networkCheck.Success) {
        Write-Log "ERROR: Cannot reach Windows Update servers" -Level Error
        foreach ($result in $networkCheck.Results) {
            Write-Log "  - $($result.Name): $($result.Status)" -Level $(if ($result.Reachable) { "Success" } else { "Error" })
        }
        $allPassed = $false
    }
    else {
        Write-Log "Network connectivity OK" -Level Success
    }
    
    # Check 4: WSUS Configuration
    Write-Log "Checking update source configuration..." -Level Info
    $wsusCheck = Test-WSUSConfiguration
    if ($wsusCheck.UsingWSUS) {
        Write-Log "System configured to use WSUS: $($wsusCheck.WSUSServer)" -Level Info
    }
    else {
        Write-Log "System configured to use Microsoft Update directly" -Level Info
    }
    
    # Check 5: Windows Update Service Status
    Write-Log "Checking Windows Update service status..." -Level Info
    $wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
    if ($wuService) {
        Write-Log "Windows Update service status: $($wuService.Status), StartType: $($wuService.StartType)" -Level Info
        if ($wuService.StartType -eq "Disabled") {
            Write-Log "WARNING: Windows Update service is disabled. Will attempt to enable." -Level Warning
            try {
                Set-Service -Name wuauserv -StartupType Manual -ErrorAction Stop
                Write-Log "Windows Update service startup type changed to Manual" -Level Success
            }
            catch {
                Write-Log "Failed to change Windows Update service startup type: $_" -Level Error
                $allPassed = $false
            }
        }
    }
    else {
        Write-Log "ERROR: Windows Update service not found" -Level Error
        $allPassed = $false
    }
    
    # Check 6: Windows Update Agent Health
    $wuaHealth = Test-WUAgentHealth
    if (-not $wuaHealth.Success) {
        Write-Log "WARNING: Windows Update Agent health check failed: $($wuaHealth.Error)" -Level Warning
        $warnings += "WUA health check failed - updates may not work properly"
    }
    
    # Check 7: Network Configuration (Proxy, VPN, Metered)
    $netConfig = Test-NetworkConfiguration
    if ($netConfig.Issues.Count -gt 0) {
        foreach ($issue in $netConfig.Issues) {
            $warnings += $issue
        }
    }
    
    # Check 8: Extended Microsoft Connectivity Tests (optional)
    if ($TestConnectivity) {
        $connectivityResult = Test-MicrosoftConnectivity -Ports @(443, 80)
        if (-not $connectivityResult.Success) {
            $warnings += "Microsoft connectivity tests failed: $($connectivityResult.Failed) of $($connectivityResult.Total) tests failed"
            if ($connectivityResult.Failed -gt ($connectivityResult.Total / 2)) {
                Write-Log "CRITICAL: More than 50% of connectivity tests failed. Updates may not work." -Level Error
                $allPassed = $false
            }
        }
    }
    
    Write-Log "Pre-flight checks completed" -Level $(if ($allPassed) { "Success" } else { "Warning" })
    
    if ($warnings.Count -gt 0) {
        Write-Log "Warnings:" -Level Warning
        foreach ($warning in $warnings) {
            Write-Log "  - $warning" -Level Warning
        }
    }
    
    # Check 6: Scan system logs for issues
    Write-Log "Scanning system logs for existing issues..." -Level Info
    $logAnalysis = Get-SystemLogIssues -HoursBack 48 -MaxEntries 50
    
    return @{
        Passed = $allPassed
        Warnings = $warnings
        RebootPending = $rebootCheck.RebootPending
        LogAnalysis = $logAnalysis
    }
}

function New-SystemRestorePoint {
    param([string]$Description = "Before Windows Update Script")
    
    Write-Log "Creating System Restore Point..." -Level Info
    
    try {
        # Check if System Protection is enabled for system drive
        $systemDrive = $env:SystemDrive
        
        # Check if System Restore is available on system drive
        $srStatus = & vssadmin list shadowstorage /for=$systemDrive 2>&1
        if ($srStatus -match "Error" -or $LASTEXITCODE -ne 0) {
            Write-Log "System Protection not configured on $systemDrive. Skipping restore point creation." -Level Warning
            Write-Log "TIP: Enable System Protection in Control Panel > System > System Protection" -Level Info
            return $false
        }
        
        # Check available shadow storage space (warn if low)
        try {
            $shadowInfo = Get-CimInstance -ClassName Win32_ShadowStorage -Filter "Volume='$systemDrive\\'" -ErrorAction SilentlyContinue
            if ($shadowInfo) {
                $usedSpaceGB = [math]::Round($shadowInfo.UsedSpace / 1GB, 2)
                $maxSpaceGB = [math]::Round($shadowInfo.MaxSpace / 1GB, 2)
                if ($maxSpaceGB -gt 0) {
                    $usedPercent = ($shadowInfo.UsedSpace / $shadowInfo.MaxSpace) * 100
                    if ($usedPercent -gt 90) {
                        Write-Log "WARNING: Shadow storage is $([math]::Round($usedPercent,1))% full ($usedSpaceGB GB of $maxSpaceGB GB)" -Level Warning
                    }
                }
            }
        }
        catch {
            # Ignore shadow storage check errors
        }
        
        # Enable System Restore if needed (on system drive)
        Enable-ComputerRestore -Drive "$systemDrive\" -ErrorAction SilentlyContinue
        
        # Create restore point
        Checkpoint-Computer -Description $Description -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
        Write-Log "System Restore Point created successfully" -Level Success
        
        # Get the restore point sequence number for reference
        $newRestorePoint = Get-ComputerRestorePoint | Sort-Object -Property Date -Descending | Select-Object -First 1
        if ($newRestorePoint) {
            Write-Log "Restore Point ID: $($newRestorePoint.SequenceNumber), Date: $($newRestorePoint.CreationTime)" -Level Info
        }
        
        return $true
    }
    catch {
        if ($_.Exception.Message -match "frequency") {
            Write-Log "A restore point was already created recently. Skipping." -Level Warning
        }
        elseif ($_.Exception.Message -match "disabled") {
            Write-Log "System Restore is disabled on this system." -Level Warning
        }
        else {
            Write-Log "Failed to create System Restore Point: $_" -Level Warning
        }
        return $false
    }
}

function Stop-UpdateServices {
    $services = @("wuauserv", "cryptsvc", "bits", "msiserver")
    
    foreach ($service in $services) {
        try {
            Write-Log "Stopping service: $service" -Level Info
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -eq 'Running') {
                Stop-Service -Name $service -Force -ErrorAction Stop
                Write-Log "Successfully stopped $service" -Level Success
            }
            elseif ($svc) {
                Write-Log "Service $service is already stopped" -Level Info
            }
            else {
                Write-Log "Service $service not found" -Level Warning
            }
        }
        catch {
            Write-Log "Failed to stop $service : $_" -Level Error
        }
    }
}

function Start-UpdateServices {
    $services = @("wuauserv", "cryptsvc", "bits", "msiserver")
    
    foreach ($service in $services) {
        try {
            Write-Log "Starting service: $service" -Level Info
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                Start-Service -Name $service -ErrorAction Stop
                Write-Log "Successfully started $service" -Level Success
            }
            else {
                Write-Log "Service $service not found" -Level Warning
            }
        }
        catch {
            Write-Log "Failed to start $service : $_" -Level Error
        }
    }
}

function Set-FolderOwnership {
    param(
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return
    }

    try {
        Write-Log "Taking ownership of $Path" -Level Info
        & takeown.exe /F $Path /A /R /D Y | Out-Null

        Write-Log "Granting Administrators full control on $Path" -Level Info
        & icacls.exe $Path /grant Administrators:F /T /C | Out-Null

        Write-Log "Ownership and permissions updated for $Path" -Level Success
    }
    catch {
        Write-Log "Failed to adjust permissions on $Path : $_" -Level Warning
    }
}

function Reset-UpdateFolders {
    $foldersToReset = @(
        @{ Path = "$env:SystemRoot\SoftwareDistribution"; BackupName = "SoftwareDistribution.old"; Recreate = $true },
        @{ Path = "$env:SystemRoot\System32\catroot2"; BackupName = "catroot2.old"; Recreate = $true }
    )
    
    foreach ($folder in $foldersToReset) {
        try {
            $originalPath = $folder.Path
            $backupPath = "$($folder.Path).$((Get-Date).ToString('yyyyMMddHHmmss')).bak"
            
            if (Test-Path $originalPath) {
                Write-Log "Renaming folder: $originalPath to $backupPath" -Level Info

                # Check if we need ownership adjustment - only for locked files
                $needsOwnership = $false
                try {
                    $testFile = Join-Path $originalPath "test.tmp"
                    [System.IO.File]::WriteAllText($testFile, "test")
                    Remove-Item $testFile -ErrorAction SilentlyContinue
                }
                catch {
                    $needsOwnership = $true
                }
                
                if ($needsOwnership) {
                    Set-FolderOwnership -Path $originalPath
                }

                # Remove old backup if exists
                $existingBackups = Get-ChildItem -Path (Split-Path $originalPath -Parent) -Filter "$($folder.BackupName)*" -ErrorAction SilentlyContinue
                foreach ($backup in $existingBackups) {
                    try {
                        Remove-Item -Path $backup.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        Write-Log "Removed old backup: $($backup.FullName)" -Level Info
                    }
                    catch {
                        Write-Log "Could not remove old backup: $($backup.FullName)" -Level Warning
                    }
                }
                
                Rename-Item -Path $originalPath -NewName (Split-Path $backupPath -Leaf) -Force -ErrorAction Stop
                Write-Log "Successfully renamed $originalPath" -Level Success
                
                # IMMEDIATELY recreate the folder to avoid race condition with services
                if ($folder.Recreate) {
                    Write-Log "Recreating folder: $originalPath" -Level Info
                    New-Item -ItemType Directory -Path $originalPath -Force | Out-Null
                    
                    # Set proper permissions on recreated folder
                    $acl = Get-Acl $originalPath
                    $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
                    $acl.SetAccessRule($systemRule)
                    $acl.SetAccessRule($adminRule)
                    Set-Acl $originalPath $acl
                    
                    Write-Log "Folder recreated with proper permissions: $originalPath" -Level Success
                }
            }
            else {
                Write-Log "Folder not found: $originalPath" -Level Warning
                # Create the folder if it doesn't exist
                if ($folder.Recreate) {
                    Write-Log "Creating missing folder: $originalPath" -Level Info
                    New-Item -ItemType Directory -Path $originalPath -Force | Out-Null
                    Write-Log "Created folder: $originalPath" -Level Success
                }
            }
        }
        catch {
            Write-Log "Failed to reset folder $($folder.Path) : $_" -Level Error
        }
    }
}

function Reset-ServiceSecurityDescriptors {
    try {
        Write-Log "Resetting BITS service security descriptor..." -Level Info
        $result = sc.exe sdset bits "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
        if ($LASTEXITCODE -eq 0) {
            Write-Log "BITS security descriptor reset successfully" -Level Success
        }
        else {
            Write-Log "BITS security descriptor reset returned: $result" -Level Warning
        }
        
        Write-Log "Resetting Windows Update service security descriptor..." -Level Info
        $result = sc.exe sdset wuauserv "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"
        if ($LASTEXITCODE -eq 0) {
            Write-Log "Windows Update security descriptor reset successfully" -Level Success
        }
        else {
            Write-Log "Windows Update security descriptor reset returned: $result" -Level Warning
        }
    }
    catch {
        Write-Log "Failed to reset security descriptors: $_" -Level Error
    }
}

function Register-WindowsUpdateDLLs {
    Write-Log "Re-registering Windows Update DLLs (OS-optimized)..." -Level Info
    
    # Detect Windows version
    $osVersion = [System.Environment]::OSVersion.Version
    $is64Bit = [Environment]::Is64BitOperatingSystem
    $sys32Path = Join-Path $env:SystemRoot "System32"
    $sysWow64Path = Join-Path $env:SystemRoot "SysWOW64"
    
    # Get additional OS info for Server detection
    $osProductType = (Get-CimInstance Win32_OperatingSystem).ProductType  # 1=Workstation, 2=DC, 3=Server
    $isServer = $osProductType -in @(2, 3)
    
    # Determine OS generation
    $osGeneration = switch ($osVersion.Major) {
        10 { 
            if ($isServer) {
                # Server builds: 14393=2016, 17763=2019, 20348=2022, 26100=2025
                switch ($osVersion.Build) {
                    { $_ -ge 26100 } { "Server2025" }
                    { $_ -ge 20348 } { "Server2022" }
                    { $_ -ge 17763 } { "Server2019" }
                    default { "Server2016" }
                }
            }
            elseif ($osVersion.Build -ge 22000) { "Windows11" } 
            else { "Windows10" }
        }
        6 { 
            if ($isServer) {
                switch ($osVersion.Minor) {
                    3 { "Server2012R2" }
                    2 { "Server2012" }
                    1 { "Server2008R2" }
                    0 { "Server2008" }
                }
            }
            else {
                switch ($osVersion.Minor) {
                    3 { "Windows81" }
                    2 { "Windows8" }
                    1 { "Windows7" }
                    0 { "WindowsVista" }
                }
            }
        }
        5 { 
            if ($isServer) { "Server2003" }
            else { "WindowsXP" }
        }
        default { "Unknown" }
    }
    
    Write-Log "Detected OS: $osGeneration (Build $($osVersion.Build))" -Level Info
    
    # Core WU DLLs that exist on all modern Windows versions
    $coreDlls = @(
        "wuapi.dll",
        "wuaueng.dll",
        "wups.dll",
        "wups2.dll"
    )
    
    # OS-specific DLLs
    $windows10_11_Dlls = @(
        "usocore.dll",
        "usocoreworker.dll",
        "wucltux.dll",
        "muweb.dll"
    )
    
    $windows11OnlyDlls = @(
        "wuuhosdeployment.dll"  # Windows 11 specific
    )
    
    $legacyDlls = @(
        "wuaueng1.dll",
        "wucltui.dll",
        "wuweb.dll",
        "wuwebv.dll",
        "qmgr.dll",
        "qmgrprxy.dll"
    )
    
    # Select appropriate DLL list based on OS
    $dllsToRegister = @()
    $skippedReasons = @()
    
    # Always include core DLLs
    foreach ($dll in $coreDlls) {
        if (Test-Path (Join-Path $sys32Path $dll)) {
            $dllsToRegister += $dll
        }
        else {
            $skippedReasons += "Core DLL not found: $dll"
        }
    }
    
    # Add OS-specific DLLs
    switch ($osGeneration) {
        { $_ -in @("Windows11", "Server2025") } {
            foreach ($dll in $windows10_11_Dlls) {
                if (Test-Path (Join-Path $sys32Path $dll)) {
                    $dllsToRegister += $dll
                }
            }
            foreach ($dll in $windows11OnlyDlls) {
                if (Test-Path (Join-Path $sys32Path $dll)) {
                    $dllsToRegister += $dll
                }
                else {
                    $skippedReasons += "Windows 11/Server 2025 DLL not found: $dll"
                }
            }
            Write-Log "Using Windows 11/Server 2025 specific DLL set ($($dllsToRegister.Count) DLLs)" -Level Info
        }
        
        { $_ -in @("Windows10", "Server2022", "Server2019", "Server2016") } {
            foreach ($dll in $windows10_11_Dlls) {
                if (Test-Path (Join-Path $sys32Path $dll)) {
                    $dllsToRegister += $dll
                }
            }
            Write-Log "Using Windows 10/Server 2016-2022 specific DLL set ($($dllsToRegister.Count) DLLs)" -Level Info
        }
        
        { $_ -in @("Windows81", "Windows8", "Windows7", "Server2012R2", "Server2012", "Server2008R2", "Server2008") } {
            foreach ($dll in $legacyDlls) {
                if (Test-Path (Join-Path $sys32Path $dll)) {
                    $dllsToRegister += $dll
                }
            }
            Write-Log "Using legacy Windows/Server 2008-2012 DLL set ($($dllsToRegister.Count) DLLs)" -Level Info
        }
        
        default {
            Write-Log "Unknown OS generation - using core DLLs only ($($dllsToRegister.Count) DLLs)" -Level Warning
        }
    }
    
    # Log skipped DLLs at verbose level
    foreach ($reason in $skippedReasons) {
        Write-Log "  Skipped: $reason" -Level Verbose
    }
    
    $successCount = 0
    $failCount = 0
    $skipCount = 0
    $detailedResults = @()
    
    foreach ($dll in $dllsToRegister) {
        # Register x64 version
        $dllPath64 = Join-Path $sys32Path $dll
        if (Test-Path $dllPath64) {
            try {
                [void](& regsvr32.exe /s $dllPath64 2>&1)
                $exitCode = $LASTEXITCODE
                
                if ($exitCode -eq 0) {
                    $successCount++
                    $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x64"; Status = "Success" }
                }
                else {
                    $failCount++
                    $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x64"; Status = "Failed"; ExitCode = $exitCode }
                    Write-Log "  Failed to register $dll (x64, exit $exitCode)" -Level Verbose
                }
            }
            catch {
                $failCount++
                $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x64"; Status = "Exception"; Error = $_.Exception.Message }
            }
        }
        else {
            $skipCount++
        }
        
        # Register x86 version on 64-bit systems
        if ($is64Bit) {
            $dllPath32 = Join-Path $sysWow64Path $dll
            if (Test-Path $dllPath32) {
                try {
                    $regsvr32x86 = Join-Path $sysWow64Path "regsvr32.exe"
                    if (Test-Path $regsvr32x86) {
                        [void](& $regsvr32x86 /s $dllPath32 2>&1)
                    }
                    else {
                        [void](& regsvr32.exe /s $dllPath32 2>&1)
                    }
                    $exitCode = $LASTEXITCODE
                    
                    if ($exitCode -eq 0) {
                        $successCount++
                        $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x86"; Status = "Success" }
                    }
                    else {
                        $failCount++
                        $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x86"; Status = "Failed"; ExitCode = $exitCode }
                        Write-Log "  Failed to register $dll (x86, exit $exitCode)" -Level Verbose
                    }
                }
                catch {
                    $failCount++
                    $detailedResults += [PSCustomObject]@{ DLL = $dll; Arch = "x86"; Status = "Exception"; Error = $_.Exception.Message }
                }
            }
        }
    }
    
    # Log summary
    $totalAttempts = $successCount + $failCount
    if ($totalAttempts -eq 0) {
        Write-Log "No DLLs required registration on this system" -Level Warning
    }
    else {
        $successRate = [math]::Round(($successCount / $totalAttempts) * 100, 1)
        Write-Log "DLL registration complete: $successCount/$totalAttempts succeeded ($successRate%), $failCount failed, $skipCount skipped" -Level $(if ($failCount -eq 0 -or $successRate -ge 90) { "Success" } else { "Warning" })
    }
    
    # Log failures at Info level if any
    $failedDlls = $detailedResults | Where-Object { $_.Status -ne "Success" }
    if ($failedDlls) {
        Write-Log "Failed registrations:" -Level Info
        foreach ($fail in $failedDlls) {
            Write-Log "  - $($fail.DLL) ($($fail.Arch)): $($fail.Status)" -Level Info
        }
    }
}

function Clear-BITSQueue {
    Write-Log "Clearing BITS transfer queue..." -Level Info
    
    try {
        $bitsJobs = Get-BitsTransfer -AllUsers -ErrorAction SilentlyContinue
        if ($bitsJobs) {
            $bitsJobs | Remove-BitsTransfer -ErrorAction SilentlyContinue
            Write-Log "Cleared $($bitsJobs.Count) BITS transfer(s)" -Level Success
        }
        else {
            Write-Log "No pending BITS transfers found" -Level Info
        }
    }
    catch {
        Write-Log "Failed to clear BITS queue: $_" -Level Warning
    }
}

function Invoke-DISMRepair {
    Write-Log "Running DISM component store repair..." -Level Info
    Write-Log "This may take 10-30 minutes depending on system state..." -Level Warning
    
    try {
        # First check health
        Write-Log "DISM: Checking component store health..." -Level Info
        $checkResult = & DISM.exe /Online /Cleanup-Image /CheckHealth 2>&1
        Write-Log "DISM CheckHealth: $($checkResult -join ' ')" -Level Info
        
        # Scan health
        Write-Log "DISM: Scanning component store..." -Level Info
        $scanResult = & DISM.exe /Online /Cleanup-Image /ScanHealth 2>&1
        foreach ($line in $scanResult) {
            if ($line -and $line.Trim()) { Write-Log "  $line" -Level Verbose -NoConsole }
        }
        Write-Log "DISM ScanHealth completed (Exit code: $LASTEXITCODE)" -Level Info
        
        # Restore health if needed
        Write-Log "DISM: Restoring component store health..." -Level Info
        $restoreResult = & DISM.exe /Online /Cleanup-Image /RestoreHealth 2>&1
        foreach ($line in $restoreResult) {
            if ($line -and $line.Trim()) { Write-Log "  $line" -Level Verbose -NoConsole }
        }
        
        if ($LASTEXITCODE -eq 0) {
            Write-Log "DISM RestoreHealth completed successfully" -Level Success
        }
        else {
            Write-Log "DISM RestoreHealth completed with exit code: $LASTEXITCODE" -Level Warning
            # Try with Windows Update as source
            Write-Log "Attempting DISM repair with Windows Update as source..." -Level Info
            $retryResult = & DISM.exe /Online /Cleanup-Image /RestoreHealth /Source:WU 2>&1
            foreach ($line in $retryResult) {
                if ($line -and $line.Trim()) { Write-Log "  $line" -Level Verbose -NoConsole }
            }
            if ($LASTEXITCODE -eq 0) {
                Write-Log "DISM repair with Windows Update source succeeded" -Level Success
            }
            else {
                Write-Log "DISM repair returned exit code: $LASTEXITCODE" -Level Warning
            }
        }
        
        # Cleanup component store
        Write-Log "DISM: Cleaning up component store..." -Level Info
        $cleanupResult = & DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase 2>&1
        foreach ($line in $cleanupResult) {
            if ($line -and $line.Trim()) { Write-Log "  $line" -Level Verbose -NoConsole }
        }
        Write-Log "Component store cleanup completed (Exit code: $LASTEXITCODE)" -Level Info
        
        return $true
    }
    catch {
        Write-Log "DISM repair failed: $_" -Level Error
        return $false
    }
}

function Invoke-SFCRepair {
    Write-Log "Running System File Checker (SFC)..." -Level Info
    Write-Log "This may take 10-20 minutes..." -Level Warning
    
    try {
        $sfcResult = & sfc.exe /scannow 2>&1
        $sfcOutput = $sfcResult -join "`n"
        
        if ($sfcOutput -match "did not find any integrity violations") {
            Write-Log "SFC: No integrity violations found" -Level Success
        }
        elseif ($sfcOutput -match "successfully repaired") {
            Write-Log "SFC: Found and repaired corrupted files" -Level Success
        }
        elseif ($sfcOutput -match "could not perform") {
            Write-Log "SFC: Could not perform the requested operation. Try running in Safe Mode." -Level Warning
        }
        else {
            Write-Log "SFC completed. Check CBS.log for details." -Level Info
        }
        
        return $true
    }
    catch {
        Write-Log "SFC scan failed: $_" -Level Error
        return $false
    }
}

function Reset-WindowsUpdatePolicy {
    Write-Log "Resetting Windows Update policies..." -Level Info
    
    try {
        # Delete Windows Update registry keys that might be corrupt
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate",
            "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                try {
                    # Just reset specific problem keys, not entire path
                    Remove-ItemProperty -Path $path -Name "SusClientId" -ErrorAction SilentlyContinue
                    Remove-ItemProperty -Path $path -Name "SusClientIdValidation" -ErrorAction SilentlyContinue
                    Write-Log "Reset problematic keys in $path" -Level Info
                }
                catch {
                    Write-Log "Could not modify $path : $_" -Level Warning
                }
            }
        }
        
        # Force Windows Update to re-authorize
        Write-Log "Forcing Windows Update re-authorization..." -Level Info
        & wuauclt.exe /resetauthorization /detectnow 2>&1 | Out-Null
        
        Write-Log "Windows Update policy reset completed" -Level Success
    }
    catch {
        Write-Log "Failed to reset Windows Update policies: $_" -Level Warning
    }
}

function Invoke-RepairPhase {
    param([switch]$Aggressive)
    
    Write-Log "========================================" -Level Info
    Write-Log "Starting Windows Update Repair Phase" -Level Info
    if ($Aggressive) {
        Write-Log "AGGRESSIVE MODE ENABLED - Running deep repairs" -Level Warning
    }
    Write-Log "========================================" -Level Info
    
    $totalSteps = if ($Aggressive) { 8 } else { 5 }
    $currentStep = 0
    
    # Step 1: Stop services
    $currentStep++
    Write-Log "Step $currentStep/$totalSteps : Stopping Windows Update related services..." -Level Info
    Stop-UpdateServices
    Start-Sleep -Seconds 3
    
    # Step 2: Clear BITS queue
    $currentStep++
    Write-Log "Step $currentStep/$totalSteps : Clearing BITS transfer queue..." -Level Info
    Clear-BITSQueue
    
    # Step 3: Reset folders
    $currentStep++
    Write-Log "Step $currentStep/$totalSteps : Resetting update cache folders..." -Level Info
    Reset-UpdateFolders
    
    # Step 4: Reset security descriptors
    $currentStep++
    Write-Log "Step $currentStep/$totalSteps : Resetting service security descriptors..." -Level Info
    Reset-ServiceSecurityDescriptors
    
    # Step 5: Re-register DLLs
    $currentStep++
    Write-Log "Step $currentStep/$totalSteps : Re-registering Windows Update DLLs..." -Level Info
    Register-WindowsUpdateDLLs
    
    if ($Aggressive) {
        # Step 6: DISM Repair
        $currentStep++
        Write-Log "Step $currentStep/$totalSteps : Running DISM component store repair..." -Level Info
        Invoke-DISMRepair
        
        # Step 7: SFC Scan
        $currentStep++
        Write-Log "Step $currentStep/$totalSteps : Running System File Checker..." -Level Info
        Invoke-SFCRepair
        
        # Step 8: Reset WU Policy
        $currentStep++
        Write-Log "Step $currentStep/$totalSteps : Resetting Windows Update policies..." -Level Info
        Reset-WindowsUpdatePolicy
    }
    
    # Final Step: Restart services
    Write-Log "Restarting Windows Update related services..." -Level Info
    Start-UpdateServices
    
    # Give services time to initialize
    Write-Log "Waiting for services to initialize..." -Level Info
    Start-Sleep -Seconds 5
    
    Write-Log "Repair phase completed!" -Level Success
    Write-Log "========================================" -Level Info
}

function Get-AvailableUpdates {
    Write-Log "========================================" -Level Info
    Write-Log "Scanning for available Windows Updates..." -Level Info
    Write-Log "========================================" -Level Info
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        
        Write-Log "Searching for updates (this may take several minutes)..." -Level Info
        $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        $updates = @()
        $exclusiveUpdates = @()
        $regularUpdates = @()
        
        foreach ($update in $searchResult.Updates) {
            # Determine update category
            $category = "Other"
            if ($update.Categories.Count -gt 0) {
                $category = $update.Categories[0].Name
            }
            
            # Check if it's an exclusive installer update (must be installed alone)
            $isExclusive = $false
            try {
                if ($update.InstallationBehavior -and $update.InstallationBehavior.CanRequestUserInput) {
                    $isExclusive = $true
                }
                # Check for common exclusive update patterns in title
                if ($update.Title -match "Servicing Stack|SSU|Cumulative Update.*Preview" -and $category -eq "Security Updates") {
                    $isExclusive = $true
                }
            }
            catch {
                # If we can't determine, treat as regular
            }
            
            $updateObj = [PSCustomObject]@{
                Index       = $updates.Count + 1
                Title       = $update.Title
                KB          = ($update.KBArticleIDs -join ", ")
                Size        = "{0:N2} MB" -f ($update.MaxDownloadSize / 1MB)
                Severity    = $update.MsrcSeverity
                IsDownloaded = $update.IsDownloaded
                Category    = $category
                IsExclusive = $isExclusive
                UpdateObject = $update
            }
            
            $updates += $updateObj
            
            # Separate exclusive vs regular updates
            if ($isExclusive) {
                $exclusiveUpdates += $updateObj
                Write-Log "  Found exclusive installer update: $($update.Title) [$category]" -Level Warning -NoConsole
            }
            else {
                $regularUpdates += $updateObj
            }
        }
        
        if ($updates.Count -eq 0) {
            Write-Log "No updates available. Your system is up to date!" -Level Success
        }
        else {
            Write-Log "Found $($updates.Count) available update(s)" -Level Success
            
            # Log breakdown by category
            $categoryGroups = $updates | Group-Object -Property Category
            foreach ($group in $categoryGroups) {
                Write-Log "  - $($group.Name): $($group.Count) update(s)" -Level Info
            }
            
            if ($exclusiveUpdates.Count -gt 0) {
                Write-Log "WARNING: $($exclusiveUpdates.Count) exclusive installer update(s) detected - these will be installed individually" -Level Warning
            }
        }
        
        # Return both full list and categorized lists
        return [PSCustomObject]@{
            All = $updates
            Exclusive = $exclusiveUpdates
            Regular = $regularUpdates
            Count = $updates.Count
        }
    }
    catch {
        Write-Log "Failed to search for updates: $_" -Level Error
        return [PSCustomObject]@{
            All = @()
            Exclusive = @()
            Regular = @()
            Count = 0
        }
    }
}

function Show-UpdateList {
    param(
        [array]$Updates
    )
    
    Write-Host "`n" -NoNewline
    Write-Host "=" * 100 -ForegroundColor Cyan
    Write-Host "Available Windows Updates" -ForegroundColor White
    Write-Host "=" * 100 -ForegroundColor Cyan
    
    foreach ($update in $Updates) {
        $downloadStatus = if ($update.IsDownloaded) { "[Downloaded]" } else { "[Pending]" }
        Write-Host ("{0,3}. {1}" -f $update.Index, $update.Title) -ForegroundColor White
        Write-Host ("     KB: {0} | Size: {1} | Severity: {2} {3}" -f $update.KB, $update.Size, $update.Severity, $downloadStatus) -ForegroundColor Gray
    }
    
    Write-Host "=" * 100 -ForegroundColor Cyan
    Write-Host "`n"
}

function Get-UserSelection {
    param(
        [array]$Updates
    )
    
    Write-Host "Enter your selection:" -ForegroundColor Yellow
    Write-Host "  - Enter 'A' or 'All' to install all updates" -ForegroundColor Gray
    Write-Host "  - Enter numbers separated by commas (e.g., 1,3,5) for specific updates" -ForegroundColor Gray
    Write-Host "  - Enter 'Q' or 'Quit' to exit" -ForegroundColor Gray
    Write-Host ""
    
    $selection = Read-Host "Selection"
    
    if ($selection -match "^[Qq](uit)?$") {
        return $null
    }
    
    if ($selection -match "^[Aa](ll)?$") {
        return $Updates
    }
    
    try {
        $indices = $selection -split "," | ForEach-Object { [int]$_.Trim() }
        $selectedUpdates = $Updates | Where-Object { $_.Index -in $indices }
        
        if ($selectedUpdates.Count -eq 0) {
            Write-Log "Invalid selection. No updates matched." -Level Warning
            return @()
        }
        
        return $selectedUpdates
    }
    catch {
        Write-Log "Invalid input. Please enter valid numbers." -Level Warning
        return @()
    }
}

function Install-SingleUpdate {
    param(
        [Parameter(Mandatory)]$Update,
        [int]$RetryCount = 0,
        [int]$MaxRetries = 3
    )
    
    $updateTitle = $Update.Title
    $attemptNum = $RetryCount + 1
    
    Write-Log "Installing update (Attempt $attemptNum/$MaxRetries): $updateTitle" -Level Info
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateDownloader = $updateSession.CreateUpdateDownloader()
        $updateInstaller = $updateSession.CreateUpdateInstaller()
        
        # Create single update collection
        $singleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
        $singleUpdate.Add($Update.UpdateObject) | Out-Null
        
        # Download if needed
        if (-not $Update.IsDownloaded) {
            Write-Log "  Downloading: $updateTitle" -Level Info
            $updateDownloader.Updates = $singleUpdate
            $downloadResult = $updateDownloader.Download()
            
            if ($downloadResult.ResultCode -notin 2, 3) {
                $errorDesc = Get-WUErrorDescription -ErrorCode $downloadResult.HResult
                Write-Log "  Download failed: $errorDesc" -Level Error -ErrorCode $downloadResult.HResult
                return @{
                    Success = $false
                    ResultCode = $downloadResult.ResultCode
                    HResult = $downloadResult.HResult
                    ErrorDescription = $errorDesc
                    RebootRequired = $false
                }
            }
        }
        
        # Install
        Write-Log "  Installing: $updateTitle" -Level Info
        $updateInstaller.Updates = $singleUpdate
        $installResult = $updateInstaller.Install()
        
        $updateResult = $installResult.GetUpdateResult(0)
        $success = $updateResult.ResultCode -eq 2
        
        if ($success) {
            Write-Log "  SUCCESS: $updateTitle installed" -Level Success
            $script:SuccessfulUpdates += $Update
        }
        else {
            $errorDesc = Get-WUErrorDescription -ErrorCode $updateResult.HResult
            Write-Log "  FAILED: $updateTitle - $errorDesc" -Level Error -ErrorCode $updateResult.HResult
        }
        
        return @{
            Success = $success
            ResultCode = $updateResult.ResultCode
            HResult = $updateResult.HResult
            ErrorDescription = if (-not $success) { $errorDesc } else { $null }
            RebootRequired = $installResult.RebootRequired
        }
    }
    catch {
        Write-Log "  Exception installing $updateTitle : $_" -Level Error
        return @{
            Success = $false
            ResultCode = -1
            HResult = $_.Exception.HResult
            ErrorDescription = $_.Exception.Message
            RebootRequired = $false
        }
    }
}

function Install-SelectedUpdates {
    param(
        [array]$SelectedUpdates,
        [int]$MaxRetries = 3,
        [switch]$RetryFailedIndividually
    )
    
    if ($SelectedUpdates.Count -eq 0) {
        Write-Log "No updates selected for installation." -Level Warning
        return @{
            Successful = @()
            Failed = @()
            RebootRequired = $false
        }
    }
    
    Write-Log "========================================" -Level Info
    Write-Log "Installing $($SelectedUpdates.Count) update(s)..." -Level Info
    Write-Log "========================================" -Level Info
    
    $successfulUpdates = @()
    $failedUpdates = @()
    $rebootRequired = $false
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateDownloader = $updateSession.CreateUpdateDownloader()
        $updateInstaller = $updateSession.CreateUpdateInstaller()
        
        # Create update collection
        $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        
        foreach ($update in $SelectedUpdates) {
            # Accept EULA if required
            if (-not $update.UpdateObject.EulaAccepted) {
                $update.UpdateObject.AcceptEula()
            }
            $updatesToInstall.Add($update.UpdateObject) | Out-Null
        }
        
        # Download Phase with retry
        Write-Log "Starting download phase..." -Level Info
        $updateDownloader.Updates = $updatesToInstall
        
        $downloadCount = 0
        foreach ($update in $SelectedUpdates) {
            $downloadCount++
            if (-not $update.IsDownloaded) {
                Write-Log "[$downloadCount/$($SelectedUpdates.Count)] Queued for download: $($update.Title)" -Level Info
            }
            else {
                Write-Log "[$downloadCount/$($SelectedUpdates.Count)] Already downloaded: $($update.Title)" -Level Info
            }
        }
        
        $downloadAttempt = 0
        $downloadSuccess = $false
        
        while (-not $downloadSuccess -and $downloadAttempt -lt $MaxRetries) {
            $downloadAttempt++
            Write-Log "Download attempt $downloadAttempt of $MaxRetries..." -Level Info
            
            try {
                $downloadResult = $updateDownloader.Download()
                
                switch ($downloadResult.ResultCode) {
                    2 { 
                        Write-Log "Download completed successfully" -Level Success
                        $downloadSuccess = $true
                    }
                    3 { 
                        Write-Log "Download completed with some errors - continuing" -Level Warning
                        $downloadSuccess = $true
                    }
                    4 { 
                        $errorDesc = Get-WUErrorDescription -ErrorCode $downloadResult.HResult
                        Write-Log "Download failed: $errorDesc" -Level Error -ErrorCode $downloadResult.HResult
                        if ($downloadAttempt -lt $MaxRetries) {
                            Write-Log "Waiting 10 seconds before retry..." -Level Info
                            Start-Sleep -Seconds 10
                        }
                    }
                    5 { 
                        Write-Log "Download aborted" -Level Error
                        break
                    }
                    default { 
                        Write-Log "Download result code: $($downloadResult.ResultCode)" -Level Warning
                    }
                }
            }
            catch {
                Write-Log "Download exception: $_" -Level Error
                if ($downloadAttempt -lt $MaxRetries) {
                    Write-Log "Waiting 10 seconds before retry..." -Level Info
                    Start-Sleep -Seconds 10
                }
            }
        }
        
        if (-not $downloadSuccess) {
            Write-Log "All download attempts failed. Trying individual downloads..." -Level Warning
            # Fall through to individual installation which will handle downloads
        }
        
        # Installation Phase
        Write-Log "Starting installation phase..." -Level Info
        $updateInstaller.Updates = $updatesToInstall
        
        $installCount = 0
        foreach ($update in $SelectedUpdates) {
            $installCount++
            Write-Log "[$installCount/$($SelectedUpdates.Count)] Queued for installation: $($update.Title)" -Level Info
        }
        
        $installResult = $updateInstaller.Install()
        
        # Process results
        switch ($installResult.ResultCode) {
            2 { Write-Log "Installation batch completed successfully" -Level Success }
            3 { Write-Log "Installation batch completed with some errors" -Level Warning }
            4 { Write-Log "Installation batch failed" -Level Error }
            5 { Write-Log "Installation batch aborted" -Level Error }
            default { Write-Log "Installation result code: $($installResult.ResultCode)" -Level Info }
        }
        
        # Check individual results and track failures
        Write-Log "Individual update results:" -Level Info
        for ($i = 0; $i -lt $SelectedUpdates.Count; $i++) {
            $updateResult = $installResult.GetUpdateResult($i)
            $currentUpdate = $SelectedUpdates[$i]
            
            $status = switch ($updateResult.ResultCode) {
                0 { "Not Started" }
                1 { "In Progress" }
                2 { "Succeeded" }
                3 { "Succeeded with Errors" }
                4 { "Failed" }
                5 { "Aborted" }
                default { "Unknown ($($updateResult.ResultCode))" }
            }
            
            if ($updateResult.ResultCode -eq 2) {
                Write-Log "  [OK] $($currentUpdate.Title)" -Level Success
                $successfulUpdates += $currentUpdate
                $script:SuccessfulUpdates += $currentUpdate
            }
            elseif ($updateResult.ResultCode -eq 3) {
                Write-Log "  [WARN] $($currentUpdate.Title): $status" -Level Warning
                $successfulUpdates += $currentUpdate
            }
            else {
                $errorDesc = Get-WUErrorDescription -ErrorCode $updateResult.HResult
                Write-Log "  [FAIL] $($currentUpdate.Title): $status - $errorDesc" -Level Error -ErrorCode $updateResult.HResult
                $failedUpdates += @{
                    Update = $currentUpdate
                    ResultCode = $updateResult.ResultCode
                    HResult = $updateResult.HResult
                    ErrorDescription = $errorDesc
                }
            }
        }
        
        $rebootRequired = $installResult.RebootRequired
        if ($rebootRequired) {
            Write-Log "A system reboot is required to complete installation." -Level Warning
            $script:RebootRequired = $true
        }
    }
    catch {
        Write-Log "Exception during batch installation: $_" -Level Error
        # Mark all as failed if batch fails completely
        foreach ($update in $SelectedUpdates) {
            if ($update -notin $successfulUpdates) {
                $failedUpdates += @{
                    Update = $update
                    ResultCode = -1
                    HResult = $_.Exception.HResult
                    ErrorDescription = $_.Exception.Message
                }
            }
        }
    }
    
    # Retry failed updates individually if requested
    if ($failedUpdates.Count -gt 0 -and $RetryFailedIndividually) {
        Write-Log "========================================" -Level Info
        Write-Log "Retrying $($failedUpdates.Count) failed update(s) individually..." -Level Info
        Write-Log "========================================" -Level Info
        
        $retriedFailed = @()
        
        foreach ($failedItem in $failedUpdates) {
            $update = $failedItem.Update
            
            for ($retry = 0; $retry -lt $MaxRetries; $retry++) {
                $result = Install-SingleUpdate -Update $update -RetryCount $retry -MaxRetries $MaxRetries
                
                if ($result.Success) {
                    $successfulUpdates += $update
                    if ($result.RebootRequired) {
                        $rebootRequired = $true
                        $script:RebootRequired = $true
                    }
                    break
                }
                
                if ($retry -lt ($MaxRetries - 1)) {
                    Write-Log "  Waiting 15 seconds before retry..." -Level Info
                    Start-Sleep -Seconds 15
                }
            }
            
            # If still failed after all retries
            if ($update -notin $successfulUpdates) {
                $retriedFailed += $failedItem
                $script:FailedUpdates += $update
            }
        }
        
        $failedUpdates = $retriedFailed
    }
    else {
        # Track failed updates in script-level variable
        foreach ($failedItem in $failedUpdates) {
            $script:FailedUpdates += $failedItem.Update
        }
    }
    
    return @{
        Successful = $successfulUpdates
        Failed = $failedUpdates
        RebootRequired = $rebootRequired
    }
}

function Invoke-VerificationPhase {
    Write-Log "========================================" -Level Info
    Write-Log "Verification Phase - Re-scanning for updates..." -Level Info
    Write-Log "========================================" -Level Info
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        
        Write-Log "Verifying installation status..." -Level Info
        $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        if ($searchResult.Updates.Count -eq 0) {
            Write-Log "Verification complete: All updates have been successfully installed!" -Level Success
        }
        else {
            Write-Log "Verification complete: $($searchResult.Updates.Count) update(s) still pending:" -Level Warning
            foreach ($update in $searchResult.Updates) {
                Write-Log "  - $($update.Title)" -Level Warning
            }
        }
        
        # Post-install log capture for CBS and WindowsUpdate logs
        Write-Log "Capturing post-install log summary..." -Level Info
        try {
            # Get recent CBS log entries
            $cbsLogPath = "$env:SystemRoot\Logs\CBS\CBS.log"
            if (Test-Path $cbsLogPath) {
                $recentCBSErrors = Get-Content $cbsLogPath -Tail 100 -ErrorAction SilentlyContinue | 
                    Where-Object { $_ -match "Error|FAIL|0x80[0-9A-F]{6}" } | 
                    Select-Object -Last 5
                if ($recentCBSErrors) {
                    Write-Log "Recent CBS activity (last 5 entries with errors):" -Level Info
                    foreach ($entry in $recentCBSErrors) {
                        Write-Log "  CBS: $entry" -Level Warning -NoConsole
                    }
                }
            }
            
            # Get Windows Update history for this session
            $updateHistory = $updateSearcher.QueryHistory(0, 50)
            $recentInstalls = $updateHistory | 
                Where-Object { $_.Date -gt (Get-Date).AddHours(-1) } | 
                Select-Object -First 10
            if ($recentInstalls) {
                Write-Log "Recent update history (last hour):" -Level Info
                foreach ($histEntry in $recentInstalls) {
                    $result = switch ($histEntry.ResultCode) {
                        2 { "Success" }
                        3 { "Partial" }
                        4 { "Failed" }
                        default { "Other ($($histEntry.ResultCode))" }
                    }
                    Write-Log "  [$result] $($histEntry.Title)" -Level $(if ($histEntry.ResultCode -eq 2) { "Success" } else { "Warning" }) -NoConsole
                }
            }
        }
        catch {
            Write-Log "Could not capture post-install logs: $_" -Level Warning
        }
        
        return $searchResult.Updates.Count
    }
    catch {
        Write-Log "Failed during verification: $_" -Level Error
        return -1
    }
}

function Get-FinalSummary {
    $duration = (Get-Date) - $script:StartTime
    
    Write-Log "========================================" -Level Info
    Write-Log "FINAL EXECUTION SUMMARY" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "Duration: $([math]::Round($duration.TotalMinutes, 2)) minutes" -Level Info
    Write-Log "Updates attempted: $($script:SuccessfulUpdates.Count + $script:FailedUpdates.Count)" -Level Info
    Write-Log "Successful: $($script:SuccessfulUpdates.Count)" -Level Success
    Write-Log "Failed: $($script:FailedUpdates.Count)" -Level $(if ($script:FailedUpdates.Count -gt 0) { "Error" } else { "Success" })
    
    if ($script:SuccessfulUpdates.Count -gt 0) {
        Write-Log "Successful updates:" -Level Success
        foreach ($update in $script:SuccessfulUpdates) {
            Write-Log "  [OK] $($update.Title)" -Level Success
        }
    }
    
    if ($script:FailedUpdates.Count -gt 0) {
        Write-Log "Failed updates:" -Level Error
        foreach ($update in $script:FailedUpdates) {
            Write-Log "  [FAIL] $($update.Title)" -Level Error
        }
        Write-Log "TIP: Run with -AggressiveRepair to perform deep system repair, then try again." -Level Warning
    }
    
    if ($script:RebootRequired) {
        Write-Log "========================================" -Level Warning
        Write-Log "REBOOT REQUIRED" -Level Warning
        Write-Log "Some updates require a system restart to complete." -Level Warning
        Write-Log "========================================" -Level Warning
    }
    
    # Final system log scan to capture any new issues from this session
    Write-Log "" -Level Info
    Write-Log "Performing final system log analysis..." -Level Info
    $null = Get-SystemLogIssues -HoursBack 1 -MaxEntries 25
    
    Write-Log "" -Level Info
    Write-Log "Log file: $LogPath" -Level Info
}

function Invoke-FallbackRepairAndRetry {
    param([array]$FailedUpdates)
    
    if ($FailedUpdates.Count -eq 0) {
        return @()
    }
    
    Write-Log "========================================" -Level Info
    Write-Log "FALLBACK REPAIR: Attempting to fix failed updates" -Level Warning
    Write-Log "========================================" -Level Info
    
    # Run quick repair
    Write-Log "Running quick repair before retry..." -Level Info
    Stop-UpdateServices
    Start-Sleep -Seconds 2
    Clear-BITSQueue
    Reset-UpdateFolders
    Start-UpdateServices
    Start-Sleep -Seconds 5
    
    # Trigger a new detection cycle to refresh metadata
    Write-Log "Triggering Windows Update detection..." -Level Info
    try {
        & UsoClient StartScan 2>&1 | Out-Null
        Start-Sleep -Seconds 10
    }
    catch {
        # UsoClient not available on older Windows, fallback to wuauclt
        & wuauclt.exe /detectnow 2>&1 | Out-Null
        Start-Sleep -Seconds 15
    }
    
    # Re-scan to get fresh update objects (with retry)
    $freshUpdates = $null
    $scanRetries = 0
    $maxScanRetries = 3
    
    while (-not $freshUpdates -and $scanRetries -lt $maxScanRetries) {
        $scanRetries++
        Write-Log "Re-scanning for updates after repair (attempt $scanRetries/$maxScanRetries)..." -Level Info
        $freshUpdates = Get-AvailableUpdates
        
        if ($freshUpdates.Count -eq 0 -and $scanRetries -lt $maxScanRetries) {
            Write-Log "Scan returned no updates, waiting before retry..." -Level Info
            Start-Sleep -Seconds 10
        }
        else {
            break
        }
    }
    
    # Find the failed updates in the fresh list
    $updatesToRetry = @()
    foreach ($failed in $FailedUpdates) {
        $failedTitle = $failed.Update.Title
        $freshMatch = $freshUpdates.All | Where-Object { $_.Title -eq $failedTitle }
        if ($freshMatch) {
            $updatesToRetry += $freshMatch
        }
        else {
            Write-Log "Update '$failedTitle' no longer found in scan results - may have been installed or superseded" -Level Info
        }
    }
    
    if ($updatesToRetry.Count -gt 0) {
        Write-Log "Retrying $($updatesToRetry.Count) update(s) after repair..." -Level Info
        $retryResult = Install-SelectedUpdates -SelectedUpdates $updatesToRetry -MaxRetries $MaxRetries -RetryFailedIndividually
        return $retryResult.Failed
    }
    else {
        Write-Log "No remaining updates to retry after repair" -Level Info
    }
    
    return $FailedUpdates
}

#endregion

#region Main Script

# Set up transcript logging if path specified
if ([string]::IsNullOrEmpty($TranscriptPath)) {
    $TranscriptPath = $LogPath -replace '\.log$', '_transcript.log'
}

try {
    Start-Transcript -Path $TranscriptPath -Force -ErrorAction SilentlyContinue | Out-Null
}
catch {
    # Transcript may already be running or not supported
}

# Banner
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                                                                              ║" -ForegroundColor Cyan
Write-Host "║           WINDOWS UPDATE REPAIR & INSTALLATION SCRIPT                        ║" -ForegroundColor Cyan
Write-Host "║                        Version 1.0                                           ║" -ForegroundColor Cyan
Write-Host "║                                                                              ║" -ForegroundColor Cyan
Write-Host "╠══════════════════════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
Write-Host "║  Capabilities:                                                               ║" -ForegroundColor Gray
Write-Host "║    • Pre-flight checks (disk, network, WUA health, pending reboots)          ║" -ForegroundColor Gray
Write-Host "║    • System restore point creation                                           ║" -ForegroundColor Gray
Write-Host "║    • Windows Update component repair (services, folders, DLLs)               ║" -ForegroundColor Gray
Write-Host "║    • DISM & SFC system repair                                                ║" -ForegroundColor Gray
Write-Host "║    • Update scanning with retry logic                                        ║" -ForegroundColor Gray
Write-Host "║    • Comprehensive logging & recommendations                                   ║" -ForegroundColor Gray
Write-Host "╚══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Check for Administrator privileges
if (-not (Test-AdminPrivileges)) {
    Write-Host ""
    Write-Host "ERROR: This script requires Administrator privileges!" -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as Administrator'." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Log "Script started with Administrator privileges" -Level Success
Write-Log "Log file: $LogPath" -Level Info
Write-Log "Transcript: $TranscriptPath" -Level Info
Write-Log "Parameters: SkipFix=$SkipFix, AutoInstall=$AutoInstall, ForceAutoInstall=$ForceAutoInstall, AggressiveRepair=$AggressiveRepair, MaxRetries=$MaxRetries" -Level Info

# Determine auto-install mode
$effectiveAutoInstall = $AutoInstall -or $ForceAutoInstall
if ($ForceAutoInstall) {
    Write-Log "WARNING: ForceAutoInstall enabled - updates will be installed without confirmation" -Level Warning
}

# Phase 0: Pre-flight checks (unless skipped)
if (-not $SkipPreflightChecks) {
    $preflightResult = Invoke-PreflightChecks
    
    if (-not $preflightResult.Passed) {
        Write-Log "Pre-flight checks failed. Please resolve issues before continuing." -Level Error
        if (-not $effectiveAutoInstall) {
            Write-Host "Continue anyway? (Y/N)" -ForegroundColor Yellow
            $continueChoice = Read-Host "Choice"
            if ($continueChoice -notmatch "^[Yy]") {
                Write-Log "Script terminated due to failed pre-flight checks." -Level Error
                exit 1
            }
        }
        else {
            Write-Log "AutoInstall mode: Continuing despite pre-flight failures..." -Level Warning
        }
    }
}
else {
    Write-Log "Pre-flight checks skipped via -SkipPreflightChecks parameter." -Level Info
}

# Phase 0.5: Create System Restore Point (unless skipped)
if (-not $SkipSystemRestore) {
    New-SystemRestorePoint -Description "Before Windows Update Script $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
}
else {
    Write-Log "System Restore Point creation skipped via -SkipSystemRestore parameter." -Level Info
}

# Phase 1: Repair (unless skipped)
if (-not $SkipFix) {
    Write-Host ""
    if (-not $effectiveAutoInstall) {
        Write-Host "Do you want to run the Windows Update repair phase? (Y/N)" -ForegroundColor Yellow
        $repairChoice = Read-Host "Choice"
        if ($repairChoice -match "^[Yy]") {
            Invoke-RepairPhase -Aggressive:$AggressiveRepair
        }
        else {
            Write-Log "Repair phase skipped by user choice." -Level Info
        }
    }
    else {
        Invoke-RepairPhase -Aggressive:$AggressiveRepair
    }
}
else {
    Write-Log "Repair phase skipped via -SkipFix parameter." -Level Info
}

# Phase 2: Scan for updates
$availableUpdatesResult = Get-AvailableUpdates
$availableUpdates = $availableUpdatesResult.All

if ($availableUpdatesResult.Count -eq 0) {
    Write-Log "No updates to install. System is up to date." -Level Success
    Get-FinalSummary
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
    exit 0
}

# Phase 3: Selection
$selectedUpdates = $null

if ($effectiveAutoInstall) {
    Write-Log "AutoInstall mode: Selecting all $($availableUpdatesResult.Count) available updates." -Level Info
    $selectedUpdates = $availableUpdates
}
else {
    Show-UpdateList -Updates $availableUpdates
    
    do {
        $selectedUpdates = Get-UserSelection -Updates $availableUpdates
        
        if ($null -eq $selectedUpdates) {
            Write-Log "User chose to exit. No updates installed." -Level Info
            Get-FinalSummary
            try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
            exit 0
        }
    } while ($selectedUpdates.Count -eq 0)
}

# Confirm selection
if (-not $effectiveAutoInstall -and $selectedUpdates.Count -gt 0) {
    Write-Host ""
    Write-Host "You have selected $($selectedUpdates.Count) update(s) to install:" -ForegroundColor Yellow
    foreach ($update in $selectedUpdates) {
        Write-Host "  - $($update.Title)" -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "Proceed with installation? (Y/N)" -ForegroundColor Yellow
    $confirm = Read-Host "Choice"
    
    if ($confirm -notmatch "^[Yy]") {
        Write-Log "Installation cancelled by user." -Level Info
        Get-FinalSummary
        try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
        exit 0
    }
}

# Phase 4: Installation with retry logic
$installResult = Install-SelectedUpdates -SelectedUpdates $selectedUpdates -MaxRetries $MaxRetries -RetryFailedIndividually

# Phase 4.5: Fallback repair and retry for persistent failures
if ($installResult.Failed.Count -gt 0 -and $effectiveAutoInstall) {
    Write-Log "Some updates failed. Attempting fallback repair..." -Level Warning
    $remainingFailed = Invoke-FallbackRepairAndRetry -FailedUpdates $installResult.Failed
    
    # If still failing and aggressive repair not yet done, try it
    if ($remainingFailed.Count -gt 0 -and -not $AggressiveRepair) {
        Write-Log "Updates still failing. Running aggressive repair..." -Level Warning
        Invoke-RepairPhase -Aggressive
        
        # One final retry
        $finalUpdates = Get-AvailableUpdates
        $finalToRetry = @()
        foreach ($failed in $remainingFailed) {
            $failedTitle = $failed.Update.Title
            $match = $finalUpdates | Where-Object { $_.Title -eq $failedTitle }
            if ($match) { $finalToRetry += $match }
        }
        
        if ($finalToRetry.Count -gt 0) {
            Write-Log "Final retry attempt for $($finalToRetry.Count) update(s)..." -Level Info
            Install-SelectedUpdates -SelectedUpdates $finalToRetry -MaxRetries 1
        }
    }
}

# Phase 5: Verification
Start-Sleep -Seconds 3
$pendingCount = Invoke-VerificationPhase

# Final Summary
Write-Host ""
Get-FinalSummary

if ($pendingCount -gt 0 -and $script:FailedUpdates.Count -gt 0) {
    Write-Host ""
    Write-Log "TROUBLESHOOTING TIPS:" -Level Warning
    Write-Log "  1. Reboot the system and run this script again" -Level Info
    Write-Log "  2. Run with -AggressiveRepair for deep system repair" -Level Info
    Write-Log "  3. Check C:\Windows\Logs\CBS\CBS.log for detailed errors" -Level Info
    Write-Log "  4. Check C:\Windows\SoftwareDistribution\ReportingEvents.log" -Level Info
    Write-Log "  5. Consider running: DISM /Online /Cleanup-Image /RestoreHealth" -Level Info
}

try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}

#endregion
