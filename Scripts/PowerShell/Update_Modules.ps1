<#
.SYNOPSIS
    Updates PowerShell modules to their latest versions with enhanced error handling and parallel processing.

.DESCRIPTION
    This script updates all or specified PowerShell modules to their latest versions. It includes:
    - Support for both PowerShell 5.1 and 7+
    - Parallel processing (PowerShell 7+) or efficient sequential processing (5.1)
    - Comprehensive error handling and logging
    - Automatic cleanup of old module versions
    - Progress reporting and detailed status information
    - WhatIf support for testing
    - Flexible filtering options

.PARAMETER Name
    Specifies module names to update. Supports wildcards. Default is '*' for all modules.

.PARAMETER AllowPrerelease
    If specified, allows updating to prerelease versions.

.PARAMETER WhatIf
    Shows what would happen without making changes.

.PARAMETER Force
    Forces updates even if modules appear up-to-date.

.PARAMETER ThrottleLimit
    Maximum concurrent operations for PowerShell 7+ (1-10). Default is 3.

.PARAMETER Scope
    Specifies the scope for module operations. Options: CurrentUser, AllUsers, Both. Default is 'Both'.

.PARAMETER SkipOldVersionCleanup
    If specified, does not remove old module versions after update.

.EXAMPLE
    .\Update_Modules_v2.ps1
    Updates all installed modules to their latest stable versions.

.EXAMPLE
    .\Update_Modules_v2.ps1 -Name "Az*" -AllowPrerelease -Verbose
    Updates all Az modules to latest versions including prereleases with verbose output.

.EXAMPLE
    .\Update_Modules_v2.ps1 -WhatIf
    Shows what modules would be updated without making changes.

.NOTES
    Author: Matthew Miles
    Version: 2.0
    Compatible: PowerShell 5.1+ (optimized for 7+)
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [string[]]$Name = @('*'),
    [switch]$AllowPrerelease,
    [switch]$Force,
    [ValidateSet('CurrentUser', 'AllUsers', 'Both')]
    [string]$Scope = 'Both',
    [ValidateRange(1, 10)]
    [int]$ThrottleLimit = 3,
    [switch]$SkipOldVersionCleanup
)

# Initialize variables
$ErrorActionPreference = 'Continue'
$ProgressPreference = 'Continue'
$Script:UpdatedModules = @()
$Script:FailedModules = @()
$Script:ErrorLog = Join-Path $env:TEMP "Update-Modules-$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Helper function for consistent logging
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $Message -ForegroundColor Cyan }
        'Warning' { Write-Warning $Message }
        'Error'   { Write-Host $Message -ForegroundColor Red }
        'Success' { Write-Host $Message -ForegroundColor Green }
    }
    
    Add-Content -Path $Script:ErrorLog -Value $logMessage -ErrorAction SilentlyContinue
}

# Helper function to update a single module
function Update-SingleModule {
    param(
        [PSCustomObject]$Module,
        [bool]$AllowPrerelease,
        [bool]$WhatIf,
        [bool]$Force,
        [bool]$SkipCleanup,
        [string]$UpdateScope
    )
    
    try {
        # Find latest version
        $findParams = @{
            Name = $Module.Name
            AllowPrerelease = $AllowPrerelease
            ErrorAction = 'Stop'
        }
        
        if ($Module.Repository) {
            $findParams.Repository = $Module.Repository
        }
        
        $latestModule = Find-Module @findParams | Select-Object -First 1
        
        # Compare versions
        $needsUpdate = $Force
        if (-not $needsUpdate) {
            try {
                $currentVersion = [version]$Module.Version
                $latestVersion = [version]$latestModule.Version
                $needsUpdate = $latestVersion -gt $currentVersion
            }
            catch {
                # Fallback to string comparison if version parsing fails
                $needsUpdate = $Module.Version -ne $latestModule.Version
            }
        }
        
        if ($needsUpdate) {
            if ($WhatIf) {
                Write-Log "Would update $($Module.Name) from $($Module.Version) to $($latestModule.Version)" -Level 'Info'
                return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $latestModule.Version; Status = 'WhatIf' }
            }
            
            # Perform update with appropriate scope
            $updateParams = @{
                Name = $Module.Name
                AllowPrerelease = $AllowPrerelease
                AcceptLicense = $true
                Force = $true
                ErrorAction = 'Stop'
            }
            
            # Determine update scope based on where module is installed
            if ($UpdateScope -ne 'Both') {
                $updateParams.Scope = $UpdateScope
            }
            elseif ($Module.InstalledLocation) {
                # Try to determine scope from installation path
                if ($Module.InstalledLocation -like "*Program Files*") {
                    $updateParams.Scope = 'AllUsers'
                }
                else {
                    $updateParams.Scope = 'CurrentUser'
                }
            }
            
            Update-Module @updateParams
            Write-Log "Updated $($Module.Name) from $($Module.Version) to $($latestModule.Version)" -Level 'Success'
            
            # Clean up old versions if requested
            if (-not $SkipCleanup) {
                try {
                    $allVersions = Get-InstalledModule -Name $Module.Name -AllVersions -ErrorAction Stop | 
                                   Sort-Object PublishedDate -Descending
                    
                    foreach ($oldVersion in ($allVersions | Select-Object -Skip 1)) {
                        try {
                            Uninstall-Module -Name $Module.Name -RequiredVersion $oldVersion.Version -Force -ErrorAction Stop
                            if ($VerbosePreference -eq 'Continue') {
                                Write-Log "Removed old version $($oldVersion.Version) of $($Module.Name)" -Level 'Info'
                            }
                        }
                        catch {
                            Write-Log "Failed to remove old version $($oldVersion.Version) of $($Module.Name): $($_.Exception.Message)" -Level 'Warning'
                        }
                    }
                }
                catch {
                    Write-Log "Failed to clean up old versions of $($Module.Name): $($_.Exception.Message)" -Level 'Warning'
                }
            }
            
            return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $latestModule.Version; Status = 'Updated' }
        }
        else {
            if ($VerbosePreference -eq 'Continue') {
                Write-Log "$($Module.Name) is up-to-date (version $($Module.Version))" -Level 'Info'
            }
            return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $Module.Version; Status = 'Current' }
        }
    }
    catch {
        $errorMsg = "Failed to update $($Module.Name): $($_.Exception.Message)"
        Write-Log $errorMsg -Level 'Error'
        return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = 'Failed'; Status = 'Error'; Error = $_.Exception.Message }
    }
}

# Main execution
try {
    Write-Log "Starting module update process..." -Level 'Info'
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" -Level 'Info'
    
    # Get installed modules
    Write-Log "Retrieving installed modules (Scope: $Scope)..." -Level 'Info'
    $installedModules = @()
    
    foreach ($namePattern in $Name) {
        $modules = Get-InstalledModule -Name $namePattern -ErrorAction SilentlyContinue |
                   Select-Object Name, Version, Repository, InstalledLocation |
                   Sort-Object Name
        
        if ($modules) {
            # Filter by scope if specified
            if ($Scope -ne 'Both') {
                $modules = $modules | Where-Object {
                    if ($Scope -eq 'AllUsers') {
                        $_.InstalledLocation -like "*Program Files*"
                    }
                    else {
                        $_.InstalledLocation -notlike "*Program Files*"
                    }
                }
            }
            $installedModules += $modules
        }
    }
    
    # Remove duplicates
    $installedModules = $installedModules | Sort-Object Name -Unique
    
    if (-not $installedModules) {
        Write-Log "No modules found matching the specified criteria." -Level 'Warning'
        return
    }
    
    $moduleCount = $installedModules.Count
    Write-Log "Found $moduleCount modules to process" -Level 'Info'
    
    $prereleaseText = if ($AllowPrerelease) { "prerelease" } else { "stable" }
    Write-Log "Updating to latest $prereleaseText versions..." -Level 'Info'
    
    # Process modules
    $results = @()
    $completedCount = 0
    
    if ($PSVersionTable.PSVersion.Major -ge 7 -and $moduleCount -gt 1) {
        # Use parallel processing for PowerShell 7+
        Write-Log "Using parallel processing (ThrottleLimit: $ThrottleLimit)" -Level 'Info'
        Write-Log "Processing modules... (this may take a few minutes)" -Level 'Info'
        
        $results = $installedModules | ForEach-Object -Parallel {
            $module = $_
            
            # Re-define the function in parallel scope
            function Update-SingleModule {
                param(
                    [PSCustomObject]$Module,
                    [bool]$AllowPrerelease,
                    [bool]$WhatIf,
                    [bool]$Force,
                    [bool]$SkipCleanup,
                    [string]$UpdateScope
                )
                
                try {
                    # Find latest version
                    $findParams = @{
                        Name = $Module.Name
                        AllowPrerelease = $AllowPrerelease
                        ErrorAction = 'Stop'
                    }
                    
                    if ($Module.Repository) {
                        $findParams.Repository = $Module.Repository
                    }
                    
                    $latestModule = Find-Module @findParams | Select-Object -First 1
                    
                    if (-not $latestModule) {
                        return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = 'Not Found'; Status = 'Error'; Error = "Module not found in any registered repository" }
                    }
                    
                    # Compare versions
                    $needsUpdate = $Force
                    if (-not $needsUpdate) {
                        try {
                            $currentVersion = [version]$Module.Version
                            $latestVersion = [version]$latestModule.Version
                            $needsUpdate = $latestVersion -gt $currentVersion
                        }
                        catch {
                            # Fallback to string comparison if version parsing fails
                            $needsUpdate = $Module.Version -ne $latestModule.Version
                        }
                    }
                    
                    if ($needsUpdate) {
                        if ($WhatIf) {
                            return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $latestModule.Version; Status = 'WhatIf' }
                        }
                        
                        # Perform update with appropriate scope
                        $updateParams = @{
                            Name = $Module.Name
                            AllowPrerelease = $AllowPrerelease
                            AcceptLicense = $true
                            Force = $true
                            ErrorAction = 'Stop'
                        }
                        
                        # Handle known problematic modules
                        $problematicModules = @('Az.ConfidentialLedger', 'Az.StorageSync', 'Microsoft.Graph.Compliance')
                        if ($Module.Name -in $problematicModules) {
                            $updateParams.Remove('AcceptLicense')
                            # Try alternative update method for these modules
                            try {
                                Install-Module -Name $Module.Name -Force -AllowClobber -Scope $updateParams.Scope -ErrorAction Stop
                                return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = 'Updated via Install'; Status = 'Updated' }
                            }
                            catch {
                                return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = 'Failed'; Status = 'Error'; Error = "Known compatibility issue: $($_.Exception.Message)" }
                            }
                        }
                        
                        # Determine update scope based on where module is installed
                        if ($UpdateScope -ne 'Both') {
                            $updateParams.Scope = $UpdateScope
                        }
                        elseif ($Module.InstalledLocation) {
                            # Try to determine scope from installation path
                            if ($Module.InstalledLocation -like "*Program Files*") {
                                $updateParams.Scope = 'AllUsers'
                            }
                            else {
                                $updateParams.Scope = 'CurrentUser'
                            }
                        }
                        
                        Update-Module @updateParams
                        
                        # Clean up old versions if requested
                        if (-not $SkipCleanup) {
                            try {
                                $allVersions = Get-InstalledModule -Name $Module.Name -AllVersions -ErrorAction Stop | 
                                               Sort-Object PublishedDate -Descending
                                
                                foreach ($oldVersion in ($allVersions | Select-Object -Skip 1)) {
                                    try {
                                        Uninstall-Module -Name $Module.Name -RequiredVersion $oldVersion.Version -Force -ErrorAction Stop
                                    }
                                    catch {
                                        # Silent fail for parallel processing
                                    }
                                }
                            }
                            catch {
                                # Silent fail for parallel processing
                            }
                        }
                        
                        return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $latestModule.Version; Status = 'Updated' }
                    }
                    else {
                        return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = $Module.Version; Status = 'Current' }
                    }
                }
                catch {
                    return @{ Name = $Module.Name; OldVersion = $Module.Version; NewVersion = 'Failed'; Status = 'Error'; Error = $_.Exception.Message }
                }
            }
            
            $result = Update-SingleModule -Module $module -AllowPrerelease $using:AllowPrerelease -WhatIf $using:WhatIfPreference -Force $using:Force -SkipCleanup $using:SkipOldVersionCleanup -UpdateScope $using:Scope
            
            # Show progress for significant actions (simplified for parallel processing)
            if ($result.Status -eq 'Updated') {
                Write-Host "+ Updated: $($result.Name) ($($result.OldVersion) -> $($result.NewVersion))" -ForegroundColor Green
            }
            elseif ($result.Status -eq 'Error') {
                Write-Host "- Failed: $($result.Name) - $($result.Error)" -ForegroundColor Red
            }
            elseif ($result.Status -eq 'Current') {
                Write-Host "= Current: $($result.Name) ($($result.OldVersion))" -ForegroundColor DarkGray
            }
            
            return $result
        } -ThrottleLimit $ThrottleLimit
        
        Write-Log "`nParallel processing completed. Analyzing results..." -Level 'Info'
    }
    else {
        # Sequential processing for PowerShell 5.1 or single module
        $counter = 0
        foreach ($module in $installedModules) {
            $counter++
            $percentComplete = [math]::Round(($counter / $moduleCount) * 100)
            Write-Progress -Activity "Updating Modules" -Status "Processing $($module.Name) ($counter of $moduleCount)" -PercentComplete $percentComplete
            
            $result = Update-SingleModule -Module $module -AllowPrerelease $AllowPrerelease -WhatIf $WhatIfPreference -Force $Force -SkipCleanup $SkipOldVersionCleanup -UpdateScope $Scope
            $results += $result
        }
        Write-Progress -Activity "Updating Modules" -Completed
    }
    
    # Process results
    $updated = $results | Where-Object { $_.Status -eq 'Updated' }
    $whatif = $results | Where-Object { $_.Status -eq 'WhatIf' }
    $current = $results | Where-Object { $_.Status -eq 'Current' }
    $failed = $results | Where-Object { $_.Status -eq 'Error' }
    
    # Display summary
    Write-Log "`n=== UPDATE SUMMARY ===" -Level 'Info'
    
    if ($whatif) {
        Write-Log "`nModules that would be updated:" -Level 'Info'
        foreach ($module in $whatif) {
            Write-Log "  * $($module.Name): $($module.OldVersion) -> $($module.NewVersion)" -Level 'Success'
        }
    }
    
    if ($updated) {
        Write-Log "`nSuccessfully updated modules:" -Level 'Success'
        foreach ($module in $updated) {
            Write-Log "  * $($module.Name): $($module.OldVersion) -> $($module.NewVersion)" -Level 'Success'
        }
    }
    
    if ($failed) {
        Write-Log "`nFailed to update modules:" -Level 'Error'
        foreach ($module in $failed) {
            Write-Log "  * $($module.Name): $($module.Error)" -Level 'Error'
        }
    }
    
    if (-not $updated -and -not $whatif) {
        Write-Log "No modules were updated. All modules are current." -Level 'Info'
    }
    
    Write-Log "`nProcessed: $moduleCount | Updated: $($updated.Count) | Current: $($current.Count) | Failed: $($failed.Count)" -Level 'Info'
    
    if (Test-Path $Script:ErrorLog) {
        Write-Log "Detailed log available at: $Script:ErrorLog" -Level 'Info'
    }
}
catch {
    Write-Log "Critical error in module update process: $($_.Exception.Message)" -Level 'Error'
    throw
}
finally {
    Write-Log "Module update process completed." -Level 'Info'
}