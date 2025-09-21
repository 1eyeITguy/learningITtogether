#Requires -RunAsAdministrator

<#
.SYNOPSIS
    OSD.Workspace Prerequisites Installation Script
    
.DESCRIPTION
    This script checks and installs all prerequisites for OSD.Workspace development.
    Based on the official OSDeploy documentation: https://github.com/OSDeploy/OSD.Workspace/wiki
    
    The script performs a comprehensive scan first, then presents a summary of what
    is already installed and what needs to be installed before proceeding.
    
    IMPORTANT: Before running this script, you must set the execution policy manually:
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
    
    Components checked and installed:
    - PowerShell execution policy and security protocols
    - NuGet package provider
    - Trusted PowerShell Gallery repository
    - PowerShellGet and PackageManagement modules
    - PowerShell 7 with full configuration
    - Git for Windows with user configuration
    - Visual Studio Code Insiders
    - Windows ADK and Windows PE add-on
    - Microsoft Deployment Toolkit
    - OSD PowerShell modules (OSD.Workspace, platyPS, OSD, OSDCloud)
    
.NOTES
    Author: Matthew Miles
    Version: 3.0
    Requires: Administrator privileges
    
.LINK
    https://github.com/OSDeploy/OSD.Workspace/wiki
#>

# ===============================================================================
# SCRIPT INITIALIZATION
# ===============================================================================

Write-Host ""
Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host "|                 OSD.Workspace Prerequisites Installer                 |" -ForegroundColor Cyan
Write-Host "|                             Version 3.0                               |" -ForegroundColor Cyan
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Global variables for tracking (these will be populated during execution)
$AllPrerequisites = @()
$PrerequisiteNames = @()

# Function to write section headers
function Write-SectionHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "+- $Title " -ForegroundColor Yellow -NoNewline
    Write-Host ("-" * (70 - $Title.Length)) -ForegroundColor Yellow
    Write-Host ""
}

# Function to write status messages
function Write-Status {
    param(
        [string]$Message,
        [string]$Status = "INFO"
    )
    
    $color = switch ($Status) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "INFO" { "White" }
        "SKIP" { "Gray" }
        "INSTALLED" { "Green" }
        "NOT_INSTALLED" { "Red" }
        "NEEDS_CONFIG" { "Yellow" }
    }
    
    $icon = switch ($Status) {
        "SUCCESS" { "[+]" }
        "WARNING" { "[!]" }
        "ERROR" { "[X]" }
        "INFO" { "[i]" }
        "SKIP" { "[-]" }
        "INSTALLED" { "[OK]" }
        "NOT_INSTALLED" { "[--]" }
        "NEEDS_CONFIG" { "[!]" }
    }
    
    Write-Host "  $icon " -ForegroundColor $color -NoNewline
    Write-Host $Message -ForegroundColor $color
}

# Function to check if a command exists
function Test-CommandExists {
    param([string]$Command)
    
    $cmd = Get-Command $Command -ErrorAction SilentlyContinue
    return $null -ne $cmd
}

# Function to get detailed prerequisite status
function Get-PrerequisiteStatus {
    param([string]$Name)
    
    $status = @{
        Name = $Name
        IsInstalled = $false
        NeedsConfiguration = $false
        CurrentVersion = $null
        Details = ""
        Action = "None"
        Priority = 0
        IsOptional = $false
    }
    
    switch ($Name) {
        "Administrator" {
            $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            $status.IsInstalled = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            $status.Details = if ($status.IsInstalled) { "Running as Administrator" } else { "Must run as Administrator" }
            $status.Priority = 1
        }
        
        "PowerShell7" {
            $psVersion = $PSVersionTable.PSVersion
            $status.CurrentVersion = $psVersion.ToString()
            $status.IsInstalled = $psVersion.Major -ge 7
            $status.Details = "Current version: $($status.CurrentVersion)"
            if (-not $status.IsInstalled) {
                $status.Action = "Install PowerShell 7"
                $status.Priority = 2
            }
        }
        
        "ExecutionPolicy" {
            $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
            $status.CurrentVersion = $currentPolicy
            $status.IsInstalled = $currentPolicy -in @("RemoteSigned", "Unrestricted", "Bypass")
            $status.Details = "Current policy: $currentPolicy"
            if (-not $status.IsInstalled) {
                $status.Action = "Set execution policy to RemoteSigned"
                $status.Priority = 1
            }
        }
        
        "TLS12" {
            $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
            $status.IsInstalled = $currentProtocol -match 'Tls12'
            $status.Details = "Current protocol: $currentProtocol"
            if (-not $status.IsInstalled) {
                $status.Action = "Configure TLS 1.2"
                $status.Priority = 2
            }
        }
        
        "NuGet" {
            $nugetProvider = Get-PackageProvider -Name 'NuGet' -ErrorAction SilentlyContinue
            $status.IsInstalled = $null -ne $nugetProvider
            if ($status.IsInstalled) {
                $status.CurrentVersion = $nugetProvider.Version.ToString()
                $status.Details = "Version: $($status.CurrentVersion)"
            } else {
                $status.Action = "Install NuGet Package Provider"
                $status.Details = "Not installed"
                $status.Priority = 3
            }
        }
        
        "PSGallery" {
            $psGallery = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
            $status.IsInstalled = ($null -ne $psGallery) -and ($psGallery.InstallationPolicy -eq 'Trusted')
            if ($psGallery) {
                $status.Details = "Installation Policy: $($psGallery.InstallationPolicy)"
                if (-not $status.IsInstalled) {
                    $status.NeedsConfiguration = $true
                    $status.Action = "Set PSGallery as Trusted"
                    $status.Priority = 3
                }
            } else {
                $status.Details = "Not found"
                $status.Action = "Configure PowerShell Gallery"
                $status.Priority = 3
            }
        }
        
        "PowerShellGet" {
            $psGetModule = Get-Module -Name PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            $requiredVersion = [Version]"2.2.5"
            $status.IsInstalled = ($null -ne $psGetModule) -and ($psGetModule.Version -ge $requiredVersion)
            if ($psGetModule) {
                $status.CurrentVersion = $psGetModule.Version.ToString()
                $status.Details = "Current version: $($status.CurrentVersion), Required: $($requiredVersion.ToString())"
            } else {
                $status.Details = "Not installed, Required: $($requiredVersion.ToString())"
            }
            if (-not $status.IsInstalled) {
                $status.Action = "Install/Update PowerShellGet"
                $status.Priority = 4
            }
        }
        
        "PackageManagement" {
            $pkgMgmtModule = Get-Module -Name PackageManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            $requiredVersion = [Version]"1.4.7"
            $status.IsInstalled = ($null -ne $pkgMgmtModule) -and ($pkgMgmtModule.Version -ge $requiredVersion)
            if ($pkgMgmtModule) {
                $status.CurrentVersion = $pkgMgmtModule.Version.ToString()
                $status.Details = "Current version: $($status.CurrentVersion), Required: $($requiredVersion.ToString())"
            } else {
                $status.Details = "Not installed, Required: $($requiredVersion.ToString())"
            }
            if (-not $status.IsInstalled) {
                $status.Action = "Install/Update PackageManagement"
                $status.Priority = 4
            }
        }
        
        "Git" {
            $status.IsInstalled = Test-CommandExists "git"
            if ($status.IsInstalled) {
                $status.CurrentVersion = (git --version 2>$null)
                $status.Details = $status.CurrentVersion
                
                # Check Git configuration
                $currentEmail = git config --global user.email 2>$null
                $currentName = git config --global user.name 2>$null
                
                if ([string]::IsNullOrWhiteSpace($currentEmail) -or $currentEmail -eq "you@example.com" -or
                    [string]::IsNullOrWhiteSpace($currentName) -or $currentName -eq "Your Name") {
                    $status.NeedsConfiguration = $true
                    $status.Action = "Configure Git user identity"
                    $status.Priority = 5
                }
            } else {
                $status.Action = "Install Git for Windows"
                $status.Details = "Not installed"
                $status.Priority = 5
            }
        }
        
        "VSCodeInsiders" {
            $status.IsInstalled = Test-CommandExists "code-insiders"
            if ($status.IsInstalled) {
                $status.Details = "Installed and in PATH"
            } else {
                $status.Action = "Install Visual Studio Code Insiders"
                $status.Details = "Not installed"
                $status.Priority = 6
            }
        }
        
        "WindowsADK" {
            $adkPath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"
            $status.IsInstalled = Test-Path $adkPath
            if ($status.IsInstalled) {
                $status.Details = "Installed at: $adkPath"
            } else {
                $status.Action = "Install Windows ADK"
                $status.Details = "Not installed"
                $status.Priority = 7
            }
        }
        
        "WindowsPE" {
            $winpePath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment"
            $status.IsInstalled = Test-Path $winpePath
            if ($status.IsInstalled) {
                $status.Details = "Installed at: $winpePath"
            } else {
                $status.Action = "Install Windows PE Add-on"
                $status.Details = "Not installed"
                $status.Priority = 8
            }
        }
        
        "WinPEx86Fix" {
            $winpeOCsPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\x86\WinPE_OCs'
            $mdtPath = "${env:ProgramFiles}\Microsoft Deployment Toolkit"
            $mdtInstalled = Test-Path $mdtPath
            
            $status.IsInstalled = Test-Path $winpeOCsPath
            $status.IsOptional = -not $mdtInstalled  # Only required if MDT is installed
            
            if ($status.IsInstalled) {
                $status.Details = "x86 OCs directory exists"
            } else {
                if ($mdtInstalled) {
                    $status.Action = "Create WinPE x86 OCs directory"
                    $status.Details = "Directory missing - required for MDT compatibility"
                    $status.Priority = 9
                } else {
                    $status.Action = "Create WinPE x86 OCs directory (if MDT is installed)"
                    $status.Details = "Directory missing - only needed if MDT is installed"
                    $status.Priority = 9
                }
            }
        }
        
        "MDT" {
            $mdtPath = "${env:ProgramFiles}\Microsoft Deployment Toolkit"
            $status.IsInstalled = Test-Path $mdtPath
            $status.IsOptional = $true  # Mark as optional
            if ($status.IsInstalled) {
                $status.Details = "Installed at: $mdtPath"
            } else {
                $status.Action = "Install Microsoft Deployment Toolkit (Optional)"
                $status.Details = "Not installed - Optional for DaRT capabilities"
                $status.Priority = 11
            }
        }
        
        "OSDModules" {
            $ps7ModulesPath = "C:\Program Files\PowerShell\Modules"
            $modules = @("OSD.Workspace", "platyPS", "OSD", "OSDCloud")
            $installedModules = @()
            $missingModules = @()
            
            foreach ($module in $modules) {
                $modulePath = Join-Path $ps7ModulesPath $module
                if (Test-Path $modulePath) {
                    $installedModules += $module
                } else {
                    $missingModules += $module
                }
            }
            
            $status.IsInstalled = $missingModules.Count -eq 0
            if ($status.IsInstalled) {
                $status.Details = "All modules installed: $($installedModules -join ', ')"
            } else {
                $status.Action = "Install missing OSD modules"
                $status.Details = "Installed: $($installedModules -join ', ') | Missing: $($missingModules -join ', ')"
                $status.Priority = 12
            }
        }
    }
    
    return $status
}

# ===============================================================================
# PREREQUISITE SCANNING PHASE
# ===============================================================================

Write-SectionHeader "Scanning System Prerequisites"

# Check administrator privileges first
$adminStatus = Get-PrerequisiteStatus "Administrator"
if (-not $adminStatus.IsInstalled) {
    Write-Status "ERROR: This script must be run as Administrator!" "ERROR"
    Write-Status "Please run PowerShell as Administrator and try again." "ERROR"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Status "Scanning system for OSDCloud Workspace prerequisites..." "INFO"
Write-Host ""

# Define all prerequisites to check
$PrerequisiteNames = @(
    "Administrator",
    "PowerShell7", 
    "ExecutionPolicy",
    "TLS12",
    "NuGet",
    "PSGallery", 
    "PowerShellGet",
    "PackageManagement",
    "Git",
    "VSCodeInsiders",
    "WindowsADK",
    "WindowsPE",
    "MDT",
    "WinPEx86Fix",
    "OSDModules"
)

# Scan all prerequisites
$AllPrerequisites = @()
foreach ($prereqName in $PrerequisiteNames) {
    Write-Host "  Checking $prereqName..." -ForegroundColor Gray
    $status = Get-PrerequisiteStatus $prereqName
    $AllPrerequisites += $status
}

# Separate installed vs needs action
$InstalledItems = $AllPrerequisites | Where-Object { $_.IsInstalled -and -not $_.NeedsConfiguration }
$NeedsInstall = $AllPrerequisites | Where-Object { -not $_.IsInstalled -and -not $_.IsOptional }
$NeedsConfig = $AllPrerequisites | Where-Object { $_.IsInstalled -and $_.NeedsConfiguration }
$OptionalItems = $AllPrerequisites | Where-Object { -not $_.IsInstalled -and $_.IsOptional }

Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host "|                        PREREQUISITE SCAN RESULTS                      |" -ForegroundColor Cyan  
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host ""

# Show what's already installed
if ($InstalledItems.Count -gt 0) {
    Write-Host "ALREADY INSTALLED:" -ForegroundColor Green
    Write-Host "-----------------" -ForegroundColor Green
    foreach ($item in $InstalledItems) {
        Write-Status "$($item.Name) - $($item.Details)" "INSTALLED"
    }
    Write-Host ""
}

# Show what needs configuration
if ($NeedsConfig.Count -gt 0) {
    Write-Host "NEEDS CONFIGURATION:" -ForegroundColor Yellow
    Write-Host "-------------------" -ForegroundColor Yellow
    foreach ($item in $NeedsConfig) {
        Write-Status "$($item.Name) - $($item.Details)" "NEEDS_CONFIG"
        Write-Status "  Action: $($item.Action)" "INFO"
    }
    Write-Host ""
}

# Show what needs installation
if ($NeedsInstall.Count -gt 0) {
    Write-Host "NEEDS INSTALLATION:" -ForegroundColor Red
    Write-Host "------------------" -ForegroundColor Red
    foreach ($item in $NeedsInstall) {
        Write-Status "$($item.Name) - $($item.Details)" "NOT_INSTALLED"
        Write-Status "  Action: $($item.Action)" "INFO"
    }
    Write-Host ""
}

# Show optional items
if ($OptionalItems.Count -gt 0) {
    Write-Host "OPTIONAL COMPONENTS:" -ForegroundColor Cyan
    Write-Host "-------------------" -ForegroundColor Cyan
    foreach ($item in $OptionalItems) {
        Write-Status "$($item.Name) - $($item.Details)" "INFO"
        Write-Status "  Action: $($item.Action)" "INFO"
    }
    Write-Host ""
}

# Summary
$TotalActions = $NeedsInstall.Count + $NeedsConfig.Count
Write-Host "SUMMARY:" -ForegroundColor Cyan
Write-Host "-------" -ForegroundColor Cyan
Write-Status "Prerequisites already satisfied: $($InstalledItems.Count)" "SUCCESS"
Write-Status "Prerequisites needing configuration: $($NeedsConfig.Count)" "WARNING"
Write-Status "Prerequisites needing installation: $($NeedsInstall.Count)" "ERROR"
Write-Status "Optional components available: $($OptionalItems.Count)" "INFO"
Write-Status "Total required actions: $TotalActions" "INFO"

# If everything is already installed
if ($TotalActions -eq 0) {
    Write-Host ""
    Write-Host "+=======================================================================+" -ForegroundColor Green
    Write-Host "|        ALL PREREQUISITES ARE ALREADY INSTALLED AND CONFIGURED!        |" -ForegroundColor Green
    Write-Host "+=======================================================================+" -ForegroundColor Green
    Write-Host ""
    Write-Status "Your system is ready for OSDCloud Workspace!" "SUCCESS"
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Cyan
    Write-Host "----------" -ForegroundColor Cyan
    Write-Host "1. Open PowerShell 7 and run: Install-OSDWorkspace" -ForegroundColor Yellow
    Write-Host "2. Navigate to: cd C:\OSDWorkspace" -ForegroundColor Yellow  
    Write-Host "3. Open in VS Code: code-insiders ." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 0
}

# Prompt to continue with installation
Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Yellow
Write-Host "|                        INSTALLATION CONFIRMATION                      |" -ForegroundColor Yellow
Write-Host "+=======================================================================+" -ForegroundColor Yellow
Write-Host ""

Write-Host "This script will now install and configure the missing prerequisites." -ForegroundColor White
Write-Host "The installation may take several minutes depending on your internet connection." -ForegroundColor White
Write-Host ""

$confirmation = Read-Host "Do you want to continue with the installation? (Y/N)"
if ($confirmation -notmatch '^[Yy]([Ee][Ss])?$') {
    Write-Status "Installation cancelled by user" "WARNING"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 0
}

# Get Git configuration if needed (only after user confirms)
$gitUserEmail = $null
$gitUserName = $null
$needsGitConfig = $AllPrerequisites | Where-Object { $_.Name -eq "Git" -and ($_.NeedsConfiguration -or -not $_.IsInstalled) }

if ($needsGitConfig) {
    Write-Host ""
    Write-Host "Git configuration is required. Please provide the following information:" -ForegroundColor Yellow
    Write-Host "Git Email Address: " -ForegroundColor Yellow -NoNewline
    $gitUserEmail = Read-Host
    Write-Host "Git User Name: " -ForegroundColor Yellow -NoNewline  
    $gitUserName = Read-Host
    Write-Host ""
}

# Ask about optional MDT installation
$installMDT = $false
$mdtOptional = $OptionalItems | Where-Object { $_.Name -eq "MDT" }

if ($mdtOptional) {
    Write-Host ""
    Write-Host "OPTIONAL: Microsoft Deployment Toolkit (MDT)" -ForegroundColor Cyan
    Write-Host "MDT provides advanced deployment capabilities including DaRT (Diagnostics and Recovery Toolset)." -ForegroundColor White
    Write-Host "Note: DaRT capabilities require Software Assurance (SA) licensing to utilize." -ForegroundColor Yellow
    Write-Host ""
    $mdtConfirmation = Read-Host "Would you like to install MDT for adding DaRT capabilities to your deployment? (Y/N)"
    $installMDT = $mdtConfirmation -match '^[Yy]([Ee][Ss])?$'
    
    if ($installMDT) {
        Write-Status "MDT will be installed as an optional component" "INFO"
    } else {
        Write-Status "MDT installation skipped - you can install it manually later if needed" "INFO"
    }
}

Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host "|                         STARTING INSTALLATION                         |" -ForegroundColor Cyan
Write-Host "+=======================================================================+" -ForegroundColor Cyan

# ===============================================================================
# INSTALLATION FUNCTIONS
# ===============================================================================

function Install-PowerShell7 {
    Write-SectionHeader "Installing PowerShell 7"
    
    try {
        Write-Status "Installing PowerShell 7 with full configuration..." "INFO"
        $process = Start-Process -FilePath "winget" -ArgumentList "install -e --id Microsoft.PowerShell --accept-source-agreements --override `"/passive ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1`"" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Status "PowerShell 7 installed successfully" "SUCCESS"
        } else {
            Write-Status "PowerShell 7 installation completed with exit code: $($process.ExitCode)" "WARNING"
        }
    } catch {
        Write-Status "Failed to install PowerShell 7: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Set-PowerShellExecutionPolicy {
    Write-SectionHeader "Configuring PowerShell Execution Policy"
    
    try {
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
        if ($currentPolicy -notin @("RemoteSigned", "Unrestricted", "Bypass")) {
            Write-Status "Setting execution policy to RemoteSigned for CurrentUser scope..." "INFO"
            Microsoft.PowerShell.Security\Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
            Write-Status "Execution policy set to RemoteSigned successfully" "SUCCESS"
        } else {
            Write-Status "Execution policy is already properly configured: $currentPolicy" "SKIP"
        }
    } catch {
        Write-Status "Failed to set execution policy: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Set-PowerShellSecurity {
    Write-SectionHeader "Configuring PowerShell Security"
    
    try {
        # Check and set security protocol
        $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
        if ($currentProtocol -notmatch 'Tls12') {
            Write-Status "Configuring TLS 1.2 security protocol..." "INFO"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Write-Status "TLS 1.2 protocol configured successfully" "SUCCESS"
        } else {
            Write-Status "TLS 1.2 protocol already configured" "SKIP"
        }
    } catch {
        Write-Status "Failed to configure PowerShell security: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-NuGetProvider {
    Write-SectionHeader "Installing NuGet Package Provider"
    
    try {
        Write-Status "Installing NuGet package provider..." "INFO"
        Install-PackageProvider -Name NuGet -ForceBootstrap -Verbose
        Write-Status "NuGet package provider installed successfully" "SUCCESS"
    } catch {
        Write-Status "Failed to install NuGet package provider: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Set-PSGalleryTrusted {
    Write-SectionHeader "Configuring PowerShell Gallery"
    
    try {
        Write-Status "Setting PowerShell Gallery as trusted repository..." "INFO"
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        Write-Status "PowerShell Gallery (PSGallery) has been set to Trusted" "SUCCESS"
    } catch {
        Write-Status "Failed to configure PowerShell Gallery: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-PowerShellModules {
    Write-SectionHeader "Installing PowerShell Core Modules"
    
    # Ensure PowerShell 7 modules path exists and is in PSModulePath
    $ps7ModulesPath = "C:\Program Files\PowerShell\Modules"
    if (!(Test-Path $ps7ModulesPath)) {
        Write-Status "Creating PowerShell 7 Modules directory..." "INFO"
        New-Item -Path $ps7ModulesPath -ItemType Directory -Force | Out-Null
    }

    # Add PS7 modules path to PSModulePath if not already there
    if ($env:PSModulePath -notlike "*$ps7ModulesPath*") {
        $env:PSModulePath = "$ps7ModulesPath;$env:PSModulePath"
        Write-Status "Added PowerShell 7 modules path to PSModulePath" "INFO"
    }

    try {
        # Check and install PowerShellGet
        $psGetModule = Get-Module -Name PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ((-not $psGetModule) -or ($psGetModule.Version -lt [Version]"2.2.5")) {
            Write-Status "Installing/updating PowerShellGet module..." "INFO"
            Save-Module -Name PowerShellGet -Path $ps7ModulesPath -Repository PSGallery -Force
            Write-Status "PowerShellGet module installed to PowerShell 7 location successfully" "SUCCESS"
        }
        
        # Check and install PackageManagement
        $pkgMgmtModule = Get-Module -Name PackageManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ((-not $pkgMgmtModule) -or ($pkgMgmtModule.Version -lt [Version]"1.4.7")) {
            Write-Status "Installing/updating PackageManagement module..." "INFO"
            Save-Module -Name PackageManagement -Path $ps7ModulesPath -Repository PSGallery -Force
            Write-Status "PackageManagement module installed to PowerShell 7 location successfully" "SUCCESS"
        }
    } catch {
        Write-Status "Failed to install PowerShell core modules: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-Git {
    param(
        [string]$UserEmail,
        [string]$UserName
    )
    
    Write-SectionHeader "Installing and Configuring Git"
    
    try {
        # Install Git if not already installed
        if (-not (Test-CommandExists "git")) {
            Write-Status "Installing Git for Windows..." "INFO"
            winget install --id Git.Git -e -h --accept-source-agreements
            Write-Status "Git for Windows installed successfully" "SUCCESS"
            
            # Refresh environment variables
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        }
        
        # Configure Git user identity
        if (![string]::IsNullOrWhiteSpace($UserEmail)) {
            Write-Status "Configuring Git user email: $UserEmail" "INFO"
            git config --global user.email "$UserEmail"
            Write-Status "Git user email configured successfully" "SUCCESS"
        }
        
        if (![string]::IsNullOrWhiteSpace($UserName)) {
            Write-Status "Configuring Git user name: $UserName" "INFO"
            git config --global user.name "$UserName"
            Write-Status "Git user name configured successfully" "SUCCESS"
        }
    } catch {
        Write-Status "Failed to install/configure Git: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-VSCodeInsiders {
    Write-SectionHeader "Installing Visual Studio Code Insiders"
    
    try {
        Write-Status "Installing Visual Studio Code Insiders..." "INFO"
        $vscodeArgs = '/SILENT /mergetasks="!runcode,addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath"'
        winget install -e --id Microsoft.VisualStudioCode.Insiders --accept-source-agreements --override $vscodeArgs
        Write-Status "Visual Studio Code Insiders installed successfully" "SUCCESS"
    } catch {
        Write-Status "Failed to install Visual Studio Code Insiders: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-WindowsADK {
    Write-SectionHeader "Installing Windows Assessment and Deployment Kit"
    
    try {
        Write-Status "Downloading and installing Windows ADK..." "INFO"
        $adkUrl = 'https://go.microsoft.com/fwlink/?linkid=2289980'
        $adkSetupPath = "$env:TEMP\adksetup.exe"
        
        # Download ADK setup
        Invoke-WebRequest -Uri $adkUrl -OutFile $adkSetupPath -UseBasicParsing
        
        # Install ADK with required features
        $adkArgs = @('/features', 'OptionId.DeploymentTools', 'OptionId.ImagingAndConfigurationDesigner', '/quiet', '/ceip', 'off', '/norestart')
        $adkProcess = Start-Process -FilePath $adkSetupPath -ArgumentList $adkArgs -Wait -PassThru
        
        if ($adkProcess.ExitCode -eq 0) {
            Write-Status "Windows ADK installed successfully" "SUCCESS"
        } else {
            Write-Status "Windows ADK installation completed with exit code: $($adkProcess.ExitCode)" "WARNING"
        }
        
        # Clean up
        Remove-Item $adkSetupPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Status "Failed to install Windows ADK: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-WindowsPE {
    Write-SectionHeader "Installing Windows PE Add-on"
    
    try {
        Write-Status "Downloading and installing Windows PE add-on..." "INFO"
        $winpeUrl = 'https://go.microsoft.com/fwlink/?linkid=2289981'
        $winpeSetupPath = "$env:TEMP\adkwinpesetup.exe"
        
        # Download Windows PE add-on setup
        Invoke-WebRequest -Uri $winpeUrl -OutFile $winpeSetupPath -UseBasicParsing
        
        # Install Windows PE add-on
        $winpeArgs = @('/features', 'OptionId.WindowsPreinstallationEnvironment', '/quiet', '/ceip', 'off', '/norestart')
        $winpeProcess = Start-Process -FilePath $winpeSetupPath -ArgumentList $winpeArgs -Wait -PassThru
        
        if ($winpeProcess.ExitCode -eq 0) {
            Write-Status "Windows PE add-on installed successfully" "SUCCESS"
        } else {
            Write-Status "Windows PE add-on installation completed with exit code: $($winpeProcess.ExitCode)" "WARNING"
        }
        
        # Clean up
        Remove-Item $winpeSetupPath -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Status "Failed to install Windows PE add-on: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-WinPEx86Fix {
    Write-SectionHeader "Applying MDT WinPE x86 Bugfix"
    
    try {
        $winpeOCsPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\x86\WinPE_OCs'
        Write-Status "Creating WinPE x86 OCs directory for MDT compatibility..." "INFO"
        New-Item -Path $winpeOCsPath -ItemType Directory -Force | Out-Null
        Write-Status "WinPE x86 OCs directory created successfully" "SUCCESS"
    } catch {
        Write-Status "Failed to create WinPE x86 OCs directory: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-MDT {
    Write-SectionHeader "Installing Microsoft Deployment Toolkit (MDT)"
    
    try {
        Write-Status "Installing Microsoft Deployment Toolkit..." "INFO"
        winget install --id Microsoft.DeploymentToolkit -e --accept-source-agreements --accept-package-agreements
        Write-Status "Microsoft Deployment Toolkit installed successfully" "SUCCESS"
        
        # After MDT is installed, we should also install the WinPE x86 fix if needed
        $winpeOCsPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\x86\WinPE_OCs'
        if (-not (Test-Path $winpeOCsPath)) {
            Write-Status "Installing WinPE x86 OCs directory (required for MDT)..." "INFO"
            Install-WinPEx86Fix
        }
    } catch {
        Write-Status "Failed to install Microsoft Deployment Toolkit: $($_.Exception.Message)" "ERROR"
        throw
    }
}

function Install-OSDModules {
    Write-SectionHeader "Installing OSD PowerShell Modules"
    
    # Ensure PowerShell 7 modules path exists and is in PSModulePath
    $ps7ModulesPath = "C:\Program Files\PowerShell\Modules"
    if (!(Test-Path $ps7ModulesPath)) {
        Write-Status "Creating PowerShell 7 Modules directory..." "INFO"
        New-Item -Path $ps7ModulesPath -ItemType Directory -Force | Out-Null
    }

    # Add PS7 modules path to PSModulePath if not already there
    if ($env:PSModulePath -notlike "*$ps7ModulesPath*") {
        $env:PSModulePath = "$ps7ModulesPath;$env:PSModulePath"
        Write-Status "Added PowerShell 7 modules path to PSModulePath" "INFO"
    }

    $modules = @(
        @{ Name = "OSD.Workspace"; Description = "The main OSDWorkspace PowerShell Module" },
        @{ Name = "platyPS"; Description = "Required for creating OSDWorkspace help files" },
        @{ Name = "OSD"; Description = "Used in some of the OSDWorkspace functions" },
        @{ Name = "OSDCloud"; Description = "Optionally used in some of the OSDWorkspace Gallery functions" }
    )

    foreach ($module in $modules) {
        try {
            # Check if module is already installed in PS7 location
            $modulePath = Join-Path $ps7ModulesPath $module.Name
            if (-not (Test-Path $modulePath)) {
                Write-Status "Installing $($module.Name) module to PowerShell 7 location - $($module.Description)..." "INFO"
                Save-Module -Name $module.Name -Path $ps7ModulesPath -Repository PSGallery -Force
                Write-Status "$($module.Name) module installed to PowerShell 7 location successfully" "SUCCESS"
            }
        } catch {
            Write-Status "Failed to install $($module.Name) module: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
}

# ===============================================================================
# SELECTIVE INSTALLATION EXECUTION
# ===============================================================================

Write-Host ""

# Sort prerequisites by priority and execute installations in order
# Include optional MDT if user chose to install it
$InstallationOrder = $AllPrerequisites | Where-Object { 
    (-not $_.IsInstalled -or $_.NeedsConfiguration) -and 
    (-not $_.IsOptional -or ($_.Name -eq "MDT" -and $installMDT))
} | Sort-Object Priority

$InstallationCount = 0
$TotalInstallations = $InstallationOrder.Count

foreach ($prerequisite in $InstallationOrder) {
    $InstallationCount++
    Write-Host ""
    Write-Host "[$InstallationCount/$TotalInstallations] Processing: $($prerequisite.Name)" -ForegroundColor Cyan
    
    try {
        switch ($prerequisite.Name) {
            "PowerShell7" {
                if (-not $prerequisite.IsInstalled) {
                    Install-PowerShell7
                }
            }
            
            "ExecutionPolicy" {
                if (-not $prerequisite.IsInstalled) {
                    Set-PowerShellExecutionPolicy
                }
            }
            
            "TLS12" {
                if (-not $prerequisite.IsInstalled) {
                    Set-PowerShellSecurity
                }
            }
            
            "NuGet" {
                if (-not $prerequisite.IsInstalled) {
                    Install-NuGetProvider
                }
            }
            
            "PSGallery" {
                if ($prerequisite.NeedsConfiguration -or -not $prerequisite.IsInstalled) {
                    Set-PSGalleryTrusted
                }
            }
            
            "PowerShellGet" {
                if (-not $prerequisite.IsInstalled) {
                    Install-PowerShellModules
                }
            }
            
            "PackageManagement" {
                if (-not $prerequisite.IsInstalled) {
                    # This is handled by Install-PowerShellModules
                    Write-Status "PackageManagement will be installed with PowerShellGet" "INFO"
                }
            }
            
            "Git" {
                Install-Git -UserEmail $gitUserEmail -UserName $gitUserName
            }
            
            "VSCodeInsiders" {
                if (-not $prerequisite.IsInstalled) {
                    Install-VSCodeInsiders
                }
            }
            
            "WindowsADK" {
                if (-not $prerequisite.IsInstalled) {
                    Install-WindowsADK
                }
            }
            
            "WindowsPE" {
                if (-not $prerequisite.IsInstalled) {
                    Install-WindowsPE
                }
            }
            
            "WinPEx86Fix" {
                # Only install if MDT is installed and the fix is needed
                $mdtPath = "${env:ProgramFiles}\Microsoft Deployment Toolkit"
                if (-not $prerequisite.IsInstalled -and (Test-Path $mdtPath)) {
                    Install-WinPEx86Fix
                } elseif (-not (Test-Path $mdtPath)) {
                    Write-Status "Skipping WinPE x86 fix - MDT is not installed" "SKIP"
                }
            }
            
            "MDT" {
                if (-not $prerequisite.IsInstalled -and $installMDT) {
                    Install-MDT
                } elseif (-not $installMDT) {
                    Write-Status "Skipping MDT installation - user chose not to install" "SKIP"
                }
            }
            
            "OSDModules" {
                if (-not $prerequisite.IsInstalled) {
                    Install-OSDModules
                }
            }
        }
        
        Write-Status "$($prerequisite.Name) processed successfully" "SUCCESS"
        
    } catch {
        Write-Status "Failed to process $($prerequisite.Name): $($_.Exception.Message)" "ERROR"
        Write-Host ""
        Write-Host "INSTALLATION FAILED!" -ForegroundColor Red
        Write-Host "The installation was interrupted due to an error with $($prerequisite.Name)." -ForegroundColor Red
        Write-Host "Please review the error message above and try running the script again." -ForegroundColor Red
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# ===============================================================================
# INSTALLATION COMPLETION AND VERIFICATION
# ===============================================================================

Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host "|                    INSTALLATION COMPLETED SUCCESSFULLY!               |" -ForegroundColor Green
Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host ""

Write-Status "All required OSDCloud Workspace prerequisites have been installed!" "SUCCESS"
Write-Status "Performing final verification..." "INFO"

# Quick verification scan
Write-Host ""
Write-Host "VERIFICATION RESULTS:" -ForegroundColor Cyan
Write-Host "--------------------" -ForegroundColor Cyan

$VerificationResults = @()
foreach ($prereqName in $PrerequisiteNames) {
    if ($prereqName -ne "Administrator") {  # Skip admin check in verification
        $verifyStatus = Get-PrerequisiteStatus $prereqName
        $VerificationResults += $verifyStatus
        
        if ($verifyStatus.IsInstalled -and -not $verifyStatus.NeedsConfiguration) {
            Write-Status "$($verifyStatus.Name) - VERIFIED" "SUCCESS"
        } elseif ($verifyStatus.NeedsConfiguration) {
            Write-Status "$($verifyStatus.Name) - NEEDS MANUAL CONFIG" "WARNING"
        } else {
            Write-Status "$($verifyStatus.Name) - FAILED VERIFICATION" "ERROR"
        }
    }
}

$VerifiedCount = ($VerificationResults | Where-Object { $_.IsInstalled -and -not $_.NeedsConfiguration }).Count
$TotalCount = $VerificationResults.Count

Write-Host ""
Write-Status "Verification Summary: $VerifiedCount/$TotalCount prerequisites successfully installed" "INFO"

if ($VerifiedCount -eq $TotalCount) {
    Write-Status "All prerequisites verified successfully!" "SUCCESS"
} else {
    Write-Status "Some prerequisites may need manual attention" "WARNING"
}

Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "----------" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Open PowerShell 7 and run:" -ForegroundColor Yellow
Write-Host "   Install-OSDWorkspace" -ForegroundColor Cyan
Write-Host "   This will create your OSD workspace directory." -ForegroundColor White
Write-Host ""
Write-Host "2. Open the workspace in VS Code Insiders:" -ForegroundColor Yellow
Write-Host "   cd C:\OSDWorkspace" -ForegroundColor Cyan
Write-Host "   code-insiders ." -ForegroundColor Cyan
Write-Host "   This will open your workspace in VS Code Insiders." -ForegroundColor White
Write-Host ""
Write-Host "3. Your OSD workspace will be located at:" -ForegroundColor Yellow
Write-Host "   C:\OSDWorkspace" -ForegroundColor Cyan
Write-Host ""

Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host "|                       INSTALLATION COMPLETE!                          |" -ForegroundColor Green
Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host ""
Write-Host "Script execution completed at $(Get-Date)" -ForegroundColor Gray
Write-Host ""
Read-Host "Press Enter to exit"
